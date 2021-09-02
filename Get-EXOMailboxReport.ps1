#region Credits
# Author: Federico Lillacci - Coesione Srl - www.coesione.net
# GitHub: https://github.com/tsmagnum
# Version: 2.0
#endregion

#region TODO

#endregion

#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="2.0.5" }

[CmdletBinding()]
$reportDate = Get-Date -UFormat %d%m%Y

#######################################
#region User-variables - General Settings
$scheduled = $false #if $true, set the Unattended Connection variables
$sortBy = "Total Size (GB)"
$reportFormat = "console" # "console","html", "ogv"
$exportCSV = $false
$archive = $true
$reportHtmlPath = "EXOL_MBX_Report_"+ $reportDate +".html"
$reportCSVPath = "EXOL_MBX_Report_"+ $reportDate +".csv"
#endregion

#region User-variables - Email Settings
$sendEmail = $false
$emailHost = "your.smtp.com"
$emailPort = 25
$emailEnableSSL = $true
$emailUser = "yourUser"
$emailPass = "yourPassword"
$emailFrom = "exolReport@microsoft.com"
$emailTo = "your@email.com"
$emailSubject = "Exchange Online Mailbox Report"
#endregion

#region User-variables - Unattended Connection (Scheduled Report)
# Documentation: https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
$certThumbPrint = "<insert here your certificate thumbprint>" #REQUIRED if you need to run the script as a scheduled task
$appID = "<insert here your Azure app id>" #REQUIRED if you need to run the script as a scheduled task
$org = "<domain>.onmicrosoft.com" #REQUIRED if you need to run the script as a scheduled task, use the default ".onmicrosoft.com" domain

#DO NOT EDIT ANYTHING BEYOND THIS LINE!
#######################################
#endregion

#region CSS Code
$header = @"
<style>
    body
  {
      background-color: White;
      font-size: 12px;
      font-family: Arial, Helvetica, sans-serif;
  }

    table {
      border: 0.5px solid;
      border-collapse: collapse;
      width: 100%;
    }

    th {
        background-color: CornflowerBlue;
        color: white;
        padding: 6px;
        border: 0.5px solid;
        border-color: #000000;
    }

    tr:nth-child(even) {
            background-color: #f5f5f5;
        }

    td {
        padding: 6px;
        margin: 0px;
        border: 1px solid;
}

    h2{
        background-color: CornflowerBlue;
        color:white;
        text-align: center;
    }
</style>
"@
#endregion

#region HTML Code
$preContent = "<h2>Exchange Online Mailbox Report</h2>"
$postContent = "<p>Total Mailbox Count: $($stats.Count) - Creation Date: $(Get-Date)<p>"
$title = "Exchange Online Mailbox Report"
#endregion

#region Functions
function reportHtml {
        
        [CmdletBinding()]
        Param(
                [Parameter(Mandatory = $true)] $rawStats,
                [Parameter(Mandatory = $false)] [switch] $generateFile
        )
        
        if ($generateFile) 
        { 
                $rawStats | `
                ConvertTo-Html `
                        -PreContent $preContent `
                        -PostContent $postContent `
                        -Title $title `
                        -Head $header | `
                        Out-File -FilePath $reportHtmlPath
        }

        else 
        {
                $rawStats | `
                ConvertTo-Html `
                        -PreContent $preContent `
                        -PostContent $postContent `
                        -Title $title `
                        -Head $header 
        }
}

function reportConsole ($rawStats){
        $rawStats | Sort-Object -Property $sortBy -Descending | Format-Table -AutoSize -Wrap
        Write-Host "Total Number of Mailboxes: " $rawStats.count -ForegroundColor Green
}

function reportOGV ($rawStats){
        $rawStats | Out-GridView
}

#endregion

# Are we already connected to Exchange Online?
$ExolPSSession = Get-PSSession 

if ($ExolPSSession.ConfigurationName -ne "Microsoft.Exchange" -and $ExolPSSession.State -ne "Opened")
{
        #Connecting to Exchange Online - if you want to run the script as a scheduled task, 
        #please see https://office365itpros.com/2020/08/13/exo-certificate-based-authentication-powershell/
        Write-Host "We need to connect first to EXOL" -ForegroundColor Yellow
        
        if ($scheduled)
        {
                Connect-ExchangeOnline `
                        -CertificateThumbprint $certThumbPrint `
                        -AppId $appID `
                        -ShowBanner:$false `
                        -Organization $org
        }
        else 
        {
                Connect-ExchangeOnline -ShowProgress $true
        }
}

# Getting mailbox data
Write-Host "Processing your mailboxes, please wait..." -ForegroundColor Green
$mailboxes = @()

if ($archive) 
{
        $mailboxes += Get-EXOMailbox -Archive -ResultSize Unlimited | Get-EXOMailboxStatistics -Archive -PropertySets All
}

$mailboxes += Get-EXOMailbox -ResultSize Unlimited | Get-EXOMailboxStatistics -PropertySets All

#Creating the raw stats 
$stats = ( $mailboxes | Select-Object -property `
                @{Label = "Name"; Expression={$_.DisplayName}},`
                @{Label = "Total Size (GB)"; Expression={[System.Math]::Round((($_.TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},`
                @{Label = "Total Items"; Expression={$_.ItemCount}},`               
                @{Label = "Deleted Items (GB)"; Expression={[System.Math]::Round((($_.TotalDeletedItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},`
                @{Label = "Last Logon Time"; Expression={$_.LastLogonTime}},`
                @{Label = "Mailbox Type"; Expression={$_.MailboxTypeDetail}},`
                @{Label = "Archive Mbx"; Expression={$_.IsArchiveMailbox}},`
                @{Label = "Storage Limit Status"; Expression={$_.StorageLimitStatus}}`
|               Sort-Object -Property $sortBy -Descending 
)

# Choosing the desired report format
switch ($reportFormat) 
{
    console { reportConsole($stats)}
    ogv { reportOGV($stats) }
    html { reportHtml -rawStats $stats -generateFile }
}

# Exporting the data in CSV, if selected
if ($exportCSV) 
{
        $stats | Export-Csv -NoTypeInformation -Path $reportCSVPath
}

# Sending the report via email, if selected
If ($sendEmail) 
{
        $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
        $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
        $smtp.EnableSsl = $emailEnableSSL
        $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
        $msg.Subject = $emailSubject
        $msg.Body = reportHtml -rawStats $stats
        $msg.isBodyhtml = $true
        $smtp.send($msg)
}