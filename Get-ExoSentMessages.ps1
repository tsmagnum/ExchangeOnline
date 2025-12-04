#region Credits
# Author: Federico Lillacci - Coesione Srl - www.coesione.net
# GitHub: https://github.com/tsmagnum
# Version: 2.0
#endregion

#region TODO
#endregion

#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="2.0.5" }

[CmdletBinding()]
Param()

#region User-variables - Dates
$observationDays = -1 #negative number of days of the observation window, -1 for the last day
$today = Get-Date
$startDate = $today.AddDays($observationDays)
#endregion

#region User-variables - General Settings
$targetUsers = @(   "user1@domain.com",`
                    "user2@domain.com",`
                    "user3@domain.com")
$scheduled = $false #if $true, set the Unattended Connection variables
$consoleReport = $true #set to $true to display report data in the console window
#endregion

#region User-variables - Email Settings
$sendEmail = $false #set to $true to send the report via email
$emailHost = "smtp.office365.com"
$emailPort = 587
$emailEnableSSL = $true
$emailUser = "user@smtpserver.com"
$emailPass = "mySecretPass"
$emailFrom = "user@smtpserver.com"
$emailTo = "recipient@smtpserver.com"
$emailSubject = "Exchange Online - Report Sent Emails"
#endregion

#region User-variables - Unattended Connection (Scheduled Report)
# Documentation: https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
# and https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps#set-up-app-only-authentication
$certThumbPrint = "yyyyyyyyy" #REQUIRED if you need to run the script as a scheduled task
$appID = "xxxxxxxx" #REQUIRED if you need to run the script as a scheduled task
$org = "myorg.onmicrosoft.com" #REQUIRED if you need to run the script as a scheduled task, use the default ".onmicrosoft.com" domain
#endregion

#region CSS Code
$header = @"
<style>
    body {
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

    tbody tr:nth-child(even) {
            background-color: #f0f0f2;
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
$preContent = "<h2>Exchange Online - Report Sent Emails</h2>"
$postContent = "<p>Report from $($startDate) to $($today)</p>"
$title = "Exchange Online - Report Sent Emails"
#endregion

#DO NOT EDIT ANYTHING BEYOND THIS LINE!
#######################################

#region Functions
function Get-MessageCount {
        # I have integrated in this function some code of a great script from Vasil Michev - www.michev.info
        # https://github.com/michevnew/PowerShell/blob/master/Get-DetailedMessageStatsV2.ps1
        $cMessages = Get-MessageTraceV2 -SenderAddress $user -ResultSize 5000 `
                -StartDate  ($startDate) -EndDate ($today) -WarningVariable MoreResultsAvailable `
                        -Verbose:$false 3>$null

        $Messages += $cMessages | Select-Object Received,SenderAddress,RecipientAddress,Size,Status

        #If more results are available, as indicated by the presence of the WarningVariable, we need to loop until we get all results
        if ($MoreResultsAvailable) {
                do {
                #As we don't have a clue how many pages we will get, proper progress indicator is not feasible.
                Write-Host "." -NoNewline

                #Handling this via Warning output is beyond annoying...
                $NextPage = ($MoreResultsAvailable -join "").TrimStart("There are more results, use the following command to get more. ")
                $ScriptBlock = [ScriptBlock]::Create($NextPage)
                $cMessages = Invoke-Command -ScriptBlock $ScriptBlock -WarningVariable MoreResultsAvailable -Verbose:$false 3>$null #MUST PASS WarningVariable HERE OR IT WILL NOT WORK
                $Messages += $cMessages | Select-Object Received,SenderAddress,RecipientAddress,Size,Status
        }
                until ($MoreResultsAvailable.Count -eq 0) #Arraylist
        }
        # end of the code from Michev script
        
        $messagesCount = $Messages.Count

        return $messagesCount

        #If no messages were found, exit
        if ($Messages.Count -eq 0) {
        Write-Error "No messages found for the specified date range. Please check your permissions or update the date range above."
        return
        }
}

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
        Write-Host -ForegroundColor Green "Exchange Online - Report Email Inviate"
        $rawStats | Format-Table -AutoSize -Wrap
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

#Getting sent emails data
$stats = @()
foreach ($user in $targetUsers)
{
    $objUser = [PSCustomObject]@{
        Nome = $user
        Messaggi = Get-MessageCount
                                }

    $stats += $objUser
}

#Displaying the report in the console window
If ($consoleReport)
{
        reportConsole($stats)
}

#Sending the report via email
If ($sendEmail) 
{
        $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
        $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
        $smtp.EnableSsl = $emailEnableSSL
        $smtp.UseDefaultCredentials = $false
        $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
        $msg.Subject = $emailSubject
        $msg.Body = reportHtml -rawStats $stats
        $msg.isBodyhtml = $true
        $smtp.send($msg)
}
