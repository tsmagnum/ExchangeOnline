<#
.SYNOPSIS
The scripts generates a detailed EXOL Mailbox report.
.DESCRIPTION
The scripts generates a detailed EXOL Mailbox report, with the information needed by EXOL admins. You can choose if you want to use MFA to connect, how you want to sort the data (default by Total Size), the report format (console or OutGrid-View) and whether to export the info in a CSV file.
.PARAMETER noMFA
Select if you have to connect without MFA.
.PARAMETER sortBy
Specify how to sort the report: choose between 'Name','Total Items','Last Logon Time'. Use single quotes to enclose the parameter.
.PARAMETER reportFormat
Select report format: console (default) or ogv.
.PARAMETER exportCSV
Select if you want to export data in CSV format. You will have later to specify where the report should be saved.
.EXAMPLE
Get-MailboxReport.ps1 
Running the script without any parameter will connect with MFA and will output a console report sorted by total size.
.EXAMPLE
Get-MailboxReport.ps1 -noMFA
Running the script with "-noMFA" will connect with MFA and will output a console report sorted by total size.
.EXAMPLE
Get-MailboxReport.ps1 -sortBy "Last Logon Time"
The script will connect with MFA and will output a console report sorted by last logon time.
.EXAMPLE
Get-MailboxReport.ps1 -reportFormat ogv -exportCSV
The script will connect with MFA and will output a OutGrid-View report, exported in CSV.
#>

#region Credits
# Author: Federico Lillacci - Coesione Srl - www.coesione.net
# GitHub: https://github.com/tsmagnum
# Version: 1.0
#endregion

#region TODO
#Include archive mailboxes.
#endregion

#region PSS EXOL v2 info
# Go to https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/exchange-online-powershell-v2/exchange-online-powershell-v2?view=exchange-ps
#endregion

[CmdletBinding()]
param (
        [Parameter(Mandatory=$false)]
        [switch]
        $noMFA,

        [Parameter(Mandatory=$false)]
        [string]
        $sortBy = "Total Size (GB)",

        [Parameter (Mandatory=$false)]
        [string]
        $reportFormat = "console",

        [Parameter(Mandatory=$false)]
        [switch]
        $exportCSV
)

$ExolPSSession = Get-PSSession 
if ($ExolPSSession.ConfigurationName -ne "Microsoft.Exchange" -or $ExolPSSession.State -ne "Opened")
{
#region Connect Exol PSS without MFA
        if ($noMFA)
{
        Write-Host "We need to connect first to EXOL" -ForegroundColor Yellow
        $creds = Get-Credential
        $ex = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell `
        -Credential $creds -Authentication basic -AllowRedirection
        Import-PSSession $ex | Out-Null
       
}
#endregion

#region Connect Exol PSS with MFA
        else 
{
        Write-Host "We need to connect first to EXOL" -ForegroundColor Yellow        
        $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") `
        -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)

        . "$MFAExchangeModule"

        $ExolPSSession = Connect-ExoPSSession -UserPrincipalName (Read-Host "Username (UPN Format)")
}
} 
#endregion

# Getting mailbox data and sorting them
$mailboxes = Get-Mailbox

$stats = ( $mailboxes | Get-MailboxStatistics | Select-Object -property `
@{Label = "Name"; Expression={$_.DisplayName}},`
@{Label = "Total Items"; Expression={$_.ItemCount}},`
@{Label = "Total Size (GB)"; Expression={[System.Math]::Round((($_.TotalItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},`
@{Label = "Deleted Items (GB)"; Expression={[System.Math]::Round((($_.TotalDeletedItemSize.Value.ToString()).Split("(")[1].Split(" ")[0].Replace(",","")/1GB),2)}},`
@{Label = "Last Logon Time"; Expression={$_.LastLogonTime}},`
@{Label = "Mailbox Type"; Expression={$_.MailboxTypeDetail}},`
@{Label = "Storage Limit Status"; Expression={$_.StorageLimitStatus}}`
| Sort-Object -Property $sortBy -Descending 
)

# Choosing the desired report format
switch ($reportFormat) 
{
    console { $stats | Format-Table -AutoSize -Wrap -RepeatHeader }
    ogv { $stats | Out-Gridview }
    html { write-warning "Sorry, this feature is not yet available: please choose another format." }
}

# Exporting the data in CSV, if selected
if ($exportCSV) 
{
        $reportPath = Read-Host "Enter the full path of the report (e.g. c:\temp\report.csv)"
        $stats | Export-Csv -NoTypeInformation -Path $reportpath
}
