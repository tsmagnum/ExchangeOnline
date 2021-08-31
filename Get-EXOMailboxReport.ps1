<#
.SYNOPSIS
The scripts generates a detailed EXOL Mailbox report.
.DESCRIPTION
The scripts generates a detailed EXOL Mailbox report, with the information needed by EXOL admins. You can choose how you want to sort the data (default by Total Size), the report format (console, html or OutGrid-View) and whether to export the info in a CSV file.
.PARAMETER sortBy
Specify how to sort the report: choose between 'Name','Total Items','Last Logon Time'. Use single quotes to enclose the parameter.
.PARAMETER reportFormat
Select report format: console (default), html or ogv.
.PARAMETER reportHtmlPath
Name and (if desired) full filesystem path for the HTML report file; if omitted, it defaults to the script execution directory and a "EXOL_MBX_Report_<date>.html" filename.
.PARAMETER reportCSVPath
Name and (if desired) full filesystem path for the CSV report file; if omitted, it defaults to the script execution directory and a "EXOL_MBX_Report_<date>.csv" filename.
.PARAMETER exportCSV
Select if you want to export data in CSV format. You can specify where the report should be saved using -reportCSVPath parameter.
.PARAMETER archive
Select if you want to process archive mailboxes instead of regular mailboxes.
.EXAMPLE
Get-EXOMailboxReport.ps1 
Running the script without any parameter will output a console report sorted by total size.
.EXAMPLE
Get-EXOMailboxReport.ps1 -sortBy "Last Logon Time"
The script will output a console report sorted by last logon time.
.EXAMPLE
Get-EXOMailboxReport.ps1 -reportFormat html -reportHtmlPath "C:\mydir\myreport.html"
The script will output a html report, named myreport.html in the c:\mydir directory.
.EXAMPLE
Get-EXOMailboxReport.ps1 -reportFormat ogv -exportCSV -reportCSVPath "C:\mydir\myreport.csv"
The script will output a OutGrid-View report, exported in CSV and named myreport.csv in the c:\mydir directory.
#>

#region Credits
# Author: Federico Lillacci - Coesione Srl - www.coesione.net
# GitHub: https://github.com/tsmagnum
# Version: 2.0
#endregion

#region TODO

#endregion

#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="2.0.5" }

[CmdletBinding()]
param (
        [Parameter(Mandatory=$false)]
        [string]
        $sortBy = "Total Size (GB)",

        [Parameter (Mandatory=$false)]
        [string]
        $reportFormat = "console",

        [Parameter(Mandatory=$false)]
        [string]
        $reportHtmlPath = "EXOL_MBX_Report_"+ (Get-Date -UFormat %d%m%Y) +".html",

        [Parameter(Mandatory=$false)]
        [string]
        $reportCSVPath = "EXOL_MBX_Report_"+ (Get-Date -UFormat %d%m%Y) +".csv",

        [Parameter(Mandatory=$false)]
        [switch]
        $exportCSV,

        [Parameter(Mandatory=$false)]
        [switch]
        $archive
)

#region Functions
function reportHtml ($rawStats){

#CSS Code
        $header = @"
<style>
    body
  {
      background-color: White;
      font-size: 14px;
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

    h1{
        background-color: CornflowerBlue;
        color:white;
        text-align: center;
    }
</style>
"@
#End CSS Code

        $rawStats | `
        ConvertTo-Html `
                -PreContent "<h1>Exchange Online Mailbox Report</h1>"`
                -PostContent "<p>Creation Date: $(Get-Date)<p>"`
                -Title "EXOL Mailbox Report" `
                -Head $header | `
        Out-File -FilePath $reportHtmlPath
}

function reportConsole ($rawStats){
        $rawStats | Sort-Object -Property $sortBy -Descending | Format-Table -AutoSize -Wrap
        Write-Host "Total Number of Mailboxes: " $mailboxes.count -ForegroundColor Green
}

function reportOGV ($rawStats){
        $rawStats | Out-GridView
}

#endregion

# Are we already connected to Exchange Online?
$ExolPSSession = Get-PSSession 

if ($ExolPSSession.ConfigurationName -ne "Microsoft.Exchange" -and $ExolPSSession.State -ne "Opened")
{
        #Connecting to Exchange Online
        Write-Host "We need to connect first to EXOL" -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowProgress $true
}

# Getting mailbox data
Write-Host "Processing your data, please wait..." -ForegroundColor Green

if ($archive) 
{
        $mailboxes = Get-EXOMailbox -Archive | Get-EXOMailboxStatistics -PropertySets All
}

else 
{
$mailboxes = Get-EXOMailbox | Get-EXOMailboxStatistics -PropertySets All
}

$stats = ( $mailboxes | Select-Object -property `
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
    console { reportConsole($stats)}
    ogv { reportOGV($stats) }
    html { reportHtml($stats) }
}

# Exporting the data in CSV, if selected
if ($exportCSV) 
{
        $stats | Export-Csv -NoTypeInformation -Path $reportCSVPath
}