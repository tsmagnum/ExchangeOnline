#region Credits
# Author: Federico Lillacci - Coesione Srl - www.coesione.net
# GitHub: https://github.com/tsmagnum
# Version: 1.1
# Date: 26/02/2024
#endregion

#region TODO

#endregion

#Storing the execution time in a variable; used to name the log and results file
$executionTime = (Get-Date).ToString('yyyyMMdd-hhmm')

#region VARIABLES
#please set these variables before running the script

#Txt file with the computers to process, one per line
$Computers = Get-Content -Path "C:\Scripts\ClientsTest.txt"

#Log file of the processed computers
$processedPCs = "C:\Scripts\$($executionTime)_ProcessedPC.txt"

#Results file
$csvPath = "C:\Scripts\$($executionTime)_PstFilesFound.csv"
#endregion

#Begin script execution

#Creating an empty array for the results
$results = @()

Set-Content $processedPCs -Value "Logging script execution - $($executionTime)"

Foreach ($Computer in $Computers)
{
   #Checking if the target computer is online: if so, the check continues
   Write-Host -ForegroundColor Cyan "Checking if $($Computer) is online"
   $pingtest = Test-Connection -ComputerName $Computer -Quiet -Count 1 -ErrorAction SilentlyContinue

   if ($pingtest)
   {
      $message = "$($Computer) is online, looking for PST files..."
      Write-Host -ForegroundColor Cyan $message
      Add-Content -Path $processedPCs -Value $message
      #Performing the search
      $pstFiles = Get-Wmiobject -namespace "root\CIMV2" -computername $Computer -Query "Select * from CIM_DataFile Where Extension = 'pst'"

      #Storing the results in PS Object
      Foreach ($file in $PstFiles)
      {
      $result = [PSCustomObject]@{
         Computer = $file.CSName
         Name = $file.Filename
         Path = $file.Description
         Size = ($file.FileSize)/1GB
         LastAccess = ($file.LastAccessed.Split("."))[0]
      }  

      Write-Host -ForegroundColor Green "PST found! Adding $($result.Name) to results"
      $results += $result
      
      #End 3rd foreach
      }

   #End 2nd foreach
   }

   #Logging offline computers
   else {
      $message = "$($Computer) is offline, skipping PST search"
      Write-Host -ForegroundColor Red $message
      Add-Content -Path $processedPCs -Value $message
   }
#End 1st foreach
}

#Saving results to a CSV file
$results | Export-Csv -Path $csvPath -NoTypeInformation