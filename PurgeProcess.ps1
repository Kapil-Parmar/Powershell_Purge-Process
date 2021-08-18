try
{
  ######################################### INPUT PARMS ###############################################
  #update input sheet path below
  $inputSheetPath =  $PSScriptRoot +'\InputSheet.xlsx'

  #update exception text file path below. Exceptions will be appended to this text file
  $exceptionLogFilePath = $PSScriptRoot +'\ExceptionFileFolder\Exception.txt'

  #update log text file path below. Log file created in this path will be emailed and then deleted by the process in each run
  $logFilePath = $PSScriptRoot + '\LogFileFolder\'
  #####################################################################################################

  #read input sheet
  $xl = New-Object -COM "Excel.Application"
  $wb = $xl.Workbooks.Open($inputSheetPath)
  $ws = $wb.Sheets.Item(1)
  $date = get-date

  <#create log file. Details of all files deleted/moved in current run will be updated in this log file 
    and will be emailed after each run#>

  $currentDate = Get-Date -Format "ddMMyyyy_HHmmss"
  $logFileName = $logFilePath + $currentDate +'.txt'
  New-Item $logFileName -ItemType File

  #iterate for each rows in excel sheet and delete/move files as per input data

  for ($i = 2; $i -le ($ws.UsedRange.Rows).count; $i++)
  {

   $folderPath = $ws.Cells.Item($i,1).Text
   $DeleteFileOlderThan = $ws.Cells.Item($i,2).Text
   $MoveFileOlderThan = $ws.Cells.Item($i,3).Text
   $MoveFolder = $ws.Cells.Item($i,4).Text

   $insideFiles = get-childitem -Path $folderPath

   #for each child item which is present inside the parent folder

       if (-not [string]::IsNullOrEmpty($DeleteFileOlderThan))
       {
         foreach($insideFile in $insideFiles)
         {
          $diff = New-TimeSpan -Start $insideFile.CreationTime -End $date
          #if file older than threshold days
          if ($diff.Days -gt $rows.DeleteFilesOlderThan)
           {
             try
             {
               #delete file
               $insideFile.Delete()
               #add to log file
               $logText = $insideFile.Name + ' deleted from ' + $folderPath
               add-content $logFileName $logText
             }
             catch
             {
               #add exception message to log file
               $logText = $PSItem.Exception.Message + ' error while deleting ' + $insideFile.Name + ' from ' + $folderPath
               add-content $logFileName $logText
             }
           }  
         }  
       }  
   
       if (-not [string]::IsNullOrEmpty($MoveFileOlderThan))
       {
         foreach($insideFile in $insideFiles)
         {
          #if file older than threshold days
          $diff = New-TimeSpan -Start $insideFile.CreationTime -End $Date
          if ($diff.Days -ge $MoveFileOlderThan)
           {
             try
             {
               #move file
               Move-Item -Path $insideFile.FullName -Destination $MoveFolder
               #add to log file
               $logText = $insideFile.Name + ' moved from ' + $folderPath + '     to      ' + $MoveFolder
               add-content $logFileName $logText
             }
             catch
             {
               #add exception message to log file
               $logText = $PSItem.Exception.Message + ' error while moving ' + $insideFile.Name + ' from ' + $rows.FolderPath
               add-content $logFileName $logText
             }
           }
         }
       }
   }
 #release excel COM object
 $wb.Close()
 $xl.Quit()
 [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
}
#below block will be executed incase of any unhandled exception in entire process
catch
{
   #append exception message to exception file
   $currentTimeStamp = get-date -Format f
   $appendText =  $currentTimeStamp + " | " +  $PSItem.Exception.Message 
   add-content -Path $exceptionLogFilePath -Value $appendText
}
