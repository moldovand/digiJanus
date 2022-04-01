## Script that crunches the MDB data file into Error Files (containing all occurencies) and then a 
## InstrNr_SF.csv (containing the sum of each type of Errors).


$AllErrors = 'AllOperationErrorsQuery'               # alternatively export Table or Report
$AllMovements = 'Total Motor Movements'

$hostname = $env:COMPUTERNAME                        # COMPUTERNAME corresponds to Janus_17...
$localdbpath = 'C:\Packard\Janus\database\'
$serverpath = '\\10.90.1.9\Janus_PE\PerkinElmer\databases\' + $hostname ## TODO

$outfile =  (Join-Path -Path $localdbpath -ChildPath ($hostname + '_all.csv'))
$outfileMove =  (Join-Path -Path $localdbpath -ChildPath ($hostname + '_allMove.csv'))
$outfileSumMov =  (Join-Path -Path $localdbpath -ChildPath ($hostname + '_sumMov.csv'))

$tempfile = (Join-Path -Path $localdbpath -ChildPath 'temp.csv')                                # temporary csv file
$tempfileMove = (Join-Path -Path $localdbpath -ChildPath 'tempMove.csv') 
$tempfileSumMov = (Join-Path -Path $localdbpath -ChildPath 'tempSumMov.csv')

$outputpath = (Join-Path -Path $localdbpath -ChildPath ($hostname + '_SF.csv'))

New-item -Path $serverpath -ItemType Directory -Force                                          # create folder on server foreach Janus

# Query for extracting AllErrors / AllMovements
#mamamia
#mamamia 2
function ExportQuery {
    param( [string]$dbpath,
           [bool]$header,
		   [string]$tempfile,
		   [string]$query)

    $access = New-Object -ComObject Access.Application                      # create new MS Access object
    $access.OpenCurrentDatabase($dbPath)                                    # open local database
    $access.Visible = $false                                                # hide MS Access window

   
    $DoCmd = $access.DoCmd
    $DoCmd.TransferText(2,             # TransferType Export delim
                    $null,             # Export Specification
                    $query,        	# Table (or Query, Report to Export)
                    $tempfile,         # File to save to
                    $header )          # boolean, header included
    $access.Quit()
}

# Look inside the "/database" folder for MDB files
$mdbfiles = Get-ChildItem ( $localdbpath + '\*_17220.mdb') -Recurse
$header = $true

# Run the two queries on each MDB that is greater than 10MB 
foreach ($mdbfile in $mdbfiles) {
	
    if ((Get-item $mdbfile).Length/1MB -gt 10) {              # SKIP new files
	
        ExportQuery -dbpath $mdbfile -header $header -tempfile $tempfile -query $AllErrors  # run our function		
		# Creating the CSV file that will hold all errors from MDB files
        if (!(Test-Path $outfile)) {
        New-Item -Path $outfile
		}
  
		# Append the content of multiple existing MDB files
		Add-Content -Path $outfile -value (Get-Content $tempfile) # concatenate tempfile to final output file
		
		Remove-Item -Path $tempfile
		
		ExportQuery -dbpath $mdbfile -header $header -tempfile $tempfileMove -query $AllMovements
		# Creating the CSV file that will hold all movements from MDB files
			if (!(Test-Path $outfileMove)) {
			New-Item -Path $outfileMove
			}
		# Append the content of multiple existing MDB files
		Add-Content -Path $outfileMove -value (Get-Content $tempfileMove) # concatenate tempfile to final output file
		
		Remove-Item -Path $tempfileMove
		$header = $false                                          # make sure header is only used for first file    
	
    }
}

# Extract the sum of Movements (2nd column in $outfileMove file)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $False
$NewWorkbook = $excel.Workbooks.open($outfileMove)
$NewWorksheet = $NewWorkbook.Worksheets.Item(1)
$NewWorksheetRange = $NewWorksheet.UsedRange
$RowCount = $NewWorksheetRange.Rows.Count
# "Number of rows: $RowCount"
$totalMoves = 0
if ($NewWorksheet.cells.Item(1,2).Value2 -eq "SumOfTotalMovements"){
	for($i = 2; $i -le $RowCount; $i++){
		# $NewWorksheet.cells.Item($i,2).Value2
		$totalMoves = $totalMoves + $NewWorksheet.cells.Item($i,2).Value2
	}
}
$NewWorkbook.Close()
$excel.Quit()

"Finished 1 out of 3"

# List of errors to extract
$errors = "65","66","119","122","271","313"

$excel = New-Object -ComObject excel.application 
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$wksht= $workbook.Worksheets.Item(1) 
$wksht.Name = 'The name you choose'

# Populate the CSV file with the number of movements
$wksht.Cells.Item(1,1) = 'Time Interval'
$wksht.Cells.Item(1,2) = '65' 
$wksht.Cells.Item(1,3) = '66' 
$wksht.Cells.Item(1,4) = '119'
$wksht.Cells.Item(1,5) = '122' 
$wksht.Cells.Item(1,6) = '271' 
$wksht.Cells.Item(1,7) = '313'
$wksht.Cells.Item(1,8) = 'sumMove'

# row index
$i = 2
# column index
$j = 2
    
# write instrument number in the Excel file
# create a new name for the file
$name = $hostname.Replace('Janus-','')
$name_new = $name.Replace('.csv',' ')

$wksht.Cells.Item($i,1) = 'From BOT'

# import the CSV original files (generated from the DB)
$CSV_file1 = $outfile
$A = Import-Csv $CSV_file1 -Delimiter ','

# ... and for each error
foreach ($error in $errors){
	#TODO - compile also some statistics based on the sum of extracted errors

	# count the number of specific errors
	$ErrorSum = @($A | Where-Object -Property ErrorCode -Like $error).Count
	
	# verbose mode
	# "Number of errors {$error}: $ErrorSum"
	
	# save the value in the Excel file
	$wksht.Cells.Item($i,$j) = $ErrorSum

	# create a new name for the file
	$name1 = $hostname.Replace('Janus-','')
	$file_new = $name1.Replace('.csv',' ')

	# Save the result for each error in its own file
	$CSV_file2 = (Join-Path -Path $localdbpath -ChildPath ($hostname + '_' + $error + '.csv'))
	$A | Where-Object -Property ErrorCode -Like $error | select ErrorCode, ErrorText, TestDatetime | Export-Csv -Path $CSV_file2 -Delimiter ',' -NoTypeInformation
	
	# column index
	$j++
	
	Move-Item $CSV_file2 -Destination $serverpath -Force
}

# $wksht.Cells.Item(2,8) = [int] ($totalMoves/42000)
$wksht.Cells.Item(2,8) = [math]::round($totalMoves/42000, 2)
$workbook.SaveAs($outputpath,[Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSVWindows) 
$workbook.Close()
$excel.Quit()

"Finished 2 out of 3"

# Move-Item $outfileSumMov -Destination $serverpath -Force
Move-Item $outfileMove -Destination $serverpath -Force
Move-Item $outfile -Destination $serverpath -Force
Move-Item $outputpath -Destination $serverpath -Force

"Finished 3 out of 3"