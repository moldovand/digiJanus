$files = Get-ChildItem -Directory -Path \\10.90.1.9\Janus_PE\PerkinElmer\databases\ -Recurse -Force
Write-Host $files

# $files = "Janus-17227.csv","Janus-17229.csv","Janus-17230.csv" .. "Janus-17250.csv","Janus-17253.csv","Janus-17255.csv"
$errors = "65","66","119","122","271","313","sumMove"
Write-Host $errors

$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$wksht= $workbook.Worksheets.Item(1)
$wksht.Name = 'Compiled Errors'

$wksht.Cells.Item(1,1) = 'Instrument'
$wksht.Cells.Item(1,2) = '65'
$wksht.Cells.Item(1,3) = '66'
$wksht.Cells.Item(1,4) = '119'
$wksht.Cells.Item(1,5) = '122'
$wksht.Cells.Item(1,6) = '271'
$wksht.Cells.Item(1,7) = '313'
$wksht.Cells.Item(1,8) = 'sumMove'
# TO DO: implement the automatic generation of Weighted values and Median
# $wksht.Cells.Item(1,9) = '66_Weight'
# $wksht.Cells.Item(1,10) = '66_Med'
# $wksht.Cells.Item(1,11) = '122_Weight'
# $wksht.Cells.Item(1,12) = '122_Med'
# $wksht.Cells.Item(1,13) = '313_Weight'
# $wksht.Cells.Item(1,14) = '313_Med'
# counter for instrument number
$i = 2
# counter for error number
$j = 2

# create a new name for the History file
$name1 = '\\10.90.1.9\Janus_PE\PerkinElmer\databases\'
$outputpath = $name1 + 'LB_HF.csv'

# The loop for each file in the list
foreach ($file in $files)
{
    $path = $file
	# verbose mode
	Write-Host $path

	# get the folder name
	# \\10.90.1.9\Janus_PE\PerkinElmer\databases\Janus-17227
	# $hostname = $path.Replace('.csv','')
	$hostname = $path
	$serverpath = '\\10.90.1.9\Janus_PE\PerkinElmer\databases\' + $hostname

	# get the number of the instrument
	$name = $path -Replace "JANUS-",""
    # $name_new = $name.Replace('.csv','')
	$name_new = $name

    # write the number of the instrument in the first column
	$wksht.Cells.Item($i,1) = $name_new

    # import the CSV files
	# \\10.90.1.9\Janus_PE\PerkinElmer\databases\Janus-17227\Janus-17227_SF.csv
	$CSV_file1 =  (Join-Path -Path $serverpath -ChildPath ('\' + $hostname + '_SF.csv'))
    $A = @(Import-Csv $CSV_file1 -Delimiter ',')

    # ... and for each error
    foreach ($error in $errors)
    {
        #TODO - compile also some statistics based on the sum of extracted errors
		Write-Host $error
        # count the number of specific errors
        $ErrorValue = $A.$error
        Write-Host $ErrorValue

		# verbose mode
		# "Number of errors {$error}: $ErrorSum"

        # save the value in the Excel file
        $wksht.Cells.Item($i,$j) = $ErrorValue
        $j++
    }
    $j = 2
    $i++
}

<# Export-Csv -Path $CSV_file2 -Delimiter ',' -NoTypeInformation #>

# Move-Item $CSV_file2 -Destination $serverpath -Force

$workbook.SaveAs($outputpath,[Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSVWindows)
$excel.Quit()

# Move-Item $outfile -Destination $serverpath -Force
# Move-Item $outputpath -Destination $serverpath -Force
