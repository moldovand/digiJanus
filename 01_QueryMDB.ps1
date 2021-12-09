<# 
Script to iterate over all .mdb files in $localdbpath, open an MS Access Object
and export $QueryName to a combined Csv file. 
#>


# for development
$workingdir = "$HOME\PycharmProjects\Project_J\databases\" 
$resultpath = "$HOME\PycharmProjects\Project_J\src\janus_dashboard\csvs\"

$Errorfile = $resultpath + 'AllErrors.csv'
$Movementfile = $resultpath + 'AllMovements.csv'

function runQuery {
    param( [string]$query,
           [string]$dbpath,
           [string]$outfile
            )

        $connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=`"$dbpath`""
        $conn = new-object System.Data.OleDb.OleDbConnection($connString)
        $conn.open()

        $cmd = new-object System.Data.OleDb.OleDbCommand($query,$conn)

        $da = new-object System.Data.OleDb.OleDbDataAdapter($cmd)

        $dt = new-object System.Data.dataTable
        [void]$da.fill($dt)

        $conn.close()

        $dt | export-csv $outfile -NoTypeInformation -Append -Delimiter ';'

}

###### MAIN #######

$error_query = @"
                SELECT
                    o.`ErrorCode`,
                    t.`TestDateTime`,
                    `'{0}`' as janus,
                    `'{1}`' as query_time
                FROM
                    `OperationErrorTbl` o
                INNER JOIN `TestTbl` t ON
                    o.`TestId` = t.`TestId`
                WHERE o.`ErrorCode` <> 0
                  AND  t.`TestDateTime` > #{2:yyyy-MM-dd HH:mm}#
                ;
"@

$movement_query = @"
                SELECT
                    mo.`TotalMovements`,
                    t.`TestDateTime`,
                    `'{0}`' as janus
                FROM
                    `MovementTbl` mo
                INNER JOIN `TestTbl` t ON
                    t.`TestId` = mo.`TestId`
                WHERE mo.`TotalMovements` <> 0
                    AND t.`TestDateTime` > #{1:yyyy-MM-dd HH:mm}#
                ;

"@

$pg_conn = New-Object System.Data.Odbc.OdbcConnection
$pg_connstring = "Driver={PostgresQL UNICODE(x64)};Server=localhost;Port=5432;Database=pipeline;Username=postgres;"
$pg_conn.ConnectionString = $pg_connstring
$pg_conn.Open()

$da = New-Object System.Data.Odbc.OdbcDataAdapter("Select * from last_query order by janus;", $pg_conn)
$dt = New-Object System.Data.DataTable
[void]$da.fill($dt)

$current_date = Get-Date -Format "yyyy-MM-dd HH:mm"
foreach ($row in $dt) {

         $hostname = "Janus-$($row.janus)"
         $pingtest = Test-Connection -ComputerName $hostname -Quiet -Count 1 -ErrorAction SilentlyContinue
         if ($pingtest) {

             $mdbfile = "\\$hostname\c$\Packard\Janus\database\Multiprobe.mdb"

             Write-Output "Querying $($row.janus) last update on $($row.query_time)"  # print progress

             ############### RUN QUERIES ###############################
           try {
                runQuery -dbpath $mdbfile -query ($error_query -f $row.janus, $current_date, $row.query_time) -outfile $Errorfile
                runQuery -dbpath $mdbfile -query ($movement_query -f $row.janus, $row.query_time) -outfile $Movementfile

                $pg_cmd = New-Object System.Data.Odbc.OdbcCommand("UPDATE last_query SET query_time = `'$current_date`' where janus = `'$($row.janus)`'", $pg_conn)
                [void]$pg_cmd.ExecuteNonQuery()
             }
             catch {
                Write-Output "Error for Janus $($row.janus):  $_"
             }

        } else {
            Write-Output "$($row.janus) offline"  # print progress
    }

}

$pg_conn.Close()
