$szDB = 'C:\\Packard\\Janus\\database\\'
$szQuery = 'SELECT TestId, TestName FROM TestTbl ORDER BY TestId DESC'
$szResults = 'database_results.csv'

$wScript = 'C:\\Windows\\SysWoW64\\wscript.exe'
$dbScript = 'database_query.vbs'

& $wScript $dbScript $szDB $szQuery $szResults