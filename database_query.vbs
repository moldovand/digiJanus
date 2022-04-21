
Dim connStr, objConn, FSO, getData

'''''''''''''''''''''''''''''''''''''
'Arguement0 is path to database
'Arguement1 is SQL query
'Arguement2 is Result File
''''''''''''''''''''''''''''''''''''''
Set args = Wscript.Arguments

'Create file system object to write results
Set FSO = CreateObject("Scripting.FileSystemObject")

'Define result file
Set ResultFile = FSO.OpenTextFile(WScript.Arguments.Item(2) ,2 , True)

'''''''''''''''''''''''''''''''''''''
'Define the driver and data source
'Access 2007, 2010, 2013, 2016 ACCDB:
'Provider=Microsoft.ACE.OLEDB.12.0
'Access 2000, 2002-2003 MDB:
'Provider=Microsoft.Jet.OLEDB.4.0
'
'Need to install provider available from Microsoft
''''''''''''''''''''''''''''''''''''''
connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & WScript.Arguments.Item(0)
 
'Define object type
Set objConn = CreateObject("ADODB.Connection")
 
'Open Connection
objConn.open connStr
 
'Define recordset and SQL query
Set rs = objConn.execute(WScript.Arguments.Item(1))
 
'While loop, loops through all available results
DO WHILE NOT rs.EOF
  getData = ""

  'loop to pickup multiple fields
  For i = 0 To rs.Fields.count - 1
     'if last field dont add comma
     If i = rs.Fields.count - 1 Then
        getData = getData & rs.Fields(i)
     Else
        getData = getData & rs.Fields(i) & ","
     End If
  Next

  ResultFile.WriteLine(getData)

  'move to next result before looping again
  'this is important
  rs.MoveNext
'continue loop
Loop
 
'Close connection and release objects
objConn.Close
Set rs = Nothing
Set objConn = Nothing

'Close result file
Set FSO= Nothing