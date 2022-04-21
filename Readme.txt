21st of April 2021:
- Short MSL example protocol that uses vbscript to query a database directly using SQL.
- The vbscript has arguements for the database, query and result file.

Note: 
- vbscript has to run using 32bit script host.
- this vbscript was setup to connect to an MS Access database and used the provider supplied by microsoft.
- installed using the AccessDatabaseEngine.exe
- if MS Access is already installed then need to install this using the command line with quiet flag to suppress errors
ie
C:\AccessDatabaseEngine.exe /quiet

Note:
- the Example queries the Multiprobe.mdb database
