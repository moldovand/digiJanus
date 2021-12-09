# Counts the number of files in the specified folder
(Get-ChildItem 'C:\PKI Service\DANIEL\Protocols_new' -Recurse -File| Measure-Object).Count
