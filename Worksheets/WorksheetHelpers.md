# Worksheet Helpers
Mostly used in `Table Split Tool`.

## Functions
```vb
IsValidSheetName(ByVal SheetName As String) As Boolean
TryRemoveSheet(ByVal Workbook As Workbook, ByVal SheetName As String) As Boolean
SheetExists(ByVal Workbook As Workbook, ByVal SheetName As String) As Boolean
AddOrGetWorksheet(ByVal worksheetName As String) As Worksheet
TryGetWorkSheet(ByVal wb As Workbook, ByVal worksheetName As String, ByRef ws As Worksheet) As Boolean
```