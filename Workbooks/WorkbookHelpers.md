# Workbook Helpers
## Functions
```vb
GetPathFromRangeText(ByVal payload As String) As String
GetFilenameFromRangeText(ByVal payload As String) As String
IsWorkbookOpen(ByVal filename As String) As Boolean
TryGetWorkbook(ByVal filename As String, ByRef wb As Workbook, Optional path As String = vbNullString) As Boolean
```