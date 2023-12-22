# Workbook Helpers
Helper functions for the Workbook object.

# Methods
## IsWorkbookOpen
Returns True if a Workbook with the given filename is currently open.
```vb
IsWorkbookOpen(filename As String) As Boolean
```
## TryGetWorkbook
Tries to get the Workbook with a given filename or optionally a given full path.
```vb
TryGetWorkbook(filename As String, wb As Workbook, Optional path As String = vbNullString) As Boolean
```
## GetFilenameFromRangeText
Returns the filename from the full qualified global address of a Range object.
```vb
GetFilenameFromRangeText(payload As String) As String
```
## GetPathFromRangeText
Returns the path from the full qualified global address of a Range object.
```vb
GetPathFromRangeText(payload As String) As String
```