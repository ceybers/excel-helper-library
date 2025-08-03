# Workbook Helpers
Helper functions for the Workbook object.

# Methods
## TryGetWorkbook
Tries to return the Workbook with the given name if it is currently open in this instance of Excel.
```vb
TryGetWorkbook(WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
```

## IsWorkbookOpen
Returns True if a Workbook with the given filename is currently open.
```vb
IsWorkbookOpen(filename As String) As Boolean
```

## TryGetWorkbook2
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