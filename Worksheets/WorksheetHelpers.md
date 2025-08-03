# Worksheet Helpers
Mostly used in `Table Split Tool`.

# Methods
## AddOrGetWorksheet
Returns the Worksheet with a given name. If it does not exist, create and return it.
```vb
AddOrGetWorksheet(Workbook As Workbook, WorksheetName As String) As Worksheet
```
## TryGetWorksheet
Tries to return the Worksheet with the given name from a Workbook.
```vb
TryGetWorksheet(Workbook As Workbook, WorksheetName As String, OutWorksheet As Worksheet) As Boolean
```
## TryRemoveWorksheet
Tries to remove a Worksheet with a given name from a Workbook.
```vb
TryRemoveWorksheet(Workbook As Workbook, WorksheetName As String) As Boolean
```
## WorksheetExists
Returns True if a Worksheet with the given name exists in a Workbook.
```vb
WorksheetExists(Workbook As Workbook, WorksheetName As String) As Boolean
```
## IsValidWorksheetName
Tests if a given string is a valid Worksheet name.
```vb
IsValidWorksheetName(SheetName As String) As Boolean
```

## GetWorksheetDatabodyRange
Returns the range of cells in a Worksheet starting from the first row beneath the header row ranging until the last row in the Used Range. Assumes header row is always Row 1. Returns Nothing if there are no rows.
```vb
GetWorksheetDatabodyRange(Worksheet As Worksheet) As Range
```