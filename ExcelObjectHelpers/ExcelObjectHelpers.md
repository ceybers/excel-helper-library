# Excel Object Helpers
## ListObject
- Tries to return the first ListObject with the given name from a Worksheet, Workbook, Collection, or an Application object.
- The Collection may only contain ListObject objects - the function will not safely type check elements.
- The Worksheets collection in a Workbook and the Workbooks collection in an Application can also be used as the Parent.
```vb
Public Function TryGetListObjectByName(ByVal Parent As Object, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
```
## Workbook
- Tries to return the Wokbook with the given name from a Collection or an Application object.
```vb
Public Function TryGetWorkbookByName(ByVal Parent As Object, ByVal WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
```
## Worksheet
- Tries to return the first Worksheet with the given name from a Workbook, Collection, or an Application object (Workbooks).
```vb
Public Function TryGetWorksheetByName(ByVal Parent As Object, ByVal WorksheetName As String, ByRef OutWorksheet As Worksheet) As Boolean
```
- Tries to return a Collection of all the Worksheets with a given name from a Collection or an Application object.
```vb
Public Function TryGetWorksheetsByName(ByVal Parent As Object, ByVal WorksheetName As String, ByRef OutWorksheets As Collection) As Boolean
```
- Returns a Dictionary of all the Worksheets in an Application object. Keys are colon delimited strings of Workbook and Worksheet name.
```vb
Public Function GetDictionaryOfWorksheets(ByVal Application As Application) As Object
```