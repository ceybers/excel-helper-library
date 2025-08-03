# ListObject Helpers
Some project-specific helper functions for ListObjects.

## Functions
```vb
' Returns a Collection containing all the ListObjects in a given Workbook.
GetAllListObjects(ByVal Workbook As Workbook) As Collection

' Tries to return the ListObject with the given name from a Collection of ListObjects.
TryGetListObjectFromCollection(ByVal TableCollection As Collection, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
```