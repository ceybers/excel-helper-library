# ListObject Helpers
Some project-specific helper functions for ListObjects. I still need to refactor these to be reusable. Lots of messy code to try and parse strings into ListObjects that may or may not be present in the referenced workbook.

## Functions
```vb
GetAllListObjects(ByVal Workbook As Workbook) As Collection
TableFromString(ByVal s As String) As ListObject
TableToString(ByVal lo As ListObject) As String
ToKey(ByVal i As Integer) As String
```

## Methods
```vb
PasteArrayIntoWorksheet(ByRef arr As Variant, ByVal ws As Worksheet, Optional ByVal row As Long = 1, Optional ByVal column As Long = 1)
```