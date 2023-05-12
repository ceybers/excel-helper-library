Attribute VB_Name = "ListObjectHelpers"
'@Folder("Helpers")
Option Explicit

'@Description "Returns a Collection containing all the ListObjects in a given Workbook"
Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
Attribute GetAllListObjects.VB_Description = "Returns a Collection containing all the ListObjects in a given Workbook"
    Set GetAllListObjects = New Collection
    
    If Workbook Is Nothing Then Exit Function
    
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Worksheet In Workbook.Worksheets
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function
