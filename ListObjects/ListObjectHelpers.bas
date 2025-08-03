Attribute VB_Name = "ListObjectHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.ListObject"
Option Explicit

'@Description "Returns a Collection containing all the ListObjects in a given Workbook"
Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
Attribute GetAllListObjects.VB_Description = "Returns a Collection containing all the ListObjects in a given Workbook"
    Set GetAllListObjects = New Collection
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        Dim ListObject As ListObject
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function

'@Description "Tries to return the ListObject with the given name from a Collection of ListObjects."
Public Function TryGetListObjectFromCollection(ByVal TableCollection As Collection, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
Attribute TryGetListObjectFromCollection.VB_Description = "Tries to return the ListObject with the given name from a Collection of ListObjects."
    Dim ListObject As ListObject
    For Each ListObject In TableCollection
        If ListObjectName = ListObject.Name Then
            Set OutListObject = ListObject
            TryGetListObjectFromCollection = True
            Exit Function
        End If
    Next ListObject
End Function