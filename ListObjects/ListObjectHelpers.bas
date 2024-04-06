Attribute VB_Name = "ListObjectHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder("Helpers")
Option Explicit

'@Description "Returns a Collection containing all the ListObjects in a given Workbook"
Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
Attribute GetAllListObjects.VB_Description = "Returns a Collection containing all the ListObjects in a given Workbook"
    Set GetAllListObjects = New Collection
    
    Dim Worksheet As Worksheets
    For Each Worksheet In Workbook.Worksheets
        Dim ListObject As ListObject
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function

Public Function TryGetListObjectFromCollection(ByVal TableCollection As Collection, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
    Dim ListObject As ListObject
    For Each ListObject In Collection
        If ListObjectName = ListObject.Name Then
            Set OutListObject = ListObject
            TryGetListObjectFromCollection = True
            Exit Function
        End If
    Next ListObject
End Function

' DEPREC
Public Function ZZZHasListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            HasListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function
