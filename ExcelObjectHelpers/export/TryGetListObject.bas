Attribute VB_Name = "TryGetListObject"
'@Folder("CommonHelpers")
Option Explicit

'@Description "Tries to return the first ListObject with the given name from a Worksheet, Workbook, Collection, or an Application object."
Public Function TryGetListObjectByName(ByVal Parent As Object, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
    If TypeOf Parent Is Worksheet Then
        TryGetListObjectByName = TryGetListObjectInEnumerableByName(Parent.ListObjects, ListObjectName, OutListObject)
    ElseIf TypeOf Parent Is Collection Then
        TryGetListObjectByName = TryGetListObjectInEnumerableByName(Parent, ListObjectName, OutListObject)
    ElseIf TypeOf Parent Is Workbook Then
        TryGetListObjectByName = TryGetListObjectInEnumerableByName(GetAllListObjectsInWorkbook(Parent.Worksheets), ListObjectName, OutListObject)
    ElseIf TypeOf Parent Is Worksheets Then
        TryGetListObjectByName = TryGetListObjectInEnumerableByName(GetAllListObjectsInWorkbook(Parent), ListObjectName, OutListObject)
    ElseIf TypeOf Parent Is Application Then
        TryGetListObjectByName = TryGetListObjectInEnumerableByName(GetAllListObjectsInApplication(Parent.Workbooks), ListObjectName, OutListObject)
    ElseIf TypeOf Parent Is Workbooks Then
        TryGetListObjectByName = TryGetListObjectInEnumerableByName(GetAllListObjectsInApplication(Parent), ListObjectName, OutListObject)
    End If
End Function

Private Function TryGetListObjectInEnumerableByName(ByVal Enumerable As Object, ByVal ListObjectName As String, ByRef OutListObject As ListObject) As Boolean
    Dim ListObject As ListObject
    For Each ListObject In Enumerable
        If ListObject.Name Like ListObjectName Then
            Set OutListObject = ListObject
            TryGetListObjectInEnumerableByName = True
            Exit Function
        End If
    Next ListObject
End Function

Private Function GetAllListObjectsInWorkbook(ByVal Enumerable As Object) As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim ListObject As ListObject
    Dim Worksheet As Worksheet
    For Each Worksheet In Enumerable
        For Each ListObject In Worksheet.ListObjects
            Result.Add ListObject
        Next ListObject
    Next Worksheet
    
    Set GetAllListObjectsInWorkbook = Result
End Function

Private Function GetAllListObjectsInApplication(ByVal Enumerable As Object) As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim ListObject As ListObject
    Dim Workbook As Workbook
    For Each Workbook In Enumerable
        For Each ListObject In GetAllListObjectsInWorkbook(Workbook.Worksheets)
            Result.Add ListObject
        Next ListObject
    Next Workbook
    
    Set GetAllListObjectsInApplication = Result
End Function
