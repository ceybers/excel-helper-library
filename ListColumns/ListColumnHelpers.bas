' Collection manipulation helpers
' Not to be confused with VarType analysis helpers

'@Description "Tries to get a ListColumn with the given name from a ListObject. If successful, returns True and sets the Out variable to the ListColumn."
Public Function TryGetListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String, ByRef OutListColumn As ListColumn) As Boolean
    If ListObject Is Nothing Then Exit Function
    If ListColumnName = vbNullString Then Exit Function
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            Set OutListColumn = ListColumn
            TryGetListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

Public Function Exists(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Exists = TryGetListColumn(ListObject, ListColumnName)
End Function