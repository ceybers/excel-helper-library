' Collection manipulation helpers
' Not to be confused with VarType analysis helpers

' Returns True and passes a ListColumn by reference if it is found in the specified ListObject
'  Query of types Integer or Long returns the ListColumn by index in the ListObject
'  Query of type String tries to find a ListColumn with that exact name (case-sensitive)
'  Query of type Range will pass the intersecting ListColumn if the range is exactly one column wide
'  Query of type ListColumn will return true if the ListColumn exists in the ListObject
Public Function TryGetListColumnByVariant(ByVal ListObject As ListObject, ByVal Query As Variant, _
    ByRef OutListColumn As ListColumn) As Boolean
    
    If IsObject(Query) Then
        If TypeOf Query Is Range Then
            Dim Range As Range
            Set Range = Query
            If Range.Columns.Count <> 1 Then Exit Function
            Dim Intersection As Range
            Set Intersection = Intersect(Range.EntireColumn, ListObject.HeaderRowRange)
            If Intersection Is Nothing Then Exit Function
            Set OutListColumn = ListObject.ListColumns.Item(Intersection.Column - ListObject.Range.Cells(1, 1).Column + 1)
            TryGetListColumnByVariant = True
            Exit Function
        ElseIf TypeOf Query Is ListColumn Then
            Dim ListColumn As ListColumn
            Set ListColumn = Query
            If ListColumn.Parent Is ListObject Then
                Set OutListColumn = ListColumn
                TryGetListColumnByVariant = True
                Exit Function
            End If
        End If
    ElseIf IsNumeric(Query) Then
        Dim ListColumnIndex As Long
        ListColumnIndex = CLng(Query)
        If ListColumnIndex < 1 Then Exit Function
        If ListColumnIndex > ListObject.ListColumns.Count Then Exit Function
        Set OutListColumn = ListObject.ListColumns.Item(ListColumnIndex)
        TryGetListColumnByVariant = True
        Exit Function
    ElseIf VarType(Query) = vbString Then
        Dim i As Long
        For i = 1 To ListObject.ListColumns.Count
            If ListObject.ListColumns.Item(i).Name = Query Then
                Set OutListColumn = ListObject.ListColumns.Item(i)
                TryGetListColumnByVariant = True
                Exit Function
            End If
        Next i
    End If
End Function

' Returns a collection of ListColumns that are headers of the specified Range.
' The Keys of the collection are the ListColumn Names.
Public Function GetListColumnsFromRange(ByVal Range As Range) As Collection
    Dim Result As Collection
    Set Result = New Collection
    Set GetListColumnsFromRange = Result
    
    If Range.ListObject Is Nothing Then Exit Function
    
    Dim IntersectedCells As Range
    Set IntersectedCells = Intersect(Range.EntireColumn, Range.ListObject.HeaderRowRange)
    
    Dim Cell As Range
    For Each Cell In IntersectedCells
        Dim Index As Long
        Index = Cell.Column - Range.ListObject.Range.Cells(1, 1).Column + 1
        Dim ListColumn As ListColumn
        Set ListColumn = Range.ListObject.ListColumns.Item(Index)
        Result.Add Item:=ListColumn, Key:=ListColumn.Name
    Next Cell
End Function

' Returns True and passes a ListColumn by reference with the specified name if it exists in the ListObject.
Public Function TryGetListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String, _
    ByRef OutListColumn As ListColumn) As Boolean
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