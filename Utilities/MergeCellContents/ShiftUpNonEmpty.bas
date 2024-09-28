Attribute VB_Name = "modShiftUpNonEmpty"
'@Folder("VBAProject")
Option Explicit

Private Const MSG_CAPTION As String = "Shift Up Non-Empty"

'@EntryPoint
Public Sub ShiftUpNonEmpty()
    Dim Range As Range
    If Not TryGetUsableSelection(MSG_CAPTION, Range) Then Exit Sub
    
    ShiftUpNonEmptyAreas Range
End Sub

Private Sub ShiftUpNonEmptyAreas(ByVal Range As Range)
    Dim Area As Range
    For Each Area In Range.Areas
        ShiftUpNonEmptyArea Range
    Next Area
End Sub

Private Sub ShiftUpNonEmptyArea(ByVal Range As Range)
    Dim vv As Variant
    vv = Range.Value2
    
    Dim RowCount As Long
    RowCount = Range.Rows.Count
    
    Dim ColumnCount As Long
    ColumnCount = Range.Columns.Count
    
    Dim Map() As Long
    ReDim Map(1 To RowCount)
    
    Dim MapCursor As Long
    MapCursor = 0
    
    Dim i As Long
    For i = 1 To RowCount
        If Not IsValueRowEmpty(vv, i) Then
            MapCursor = MapCursor + 1
            Map(i) = MapCursor
        End If
    Next i
    
    Dim ColumnIndex As Long
    For ColumnIndex = 1 To ColumnCount
        Dim RowIndex As Long
        For RowIndex = 1 To RowCount
            If Map(RowIndex) > 0 Then
                vv(Map(RowIndex), ColumnIndex) = vv(RowIndex, ColumnIndex)
            End If
        Next RowIndex
        
        For RowIndex = (MapCursor + 1) To RowCount
            vv(RowIndex, ColumnIndex) = vbNullString
        Next RowIndex
    Next ColumnIndex
    
    Range.Value2 = vv
End Sub

Private Function IsValueRowEmpty(ByVal vv As Variant, ByVal RowIndex As Long) As Boolean
    Dim i As Long
    For i = 1 To UBound(vv, 2)
        If Not IsEmpty(vv(RowIndex, i)) Then
            IsValueRowEmpty = False
            Exit Function
        End If
    Next i
    
    IsValueRowEmpty = True
End Function
