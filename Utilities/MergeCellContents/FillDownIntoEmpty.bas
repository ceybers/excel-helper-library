Attribute VB_Name = "modFillDownIntoEmpty"
'@Folder("VBAProject")
Option Explicit

Private Const MSG_CAPTION As String = "Fill Down Into Empty"

'@EntryPoint
Public Sub FillDownIntoEmpty()
    Dim Range As Range
    If Not TryGetUsableSelection(MSG_CAPTION, Range) Then Exit Sub
    
    FillDownIntoEmptyAreas Range
End Sub

Private Sub FillDownIntoEmptyAreas(ByVal Range As Range)
    Dim Area As Range
    For Each Area In Selection.Areas
        FillDownIntoEmptyArea Area
    Next Area
End Sub

Private Sub FillDownIntoEmptyArea(ByVal Range As Range)
    Dim Column As Range
    For Each Column In Range.Columns
        FillDownIntoEmptyColumn Column
    Next Column
End Sub

Private Sub FillDownIntoEmptyColumn(ByVal Range As Range)
    Dim vv As Variant
    vv = Range.Value2
    
    Dim RowIndex As Long
    For RowIndex = 2 To UBound(vv, 1)
        If Not IsError(vv(RowIndex, 1)) And Not IsError(vv(RowIndex - 1, 1)) Then
            If vv(RowIndex, 1) = vbNullString Then
                vv(RowIndex, 1) = vv(RowIndex - 1, 1)
            End If
        End If
    Next RowIndex
    
    Range.Value2 = vv
End Sub
