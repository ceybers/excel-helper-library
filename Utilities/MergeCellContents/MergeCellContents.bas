Attribute VB_Name = "modMergeCellContents"
'@Folder("VBAProject")
Option Explicit

Private Const MSG_CAPTION As String = "Merge Cell Contents"

'@EntryPoint
Public Sub MergeCellContents()
    Dim Range As Range
    If Not TryGetUsableSelection(MSG_CAPTION, Range) Then Exit Sub
    
    MergeCellContentsAreas Range
End Sub

Private Sub MergeCellContentsAreas(ByVal Range As Range)
    Dim Area As Range
    For Each Area In Range.Areas
        MergeCellContentsArea Area
    Next Area
End Sub

Private Sub MergeCellContentsArea(ByVal Area As Range)
    If Area.Rows.Count = 1 Then Exit Sub
    
    Dim vv As Variant
    vv = Area.Value2
    
    Dim RowCount As Long
    RowCount = Area.Rows.Count
    
    Dim ColumnCount As Long
    ColumnCount = Area.Columns.Count
    
    Dim Column As Long
    For Column = 1 To ColumnCount
        Dim RowValues() As Variant
        ReDim RowValues(1 To RowCount)
    
        Dim Row As Long
        For Row = 1 To RowCount
            RowValues(Row) = vv(Row, Column)
            vv(Row, Column) = vbNullString
            
            If VarType(RowValues(Row)) = vbError Then
                RowValues(Row) = CStr(RowValues(Row))
            End If
        Next Row
        
        vv(1, Column) = Join$(RowValues, vbCrLf)
    Next Column
    
    Area.Value2 = vv
End Sub
