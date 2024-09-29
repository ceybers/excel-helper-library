Attribute VB_Name = "RangeHelpers"
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder("HelperFunctions")
Option Explicit

Private Sub TestAppendRange()
    Dim runningRange As Range
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    Dim sel As Range
    
    Set sel = ws.Range("A2")
    Debug.Print sel.Value2
    AppendRange sel, runningRange

    Set sel = ws.Range("b6")
    Debug.Print sel.Value2
    AppendRange sel, runningRange

    Set sel = ThisWorkbook.Worksheets(2).Range("b7")
    Debug.Print sel.Value2
    AppendRange sel, runningRange
    
    Debug.Print runningRange.Address
End Sub

Public Sub AppendRange(ByVal rangeToAppend As Range, ByRef unionRange As Range)
    If rangeToAppend Is Nothing Then Exit Sub
    
    If unionRange Is Nothing Then
        Set unionRange = rangeToAppend
        Exit Sub
    End If
    
    If Not rangeToAppend.parent Is unionRange.parent Then Exit Sub
    
    Set unionRange = Application.Union(unionRange, rangeToAppend)
End Sub

'@Description "Returns True if SpecialCells would have returned a Range. Returns False if no cells were selected."
Public Function HasSpecialCells(ByVal Range As Range, ByVal CellType As XlCellType, _
    Optional ByVal Value As XlSpecialCellsValue) As Boolean
    If Range Is Nothing Then Exit Function

    Dim Result As Range
    On Error Resume Next
    If Value = 0 Then
        Set Result = Range.SpecialCells(CellType)
    Else
        Set Result = Range.SpecialCells(CellType, Value)
    End If
    
    HasSpecialCells = (Not Result Is Nothing)
End Function

'@Description "Tries to get the Range that contains all the hidden cells in a Range."
Public Function TryGetHiddenCellsInRange(ByVal InputRange As Range, ByRef OutputRange As Range) As Boolean
    If InputRange Is Nothing Then Exit Function
    
    Dim VisibleRange As Range
    Set VisibleRange = InputRange.SpecialCells(xlCellTypeVisible)
    If InputRange.Cells.Count = VisibleRange.Cells.Count Then Exit Function
    
    Dim Result As Range
    Dim Cell As Range
    For Each Cell In InputRange.Cells
        If Cell.ColumnWidth = 0 Or Cell.RowHeight = 0 Then
            If Result Is Nothing Then
                Set Result = Cell
            Else
                Set Result = Application.Union(Result, Cell)
            End If
        End If
    Next Cell
    
    Set OutputRange = Result
    TryGetHiddenCellsInRange = True
End Function