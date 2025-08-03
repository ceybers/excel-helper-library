Attribute VB_Name = "RangeHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.Range"
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

' Returns a range with the same shape as the specified 2-dimensional InputArray, starting
' from the top-most cell in the specified InputRange.
Public Function ResizeRangeToArray(ByVal InputRange As Range, ByVal InputArray As Variant) As Range
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputArray)
    Set ResizeRangeToArray = InputRange.Cells.Item(1, 1).Resize(UBound(InputArray, 1), UBound(InputArray, 2))
End Function

' Updates the .Value2 property of all the cells in the InputRange with the Variant Values
' in the specified 2-dimensional InputArray.
Public Sub RangeSetValueFromVariant(ByVal InputRange As Range, ByVal InputVariant As Variant)
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputVariant)
    InputRange.Cells.Item(1, 1).Resize(UBound(InputVariant, 1), UBound(InputVariant, 2)).Value2 = InputVariant
End Sub

' Returns a range with the offset and size of the specified input parameters, starting from the
' top-most cell in the InputRange. Row = 1 and Column = 1 start the box from the top-left cell.
' e.g., RangeBox(Range("A1"), 1, 2, 4, 8).Address = B1:I4
Public Function RangeBox(ByVal InputRange As Range, ByVal Row As Long, ByVal Column As Long, _
    ByVal Rows As Long, ByVal Columns As Long) As Range
    Debug.Assert Not InputRange Is Nothing
    Debug.Assert Row > 0
    Debug.Assert Column > 0
    Debug.Assert Rows > 0
    Debug.Assert Columns > 0
    
    Set RangeBox = InputRange.Cells.Item(1, 1).Offset(Row - 1, Column - 1).Resize(Rows, Columns)
End Function