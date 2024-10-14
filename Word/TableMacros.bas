Option Explicit

'@Description "Sets all tables to not float or overlap content."
'@EntryPoint
Public Sub NoFloatingTables()
    Dim Table As Table
    For Each Table In ActiveDocument.Tables
        With Table.Rows
            .AllowOverlap = False
            .Alignment = wdAlignRowLeft
            .WrapAroundText = False
        End With
    Next Table
End Sub

'@Description "Expands selection to the entire table."
'@EntryPoint
Public Sub SelectTable()
    If Selection.Tables.Count <> 1 Then Exit Sub

    Selection.Tables.Item(1).Select
End Sub

'@Description "Moves the cursor to immediately after the table."
'@EntryPoint
Public Sub MoveCursorAfterTable()
    If Selection.Tables.Count <> 1 Then Exit Sub

    Dim Range As Range
    Set Range = Selection.Tables.Item(1).Range
    
    Range.Collapse wdCollapseEnd
    Range.Select
End Sub

'@Description "Unmerges all the cells in a table."
'@EntryPoint
Public Sub UnmergeCellsInTables()
    Dim Tables As Tables
    If Application.Selection.Tables.Count > 1 Then
        Set Tables = Application.Selection.Tables
    Else
        Set Tables = ActiveDocument.Tables
    End If
    
    Dim Table As Table
    For Each Table In Tables
        UnmergeCellsInTable Table
    Next Table
End Sub

Private Sub UnmergeCellsInTable(ByVal Table As Table)
    Dim Column As Column
    For Each Column In Table.Columns
        UnmergeCellsInColumn Table, Column
    Next Column
End Sub

Private Sub UnmergeCellsInColumn(ByVal Table As Table, ByVal Column As Column)
    Dim TotalRows As Long
    TotalRows = Table.Rows.Count
    
    Dim ColumnRows As Long
    ColumnRows = Column.Cells.Count
    
    If TotalRows = ColumnRows Then Exit Sub
    
    Dim Bitmap() As Long
    ReDim Bitmap(1 To TotalRows)
    
    ' Create a 1-dimensional array containing each row in the table.
    ' Set an element to 1 if this column has a cell in that row.
    ' Leave element as 0 if this column doesn't have a cell
    ' (i.e., it is merged with the cell above)
    Dim i As Long
    For i = 1 To ColumnRows
        Bitmap(Column.Cells.Item(i).RowIndex) = 1
    Next i

    Dim Cursor As Long
    Cursor = 1
    Dim Counter As Long
    Counter = 1
    Dim SplitMap() As Long
    ReDim SplitMap(1 To ColumnRows)
    
    ' Create a 1-dimensional array for each of the cells present in this column.
    ' Set the element to the number of rows this cell spans relative to the table's rows.
    ' Value of 1 means that none of the cells in this row are merged.
    ' Values of greater than 1 mean that this cell spans multiple rows.
    For i = 2 To TotalRows
        If Bitmap(i) = 1 Then
            SplitMap(Cursor) = Counter
            Cursor = Cursor + 1
            Counter = 1
        Else
            Counter = Counter + 1
        End If
    Next i
    SplitMap(Cursor) = Counter
    
    ' Split each cell in this column based on how many rows it spans.
    For i = ColumnRows To 1 Step -1
        If SplitMap(i) > 1 Then
            Column.Cells.Item(i).Split SplitMap(i), 1
        End If
    Next i
End Sub