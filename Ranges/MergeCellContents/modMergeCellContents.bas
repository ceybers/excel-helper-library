Attribute VB_Name = "modMergeCellContents"
'@Folder("VBAProject")
Option Explicit

Private Const MSG_CAPTION As String = "Merge Cell Contents"
Private Const MSG_SHEET_PROTECTED As String = "This tool cannot work while the worksheet is protected." _
    & vbCrLf & vbCrLf & "Unprotect the worksheet and try again."
Private Const MSG_SHEET_NOT_RANGE As String = "Select the cells whose contents you wish to merge."
Private Const MSG_SHEET_ONE_ROW As String = "This tool merges the contents of all the cells in a selection into the top-most cells." _
    & vbCrLf & vbCrLf & "Select cells spanning two or more rows and try again."
Private Const MSG_SHEET_ENTIRE_ROW As String = "This tool cannot be applied to all the cells in an entire row." & _
    vbCrLf & vbCrLf & "Select only the cells you want to merge and try again."
Private Const MSG_SHEET_ENTIRE_COLUMN As String = "This tool cannot be applied to all the cells in an entire column." & _
    vbCrLf & vbCrLf & "Select only the cells you want to merge and try again."

'@EntryPoint
Public Sub MergeCellContents()
    If Selection Is Nothing Then Exit Sub
    If Not TypeOf Selection Is Range Then
        MsgBox MSG_SHEET_NOT_RANGE, vbExclamation + vbOKOnly, MSG_CAPTION
        Exit Sub
    End If
    
    Dim Range As Range
    Set Range = Selection
    
    Dim Worksheet As Worksheet
    Set Worksheet = Range.Parent
    
    If Worksheet.ProtectContents = True Then
        MsgBox MSG_SHEET_PROTECTED, vbExclamation + vbOKOnly, MSG_CAPTION
        Exit Sub
    End If
    
    If Not HasAreaTallerThanOneRow(Range) Then
        MsgBox MSG_SHEET_ONE_ROW, vbExclamation + vbOKOnly, MSG_CAPTION
        Exit Sub
    End If
    
    If Not TestAreasEntireRowOrColumn(Range) Then
        Exit Sub
    End If
    
    Dim Area As Range
    For Each Area In Range.Areas
        MergeCellContentsArea Area
    Next Area
End Sub

Private Function HasAreaTallerThanOneRow(ByVal Range As Range) As Boolean
    Dim Area As Range
    For Each Area In Range.Areas
        If Area.Rows.Count > 1 Then
            HasAreaTallerThanOneRow = True
            Exit Function
        End If
    Next Area
End Function

Private Function TestAreasEntireRowOrColumn(ByVal Range As Range) As Boolean
    Dim Area As Range
    For Each Area In Range
        If IsEntireColumn(Area) Then
            MsgBox MSG_SHEET_ENTIRE_COLUMN, vbExclamation + vbOKOnly, MSG_CAPTION
            Exit Function
        End If
        
        If IsEntireRow(Area) Then
            MsgBox MSG_SHEET_ENTIRE_ROW, vbExclamation + vbOKOnly, MSG_CAPTION
            Exit Function
        End If
    Next Area
    
    TestAreasEntireRowOrColumn = True
End Function

Private Function IsEntireColumn(ByVal Range As Range) As Boolean
    IsEntireColumn = (Range.Rows.Count = Range.EntireColumn.Rows.Count)
End Function

Private Function IsEntireRow(ByVal Range As Range) As Boolean
    IsEntireRow = (Range.Columns.Count = Range.EntireRow.Columns.Count)
End Function

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
