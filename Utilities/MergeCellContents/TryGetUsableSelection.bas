Attribute VB_Name = "modTryGetUsableSelection"
'@Folder("VBAProject")
Option Explicit

Private Const ERR_SRC As String = "TryGetUsableSelection"

Private Const ERR_NUM_SHEET_PROTECTED As Long = vbObjectError + 1
Private Const ERR_MSG_SHEET_PROTECTED As String = "This tool cannot work while the worksheet is protected." _
    & vbCrLf & vbCrLf & "Unprotect the worksheet and try again."

Private Const ERR_NUM_SHEET_NOT_RANGE As Long = vbObjectError + 2
Private Const ERR_MSG_SHEET_NOT_RANGE As String = "Select the cells whose contents you wish to merge."

Private Const ERR_NUM_SHEET_ONE_ROW As Long = vbObjectError + 3
Private Const ERR_MSG_SHEET_ONE_ROW As String = "This tool merges the contents of all the cells in a selection into the top-most cells." _
    & vbCrLf & vbCrLf & "Select cells spanning two or more rows and try again."

Private Const ERR_NUM_SHEET_ENTIRE_ROW As Long = vbObjectError + 4
Private Const ERR_MSG_SHEET_ENTIRE_ROW As String = "This tool cannot be applied to all the cells in an entire row." & _
    vbCrLf & vbCrLf & "Select only the cells you want to merge and try again."

Private Const ERR_NUM_SHEET_ENTIRE_COLUMN As Long = vbObjectError + 5
Private Const ERR_MSG_SHEET_ENTIRE_COLUMN As String = "This tool cannot be applied to all the cells in an entire column." & _
    vbCrLf & vbCrLf & "Select only the cells you want to merge and try again."

Public Function TryGetUsableSelection(ByVal Caption As String, ByRef OutRange As Range) As Boolean
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then Exit Function
    If Not TypeOf Selection Is Range Then
        Err.Raise ERR_NUM_SHEET_NOT_RANGE, ERR_SRC, ERR_MSG_SHEET_NOT_RANGE
    End If
    
    Dim Range As Range
    Set Range = Selection
    
    Dim Worksheet As Worksheet
    Set Worksheet = Range.Parent
    
    If Worksheet.ProtectContents = True Then
        Err.Raise ERR_NUM_SHEET_PROTECTED, ERR_SRC, ERR_MSG_SHEET_PROTECTED
    End If
    
    If Not HasAreaTallerThanOneRow(Range) Then
        Err.Raise ERR_NUM_SHEET_ONE_ROW, ERR_SRC, ERR_MSG_SHEET_ONE_ROW
    End If
    
    If Not TestAreasEntireRowOrColumn(Range) Then
        Exit Function
    End If
    
    Set OutRange = Range
    TryGetUsableSelection = True
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation + vbOKOnly, Caption
End Function

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
    For Each Area In Range.Areas
        If IsEntireColumn(Area) Then
            Err.Raise ERR_NUM_SHEET_ENTIRE_COLUMN, ERR_SRC, ERR_MSG_SHEET_ENTIRE_COLUMN
        End If
        
        If IsEntireRow(Area) Then
            Err.Raise ERR_NUM_SHEET_ENTIRE_ROW, ERR_SRC, ERR_MSG_SHEET_ENTIRE_ROW
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
