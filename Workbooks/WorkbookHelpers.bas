Attribute VB_Name = "WorkbookHelpers"
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder("HelperFunctions")
Option Explicit

'@Description "Tries to return the Workbook with the given name if it is currently open in this instance of Excel."
Public Function TryGetWorkbook(ByVal WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
    If WorkbookName = vbNullString Then Exit Function
    
    Dim Workbook As Workbook
    For Each Workbook In Application.Workbooks
        If Workbook.Name = WorkbookName Then
            Set OutWorkbook = Workbook
            TryGetWorkbook = True
            Exit Function
        End If
    Next Workbook
End Function

Public Function GetPathFromRangeText(ByVal payload As String) As String
    Dim a As Integer
    Dim b As Integer
    a = InStr(payload, "'")
    b = InStr(payload, "[")
    If a = 0 Or b = 0 Then Exit Function
    GetPathFromRangeText = Mid$(payload, a + 1, b - a - 1)
End Function

Public Function GetFilenameFromRangeText(ByVal payload As String) As String
    Dim b As Integer
    Dim c As Integer
    b = InStr(payload, "[")
    c = InStr(payload, "]")
    If b = 0 Or c = 0 Then Exit Function
    GetFilenameFromRangeText = Mid$(payload, b + 1, c - b - 1)
End Function

Public Function IsWorkbookOpen(ByVal filename As String) As Boolean
    Dim wb As Workbook
    If filename = vbNullString Then Exit Function
    
    For Each wb In Application.Workbooks
        If wb.Name = filename Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

Public Function TryGetWorkbook2(ByVal filename As String, ByRef wb As Workbook, Optional path As String = vbNullString) As Boolean
    Dim curWB As Workbook
    For Each curWB In Application.Workbooks
        If path = vbNullString Then
            If curWB.Name = filename Then
                Set wb = curWB
                TryGetWorkbook = True
                Exit Function
            End If
        Else
            If curWB.fullname = path & filename Then
                Set wb = curWB
                TryGetWorkbook = True
                Exit Function
            End If
        End If
    Next curWB
End Function