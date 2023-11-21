Attribute VB_Name = "WorksheetHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers"
Option Explicit

' Reference: https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9

'@Description "Tests if a given string is a valid Worksheet name"
Public Function IsValidSheetName(ByVal SheetName As String) As Boolean
Attribute IsValidSheetName.VB_Description = "Tests if a given string is a valid Worksheet name"
    If SheetName = vbNullString Then Exit Function
    If Len(SheetName) > 31 Then Exit Function
    If UCase$(SheetName) = "HISTORY" Then Exit Function
    If Left$(SheetName, 1) = "'" Then Exit Function

    Dim InvalidChars As Variant
    InvalidChars = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i As Long

    For i = 1 To UBound(InvalidChars)
        If InStr(SheetName, InvalidChars(i)) > 0 Then Exit Function
    Next i

    IsValidSheetName = True
End Function

'@Description "Tries to remove a Worksheet with a given name from a Workbook."
Public Function TryRemoveSheet(ByVal Workbook As Workbook, ByVal SheetName As String) As Boolean
Attribute TryRemoveSheet.VB_Description = "Tries to remove a Worksheet with a given name from a Workbook."
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        If Worksheet.Name = SheetName Then
            Application.DisplayAlerts = False
            Worksheet.Delete
            Application.DisplayAlerts = True
            TryRemoveSheet = True
            Exit Function
        End If
    Next Worksheet
End Function

Public Function SheetExists(ByVal Workbook As Workbook, ByVal SheetName As String) As Boolean
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        If Worksheet.Name = SheetName Then
            SheetExists = True
            Exit Function
        End If
    Next Worksheet
End Function

Public Function AddOrGetWorksheet(ByVal worksheetName As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim prevWS As Worksheet
    
    Set wb = ActiveWorkbook
    Set prevWS = ActiveSheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = worksheetName
    End If
    
    prevWS.Activate
    
    'ws.Visible = xlSheetVeryHidden
    
    Set AddOrGetWorksheet = ws
End Function

Public Function TryGetWorkSheet(ByVal wb As Workbook, ByVal worksheetName As String, ByRef ws As Worksheet) As Boolean
    Dim curWS As Worksheet
    For Each curWS In wb.Worksheets
        If curWS.Name = worksheetName Then
            Set ws = curWS
            TryGetWorkSheet = True
        End If
    Next curWS
End Function