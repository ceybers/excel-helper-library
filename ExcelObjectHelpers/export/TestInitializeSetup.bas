Attribute VB_Name = "TestInitializeSetup"
'@Folder("Tests")
Option Explicit

Private Const SECOND_TEST_FILENAME As String = "C:\Users\User\Documents\Work\excel-helper-library\ExcelObjectHelpers\SecondTestSheet.xlsx"

Public Sub SetupTest()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    wb.Worksheets.Add wb.Worksheets.Item(1)
    
    Dim i As Long
    For i = wb.Worksheets.Count To 2 Step -1
        Application.DisplayAlerts = False
        wb.Worksheets.Item(i).Delete
        Application.DisplayAlerts = True
    Next i
    
    wb.Worksheets.Item(1).Name = "Sheet1"
    
    With wb.Worksheets(1).Range("A1:C3")
        .Value2 = "A"
        wb.Worksheets(1).ListObjects.Add(xlSrcRange, wb.Worksheets(1).Range("A1:C3"), , xlNo).Name = _
        "Table1"
    End With
    
    wb.Worksheets.Add After:=wb.Worksheets.Item(wb.Worksheets.Count)
    wb.Worksheets.Item(2).Name = "Sheet2"
    With wb.Worksheets(2).Range("A1:C3")
        .Value2 = "A"
        wb.Worksheets(2).ListObjects.Add(xlSrcRange, wb.Worksheets(2).Range("A1:C3"), , xlNo).Name = _
        "Table1"
    End With
    
    wb.Worksheets.Add After:=wb.Worksheets.Item(wb.Worksheets.Count)
    wb.Worksheets.Item(3).Name = "Sheet3"
    With wb.Worksheets(3).Range("A1:C3")
        .Value2 = "A"
        wb.Worksheets(3).ListObjects.Add(xlSrcRange, wb.Worksheets(3).Range("A1:C3"), , xlNo).Name = _
        "Table1"
    End With
    
    Workbooks.Open Filename:=SECOND_TEST_FILENAME
End Sub
