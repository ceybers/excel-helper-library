Attribute VB_Name = "TestHashCells"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Dim Worksheet As Worksheet
Dim TargetRange As Range

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Worksheet = ThisWorkbook.Worksheets("HashBenchmark")
    Set TargetRange = Worksheet.Range("D1:F3")
    
    With Worksheet
        .Range("D1:F1").Value2 = Array("abc", "def", "ghi")
        .Range("D2:F2").Value2 = Array("jkl", "", "pqr")
        .Range("D3:F3").Value2 = Array("stu", "vwx", "yz0")
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
With Worksheet
        .Range("D1:F1").Value2 = vbNullString
        .Range("D2:F2").Value2 = vbNullString
        .Range("D3:F3").Value2 = vbNullString
    End With
End Sub

'@TestMethod("TestHashCells")
Private Sub TestHashCells()
    On Error GoTo TestFail
    
    'Arrange:
    'Act:
    If HashCellsSHA1(TargetRange) <> "AA522EC60B4CAC5433A0415FADACD0674FF2735D" Then
        Err.Description = "HashCellsSHA1(TargetRange) FAIL"
        GoTo TestFail
    End If
    
    If HashCellsSHA256(TargetRange) <> "453C3ABA3D3B08D8383625ABB0B0D063BA7C025FF636EB547E16776790226C4A" Then
        Err.Description = "HashCellsSHA256(TargetRange) FAIL"
        GoTo TestFail
    End If
    
    If HashCellsMD5(TargetRange) <> "49EDBF9D44C39DA7A9BEE47CEE66E7F0" Then
        Err.Description = "HashCellsMD5(TargetRange) FAIL"
        GoTo TestFail
    End If
    
    If HashCellsSHA1(Worksheet.Range("E2")) <> "9ADF325316600097106AE2B76BE92E8BA2FCC8DC" Then
        Err.Description = "HashCellsSHA1(TargetRange) on 1x empty cell FAIL"
        GoTo TestFail
    End If

    'Assert:
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
