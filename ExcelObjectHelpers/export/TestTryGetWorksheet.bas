Attribute VB_Name = "TestTryGetWorksheet"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

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
    SetupTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestTryGetWorksheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Result As Worksheet
    
    'Act:
    If Not TryGetWorksheetByName(ThisWorkbook, "Sheet3", Result) Then
        Err.Description = "Get Worksheet in Workbook positive fail"
        GoTo TestFail
    End If
    
    If TryGetWorksheetByName(ThisWorkbook, "Sheet9", Result) Then
        Err.Description = "Get Worksheet in Workbook negative fail"
        GoTo TestFail
    End If
    
    Dim Collection As Collection
    Set Collection = New Collection
    Collection.Add ThisWorkbook.Worksheets.Item(1)
    Collection.Add ThisWorkbook.Worksheets.Item(2)
    
    If Not TryGetWorksheetByName(Collection, "Sheet1", Result) Then
        Err.Description = "Get Worksheet in Collection positive fail"
        GoTo TestFail
    End If
    
    If TryGetWorksheetByName(Collection, "Sheet9", Result) Then
        Err.Description = "Get Worksheet in Collection negative fail"
        GoTo TestFail
    End If
    
    If Not TryGetWorksheetByName(Application, "Sheet1", Result) Then
        Err.Description = "Get Worksheet in Application positive fail"
        GoTo TestFail
    End If
    
    If TryGetWorksheetByName(Application, "Sheet9", Result) Then
        Err.Description = "Get Worksheet in Application negative fail"
        GoTo TestFail
    End If
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

