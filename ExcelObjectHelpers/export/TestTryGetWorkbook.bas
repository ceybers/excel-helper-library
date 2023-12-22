Attribute VB_Name = "TestTryGetWorkbook"
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
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestTryGetWorkbook()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Result As Workbook
    
    'Act:
    If Not TryGetWorkbookByName(Application, ThisWorkbook.Name, Result) Then
        Err.Description = "TryGetWorkbookInApplication positive failed"
        GoTo TestFail
    End If
    
    If TryGetWorkbookByName(Application, "zzz", Result) Then
        Err.Description = "TryGetWorkbookInApplication negative failed"
        GoTo TestFail
    End If
    
    Dim Collection As Collection
    Set Collection = New Collection
    Collection.Add ThisWorkbook
    Collection.Add Application.Workbooks.Item(2)
    
    If Not TryGetWorkbookByName(Collection, ThisWorkbook.Name, Result) Then
        Err.Description = "TryGetWorkbookInCollection positive failed"
        GoTo TestFail
    End If
    
    If TryGetWorkbookByName(Collection, "zzz", Result) Then
        Err.Description = "TryGetWorkbookInCollection negative failed"
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
