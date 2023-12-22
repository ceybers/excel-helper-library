Attribute VB_Name = "TestTryGetListObject"
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
Private Sub TestTryGetListObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Result As ListObject
    
    'Act:
    If Not TryGetListObjectByName(ThisWorkbook.Worksheets.Item(1), "Table1", Result) Then
        Err.Description = "TryGetListObjectInWorksheet positive fail"
        GoTo TestFail
    End If
    
    If TryGetListObjectByName(ThisWorkbook.Worksheets.Item(1), "Table2", Result) Then
        Err.Description = "TryGetListObjectInWorksheet negative fail"
        GoTo TestFail
    End If
    
    Dim Collection As Collection
    Set Collection = New Collection
    Collection.Add ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Collection.Add ThisWorkbook.Worksheets.Item(2).ListObjects.Item(1)
    
    If Not TryGetListObjectByName(Collection, "Table1", Result) Then
        Err.Description = "Get ListObject in Collection positive fail"
        GoTo TestFail
    End If
    
    If TryGetListObjectByName(Collection, "Table9", Result) Then
        Err.Description = "Get ListObject in Collection negative fail"
        GoTo TestFail
    End If
    
    If Not TryGetListObjectByName(ThisWorkbook, "Table1", Result) Then
        Err.Description = "Get ListObject in Workbook positive fail"
        GoTo TestFail
    End If
    
    If TryGetListObjectByName(ThisWorkbook, "Table9", Result) Then
        Err.Description = "Get ListObject in Workbook negative fail"
        GoTo TestFail
    End If
    
    If Not TryGetListObjectByName(Application, "Table1", Result) Then
        Err.Description = "Get ListObject in Application positive fail"
        GoTo TestFail
    End If
    
    If TryGetListObjectByName(Application, "Table9", Result) Then
        Err.Description = "Get ListObject in Application negative fail"
        GoTo TestFail
    End If
    
    If Not TryGetListObjectByName(Application.Workbooks, "Table1", Result) Then
        Err.Description = "Get ListObject in Application positive fail"
        GoTo TestFail
    End If
    
    If TryGetListObjectByName(Application.Workbooks, "Table9", Result) Then
        Err.Description = "Get ListObject in Application negative fail"
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
