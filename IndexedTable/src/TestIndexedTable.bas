Attribute VB_Name = "TestIndexedTable"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Private ListObject As ListObject

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
    Set ListObject = ActiveSheet.ListObjects.Item(1)
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("Setter Getters")
Private Sub TestInit()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestIndexedTable As IndexedTable
    Set TestIndexedTable = New IndexedTable
    TestIndexedTable.Load ListObject, "Key Column"
    
    If TestIndexedTable.IsValid = False Then
        Err.Description = ".IsValid = False"
        GoTo TestFail
    End If
    
    If TestIndexedTable.HasKey("A3") = False Then
        Err.Description = ".HasKey = False"
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

'@TestMethod("Setter Getters")
Private Sub TestGetValue()
    On Error GoTo TestFail
    
    'Arrange:
    ListObject.DataBodyRange.Cells(2, 1).Value2 = "A3"
    ListObject.DataBodyRange.Cells(2, 2).Value2 = "B3"
    ListObject.DataBodyRange.Cells(2, 3).Value2 = "C3"
    
    Dim TestIndexedTable As IndexedTable
    Set TestIndexedTable = New IndexedTable
    TestIndexedTable.Load ListObject, "Key Column"

    If TestIndexedTable("A3", "Foo") <> "B3" Then
        Err.Description = ".Item() failed"
        GoTo TestFail
    End If
    
    Dim OutValue As Variant
    If TestIndexedTable.TryGetValue("A3", "Foo", OutValue) = False Then
        Err.Description = "TryGetValue failed"
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

'@TestMethod("Setter Getters")
Private Sub TestSetValue()
    On Error GoTo TestFail
    
    'Arrange:
    ListObject.DataBodyRange.Cells(2, 1).Value2 = "A3"
    ListObject.DataBodyRange.Cells(2, 2).Value2 = "B3"
    ListObject.DataBodyRange.Cells(2, 3).Value2 = "C3"
    
    Dim TestIndexedTable As IndexedTable
    Set TestIndexedTable = New IndexedTable
    TestIndexedTable.Load ListObject, "Key Column"
    
    TestIndexedTable("A3", "Foo") = "foobar"
    If TestIndexedTable("A3", "Foo") <> "foobar" Then
        Err.Description = ".Item() failed"
        GoTo TestFail
    End If
    
    Dim NewValue As Variant
    NewValue = "barfoo"
    
    If TestIndexedTable.TrySetValue("A3", "Foo", NewValue) = False Then
        Err.Description = "TrySetValue failed"
        GoTo TestFail
    End If
    If TestIndexedTable("A3", "Foo") <> NewValue Then
        Err.Description = "TrySetValue did not set"
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


'@TestMethod("Setter Getters")
Private Sub TestGetRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestIndexedTable As IndexedTable
    Set TestIndexedTable = New IndexedTable
    TestIndexedTable.Load ListObject, "Key Column"
    
    Dim RefRange As Range
    Set RefRange = ActiveSheet.Range("B3")
    
    Dim TestRange As Range
    Set TestRange = TestIndexedTable.Range("A3", "Foo")
    
    If RefRange <> TestRange Then
        Err.Description = ".Range() failed"
        GoTo TestFail
    End If
    
    If TestIndexedTable.TryGetRange("A3", "Foo", TestRange) = False Then
        Err.Description = "TryGetRange failed"
        GoTo TestFail
    End If
    If RefRange <> TestRange Then
        Err.Description = "TryGetRange got wrong range"
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
