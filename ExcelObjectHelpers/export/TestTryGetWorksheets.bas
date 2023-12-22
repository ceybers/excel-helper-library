Attribute VB_Name = "TestTryGetWorksheets"
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
Private Sub TestTryGetWorksheets()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Result As Collection
    
    'Act:
    Dim Collection As Collection
    Set Collection = New Collection
    Collection.Add ThisWorkbook.Worksheets.Item(1)
    Collection.Add ThisWorkbook.Worksheets.Item(2)
    
    If Not TryGetWorksheetsByName(Collection, "Sheet1", Result) Then
        If Result.Count <> 2 Then
            Err.Description = "Get Worksheet in Collection positive fail"
            GoTo TestFail
        End If
    End If
    
    If TryGetWorksheetsByName(Collection, "Sheet9", Result) Then
        Err.Description = "Get Worksheet in Collection negative fail"
        GoTo TestFail
    End If
    
    If Not TryGetWorksheetsByName(Application, "Sheet1", Result) Then
        If Result.Count <> 2 Then
            Err.Description = "Get Worksheet in Application positive fail"
            GoTo TestFail
        End If
    End If
    
    If TryGetWorksheetsByName(Application, "Sheet9", Result) Then
        Err.Description = "Get Worksheet in Application negative fail"
        GoTo TestFail
    End If
    
    Dim Dictionary As Object
    Set Dictionary = GetDictionaryOfWorksheets(Application)
    If Dictionary.Count <> 5 Then
        Err.Description = "GetDictionaryOfWorksheets fail"
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

