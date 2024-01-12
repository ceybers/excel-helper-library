Attribute VB_Name = "TestRegularExpressions"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

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

'@TestMethod("RegEx")
Private Sub TestRegularExpressionsA()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    
    If Regex("foobar", "foo") <> "foo" Then
        Err.Description = "match failed"
        GoTo TestFail
    End If
    
    If Regex("ar", "foo") <> False Then
        Err.Description = "non-match failed"
        GoTo TestFail
    End If
    
    If Regex("foobar", "^foobar$") <> "foobar" Then
        Err.Description = "^$ match failed"
        GoTo TestFail
    End If
    
    If Regex("foo", "^foobar$") <> False Then
        Err.Description = "^$ non-match failed"
        GoTo TestFail
    End If
        
    If Regex("foobar", "^foo$") <> False Then
        Err.Description = "^$ non-match failed"
        GoTo TestFail
    End If
    
    If Regex("foo123bar", "([a-z]+)([0-9]+)([a-z]+)", "$2") <> "123" Then
        GoTo TestFail
    End If
    
    If Regex("foo123bar", "([a-z]+)([0-9]+)([a-z]+)", "$2$3") <> "123bar" Then
        GoTo TestFail
    End If
    
    If Regex("foo123bar", "([a-z]+)([0-9]+)([a-z]+)", "$2test$3") <> "123testbar" Then
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
