Attribute VB_Name = "TestHashers"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Dim HasherSHA1 As IHasher
Dim HasherSHA256 As IHasher
Dim HasherMD5 As IHasher

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
    Set HasherSHA1 = New SHA1Hasher
    Set HasherSHA256 = New SHA256Hasher
    Set HasherMD5 = New MD5Hasher
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set HasherSHA1 = Nothing
    Set HasherSHA256 = Nothing
    Set HasherMD5 = Nothing
End Sub

'@TestMethod("TestHashers")
Private Sub TestHasherSHA1()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    If HasherSHA1.ComputeHash("foobar") <> "8843D7F92416211DE9EBB963FF4CE28125932878" Then
        Err.Description = "HasherSHA1.ComputeHash"
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

'@TestMethod("TestHashers")
Private Sub TestHasherSHA256()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    If HasherSHA256.ComputeHash("foobar") <> "C3AB8FF13720E8AD9047DD39466B3C8974E592C2FA383D4A3960714CAEF0C4F2" Then
        Err.Description = "HasherSHA256.ComputeHash"
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

'@TestMethod("TestHashers")
Private Sub TestHasherMD5()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    If HasherMD5.ComputeHash("foobar") <> "3858F62230AC3C915F300C664312C63F" Then
        Err.Description = "HasherMD5.ComputeHash"
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
