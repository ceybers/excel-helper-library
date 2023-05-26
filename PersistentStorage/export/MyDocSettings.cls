VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyDocSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PersistentStorage.MyDocSettings"
Option Explicit
Implements ISettings

Private Const DO_DEBUG As Boolean = False

Private Const SUBFOLDER As String = "Documents\Excel Macros"
Private Const Filename As String = "tablesplitcolumn.ini"
Private Const UUID As String = "{d5e9474a-1aa2-4bfb-aa89-ae0e69b6394f}"
Private Const FLAGS_HEADER As String = "[flags]"
Private Const TRUE_CONST As String = "TRUE"
Private Const FALSE_CONST As String = "FALSE"
Private Const DELIMITER_CHAR As String = "="
Private Const COMMENT_CHAR As String = ";"

Private Type TState
    Lines() As String
    Flags As Scripting.Dictionary
    AutoSave As Boolean
End Type
Private This As TState

Private Function GetFolder() As String
    GetFolder = Environ$("USERPROFILE") & "\" & SUBFOLDER
End Function

Private Function GetFullPath() As String
    GetFullPath = Environ$("USERPROFILE") & "\" & SUBFOLDER & "\" & Filename
End Function

Private Sub LoadFromFile()
    Open GetFullPath() For Binary As #1
    
    Dim FullData As String
    '@Ignore AssignmentNotUsed
    FullData = Space$(LOF(1))
    Get #1, , FullData
    Close #1
    
    This.Lines = Split(FullData, vbCrLf)
    
    If FullData = vbNullString Then
        SetDefaultFlags
    End If
End Sub

Private Sub SaveData()
    Open GetFullPath() For Output As #1
    Print #1, COMMENT_CHAR; " "; UUID
    Print #1, FLAGS_HEADER
    
    Dim FlagKey As Variant
    For Each FlagKey In This.Flags.Keys
        Print #1, FlagKey; DELIMITER_CHAR; This.Flags.Item(FlagKey)
    Next FlagKey
    Close #1
End Sub

Private Sub CheckIfEmpty()
    Dim FileSystemObject As Object
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    If Not FileSystemObject.FolderExists(GetFolder()) Then
        FileSystemObject.CreateFolder GetFolder()
    End If
    
    If Not FileSystemObject.FileExists(GetFullPath()) Then
        FileSystemObject.CreateTextFile GetFullPath()
        SetDefaultFlags
        SaveData
    End If
    
    Set FileSystemObject = Nothing
End Sub

Private Sub Class_Initialize()
    Set This.Flags = New Scripting.Dictionary
    This.AutoSave = True
    Load
End Sub

Private Function IsMySettings() As Boolean
    If UBound(This.Lines) = -1 Then Exit Function
    IsMySettings = (This.Lines(0) = (COMMENT_CHAR & " " & UUID))
End Function

Private Sub LoadFlags()
    This.Flags.RemoveAll
    
    Dim i As Long
    Dim ThisLine() As String
    
    If DO_DEBUG Then Debug.Print "Loaded "; UBound(This.Lines); " lines"
    
    Dim j As Long
    For j = 1 To UBound(This.Lines)
        If This.Lines(j) = FLAGS_HEADER Then Exit For
    Next j
    If j = UBound(This.Lines) Then Exit Sub
    
    For i = (j + 1) To UBound(This.Lines)
        Dim Delimiter As Long
        Delimiter = InStr(This.Lines(i), DELIMITER_CHAR)
        If Left$(Trim$(This.Lines(i)), 1) = COMMENT_CHAR Then Delimiter = 0
        If Left$(Trim$(This.Lines(i)), 1) = "[" Then Exit Sub
        
        If Delimiter > 0 Then
            Dim Key As String
            Dim Value As String
            Key = Trim$(Left$(This.Lines(i), Delimiter - 1))
            Value = Trim$(Mid$(This.Lines(i), Delimiter + 1))
            This.Flags.Add Key:=Key, Item:=Value
        End If
    Next i
End Sub

Private Sub SetDefaultFlags()
    Set This.Flags = New Scripting.Dictionary
End Sub

Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    If Not IsKeyValid(FlagName) Then Exit Function
    ISettings_GetFlag = (UCase$(This.Flags.Item(FlagName)) = TRUE_CONST)
    If DO_DEBUG Then Debug.Print "ISettings_GetFlag("; FlagName; ") = "; ISettings_GetFlag
End Function

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal FlagValue As Boolean)
    If Not IsKeyValid(FlagName) Then Exit Sub
    
    If DO_DEBUG Then Debug.Print "ISettings_SetFlag("; FlagName; ") = "; FlagValue
    
    If FlagValue = True Then
        This.Flags.Item(FlagName) = TRUE_CONST
    Else
        This.Flags.Item(FlagName) = FALSE_CONST
    End If
    
    If This.AutoSave Then SaveData
End Sub

Private Sub Load()
    If DO_DEBUG Then Debug.Print "MyDocSettings.Load()"
    CheckIfEmpty
    LoadFromFile
    If Not IsMySettings Then Exit Sub
    LoadFlags
End Sub

Private Sub Save()
    SaveData
End Sub

Private Sub ISettings_Reset()
    SetDefaultFlags
    SaveData
End Sub

Private Function ISettings_GetSetting(ByVal SettingName As String) As String
    Err.Raise Number:=42, Description:="NYI ISettings_GetSetting"
End Function

Private Sub ISettings_SetSetting(ByVal SettingName As String, ByVal Value As String)
    Err.Raise Number:=42, Description:="NYI ISettings_SetSetting"
End Sub

Private Function ISettings_GetCollection(ByVal CollectionName As String) As Collection
    Err.Raise Number:=42, Description:="NYI ISettings_GetCollection"
End Function

Private Sub ISettings_SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
    Err.Raise Number:=42, Description:="NYI ISettings_SetCollection"
End Sub

Private Function IsKeyValid(ByVal Key As String) As Boolean
    If Key = vbNullString Then Exit Function
    If InStr(Key, "=") > 0 Then Exit Function
    If InStr(Key, ";") > 0 Then Exit Function
    If InStr(Key, vbCr) > 0 Then Exit Function
    If InStr(Key, vbLf) > 0 Then Exit Function
    If Key <> Trim$(Key) Then Exit Function
    
    IsKeyValid = True
End Function
