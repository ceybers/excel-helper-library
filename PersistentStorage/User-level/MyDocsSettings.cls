Option Explicit
Implements ISettings

Private Const SUBFOLDER As String = "Documents\Excel Macros"
Private Const Filename As String = "tablesplitcolumn.ini"
Private Const UUID As String = "{3ea2eff7-2549-4cf3-b8c4-2c3c7a3451f5}"
Private Const FLAGS_HEADER As String = "FLAGS"

Private Type TState
    Lines() As String
    Flags As Scripting.Dictionary
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
    Print #1, UUID
    Print #1, FLAGS_HEADER
    
    Dim FlagKey As Variant
    For Each FlagKey In This.Flags.Keys
    Print #1, vbTab; FlagKey; vbTab; This.Flags.Item(FlagKey)
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
End Sub

Private Function IsMySettings() As Boolean
    If UBound(This.Lines) = -1 Then Exit Function
    IsMySettings = (This.Lines(0) = UUID)
End Function

Private Sub LoadFlags()
    Dim i As Long
    Dim ThisLine() As String
    Dim Extracting As Boolean
    For i = 1 To UBound(This.Lines)
        ThisLine = Split(This.Lines(i), vbTab)
        If Extracting Then
            If UBound(ThisLine) = 2 Then
                This.Flags.Add Key:=ThisLine(1), Item:=ThisLine(2)
            Else
                Extracting = False
            End If
        Else
            If ThisLine(0) = FLAGS_HEADER Then Extracting = True
        End If
    Next i
End Sub

Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    ISettings_GetFlag = UCase$(This.Flags.Item(FlagName)) = "TRUE"
End Function

Private Sub SetDefaultFlags()
    Set This.Flags = New Scripting.Dictionary
    This.Flags.Add Key:=SHOW_HIDDEN_FLAG, Item:="FALSE"
    This.Flags.Add Key:=SHOW_UNSUITABLE_FLAG, Item:="FALSE"
    This.Flags.Add Key:=REPLACE_EXISTING_FLAG, Item:="TRUE"
    This.Flags.Add Key:=REPLACE_OTHER_FLAG, Item:="FALSE"
End Sub

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal FlagValue As Boolean)
    If FlagValue = True Then
        This.Flags.Item(FlagName) = "TRUE"
    Else
        This.Flags.Item(FlagName) = "FALSE"
    End If
End Sub

Private Sub ISettings_Load()
    CheckIfEmpty
    LoadFromFile
    If Not IsMySettings Then Exit Sub
    LoadFlags
End Sub

Private Sub ISettings_Save()
    SaveData
End Sub

Private Sub ISettings_Reset()
    SetDefaultFlags
    SaveData
End Sub
