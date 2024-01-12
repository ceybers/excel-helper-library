Attribute VB_Name = "RegularExpressions"
'@Folder("VBAProject")
Option Explicit

Private Const TOKEN_OUTPUT_PATTERN As String = "\$(\d+)"
Private Const DEFAULT_OUTPUT_PATTERN As String = "$0"

Public Function Regex(ByVal Value As Variant, ByVal Expression As Variant, _
    Optional ByVal OutputPattern As Variant = DEFAULT_OUTPUT_PATTERN) As Variant
    Static InputObject As Object
    If InputObject Is Nothing Then Set InputObject = GetRegExpObject(True, False, True)
    InputObject.Pattern = Expression
    
    Dim InputMatches As Object
    Set InputMatches = InputObject.Execute(Value)
    
    If InputMatches.Count = 0 Then
        Regex = False
        Exit Function
    End If
    
    If OutputPattern = DEFAULT_OUTPUT_PATTERN Then
        Regex = InputMatches.Item(0).Value
        Exit Function
    End If
    
    Static OutputObject As Object
    If OutputObject Is Nothing Then Set OutputObject = GetRegExpObject(True, True, False, TOKEN_OUTPUT_PATTERN)
    
    Static ReplaceObject As Object
    If ReplaceObject Is Nothing Then Set ReplaceObject = GetRegExpObject(True, True, False)
    
    Dim Result As String
    Result = OutputPattern
    
    Dim ReplaceMatch As Object
    For Each ReplaceMatch In OutputObject.Execute(OutputPattern)
        Dim TokenIndex As Long
        TokenIndex = CLng(ReplaceMatch.SubMatches.Item(0))
        ReplaceObject.Pattern = "\$" & TokenIndex
        
        If TokenIndex = 0 Then
            Result = ReplaceObject.Replace(Result, InputMatches.Item(0).Value)
        ElseIf TokenIndex <= InputMatches.Item(0).SubMatches.Count Then
            Result = ReplaceObject.Replace(Result, InputMatches.Item(0).SubMatches(TokenIndex - 1))
        Else
            Regex = CVErr(xlErrValue)
            Exit Function
        End If
    Next ReplaceMatch

    Regex = Result
End Function

Private Function GetRegExpObject(ByVal GlobalFlag As Boolean, ByVal MultiLine As Boolean, _
    ByVal IgnoreCase As Boolean, Optional ByVal Pattern As String) As Object
    Set GetRegExpObject = CreateObject("VBScript.RegExp")
    With GetRegExpObject
        .Global = GlobalFlag
        .MultiLine = MultiLine
        .IgnoreCase = IgnoreCase
        If Pattern <> vbNullString Then .Pattern = Pattern
    End With
End Function

