Attribute VB_Name = "TestStack"
'@Folder("Helpers.CollectionEx.Tests")
Option Explicit

Public Sub TestStack()
    Debug.Print "Testing Stack class."
    
    Dim s As IStack
    Set s = New Stack
    
    Debug.Assert s.Top = Empty
    
    Dim Element As Variant
    For Each Element In Array("Alpha", "Bravo", "Charlie")
        s.Push Element
    Next Element
    
    Debug.Assert s.Count = 3
    Debug.Assert s.Top = "Charlie"
    Debug.Assert s.Pop = "Charlie"
    Debug.Assert s.Top = "Bravo"
    Debug.Assert s.Count = 2
    Debug.Assert s.IsEmpty = False
    s.Clear
    Debug.Assert s.Count = 0
    Debug.Assert s.IsEmpty = True
    
    Debug.Print "Asserts passed."
End Sub
