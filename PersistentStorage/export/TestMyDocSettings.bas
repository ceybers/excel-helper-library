Attribute VB_Name = "TestMyDocSettings"
'@Folder("PersistentStorage")
Option Explicit

Public Sub TestMyDocs()
    Dim s As ISettings
    Set s = New MyDocSettings
    
    Debug.Print s.GetFlag("Foo")
    s.SetFlag "Foo", True
    'Debug.Print s.GetFlag("Foo")
    
    s.SetFlag "Bar", False
End Sub
