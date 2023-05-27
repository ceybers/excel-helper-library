Attribute VB_Name = "modTestMyDocSettings"
'@Folder "PersistentStorage.Tests"
Option Explicit

Public Sub TestMyDocs()
    Dim s As ISettings
    Set s = MyDocSettings.Create("{20c36365-786c-455b-86b0-6a942560899d}", "persistentstoragetest.ini")
    
    Debug.Print s.GetFlag("Foo")
    s.SetFlag "Foo", True
    
    Debug.Print s.GetSetting("SettingA")
    s.SetSetting "SettingA", "Foobar"
    Debug.Print s.GetSetting("SettingA")
End Sub
