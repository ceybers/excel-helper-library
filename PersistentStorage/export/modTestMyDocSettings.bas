Attribute VB_Name = "modTestMyDocSettings"
'@Folder "PersistentStorage.Tests"
Option Explicit

Public Sub TestMyDocs()
    Debug.Print "TestMyDocs..."
    
    Dim s As ISettings
    Set s = MyDocSettings.Create("{20c36365-786c-455b-86b0-6a942560899d}", "persistentstoragetest.ini")
    
    'Debug.Print s.GetFlag("Foo") = True
    s.SetFlag "Foo", True
    Debug.Assert s.GetFlag("Foo") = True
    
    Debug.Assert s.GetSetting("SettingA") = "Foobar"
    s.SetSetting "SettingA", "Foobar"
    Debug.Assert s.GetSetting("SettingA") = "Foobar"
    
    Dim coll As Collection
    Set coll = s.GetCollection("alpha")
    Debug.Assert coll.Count = 3
    Debug.Assert coll.Item(1) = "alpha"
    Debug.Assert coll.Item(2) = "bravo"
    Debug.Assert coll.Item(3) = "charlie"
    ' GetCollection working
    
    Set coll = New Collection
    coll.Add "xray"
    coll.Add "yoyo"
    coll.Add "zebra"
    s.SetCollection "bravo", coll
    Set coll = Nothing
    Set coll = s.GetCollection("bravo")
    Debug.Assert coll.Count = 3
    Debug.Assert coll.Item(1) = "xray"
    Debug.Assert coll.Item(2) = "yoyo"
    Debug.Assert coll.Item(3) = "zebra"
    
    Debug.Print "Asserts passed."
End Sub
