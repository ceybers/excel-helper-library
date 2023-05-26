Attribute VB_Name = "TestXMLSettings"
'@Folder("PersistentStorage")
Option Explicit

Public Sub TestSettingsModel()
    Dim sm As ISettingsModel
    Set sm = SettingsModel.Create(ThisWorkbook, "TableTransferTool")
    Debug.Print "sm.User.GetFlag('Foo') = "; sm.User.GetFlag("Foo")
    'sm.User.SetFlag "Foo", True
    Debug.Print "sm.User.GetFlag('Foo') = "; sm.User.GetFlag("Foo")
    Debug.Print "sm.User.GetFlag('Foo') = "; sm.User.GetFlag("Bar")
End Sub
