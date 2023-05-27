Attribute VB_Name = "modTestSettingsModel"
'@Folder("PersistentStorage.Tests")
Option Explicit

Public Sub DoTestSettingsModel()
    Dim UserSettings As ISettings
    Set UserSettings = MyDocSettings.Create( _
        UUID:="{20c36365-786c-455b-86b0-6a942560899d}", _
        Filename:="persistentstoragetest.ini")
        
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
        Workbook:=ThisWorkbook, _
        RootNode:="TestPersistentStorage")
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
        .AddUserSettings(UserSettings) _
        .AddWorkbookSettings(WorkbookSettings)
    
    With ASettingsModel.User
        Debug.Print .GetFlag("Foobar")
        .SetFlag "Foobar", True
        Debug.Print .GetFlag("Foobar")
        
        Debug.Print .GetSetting("SettingA")
        .SetSetting "SettingA", "Fiat lux"
        Debug.Print .GetSetting("SettingA")
    End With
    
    With ASettingsModel.Workbook
        Debug.Print .GetFlag("SourceWorkbook")
        .SetFlag "SourceWorkbook", True
        Debug.Print .GetFlag("SourceWorkbook")
        '.Reset
    End With
    
    With ASettingsModel.Table("Table1")
        Debug.Print .GetFlag("FavouriteTable")
        .SetFlag "FavouriteTable", True
        Debug.Print .GetFlag("FavouriteTable")
    End With
    
    Dim aXMl
    
    WorkbookSettings.DebugPrint
End Sub

