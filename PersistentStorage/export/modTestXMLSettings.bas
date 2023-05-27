Attribute VB_Name = "modTestXMLSettings"
'@Folder "PersistentStorage.Tests"
Option Explicit

Public Sub TestXMLSettings()
    Dim mXMLSettings As ISettings
    Set mXMLSettings = XMLSettingsFactory.CreateWorkbookSettings( _
        Workbook:=ThisWorkbook, _
        RootNode:="TestPersistentStorage")
        
    With mXMLSettings
        Debug.Print .GetFlag("Foobar")
        .SetFlag "Foobar", True
        Debug.Print .GetFlag("Foobar")
    End With
    
    Dim debugSetting As XMLSettings
    Set debugSetting = mXMLSettings
    debugSetting.DebugPrint
    
    Dim mTableSettings As ISettings
    Set mTableSettings = XMLSettingsFactory.CreateTableSettings( _
        WorkbookSettings:=mXMLSettings, _
        TableName:="Table1")
    
    With mTableSettings
        Debug.Print .GetFlag("Barfoo")
        .SetFlag "Barfoo", True
        Debug.Print .GetFlag("Barfoo")
    End With
    
    debugSetting.DebugPrint
    
    Stop
End Sub
