Attribute VB_Name = "TestDebugEx"
'@IgnoreModule ImplicitActiveSheetReference
'@Folder "Logging"
Option Explicit

Private Log As IDebugEx

'@EntryPoint "DoTest"
Public Sub DoTest()
    Set Log = DebugEx.Create()
    Log.AddProvider ImmediateLoggingProvider.Create
    Log.AddProvider FileLoggingProvider.Create
    Log.StartLogging
    
    Log.Message "Hello"
    WasteTime
    
    Dim aVariant As Variant
    aVariant = CDbl(1.23)
    Log.Variable aVariant
    WasteTime
    
    Dim aObject As Object
    Set aObject = ActiveWorkbook
    Log.Variable aObject
    WasteTime
    
    Dim anArray As Variant
    anArray = Array(1, 2, 3)
    Log.Variable anArray
    WasteTime
    
    Log.Variable Range("A1:B2")
    WasteTime
    
    Log.Variable Range("A1").Value2
    WasteTime
    
    aVariant = Range("A1").Value2
    Log.Variable aVariant
    WasteTime
    
    Log.Variable Range("A1:B2").Value2
    WasteTime
    
    Log.Variable Array(1, 2, 3)
    Log.Variable Array("A", "B", "C", "D")
    Log.LogHR
    
    Log.Message "This is a warning", LogLevel:=Warning_Level
    Log.Message "This is an error", LogLevel:=Error_level
    
    'd.LogStop "STOP HERE", LogLevel:=Warning_Level
    
    Log.Message "Setting the level to debug"
    Log.SetDefaultLevel Debug_Level
    Log.Message "We are now debugging"
    
    Log.Message "This one is verbose", LogLevel:=Verbose_Level
    Log.LogClear
    
    Log.Many Array("A", "B", "C", "D")
    
    Log.Many Range("A1:Z10").Value, "A Long Topic", Verbose_Level
    
    Log.Message "Goodbye"
    
    Log.StopLogging
End Sub

Private Sub WasteTime()
    Exit Sub
    
    Dim i As Long
    For i = 1 To 10
        Range("A1").Copy Range("A2")
    Next i
End Sub
