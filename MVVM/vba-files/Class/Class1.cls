VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Helpers.Collection")
Option Explicit

Public Sub DoPrint(ByVal Object)
    Debug.Print "zzz "; CStr(Object); " yyy"
End Sub