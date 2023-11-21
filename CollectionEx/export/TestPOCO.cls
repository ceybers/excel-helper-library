VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPOCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Helpers.CollectionEx.Tests"
Option Explicit

Private Type TState
    Name As String
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
End Property

Public Function Create(ByVal Name As String) As TestPOCO
    Set Create = New TestPOCO
    Create.Name = Name
End Function

Public Sub HandlePOCO(ByVal TestPOCO As TestPOCO)
    Debug.Print "HandlePOCO = "; TestPOCO.Name
End Sub

