VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder "Helpers.CollectionEx.Abstract"
'@Interface

Public Sub Push(ByVal Value As Variant)
End Sub

Public Function TryPop(ByRef OutValue As Variant) As Boolean
End Function

Public Function Pop() As Variant
End Function

Public Function Top() As Variant
End Function

Public Function IsEmpty() As Boolean
End Function

Public Function Count() As Long
End Function

Public Sub Clear()
End Sub
