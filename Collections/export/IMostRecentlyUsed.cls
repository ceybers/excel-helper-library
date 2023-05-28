VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMostRecentlyUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("Helpers.CollectionEx")
'@Interface

Public Sub SetMaximumLength(ByVal Length As Long)
End Sub

Public Sub Add(ByVal Value As Variant)
End Sub

Public Sub Remove(ByVal Value As Variant)
End Sub

Public Sub RemoveAt(ByVal Index As Variant)
End Sub

Public Function Item(ByVal Index As Long) As Variant
End Function

Public Function Count() As Long
End Function

Public Sub Clear()
End Sub

