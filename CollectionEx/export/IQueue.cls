VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("Helpers.CollectionEx")
'@Interface

Public Sub Enqueue(ByVal Value As Variant)
End Sub

Public Function TryDequeue(ByRef OutValue As Variant) As Boolean
End Function

Public Function Dequeue() As Variant
End Function

Public Function IsEmpty() As Boolean
End Function

Public Function Count() As Long
End Function

Public Sub Clear()
End Sub
