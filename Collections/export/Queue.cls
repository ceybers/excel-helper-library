VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Helpers.CollectionEx")
Option Explicit
Implements IQueue

Private Const INITIAL_QUEUE_LENGTH As Long = 8

Private Type TState
    Arr As Variant
    Head As Long
    Tail As Long
End Type
Private This As TState

Private Sub Class_Initialize()
    This.Arr = Array()
    ReDim This.Arr(0 To (INITIAL_QUEUE_LENGTH - 1))
    This.Head = -1
    This.Tail = -1
End Sub

Private Sub IQueue_Enqueue(ByVal Value As Variant)
    Debug.Assert Not Me Is Queue
    This.Tail = This.Tail + 1
    If This.Tail > UBound(This.Arr) Then
        ReDim Preserve This.Arr(LBound(This.Arr) To (UBound(This.Arr) + INITIAL_QUEUE_LENGTH))
    End If
    This.Arr(This.Tail) = Value
End Sub

Private Function IQueue_Dequeue() As Variant
    Debug.Assert Not Me Is Queue
    If This.Head >= This.Tail Then Exit Function
    This.Head = This.Head + 1
    IQueue_Dequeue = This.Arr(This.Head)
     This.Arr(This.Head) = Empty
End Function

Private Function IQueue_IsEmpty() As Boolean
    Debug.Assert Not Me Is Queue
    IQueue_IsEmpty = (IQueue_Count = 0)
End Function

Private Function IQueue_Count() As Long
    Debug.Assert Not Me Is Queue
    IQueue_Count = This.Tail - This.Head
End Function

Private Sub IQueue_Clear()
    Debug.Assert Not Me Is Queue
    Dim i As Long
    For i = UBound(This.Arr) To LBound(This.Arr) Step -1
        This.Arr(i) = emtpy
    Next i
    This.Head = -1
    This.Tail = -1
End Sub
