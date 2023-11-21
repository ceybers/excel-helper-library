VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Helpers.CollectionEx")
Option Explicit
Implements IStack

Private Const INITIAL_STACK_LENGTH As Long = 8

Private Type TState
    Arr As Variant
    Cursor As Long
End Type
Private This As TState

Private Sub Class_Initialize()
    This.Arr = Array()
    ReDim This.Arr(0 To (INITIAL_STACK_LENGTH - 1))
    This.Cursor = -1
End Sub

Private Sub IStack_Push(ByVal Value As Variant)
    Debug.Assert Not Me Is Stack
    This.Cursor = This.Cursor + 1
    If This.Cursor > UBound(This.Arr) Then
        ReDim Preserve This.Arr(LBound(This.Arr) To (UBound(This.Arr) + INITIAL_STACK_LENGTH))
    End If
    This.Arr(This.Cursor) = Value
End Sub

Private Function IStack_TryPop(ByRef OutValue As Variant) As Boolean
    Debug.Assert Not Me Is Stack
    If This.Cursor < 0 Then Exit Function
    OutValue = IStack_Pop
    IStack_TryPop = True
End Function

Private Function IStack_Pop() As Variant
    Debug.Assert Not Me Is Stack
    If This.Cursor < 0 Then Exit Function
    IStack_Pop = This.Arr(This.Cursor)
    This.Arr(This.Cursor) = Empty
    This.Cursor = This.Cursor - 1
End Function

Private Function IStack_Top() As Variant
    If This.Cursor < 0 Then Exit Function
    IStack_Top = This.Arr(This.Cursor)
End Function

Private Function IStack_IsEmpty() As Boolean
    Debug.Assert Not Me Is Stack
    IStack_IsEmpty = (IStack_Count = 0)
End Function

Private Function IStack_Count() As Long
    Debug.Assert Not Me Is Stack
    IStack_Count = This.Cursor + 1
End Function

Private Sub IStack_Clear()
    Debug.Assert Not Me Is Stack
    Dim i As Long
    For i = This.Cursor To 0 Step -1
        This.Arr(i) = Empty
    Next i
    This.Cursor = -1
End Sub
