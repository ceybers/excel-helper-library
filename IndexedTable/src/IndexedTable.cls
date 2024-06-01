VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndexedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private Const ERR_SRC As String = "IndexedTable"
Private Const ERR_NUM_INVALID As Long = 1
Private Const ERR_NUM_NO_KEY As Long = 2
Private Const ERR_NUM_NO_FLD As Long = 3
Private Const ERR_MSG_INVALID  As String = "No IndexedTable set"
Private Const ERR_MSG_NO_KEY As String = "Key not found in IndexedTable"
Private Const ERR_MSG_NO_FLD As String = "Field not found in IndexedTable"

Private Type TState
    ListObject As ListObject
    KeyColumnName As String
    IsValid As Boolean
    Keys As Scripting.Dictionary
    Fields As Scripting.Dictionary
    Cache As Variant
End Type
Private This As TState

'@DefaultMember
Public Property Get Item(ByVal Key As Variant, ByVal Field As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If This.IsValid = False Then Err.Raise vbObjectError + ERR_NUM_INVALID, ERR_SRC, ERR_MSG_INVALID
    Item = GetValue(Key, Field)
End Property

Public Property Let Item(ByVal Key As Variant, ByVal Field As Variant, ByVal vNewValue As Variant)
    If This.IsValid = False Then Err.Raise vbObjectError + ERR_NUM_INVALID, ERR_SRC, ERR_MSG_INVALID
    SetValue Key, Field, vNewValue
End Property

Public Sub Load(ByVal ListObject As ListObject, ByVal KeyColumnName As String)
    Set This.ListObject = ListObject
    This.KeyColumnName = KeyColumnName
    
    This.Cache = This.ListObject.DataBodyRange.Value2
    TrySetFields
    If Not TrySetKeys Then Exit Sub
    
    This.IsValid = True
End Sub

Private Function TrySetKeys() As Boolean
    If Not (This.Fields.Exists(This.KeyColumnName)) Then Exit Function
    Set This.Keys = New Scripting.Dictionary
    
    Dim Keys As Variant
    Keys = This.ListObject.ListColumns(This.KeyColumnName).DataBodyRange.Value2
    Dim i As Long
    For i = 1 To UBound(Keys, 1)
        If Not (This.Keys.Exists(Keys(i, 1))) Then
            This.Keys.Add Keys(i, 1), i
        End If
    Next i
    
    TrySetKeys = True
End Function

Private Sub TrySetFields()
    Set This.Fields = New Scripting.Dictionary
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        This.Fields.Add ListColumn.Name, ListColumn.Index
    Next ListColumn
End Sub

Private Function GetValue(ByVal Key As Variant, ByVal Field As Variant) As Variant
    If Not (This.Keys.Exists(Key)) Then Err.Raise vbObjectError + ERR_NUM_NO_KEY, ERR_SRC, ERR_MSG_NO_KEY
    If Not (This.Fields.Exists(Field)) Then Err.Raise vbObjectError + ERR_NUM_NO_FLD, ERR_SRC, ERR_MSG_NO_FLD
    
    GetValue = This.Cache(This.Keys.Item(Key), This.Fields.Item(Field))
End Function

Public Function TryGetValue(ByVal Key As Variant, ByVal Field As Variant, ByRef OutValue As Variant) As Boolean
    If Not (This.Keys.Exists(Key)) Then Exit Function
    If Not (This.Fields.Exists(Field)) Then Exit Function
    
    OutValue = This.Cache(This.Keys.Item(Key), This.Fields.Item(Field))
    TryGetValue = True
End Function

Private Function SetValue(ByVal Key As Variant, ByVal Field As Variant, ByVal vNewValue As Variant)
    If Not (This.Keys.Exists(Key)) Then Err.Raise vbObjectError + ERR_NUM_NO_KEY, ERR_SRC, ERR_NUM_NO_KEY
    If Not (This.Fields.Exists(Field)) Then Err.Raise vbObjectError + ERR_NUM_NO_FLD, ERR_SRC, ERR_MSG_NO_FLD
    
    This.ListObject.DataBodyRange.Cells(This.Keys.Item(Key), This.Fields.Item(Field)).Value2 = vNewValue
    This.Cache(This.Keys.Item(Key), This.Fields.Item(Field)) = vNewValue
End Function

Public Function TrySetValue(ByVal Key As Variant, ByVal Field As Variant, ByVal vNewValue As Variant) As Boolean
    If Not (This.Keys.Exists(Key)) Then Exit Function
    If Not (This.Fields.Exists(Field)) Then Exit Function
    
    This.ListObject.DataBodyRange.Cells(This.Keys.Item(Key), This.Fields.Item(Field)).Value2 = vNewValue
    This.Cache(This.Keys.Item(Key), This.Fields.Item(Field)) = vNewValue
    TrySetValue = True
End Function
