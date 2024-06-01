VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndexedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("VBAProject")
Option Explicit

Private Const ERR_SRC As String = "IndexedTable"
Private Const ERR_NUM_INVALID As Long = 1
Private Const ERR_NUM_NO_KEY As Long = 2
Private Const ERR_NUM_NO_FLD As Long = 3
Private Const ERR_NUM_TBL_PROT As Long = 4
Private Const ERR_MSG_INVALID  As String = "IndexedTable not loaded with a valid ListObject and KeyColumnName."
Private Const ERR_MSG_NO_KEY As String = "Key not found in IndexedTable."
Private Const ERR_MSG_NO_FLD As String = "Field not found in IndexedTable."
Private Const ERR_MSG_TBL_PROT As String = "Table is protected and cannot be written to."

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
    GuardObjectIsValid
    Item = GetValue(Key, Field)
End Property

Public Property Let Item(ByVal Key As Variant, ByVal Field As Variant, ByVal vNewValue As Variant)
    GuardObjectIsValid
    SetValue Key, Field, vNewValue
End Property

Public Property Get Range(ByVal Key As Variant, ByVal Field As Variant) As Variant
    GuardObjectIsValid
    Set Range = GetRange(Key, Field)
End Property

Public Property Get HasKey(ByVal Key As Variant) As Boolean
    GuardObjectIsValid
    HasKey = This.Keys.Exists(Key)
End Property

Public Property Get IsValid() As Boolean
    IsValid = This.IsValid
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
    If Not This.Fields.Exists(This.KeyColumnName) Then Exit Function
    
    Set This.Keys = New Scripting.Dictionary
    
    Dim Keys As Variant
    Keys = This.ListObject.ListColumns.Item(This.KeyColumnName).DataBodyRange.Value2
    Dim i As Long
    For i = 1 To UBound(Keys, 1)
        If Not This.Keys.Exists(Keys(i, 1)) Then
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
    GuardKeyIsValid Key
    GuardFieldIsValid Field
    
    GetValue = This.Cache(This.Keys.Item(Key), This.Fields.Item(Field))
End Function

Public Function TryGetValue(ByVal Key As Variant, ByVal Field As Variant, ByRef OutValue As Variant) As Boolean
    If Not This.Keys.Exists(Key) Then Exit Function
    If Not This.Fields.Exists(Field) Then Exit Function
    
    OutValue = This.Cache(This.Keys.Item(Key), This.Fields.Item(Field))
    TryGetValue = True
End Function

Private Sub SetValue(ByVal Key As Variant, ByVal Field As Variant, ByVal vNewValue As Variant)
    GuardKeyIsValid Key
    GuardFieldIsValid Field
    GuardListObjectIsProtected
    
    This.ListObject.DataBodyRange.Cells.Item(This.Keys.Item(Key), This.Fields.Item(Field)).Value2 = vNewValue
    This.Cache(This.Keys.Item(Key), This.Fields.Item(Field)) = vNewValue
End Sub

Public Function TrySetValue(ByVal Key As Variant, ByVal Field As Variant, ByVal vNewValue As Variant) As Boolean
    If Not This.Keys.Exists(Key) Then Exit Function
    If Not This.Fields.Exists(Field) Then Exit Function
    If IsProtected Then Exit Function
    
    This.ListObject.DataBodyRange.Cells.Item(This.Keys.Item(Key), This.Fields.Item(Field)).Value2 = vNewValue
    This.Cache(This.Keys.Item(Key), This.Fields.Item(Field)) = vNewValue
    TrySetValue = True
End Function

Private Function GetRange(ByVal Key As Variant, ByVal Field As Variant) As Range
    GuardKeyIsValid Key
    GuardFieldIsValid Field
    
    Set GetRange = This.ListObject.DataBodyRange.Cells.Item(This.Keys.Item(Key), This.Fields.Item(Field))
End Function

Public Function TryGetRange(ByVal Key As Variant, ByVal Field As Variant, ByRef OutRange As Range) As Boolean
    If Not This.Keys.Exists(Key) Then Exit Function
    If Not This.Fields.Exists(Field) Then Exit Function
    
    Set OutRange = This.ListObject.DataBodyRange.Cells.Item(This.Keys.Item(Key), This.Fields.Item(Field))
    TryGetRange = True
End Function

Private Sub GuardObjectIsValid()
    If This.IsValid = False Then Err.Raise vbObjectError + ERR_NUM_INVALID, ERR_SRC, ERR_MSG_INVALID
End Sub

Private Sub GuardKeyIsValid(ByVal Key As Variant)
    If Not This.Keys.Exists(Key) Then Err.Raise vbObjectError + ERR_NUM_NO_KEY, ERR_SRC, ERR_MSG_NO_KEY
End Sub

Private Sub GuardFieldIsValid(ByVal Field As Variant)
    If Not This.Fields.Exists(Field) Then Err.Raise vbObjectError + ERR_NUM_NO_FLD, ERR_SRC, ERR_MSG_NO_FLD
End Sub

Public Function IsProtected() As Boolean
    IsProtected = This.ListObject.Parent.ProtectContents
End Function

Private Sub GuardListObjectIsProtected()
    If IsProtected Then Err.Raise vbObjectError + ERR_NUM_TBL_PROT, ERR_SRC, ERR_MSG_TBL_PROT
End Sub
