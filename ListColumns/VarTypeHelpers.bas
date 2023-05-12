Attribute VB_Name = "VarTypeHelpers"
'@Folder("Helpers")
Option Explicit

Private Const UNDEFINED_CONSTANT As String = "undefined"
Private Const ARRAY_SUFFIX As String = " (Array)"

'@Description "Returns the string description of a VarType result"
Public Function VarTypeToString(ByVal VarTypeValue As Long) As String
Attribute VarTypeToString.VB_Description = "Returns the string description of a VarType result"
    Dim Result As String
    Dim IsArray As Boolean
    Dim VarTypeConstants As Variant
    VarTypeConstants = Array("vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble", "vbCurrency", "vbDate", _
        "vbString", "vbObject", "vbError", "vbBoolean", "vbVariant", "vbDataObject", "vbDecvimal", "vbByte", _
        UNDEFINED_CONSTANT, UNDEFINED_CONSTANT, "vbLongLong")
    
    If VarTypeValue > vbArray Then
        IsArray = True
    End If
    
    Dim VarTypeValueLessArray As Long
    VarTypeValueLessArray = VarTypeValue - vbArray
    
    If VarTypeValueLessArray >= vbEmpty And VarTypeValueLessArray <= vbLongLong Then
        Result = VarTypeConstants(VarTypeValueLessArray)
    ElseIf VarTypeValueLessArray = vbUserDefinedType Then
        Result = "vbUserDefinedType"
    Else
        Result = UNDEFINED_CONSTANT
    End If
    
    If IsArray Then
        Result = Result & ARRAY_SUFFIX
    End If
    
    VarTypeToString = Result
End Function
