Attribute VB_Name = "ArrayFlatten"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.Array"
Option Explicit

' Input Variant(1 to n, 1 to 1), Output new (Variant 1 to n)
Public Function FlattenArray(ByVal InputArray As Variant) As Variant
    Dim RowCount As Long
    RowCount = UBound(InputArray) - LBound(InputArray) + 1
    
    Dim Result() As Variant
    ReDim Result(1 To RowCount)
    
    Dim i As Long
    For i = 1 To RowCount
        Result(i) = InputArray(i, 1)
    Next i
    
    FlattenArray = Result
End Function
