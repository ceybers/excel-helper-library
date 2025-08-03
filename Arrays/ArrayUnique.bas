Attribute VB_Name = "ArrayUnique"
'@Folder("Helpers.Array")
Option Explicit

' Returns a 1-dimensional array that contains all the unique values in the
' InputArray. Not particularly performant. Lacks error handling for edge cases.
Public Function Unique(ByVal InputArray As Variant) As Variant
    If VarType(InputArray) < vbArray Then Exit Function
    
    Dim SortedArray As Variant
    SortedArray = InputArray
    ArraySort.QuickSort SortedArray
    
    Dim OutputArray As Variant
    ReDim OutputArray(LBound(SortedArray) To UBound(SortedArray))
    
    Dim c As Long
    c = LBound(SortedArray)
    
    OutputArray(c) = SortedArray(LBound(SortedArray))
    c = c + 1
    
    Dim i As Long
    For i = LBound(SortedArray) + 1 To UBound(SortedArray)
        If SortedArray(i) <> SortedArray(i - 1) Then
            OutputArray(c) = SortedArray(i)
            c = c + 1
        End If
    Next i
    
    ReDim Preserve OutputArray(LBound(OutputArray) To (c - 1))
    
    Unique = OutputArray
End Function
