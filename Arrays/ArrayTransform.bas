Attribute VB_Name = "ArrayTransform"
'@IgnoreModule ProcedureNotUsed
'@Folder("Helpers.Array")
Option Explicit

' Transforms a 2-dimensional array of shape (many, one) into a 1-dimensional array.
' Useful for .Value2 of a Range that is 1-column wide.
'    (1 to m, 1 to 1) to (1 to m)
Public Function ColumnToArray(ByVal InputArray As Variant) As Variant
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputArray)
    
    Dim Result As Variant
    ReDim Result(1 To UBound(InputArray, 1))
    
    Dim i As Long
    For i = 1 To UBound(InputArray, 1)
        Result(i) = InputArray(i, 1)
    Next i
    
    ColumnToArray = Result
End Function

' Transforms a 2-dimensional array of shape (one, many) into a 1-dimensional array.
' Useful for .Value2 of a Range that is 1-row tall.
'    (1 to 1, 1 to m) to (1 to m)
Public Function RowToArray(ByVal InputArray As Variant) As Variant
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputArray)
    
    Dim Result As Variant
    ReDim Result(1 To UBound(InputArray, 2))
    
    Dim i As Long
    For i = 1 To UBound(InputArray, 2)
        Result(i) = InputArray(1, i)
    Next i
    
    RowToArray = Result
End Function

' Transforms a 1-dimensional array into a 2-dimensional array of of shape (one, many).
' Useful for preparing an array to be placed into the .Value2 of a Range that is 1-column wide.
'    (1 to m) to (1 to m, 1 to 1)
Public Function ArrayToColumn(ByVal InputArray As Variant) As Variant
    Dim Lower As Long
    Lower = LBound(InputArray)
    
    Dim Count As Long
    Count = UBound(InputArray) - Lower + 1

    Dim Result As Variant
    ReDim Result(1 To Count, 1 To 1)
    
    Dim i As Long
    For i = 1 To Count
        Result(i, 1) = InputArray(i + Lower - 1)
    Next i
    
    ArrayToColumn = Result
End Function

' Transforms a 1-dimensional array into a 2-dimensional array of of shape (many, one).
' Useful for preparing an array to be placed into the .Value2 of a Range that is 1-row tall.
'    (1 to m) to (1 to 1, 1 to m)
Public Function ArrayToRow(ByVal InputArray As Variant) As Variant
    Dim Lower As Long
    Lower = LBound(InputArray)
    
    Dim Count As Long
    Count = UBound(InputArray) - Lower + 1

    Dim Result As Variant
    ReDim Result(1 To 1, 1 To Count)
    
    Dim i As Long
    For i = 1 To Count
        Result(1, i) = InputArray(i + Lower - 1)
    Next i
    
    ArrayToRow = Result
End Function

' Transforms a 1-dimensional zero-based array into a one-based array.
'    (0 to n) to (1 to n+1)
Public Function ZeroBasedToOne(ByVal InputArray As Variant) As Variant
    Dim Lower As Long
    Lower = LBound(InputArray)
    
    Dim Count As Long
    Count = UBound(InputArray) - Lower + 1

    Dim Result As Variant
    ReDim Result(1 To Count)
    
    Dim i As Long
    For i = 1 To Count
        Result(i) = InputArray(i - 1)
    Next i
    
    ZeroBasedToOne = Result
End Function

' Transforms a 1-dimensional one-based array into a zero-based array.
'    (1 to n) to (0 to n-1)
Public Function OneBasedToZero(ByVal InputArray As Variant) As Variant
    Dim Lower As Long
    Lower = LBound(InputArray)
    
    Dim Count As Long
    Count = UBound(InputArray) - Lower + 1

    Dim Result As Variant
    ReDim Result(0 To Count - 1)
    
    Dim i As Long
    For i = 0 To Count - 1
        Result(i) = InputArray(i + 1)
    Next i
    
    OneBasedToZero = Result
End Function

' Transpose a 2-dimensional array.
Public Function Transpose(ByVal InputArray As Variant) As Variant
    Debug.Assert ArrayCheck.IsTwoDimensionalOneBasedArray(InputArray)
    
    Dim Result As Variant
    ReDim Result(LBound(InputArray, 2) To UBound(InputArray, 2), LBound(InputArray, 1) To UBound(InputArray, 1))
    
    Dim RowIndex As Long
    For RowIndex = LBound(InputArray, 2) To UBound(InputArray, 2)
        Dim ColIndex As Long
        For ColIndex = LBound(InputArray, 1) To UBound(InputArray, 1)
            Result(RowIndex, ColIndex) = InputArray(ColIndex, RowIndex)
        Next ColIndex
    Next RowIndex
    
    Transpose = Result
End Function

' Returns a 2-dimensional array of size (1 to 1, 1 to 1) when the input is a single variant.
' If input is already an array, returns the same array.
Public Function ForceTwoDimensional(ByVal InputVariant As Variant) As Variant
    If VarType(InputVariant) >= vbArray Then
        ForceTwoDimensional = InputVariant
        Exit Function
    End If
    
    Dim Result As Variant
    ReDim Result(1 To 1, 1 To 1)
    Result(1, 1) = InputVariant
    ForceTwoDimensional = Result
End Function

' Returns a collection containing the items in a 1-dimensional array. Key is not set.
Public Function ToCollection(ByVal InputArray As Variant) As Collection
    Dim Result As Collection
    Set Result = New Collection
    Dim i As Long
    For i = LBound(InputArray) To UBound(InputArray)
        Result.Add Item:=InputArray(i)
    Next i
    Set ToCollection = Result
End Function
