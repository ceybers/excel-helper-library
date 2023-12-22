Option Explicit

' Probably faster to use .SpecialCells
Public Function RangeIsEmpty(ByVal rng As Range) As Boolean
    Dim cell As Range
    For Each cell In rng
        If Not IsEmpty(cell.Value2) Then
            RangeIsEmpty = False
            Exit Function
        End If
    Next cell
    
    RangeIsEmpty = True
End Function

Public Function RangeHasValidation(ByVal rng As Range) As Boolean
    Dim valType As Long
    Dim cell As Range
    For Each cell In rng
        valType = -1
        
        On Error Resume Next
        valType = cell.Validation.Type
        On Error GoTo 0
        
        If valType <> -1 Then
            RangeHasValidation = True
            Exit Function
        End If
    Next cell
    
    RangeHasValidation = False
End Function

Public Function RangeHasFormatConditions(ByVal rng As Range) As Boolean
    RangeHasFormatConditions = rng.FormatConditions.Count > 0
End Function

Public Function GetListColumnsFromSelection(ByVal rng As Range) As Collection
    Set GetListColumnsFromSelection = New Collection
    
    If rng.ListObject Is Nothing Then Exit Function
    
    Dim lo As ListObject
    Set lo = Selection.ListObject

    Dim tryIntersect As Range
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        Set tryIntersect = Application.Intersect(lc.Range, rng)
        If Not tryIntersect Is Nothing Then
            GetListColumnsFromSelection.Add lc
        End If
    Next lc
    
End Function