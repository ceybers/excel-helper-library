# Range Helpers

## ResizeRangeToArray
Returns a range with the same shape as the specified 2-dimensional InputArray, starting from the top-most cell in the specified InputRange.
```vb
ResizeRangeToArray(ByVal InputRange As Range, ByVal InputArray As Variant) As Range
```

## RangeSetValueFromVariant
Updates the .Value2 property of all the cells in the InputRange with the Variant Values in the specified 2-dimensional InputArray.
```vb 
RangeSetValueFromVariant(ByVal InputRange As Range, ByVal InputVariant As Variant)
```

## RangeBox
Returns a range with the offset and size of the specified input parameters, starting from the top-most cell in the InputRange. Row = 1 and Column = 1 start the box from the top-left cell. 
e.g., RangeBox(Range("A1"), 1, 2, 4, 8).Address = B1:I4
```vb
RangeBox(ByVal InputRange As Range, _
    ByVal Row As Long, ByVal Column As Long, _
    ByVal Rows As Long, ByVal Columns As Long) As Range
```

## Functions
```vb
RangeIsEmpty(ByVal rng As Range) As Boolean
RangeHasValidation(ByVal rng As Range) As Boolean
RangeHasFormatConditions(ByVal rng As Range) As Boolean
GetListColumnsFromSelection(ByVal rng As Range) As Collection
TryGetHiddenCellsInRange(ByVal InputRange As Range, ByRef OutputRange As Range) As Boolean
```

## Methods
```vb
AppendRange(ByVal rangeToAppend As Range, ByRef unionRange As Range)
```