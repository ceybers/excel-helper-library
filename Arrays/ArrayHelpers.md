# Array Helpers
## ArrayToFilteredRange (method)
```vb
Public Sub ArrayToFilteredRange(ByVal rng As Range, ByVal arr As Variant)
```
Copies the visible cells in a disjoint (filtered) column into a Variant array. Must be exactly one column wide.
## VisibleRangeToArray (function)
```vb
Function VisibleRangeToArray(ByVal rng As Range) As Variant
```
Does the same thing as `ArrayToFilteredRange`?

## Analysis Functions
These were for matching key columns in `Table Transfer Tool` and `Inline Regex Tool`, as well as Column Quality tests similar to Power BI.

### ArrayAnalyseOne (function)
```vb
Public Function ArrayAnalyseOne(arr As Variant) As ArrayExAnalyseOne
```

### ArrayAnalyseTwo (function)
```vb
Public Function ArrayAnalyseTwo(lhs As Variant, rhs As Variant) As ArrayExAnalyseTwo
```

### ArrayErrorCount (function)
```vb
Public Function ArrayErrorCount(arr As Variant) As Integer
```

### ArrayBlankCount (function)
```vb
Public Function ArrayBlankCount(arr As Variant) As Integer
```

### ArrayFilterTextOnly (function)
```vb
Public Function ArrayFilterTextOnly(arr As Variant) As Variant
```

### ArrayUnique (function)
```vb
Public Function ArrayUnique(arr As Variant) As Variant
```
Returns a unique copy of the array.
Only items that appear exactly once are included.
Duplicates (including first instance), blanks and errors are excluded.

### ArrayDistinct (function)
```vb
Public Function ArrayDistinct(arr As Variant) As Variant
```
Returns a distinct copy of the array.
In the case of duplicate values, only one instance is returned.
Blanks and errors are excluded.

### ArrayIntersect (function)
```vb
Public Function ArrayIntersect(lhs As Variant, rhs As Variant) As Variant
```
Returns a list of items in both lhs and rhs.
Excludes blanks and errors.
Only checks 1st instance of each duplicate.

### ArrayLength (function)
```vb
Public Function ArrayLength(arr As Variant) As Integer
```
Returns the length of the first dimension of an array
i.e. the number of rows in an array that was created from a single column (nx1) range.

### ArrayTrim (function)
```vb
Public Function ArrayTrim(arr As Variant, Length As Integer) As Variant
```
Returns a copy of the array, retaining only the first n items.
If the length is longer than the provided array,
the provided array is returned (length is ignored)
If the length is zero or negative, the provided array is returned.

### ArrayAntiJoinLeft (function)
```vb
Public Function ArrayAntiJoinLeft(lhs As Variant, rhs As Variant) As Variant
```
Returns an array(n, 1) of all the items that are in the lhs array but not in the rhs array. Excludes blanks and errors.

### ArraySubset (function)
```vb
Public Function ArraySubset(lhs As Variant, rhs As Variant) As Boolean
```
Check if every item in lhs exists in rhs.
Does not check if rhs items all exist in lhs.
Ignores blanks and errors.

## ArrayMatch (function)
```vb 
'Example:
Public Function ArrayMatch(lhs As Variant, rhs As Variant) As Boolean
```
Checks if all items exist in both lhs and rhs

### ArrayFind (function)
```vb
Public Function ArrayFind(Match As Variant, arr As Variant) As Integer
```
Checks if the provided variant exists in the array.
- -1 means no match
- -2 means a blank (string "") was provided
- -3 means an error was provided
