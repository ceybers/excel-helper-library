# ListColumn Helpers
## Functions
```vb
' Returns True and passes a ListColumn by reference if it is found in the specified ListObject
'  Query of types Integer or Long returns the ListColumn by index in the ListObject
'  Query of type String tries to find a ListColumn with that exact name (case-sensitive)
'  Query of type Range will pass the intersecting ListColumn if the range is exactly one column wide
'  Query of type ListColumn will return true if the ListColumn exists in the ListObject
TryGetListColumnByVariant(ByVal ListObject As ListObject, ByVal Query As Variant, ByRef OutListColumn As ListColumn) As Boolean

' Returns a collection of ListColumns that are headers of the specified Range.
' The Keys of the collection are the ListColumn Names.
GetListColumnsFromRange(ByVal Range As Range) As Collection

' Returns True and passes a ListColumn by reference with the specified name if it exists in the ListObject.
TryGetListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String, ByRef OutListColumn As ListColumn) As Boolean
```

# ListColumn Analyzers
## Functions
```vb
GetR1C1(ByVal ListColumn As ListColumn) As String
ColumnHasBlanks(ByVal ListColumn As ListColumn) As Result
ColumnHasErrors(ByVal ListColumn As ListColumn) As Result
ColumnHasFormulae(ByVal ListColumn As ListColumn) As Result
ColumnHasValidation(ByVal ListColumn As ListColumn) As Result
ColumnIsLocked(ByVal ListColumn As ListColumn) As Result
EnumToString(ByVal EnumValue As Result) As String
ColumnIsUnique(ByVal ListColumn As ListColumn) As Result
GetVarTypeOfColumnRange(ByVal Range As Range) As Long
```

## ListColumnHelpers.Result (enum)
| ID | Description |
| -- | ----------- |
| 0  | None        |
| 1  | Some        |
| 2  | All         |
