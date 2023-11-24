# ListColumn Helpers
## Functions
```vb
TryGetListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String, ByRef OutListColumn As Public Function Exists(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
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
