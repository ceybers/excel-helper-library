# TableHash
## Functions
### FromListColumn
Returns the hash of the cells in the DataBodyRange of the ListColumn.
Ignores the name of the column header.
### FromHeaderNames
Returns the hash of the headers of a table.
By default the hash changes if the headers are rearranged.
If the flag IgnoreOrder is set to True, the order of the headers is ignored.
### FromListRow
Returns the hash of the cells in the Range of the ListRow.
Adding or removing columns will change the hash, even if the cells affected are blank.
### FromListObject
Returns the hash of all the cells in the DataBodyRange of the ListObject.
Changing the name of a header will not change the hash.
Rearranging the data into a different shape (rows and columns) will change the hash.
By default, rearranging columns and rows (while retaining the same shape) in the table will change the hash.
If the flag IgnoreRowSortOrder is set to True, then changing the order of rows (e.g., sorting) will not change the hash.
If the flag IgnoreRearrangedHeaders is set to True, then changing the order (index) of the headers will not change the hash. However, changing the name of a header will change the hash.
## User Defined Function
`TableHash(Range, IgnoreOrder)`
If the Range does not overlap with a Table, an error will be returned.
If the Range is one row high and one row wide, it will return the hash for the Table's entire DataBodyRange.
If the range is one row high and overlaps with the header row, it will return the hash for the Header names.
If the range is one row high and overlaps with the DataBodyRange, it will return the hash for that entire row.
If the range is one column wide, it will return the hash for the DataBodyRange in that entire column. 
If the range is more than one row high or one column wide, it will return the fash for the Table's entire DataBodyRange. Using IgnoreOrder = True in this mode will ignore both the row and column order.

## Examples
```vb
ComputeHash(ByVal Value As Variant) As String
FromListColumn(ByVal ListColumn As ListColumn) As String
FromListRow(ByVal ListRow As ListRow) As String
FromListObject(ByVal ListObject As ListObject, _
    Optional ByVal IgnoreRowSortOrder As Boolean = False, _
    Optional ByVal IgnoreRearrangedHeaders As Boolean = False) As String
Public Function FromHeaderNames(ByVal ListObject As ListObject, _
    Optional ByVal IgnoreOrder As Boolean = False) As String
```