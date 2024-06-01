# IndexedTable
- Wrapper to get/set cells from a ListObject by referencing a Key and a Field (ListColumn).

# Example
```vb
Dim TestIndexedTable As IndexedTable
Set TestIndexedTable = New IndexedTable
TestIndexedTable.Load ListObject, "Key Column"

If Not(TestIndexed.IsValid) Then Exit Sub
' Exits if the key column could not be set.

Debug.Print TestIndexedTable.Item("KEY-123", "FieldFooBar") 
' Returns "Original Value"

TestIndexedTable.Item("KEY-123", "FieldFooBar") = "Some Value"
' Sets the value in the table immediately.

Debug.Print TestIndexedTable("KEY-123", "FieldFooBar") 
' Returns "Some Value"

TestIndexedTable.Range("KEY-123", "FieldFooBar").Interior.Color = RGB(0,128,255)
' Sets the background color of the referenced cell to light blue.
```

# Properties
```vb
' Retrieve a cell from the table
Get Item(Key As Variant, Field As Variant) As Variant

' Update a cell in the table
Let Item(Key As Variant, Field As Variant, vNewValue As Variant)

' Get Range object of the cell referred to by Key × Field
Get Range(Key As Variant, Field As Variant) As Range

' Returns TRUE if the table contains the given Key
Get HasKey(Key As Variant)

' Returns TRUE if the IndexedTable object is valid (ListObject loaded, KeyColumn found)
Get IsValid() As Boolean
```

# Methods
```vb
' Initialize a new IndexedTable object
Load(ListObject As ListObject, KeyColumnName As String)

' Try and retrieve a cell. Returns FALSE if it cannot find the key or the field.
TryGetValue( Key As Variant, Field As Variant, ByRef OutValue As Variant) As Boolean

' Try and update a cell. Returns FALSE if it cannot find the key or the field.
TrySetValue(Key As Variant, Field As Variant, vNewValue As Variant) As Boolean

' Try and get the Range object of the cell referred to by Key × Field.
TryGetRange(Key As Variant, Field As Variant, ByRef OutRange As Range) As Boolean

' Returns TRUE if the Worksheet containing the table is protected.
IsProtected()
```

# Notes
- In the case of duplicate keys, the first instance will be referred to.
- Keys are stored as Variants, including non-text and errors. 
  - Blank cells (including single apostrophe) are stored as `Empty` of type `Variant/Empty`.
- Property `.Item()` will throw an error if it cannot find the key or the field name. Use the `TryGetValue`/`TrySetValue` if you want the `IndexedTable` class to handle the errors.
- Property `.Item()` is the default member for the class and will be accessed if no other property is specified.
  - e.g., `IndexedTable('foo', 'bar')` vs `IndexedTable.Item('foo', 'bar')`
- Changes to the source table are committed immediately. 
  - Changes are done on a cell-by-cell basis and are not particularly performant.
- `TrySetValue()` will gracefully handle protected Worksheets. 
  - It does not take into account cases where the Worksheet is protected but the specific cell being written to is Unlocked.