# IndexedTable
- Wrapper to get/set cells from a ListObject by referencing a Key and a Field (ListColumn).

# Example
```vb
Dim TestIndexedTable As IndexedTable
Set TestIndexedTable = New IndexedTable
TestIndexedTable.Load ListObject, "Key Column"

' Do Test
Debug.Print TestIndexedTable.Item("KEY-123", "FieldFooBar") ' Returns "Original Value"
TestIndexedTable.Item("KEY-123", "FieldFooBar") = "Some Value"
Debug.Print TestIndexedTable("KEY-123", "FieldFooBar") ' Returns "Some Value"
```

# Properties
```vb
' Retrieve a cell from the table
Get Item(Key As Variant, Field As Variant) As Variant
' Update a cell in the table
Let Item(Key As Variant, Field As Variant, vNewValue As Variant)
```

# Method
```vb
' Initialize a new IndexedTable object
Load(ListObject As ListObject, KeyColumnName As String)

' Try and retrieve a cell. Returns FALSE if it cannot find the key or the field.
TryGetValue( Key As Variant, Field As Variant, ByRef OutValue As Variant) As Boolean

' Try and update a cell. Returns FALSE if it cannot find the key or the field.
TrySetValue(Key As Variant, Field As Variant, vNewValue As Variant) As Boolean
```

# Notes
- Property `.Item()` will throw an error if it cannot find the key or the field name. Use the `TryGetValue`/`TrySetValue` if you want the IndexedTable to handle the errors.
- Property `.Item()` is the default member for the class and will be accessed if no other property is specified.
  - e.g., `IndexedTable('foo', 'bar')` vs `IndexedTable.Item('foo', 'bar')`