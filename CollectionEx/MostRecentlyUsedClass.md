# `MostRecentlyUsed` class
# Methods
## Constructor
```vb
Dim MRU as MostRecentlyUsed
Set MRU = New MostRecentlyUsed
```
## Query operations
```vb
Function Item(Index As Long) As Variant
Function Count() As Long
```
## Modification operations
```vb
Sub SetMaximumLength(Length As Long)
Sub Add(Value As Variant)
Sub Remove(Value As Variant)
Sub RemoveAt(Index As Variant)
Sub Clear()
```
## Conversion operations
```vb
Function ToCollection() As Collection
Sub FromCollection(Collection As Collection)
```
# Examples
```vb
Dim MRU as MostRecentlyUsed
Set MRU = New MostRecentlyUsed
MRU.Add("foobar")
Debug.Assert MRU.Count = 1
Debug.Print MRU.Item(1)
MRU.Remove("foobar")
Debug.Assert MRU.Count = 0
```