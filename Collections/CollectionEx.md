# CollectionEx Interface
- Attempt at creating a generic interface for Collections, Dictionaries, and Arrays.
- Not sure what I want to actually name it yet.
- Need to decide how to handle situations where there are no Keys (Arrays), or we can't get a list of the Keys (Collections).
- Using the C# IList interface as inspiration
  - See also: https://learn.microsoft.com/en-us/dotnet/api/system.collections.ilist?view=net-7.0
- Implementing the C# style Try-pattern with out parameter

# CollectionEx methods (old, to clean up)
## From
```vb
Public Function From(ByVal Collection As Collection) As CollectionEx
```
Used to get a reference to a collection wrapped in CollectionEx.

## Exists
```vb
Public Function Exists(ByVal Value As Variant) As Boolean
```

Tests if a Value exists in a Collection as an Item. i.e., Not as a Key.

```vb
Debug.Print "Exists Delta   = "; CollectionEx.From(coll).Exists("Delta")
```

## ExistsByProperty
```vb
Public Function ExistsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
```

Tests each Item in a Collection, by comparing one of its properties versus a given Value.

```vb
' PocoColl contains items of Class 'TestPOCO' that have a Property named 'Name'
Debug.Print "ExistsByProperty Llama = "; CollectionEx.From(PocoColl).ExistsByProperty("Name", "Llama")
```

## IndexOf
```vb
Public Function IndexOf(ByVal Value As Variant) As Long
```

Gets the Index of an Item in a Collection.

## TryRemove
```vb
Public Function TryRemove(ByVal Value As Variant) As Boolean
```

Tries to remove an Item from a Collection based on the Item Value. The default Remove works by Index (and Key?). Returns true if it succeeds in finding and removing the Item.

## Clear
```vb
Public Sub Clear()
```

Loops through each Item in the Collection and removes it.

## TryAdd
```vb
Public Function TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
```

Tests if an Item with the same Key already exists. If not, adds a new Item with the given Key and value and returns True.

## TryGet
```vb
Public Function TryGet(ByVal Key As Variant, ByRef outValue As Variant) As Boolean
```

Tries to retrieve an Item using the given Key. If it succeeds, it returns True, and assigns the result to the given outValue variable.

## ForEach
```vb
Public Function ForEach(ByVal Object As Object, ByVal MethodName As Variant) As Boolean
```
Loops through each Item in the Collection. It compares the Type of the Item to the given Object type, and if it matches, it calls the given method, passing the Item as a parameter.

```vb
' TestPOCO is a class with a PreclaredId, therefore it has a default instance with the same name.
' It has a method that accepts an instantiation of the class as a parameter.
'    Public Sub HandlePOCO(ByVal TestPOCO As TestPOCO
CollectionEx.From(PocoColl).ForEach TestPOCO, "HandlePOCO"
```

## DebugPrint
```vb
Public Sub DebugPrint()
```
Loops through each Item in the Collection, printing out the Index and the Item into the Immediate Window.

---

Notes: Could create a similar method that accepts a predicate interface, which would let us filter the collection and perform the method conditionally.

