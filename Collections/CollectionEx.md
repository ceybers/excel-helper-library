# CollectionEx Interface
- Attempt at creating a generic interface for Collections, Dictionaries, and Arrays.
- Not sure what I want to actually name it yet.
- Need to decide how to handle situations where there are no Keys (Arrays), or we can't get a list of the Keys (Collections).
- Using the C\# `IList` interface as inspiration
  - See also: https://learn.microsoft.com/en-us/dotnet/api/system.collections.ilist?view=net-7.0
- Implementing the C\# style `Try-pattern` with `out parameter`.

# CollectionEx methods
- Read-only operations
  - Count() As Long
  - Contains(ByVal Value As Variant) As Boolean
  - ContainsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
  - IndexOf(ByVal Value As Variant) As Long
- Modification operations 
  - Add(ByVal Key As Variant, ByVal Value As Variant)
  - Insert(ByVal Index As Long, ByVal Value As Variant)
  - GetByIndex(ByVal Index As Long) As Variant
  - GetByKey(ByVal Key As Variant) As Variant
  - Remove(ByVal Value As Variant)
  - RemoveAt(ByVal Index As Long)
- Modification operations (try)
  - TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
  - TryInsert(ByVal Index As Long, ByVal Value As Variant) As Boolean
  - TryGetByIndex(ByVal Index As Long, ByRef outValue As Variant) As Boolean
  - TryGetByKey(ByVal Key As Variant, ByRef outValue As Variant) As Boolean
  - TryRemove(ByVal Value As Variant) As Boolean
  - TryRemoveAt(ByVal Index As Long) As Boolean
- Operation operations
  - ForEach(ByVal Object As Object, ByVal MethodName As Variant)
  - Clear()
- Conversion operations
  - ToArray() As Variant
  - ToCollection() As Collection
  - ToDictionary() As Scripting.Dictionary

