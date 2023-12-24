# `CollectionEx` Interface
- Attempt at creating a generic interface for Collections, Dictionaries, and Arrays.
- Not sure what I want to actually name it yet.
- Need to decide how to handle situations where there are no Keys (Arrays), or we can't get a list of the Keys (Collections).
- Using the C\# `IList` interface as inspiration
- Implementing the C\# style `Try-pattern` with `out parameter`.
# Methods
## Constructors
```vb
Dim SomeCollectionEx as ICollectionEx
Set SomeCollectionEx = ArrayEx.From(AnArray)
Set SomeCollectionEx = CollectionEx.From(ACollection)
Set SomeCollectionEx = DictionaryEx.From(ADictionary)
```
## Read-only operations
```vb
Count() As Long
Contains(Value As Variant) As Boolean
ContainsByProperty(PropertyName As String, Value As Variant) As Boolean
IndexOf(Value As Variant) As Long
```
## Query operations
```vb
GetByIndex(Index As Long) As Variant
GetByKey(Key As Variant) As Variant
TryGetByIndex(Index As Long, outValue As Variant) As Boolean
TryGetByKey(Key As Variant, outValue As Variant) As Boolean
```
## Modification operations 
```vb
Add(Key As Variant, Value As Variant)
Insert(Index As Long, Value As Variant)
Remove(Value As Variant)
RemoveAt(Index As Long)
Clear()
```
## Modification operations (try)
```vb
TryAdd(Key As Variant, Value As Variant) As Boolean
TryInsert(Index As Long, Value As Variant) As Boolean
TryRemove(Value As Variant) As Boolean
TryRemoveAt(Index As Long) As Boolean
```
## Enumerator operations
```vb
ForEach(Object As Object, MethodName As Variant)
```
## Conversion operations
```vb
ToArray() As Variant
ToCollection() As Collection
ToDictionary() As Scripting.Dictionary
ToRange(Range as Range)
```
# Examples
## Convert Collection to Dictionary
```vb
Dim Dictionary as Scripting.Dictionary
Set Dictionary = CollectionEx.From(ACollection).ToDictionary()
```
# See also
- [IList Interface (System.Collections) | Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/system.collections.ilist?view=net-7.0)
- [api - Building a better Collection. Enumerable in VBA - Code Review Stack Exchange](https://codereview.stackexchange.com/questions/60504/building-a-better-collection-enumerable-in-vba)
- [How to use the Implements in Excel VBA - Stack Overflow](https://stackoverflow.com/questions/19373081/how-to-use-the-implements-in-excel-vba/19379641#19379641)
- [Default Member Of A Class](http://www.cpearson.com/excel/DefaultMember.aspx)
- [[SOLVED] Attribute statements of VBA Classes](http://www.excelforum.com/excel-programming-vba-macros/562915-solved-attribute-statements-of-vba-classes.html)
- [How To: Collections in VBA in Excel and probably other MS Office 2003 applications. | PC Review](http://www.pcreview.co.uk/forums/collections-vba-excel-and-probably-other-ms-office-2003-applications-t2293368.html)