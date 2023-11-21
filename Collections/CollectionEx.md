# `CollectionEx` Interface
- Attempt at creating a generic interface for Collections, Dictionaries, and Arrays.
- Not sure what I want to actually name it yet.
- Need to decide how to handle situations where there are no Keys (Arrays), or we can't get a list of the Keys (Collections).
- Using the C\# `IList` interface as inspiration
- Implementing the C\# style `Try-pattern` with `out parameter`.

# Creating CollectionEx classes
```vb
Dim SomeCollectionEx as ICollectionEx
Set SomeCollectionEx = ArrayEx.From(AnArray)
Set SomeCollectionEx = CollectionEx.From(ACollection)
Set SomeCollectionEx = DictionaryEx.From(ADictionary)
```
# `CollectionEx` methods
## Read-only operations
```vb
Count() As Long
Contains(ByVal Value As Variant) As Boolean
ContainsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
IndexOf(ByVal Value As Variant) As Long
```
## Modification operations 
```vb
Add(ByVal Key As Variant, ByVal Value As Variant)
Insert(ByVal Index As Long, ByVal Value As Variant)
GetByIndex(ByVal Index As Long) As Variant
GetByKey(ByVal Key As Variant) As Variant
Remove(ByVal Value As Variant)
RemoveAt(ByVal Index As Long)
```
## Modification operations (try)
```vb
TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
TryInsert(ByVal Index As Long, ByVal Value As Variant) As Boolean
TryGetByIndex(ByVal Index As Long, ByRef outValue As Variant) As Boolean
TryGetByKey(ByVal Key As Variant, ByRef outValue As Variant) As Boolean
TryRemove(ByVal Value As Variant) As Boolean
TryRemoveAt(ByVal Index As Long) As Boolean
```
## Operation operations
```vb
ForEach(ByVal Object As Object, ByVal MethodName As Variant)
Clear()
```
## Conversion operations
```vb
ToArray() As Variant
ToCollection() As Collection
ToDictionary() As Scripting.Dictionary
```

# Stacks, Queues, and Most Recently Used
## `Stack` methods
```vb
Sub Push(ByVal Value As Variant)
Function Pop() As Variant
Function TryPop(ByRef OutValue As Variant) As Boolean
Function Top() As Variant
Function IsEmpty() As Boolean
Function Count() As Long
Sub Clear()
```

## `Queue` methods
```vb
Sub Enqueue(ByVal Value As Variant)
Function Dequeue() As Variant
Function TryDequeue(ByRef OutValue As Variant) As Boolean
Function IsEmpty() As Boolean
Function Count() As Long
Sub Clear()
```

## `MostRecentlyUsed` methods
```vb
Sub SetMaximumLength(ByVal Length As Long)
Sub Add(ByVal Value As Variant)
Sub Remove(ByVal Value As Variant)
Sub RemoveAt(ByVal Index As Variant)
Function Item(ByVal Index As Long) As Variant
Function Count() As Long
Sub Clear()
Function ToCollection() As Collection
Sub FromCollection(ByVal Collection As Collection)
```

# See also
- [IList Interface (System.Collections) | Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/system.collections.ilist?view=net-7.0)
- [api - Building a better Collection. Enumerable in VBA - Code Review Stack Exchange](https://codereview.stackexchange.com/questions/60504/building-a-better-collection-enumerable-in-vba)
- [How to use the Implements in Excel VBA - Stack Overflow](https://stackoverflow.com/questions/19373081/how-to-use-the-implements-in-excel-vba/19379641#19379641)
- [Default Member Of A Class](http://www.cpearson.com/excel/DefaultMember.aspx)
- [[SOLVED] Attribute statements of VBA Classes](http://www.excelforum.com/excel-programming-vba-macros/562915-solved-attribute-statements-of-vba-classes.html)
- [How To: Collections in VBA in Excel and probably other MS Office 2003 applications. | PC Review](http://www.pcreview.co.uk/forums/collections-vba-excel-and-probably-other-ms-office-2003-applications-t2293368.html)