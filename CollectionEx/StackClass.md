# Stack class
Implementation of a First In Last Out collection.
# Methods
## Constructor
```vb
Dim ExampleStack as Stack
Set ExampleStack = New Stack
```
## Query operations
```vb
Function IsEmpty() As Boolean
Function Count() As Long
Function Top() As Variant
```
## Modification operations
```vb
Sub Push(Value As Variant)
Function Pop() As Variant
Function TryPop(OutValue As Variant) As Boolean
Sub Clear()
```
# Examples
```vb
Dim ExampleStack as Stack
Set ExampleStack = New Stack
ExampleStack.Push("foo")
ExampleStack.Push("bar")
Debug.Print ExampleStack.Pop ' "bar"
Debug.Assert ExampleStack.Count = 1
Debug.Assert ExampleStack.Pop = "foo"
Debug.Assert ExampleStack.IsEmpty = True
```