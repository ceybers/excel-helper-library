# Queue class
Implementation of a First In First Out collection.
# Methods
## Constructor
```vb
Dim ExampleQueue as Queue
Set ExampleQueue = New Queue
```
## Query operations
```vb
Function IsEmpty() As Boolean
Function Count() As Long
```
## Modification operations
```vb
Sub Enqueue(Value As Variant)
Function Dequeue() As Variant
Function TryDequeue(OutValue As Variant) As Boolean
Sub Clear()
```
# Examples
```vb
Dim ExampleQueue as Queue
Set ExampleQueue = New Queue
ExampleQueue.Enqueue("foo")
ExampleQueue.Enqueue("bar")
Debug.Print ExampleQueue.Dequeue ' "foo"
Debug.Assert ExampleQueue.Count = 1
Debug.Assert ExampleQueue.Dequeue = "bar"
```