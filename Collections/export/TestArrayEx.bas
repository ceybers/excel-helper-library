Attribute VB_Name = "TestArrayEx"
'@Folder("VBAProject")
Option Explicit

Public Sub DoTestArrayEx()
    Dim Result As Variant

    Dim arr As Variant
    arr = Array("Alpha", "Bravo", "Charlie")
    
    Debug.Assert ArrayEx.From(arr).Contains("Charlie") = True
    Debug.Assert ArrayEx.From(arr).Contains("NonexistingKey") = False
    Debug.Print "Contains() OK"
    
    Debug.Assert ArrayEx.From(arr).IndexOf("Charlie") = 2
    Debug.Assert ArrayEx.From(arr).IndexOf("NonexistingKey") = -1
    Debug.Print "IndexOf() OK"
    
    ' Arrays have no keys, and can have duplicates
    ' When do we return false for TryAdd?
    With ArrayEx.From(arr)
        Debug.Assert .TryAdd(vbNullString, "Charlie") = True
        Debug.Assert .TryAdd(vbNullString, "DeltaItem") = True
        Debug.Assert .Contains("DeltaItem") = True
    End With
    Debug.Print "TryAdd() OK"
    
    'Debug.Assert ArrayEx.From(arr).TryGetByKey(vbNullString, Result) = False
    'Debug.Assert ArrayEx.From(arr).TryGetByKey(vbNullString, Result) = True
    'Debug.Print "TryGetByKey() Cannot Impl"

    Debug.Assert ArrayEx.From(arr).TryRemove("Alpha") = True
    Debug.Assert ArrayEx.From(arr).TryRemove("NonexistingKey") = False
    Debug.Print "TryRemove() OK"
    
    'Debug.Assert ArrayEx.From(arr).GetByIndex(2) = "Charlie"
    'Debug.Assert ArrayEx.From(arr).TryInsert(2, "InsertTest")
    'Debug.Assert ArrayEx.From(arr).GetByIndex(2) = "InsertTest"
    'Debug.Print "GetByIndex() OK"
    'Debug.Print "TryInsert() OK"
    
    With ArrayEx.From(arr)
        Debug.Assert .Count = 3
        .Clear
        Debug.Assert .Count = 0
    End With
    Debug.Print "Clear() OK"
    Debug.Print "Count() OK"
    
    Debug.Print "Asserts passed."
    
    Dim outArray As Variant
    outArray = ArrayEx.From(arr).ToArray
    Debug.Assert TypeName(outArray) = "Variant()" 'Variant/Variant(0 to 2)
    ' each element is Variant/String
    Stop
End Sub



