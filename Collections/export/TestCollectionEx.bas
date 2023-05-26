Attribute VB_Name = "TestCollectionEx"
'@Folder("VBAProject")
Option Explicit

Public Sub DoTestCollectionEx()
    Dim Result As Variant
    
    Dim coll As Collection
    Set coll = New Collection
    
    With coll
        .Add Item:="Alpha", Key:="Alpha"
        .Add Item:="Bravo", Key:="Bravo"
        .Add Item:="Charlie", Key:="Charlie"
    End With
    
    Debug.Assert CollectionEx.From(coll).Contains("Charlie") = True
    Debug.Assert CollectionEx.From(coll).Contains("NonexistingKey") = False
    Debug.Print "Contains() OK"
    
    Debug.Assert CollectionEx.From(coll).IndexOf("Charlie") = 3
    Debug.Assert CollectionEx.From(coll).IndexOf("NonexistingKey") = -1
    Debug.Print "IndexOf() OK"
       
    Debug.Assert CollectionEx.From(coll).TryAdd("Charlie", "Charlie") = False
    Debug.Assert CollectionEx.From(coll).TryAdd("DeltaKey", "DeltaItem") = True
    Debug.Assert CollectionEx.From(coll).Contains("DeltaItem") = True
    Debug.Print "TryAdd() OK"
    
    Debug.Assert CollectionEx.From(coll).TryGetByKey("NonexistingKey", Result) = False
    Debug.Assert CollectionEx.From(coll).TryGetByKey("DeltaKey", Result) = True
    Debug.Print "TryGetByKey() OK"

    Debug.Assert CollectionEx.From(coll).TryRemove("Alpha") = True
    Debug.Assert CollectionEx.From(coll).TryRemove("NonexistingKey") = False
    Debug.Print "TryRemove() OK"
    
    Debug.Assert CollectionEx.From(coll).GetByIndex(2) = "Charlie"
    Debug.Assert CollectionEx.From(coll).TryInsert(2, "InsertTest")
    Debug.Assert CollectionEx.From(coll).GetByIndex(2) = "InsertTest"
    Debug.Print "GetByIndex() OK"
    Debug.Print "TryInsert() OK"
    
    Dim outArray As Variant
    outArray = CollectionEx.From(coll).ToArray
    Debug.Assert TypeName(outArray) = "Variant()"
    Debug.Assert LBound(outArray) = 0
    Debug.Assert UBound(outArray) = 3
    Debug.Print "ToArray() OK"
    
    Debug.Assert CollectionEx.From(coll).Count = 4
    CollectionEx.From(coll).Clear
    Debug.Assert CollectionEx.From(coll).Count = 0
    Debug.Print "Clear() OK"
    Debug.Print "Count() OK"
    
    Debug.Print "Asserts passed."
End Sub


