Attribute VB_Name = "TestDictionaryEx"
'@Folder("VBAProject")
Option Explicit

Public Sub DoTestDictionaryEx()
    Dim Result As Variant
    
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    With dict
        .Add Item:="Alpha", Key:="Alpha"
        .Add Item:="Bravo", Key:="Bravo"
        .Add Item:="Charlie", Key:="Charlie"
    End With
    
    Debug.Assert DictionaryEx.From(dict).Contains("Charlie") = True
    Debug.Assert DictionaryEx.From(dict).Contains("NonexistingKey") = False
    Debug.Print "Contains() OK"
    
    Debug.Assert DictionaryEx.From(dict).IndexOf("Charlie") = 2
    Debug.Assert DictionaryEx.From(dict).IndexOf("NonexistingKey") = -1
    Debug.Print "IndexOf() OK"
       
    Debug.Assert DictionaryEx.From(dict).TryAdd("Charlie", "Charlie") = False
    Debug.Assert DictionaryEx.From(dict).TryAdd("DeltaKey", "DeltaItem") = True
    Debug.Assert DictionaryEx.From(dict).Contains("DeltaItem") = True
    Debug.Print "TryAdd() OK"
    
    Debug.Assert DictionaryEx.From(dict).TryGetByKey("NonexistingKey", Result) = False
    Debug.Assert DictionaryEx.From(dict).TryGetByKey("DeltaKey", Result) = True
    Debug.Print "TryGetByKey() OK"

    'Debug.Assert DictionaryEx.From(dict).TryRemove("Alpha") = True
    'Debug.Assert DictionaryEx.From(dict).TryRemove("NonexistingKey") = False
    Debug.Print "TryRemove() NYI"
    
    Debug.Assert DictionaryEx.From(dict).GetByIndex(2) = "Charlie"
    'Debug.Assert DictionaryEx.From(dict).TryInsert(2, "InsertTest")
    'Debug.Assert DictionaryEx.From(dict).GetByIndex(2) = "InsertTest"
    Debug.Print "GetByIndex() OK"
    Debug.Print "TryInsert() NYI"
    
    Dim outArray As Variant
    outArray = DictionaryEx.From(dict).ToArray
    Debug.Assert TypeName(outArray) = "Variant()"
    Debug.Assert LBound(outArray) = 0
    Debug.Assert UBound(outArray) = 3
    Debug.Print "ToArray() OK"
    
    Debug.Assert DictionaryEx.From(dict).Count = 4
    DictionaryEx.From(dict).Clear
    Debug.Assert DictionaryEx.From(dict).Count = 0
    Debug.Print "Clear() OK"
    Debug.Print "Count() OK"
    
    Debug.Print "Asserts passed."
End Sub

