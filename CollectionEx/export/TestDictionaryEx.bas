Attribute VB_Name = "TestDictionaryEx"
'@IgnoreModule UseMeaningfulName
'@Folder "Helpers.CollectionEx.Tests"
Option Explicit

Public Sub DoTestDictionaryEx()
    Dim Result As Variant
    
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    With dict
        .Add Item:="AlphaItem", Key:="Alpha"
        .Add Item:="BravoItem", Key:="Bravo"
        .Add Item:="CharlieItem", Key:="Charlie"
    End With
    
    Debug.Assert DictionaryEx.From(dict).Contains("Charlie") = True
    Debug.Assert DictionaryEx.From(dict).Contains("NonexistingKey") = False
    Debug.Print "Contains() OK"
    
    Debug.Assert DictionaryEx.From(dict).IndexOf("Charlie") = 2
    Debug.Assert DictionaryEx.From(dict).IndexOf("NonexistingKey") = -1
    Debug.Print "IndexOf() OK"
       
    Debug.Assert DictionaryEx.From(dict).TryAdd("Charlie", "Charlie") = False
    Debug.Assert DictionaryEx.From(dict).TryAdd("DeltaKey", "DeltaItem") = True
    Debug.Assert DictionaryEx.From(dict).Contains("DeltaKey") = True
    Debug.Print "TryAdd() OK"
    
    Debug.Assert DictionaryEx.From(dict).TryGetByKey("NonexistingKey", Result) = False
    Debug.Assert DictionaryEx.From(dict).TryGetByKey("DeltaKey", Result) = True
    Debug.Print "TryGetByKey() OK"

    'Debug.Assert DictionaryEx.From(dict).TryRemove("Alpha") = True
    'Debug.Assert DictionaryEx.From(dict).TryRemove("NonexistingKey") = False
    Debug.Print "TryRemove() NYI"
    
    Debug.Assert DictionaryEx.From(dict).GetByIndex(2) = "CharlieItem"
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
        
    Dim outCollection As Collection
    Set outCollection = DictionaryEx.From(dict).ToCollection
    Debug.Assert TypeName(outCollection) = "Collection"
    Debug.Assert outCollection.Count = 4
    Debug.Print "ToCollection() OK"
    
    Dim outDictionary As Scripting.Dictionary
    Set outDictionary = DictionaryEx.From(dict).ToDictionary
    Debug.Assert TypeName(outDictionary) = "Dictionary"
    Debug.Assert outDictionary.Count = 4
    Debug.Print "ToDictionary() OK"
    
    Dim testRng As Range
    Set testRng = ThisWorkbook.Worksheets.Item(1).Range("A1")
    testRng.Parent.UsedRange.Value2 = vbNullString
    DictionaryEx.From(dict).ToRange testRng
    Debug.Assert testRng.Cells.Item(1, 1).Value2 = dict.Keys(0)
    Debug.Assert testRng.Cells.Item(2, 1).Value2 = dict.Keys(1)
    Debug.Assert testRng.Cells.Item(3, 1).Value2 = dict.Keys(2)
    Debug.Assert testRng.Cells.Item(1, 2).Value2 = dict.Items(0)
    Debug.Assert testRng.Cells.Item(2, 2).Value2 = dict.Items(1)
    Debug.Assert testRng.Cells.Item(3, 2).Value2 = dict.Items(2)
    testRng.Parent.UsedRange.Value2 = vbNullString
    Debug.Print "ToRange() OK"
    
    Debug.Assert DictionaryEx.From(dict).Count = 4
    DictionaryEx.From(dict).Clear
    Debug.Assert DictionaryEx.From(dict).Count = 0
    Debug.Print "Clear() OK"
    Debug.Print "Count() OK"
    
    Debug.Print "Asserts passed."
End Sub

