Attribute VB_Name = "TestArrayEx"
'@IgnoreModule UseMeaningfulName
'@Folder "Helpers.CollectionEx.Tests"
Option Explicit

Public Sub DoTestArrayEx()
    'Dim Result As Variant

    Dim Arr As Variant
    Arr = Array("Alpha", "Bravo", "Charlie")
    
    Debug.Assert ArrayEx.From(Arr).Contains("Charlie") = True
    Debug.Assert ArrayEx.From(Arr).Contains("NonexistingKey") = False
    Debug.Print "Contains() OK"
    
    Debug.Assert ArrayEx.From(Arr).IndexOf("Charlie") = 2
    Debug.Assert ArrayEx.From(Arr).IndexOf("NonexistingKey") = -1
    Debug.Print "IndexOf() OK"
    
    ' Arrays have no keys, and can have duplicates
    ' When do we return false for TryAdd?
    With ArrayEx.From(Arr)
        Debug.Assert .TryAdd(vbNullString, "Charlie") = True
        Debug.Assert .TryAdd(vbNullString, "DeltaItem") = True
        Debug.Assert .Contains("DeltaItem") = True
    End With
    Debug.Print "TryAdd() OK"
    
    'Debug.Assert ArrayEx.From(arr).TryGetByKey(vbNullString, Result) = False
    'Debug.Assert ArrayEx.From(arr).TryGetByKey(vbNullString, Result) = True
    'Debug.Print "TryGetByKey() Cannot Impl"

    Debug.Assert ArrayEx.From(Arr).TryRemove("Alpha") = True
    Debug.Assert ArrayEx.From(Arr).TryRemove("NonexistingKey") = False
    Debug.Print "TryRemove() OK"
    
    'Debug.Assert ArrayEx.From(arr).GetByIndex(2) = "Charlie"
    'Debug.Assert ArrayEx.From(arr).TryInsert(2, "InsertTest")
    'Debug.Assert ArrayEx.From(arr).GetByIndex(2) = "InsertTest"
    'Debug.Print "GetByIndex() OK"
    'Debug.Print "TryInsert() OK"
    
    Dim outArray As Variant
    outArray = ArrayEx.From(Arr).ToArray
    Debug.Assert TypeName(outArray) = "Variant()"
    Debug.Print "ToArray() OK"
    
    Dim outCollection As Collection
    Set outCollection = ArrayEx.From(Arr).ToCollection
    Debug.Assert TypeName(outCollection) = "Collection"
    Debug.Assert outCollection.Count = 3
    Debug.Print "ToCollection() OK"
    
    Dim outDictionary As Scripting.Dictionary
    Set outDictionary = ArrayEx.From(Arr).ToDictionary
    Debug.Assert TypeName(outDictionary) = "Dictionary"
    Debug.Assert outDictionary.Count = 3
    Debug.Print "ToDictionary() OK"
    
    Dim testRng As Range
    Set testRng = ThisWorkbook.Worksheets.Item(1).Range("A1")
    testRng.Parent.UsedRange.Value2 = vbNullString
    ArrayEx.From(Arr).ToRange testRng
    Debug.Assert testRng.Cells.Item(1, 1).Value2 = Arr(0)
    Debug.Assert testRng.Cells.Item(2, 1).Value2 = Arr(1)
    Debug.Assert testRng.Cells.Item(3, 1).Value2 = Arr(2)
    testRng.Parent.UsedRange.Value2 = vbNullString
    Debug.Print "ToRange() OK"
    
    With ArrayEx.From(Arr)
        Debug.Assert .Count = 3
        .Clear
        Debug.Assert .Count = 0
    End With
    Debug.Print "Clear() OK"
    Debug.Print "Count() OK"
    
    Debug.Print "Asserts passed."
End Sub

