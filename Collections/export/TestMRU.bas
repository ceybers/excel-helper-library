Attribute VB_Name = "TestMRU"
'@Folder("Helpers.CollectionEx.Tests")
Option Explicit

Public Sub TestMRU()
    Debug.Print "Testing MRU class."
    
    Dim mmru As MostRecentlyUsed
    
    Dim mru As IMostRecentlyUsed
    Set mru = New MostRecentlyUsed
    Set mmru = mru
    
    Debug.Assert mru.Count = 0
    mru.Add "Alpha"
    Debug.Assert mru.Count = 1
    mru.Add "Bravo"
    Debug.Assert mru.Count = 2
    mru.Add "Charlie"
    mru.Add "Delta"
    Debug.Assert mru.Count = 4
    mru.Add "Echo"
    Debug.Assert mru.Count = 4
    ' Add (new) and Count works
    
    Debug.Assert mru.Item(0) = "Echo"
    mru.Add "Delta"
    Debug.Assert mru.Item(0) = "Delta"
    ' Add (existing) works
    
    Debug.Assert mru.Item(0) = "Delta"
    Debug.Assert mru.Item(1) = "Echo"
    mru.RemoveAt 1
    Debug.Assert mru.Item(0) = "Delta"
    Debug.Assert mru.Item(1) = "Charlie"
    mru.Remove "Charlie"
    Debug.Assert mru.Item(0) = "Delta"
    Debug.Assert mru.Item(1) = "Bravo"
    ' Remove works

    Dim coll As Collection
    Set coll = mru.ToCollection
    Debug.Assert coll.Count = 2
    Debug.Assert coll.Item(1) = "Delta"
    Debug.Assert coll.Item(2) = "Bravo"
    ' ToCollection works

    mru.Clear
    Debug.Assert mru.Count = 0
    ' Clear works
    
    Set coll = New Collection
    With coll
        .Add "Zebra"
        .Add "Xray"
        .Add "Yoyo"
        .Add "Zebra"
    End With
    
    mru.FromCollection coll
    
    Debug.Assert mru.Count = 3
    Debug.Assert mru.Item(0) = "Zebra"
    Debug.Assert mru.Item(1) = "Xray"
    Debug.Assert mru.Item(2) = "Yoyo"
    ' FromCollection works
    
    Debug.Print "Asserts passed."
End Sub

