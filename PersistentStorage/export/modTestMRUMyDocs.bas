Attribute VB_Name = "modTestMRUMyDocs"
'@Folder("PersistentStorage.Tests")
Option Explicit

Public Sub TestMRUMyDocs()
    Dim mru As IMostRecentlyUsed
    Set mru = New MostRecentlyUsed
     
    Dim s As ISettings
    Set s = MyDocSettings.Create("{20c36365-786c-455b-86b0-6a942560899d}", "persistentstoragetest.ini")
    
    mru.Clear
    mru.FromCollection s.GetCollection("bravo")
    
    PrintMRU mru
    
    mru.Add "zebra"
    
    PrintMRU mru
    
    s.SetCollection "bravo", mru.ToCollection
    
    Debug.Assert s.GetCollection("bravo").Item(1) = "zebra"
    
    Debug.Print "Tests done."
End Sub

Private Sub PrintMRU(ByVal mru As IMostRecentlyUsed)
    Debug.Print "MRU:"
    Dim i As Long
    For i = 0 To mru.Count - 1
        Debug.Print " "; i; ": "; mru.Item(i)
    Next i
    Debug.Print "---"
End Sub
