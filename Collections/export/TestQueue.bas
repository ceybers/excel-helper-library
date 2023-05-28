Attribute VB_Name = "TestQueue"
'@Folder("Helpers.CollectionEx.Tests")
Option Explicit

Public Sub TestQueue()
    Debug.Print "Testing Queue class."
    
    Dim qq As Queue
    Dim q As IQueue
    Set q = New Queue
    Set qq = q
    
    Debug.Assert q.Count = 0
    
    Dim Element As Variant
    For Each Element In Array("Alpha", "Bravo", "Charlie")
        q.Enqueue Element
    Next Element
    
    Debug.Assert q.Count = 3
    Debug.Assert q.Dequeue = "Alpha"
    Debug.Assert q.Count = 2
    q.Enqueue "Delta"
    Debug.Assert q.Count = 3
    Debug.Assert q.Dequeue = "Bravo"
    Debug.Assert q.Count = 2
    Debug.Assert q.IsEmpty = False
    Debug.Assert q.Dequeue = "Charlie"
    Debug.Assert q.Dequeue = "Delta"
    Debug.Assert q.Dequeue = vbNullString
    Debug.Assert q.IsEmpty = True
    q.Enqueue "Foxtrot"
    Debug.Assert q.Count = 1
    
    q.Clear
    
    For Each Element In Array("Xray", "Yoyo", "Zebra")
        q.Enqueue Element
    Next Element
    
    Do While q.TryDequeue(Element)
        Debug.Print Element
    Loop
    Debug.Print "Asserts passed."
End Sub

