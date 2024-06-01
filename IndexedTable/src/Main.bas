Attribute VB_Name = "Main"
'@Folder("VBAProject")
Option Explicit

Public Sub AAA()
    Dim ListObject As ListObject
    Set ListObject = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    ' Test Start
    ListObject.DataBodyRange.Cells(2, 3) = "C3"
    
    Dim TestIndexedTable As IndexedTable
    Set TestIndexedTable = New IndexedTable
    TestIndexedTable.Load ListObject, "a"
    
    ' Do Test
    Debug.Print TestIndexedTable("A3", "c")
    TestIndexedTable("A3", "c") = "zzz"
    Debug.Print TestIndexedTable("A3", "c")
    
    ' Test Done
    ListObject.DataBodyRange.Cells(2, 3) = "C3"
End Sub
