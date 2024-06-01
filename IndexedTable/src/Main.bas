Attribute VB_Name = "Main"
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess
'@Folder("VBAProject")
Option Explicit

'@EntryPoint
Public Sub RunIndexedTable()
    Dim ListObject As ListObject
    Set ListObject = ActiveSheet.ListObjects.Item(1)
    
    ' Create IndexedTable object
    Dim TestIndexedTable As IndexedTable
    Set TestIndexedTable = New IndexedTable
    TestIndexedTable.Load ListObject, "Key Column"
    
    Debug.Print "Value at 'A3' x 'Foo' is: "; TestIndexedTable("A3", "Foo")
    
    Debug.Print "Setting value at 'A3' x 'Foo' to 'zzz'"
    TestIndexedTable("A3", "Foo") = "zzz"
       
    Dim Result As Variant
    If Not TestIndexedTable.TryGetValue("A3", "Foo", Result) Then Exit Sub
    Debug.Print "Value at 'A3' x 'Foo' is now: "; Result
    
    Debug.Print "Change the background color of 'A3' x 'Foo' to green"
    TestIndexedTable.Range("A3", "Foo").Interior.Color = vbGreen
End Sub
