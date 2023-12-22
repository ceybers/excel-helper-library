Attribute VB_Name = "TryGetWorkbook"
'@Folder("CommonHelpers")
Option Explicit

'@Description "Tries to return the Wokbook with the given name from a Collection or an Application object."
Public Function TryGetWorkbookByName(ByVal Parent As Object, ByVal WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
    If TypeOf Parent Is Application Then
        TryGetWorkbookByName = TryGetWorkbookInCollectionByName(Parent.Workbooks, WorkbookName, OutWorkbook)
    ElseIf TypeOf Parent Is Collection Then
        TryGetWorkbookByName = TryGetWorkbookInCollectionByName(Parent, WorkbookName, OutWorkbook)
   End If
End Function

Private Function TryGetWorkbookInCollectionByName(ByVal Collection As Object, ByVal WorkbookName As String, ByRef OutWorkbook As Workbook) As Boolean
    Dim Workbook As Workbook
    For Each Workbook In Collection
        If Workbook.Name Like WorkbookName Then
            Set OutWorkbook = Workbook
            TryGetWorkbookInCollectionByName = True
            Exit Function
        End If
    Next Workbook
End Function
