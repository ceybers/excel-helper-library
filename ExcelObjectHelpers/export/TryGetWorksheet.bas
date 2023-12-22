Attribute VB_Name = "TryGetWorksheet"
'@Folder "CommonHelpers"
Option Explicit

'@Description "Tries to return the first Worksheet with the given name from a Workbook, Collection, or an Application object (Workbooks)."
Public Function TryGetWorksheetByName(ByVal Parent As Object, ByVal WorksheetName As String, ByRef OutWorksheet As Worksheet) As Boolean
Attribute TryGetWorksheetByName.VB_Description = "Tries to return the Worksheet with the given name from a Workbook, Collection, or an Application object (Workbooks)."
    If TypeOf Parent Is Workbook Then
        TryGetWorksheetByName = TryGetWorksheetInWorkbookByName(Parent, WorksheetName, OutWorksheet)
    ElseIf TypeOf Parent Is Collection Then
        TryGetWorksheetByName = TryGetWorksheetInCollectionByName(Parent, WorksheetName, OutWorksheet)
    ElseIf TypeOf Parent Is Application Then
        TryGetWorksheetByName = TryGetWorksheetInCollectionByName(GetAllWorksheetsInApplication(Parent), WorksheetName, OutWorksheet)
    End If
End Function

'@Description "Returns a Dictionary of all the Worksheets in an Application object. Keys are colon delimited strings of Workbook and Worksheet name."
Public Function GetDictionaryOfWorksheets(ByVal Application As Application) As Object
    Dim Result As Object
    Set Result = CreateObject("Scripting.Dictionary")
    
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    
    For Each Workbook In Application.Workbooks
        For Each Worksheet In Workbook.Worksheets
            Result.Add Key:=Workbook.Name & ":" & Worksheet.Name, Item:=Worksheet
        Next Worksheet
    Next Workbook
    
    Set GetDictionaryOfWorksheets = Result
End Function

Private Function TryGetWorksheetInWorkbookByName(ByVal Workbook As Workbook, ByVal WorksheetName As String, ByRef OutWorksheet As Worksheet) As Boolean
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        If Worksheet.Name Like WorksheetName Then
            Set OutWorksheet = Worksheet
            TryGetWorksheetInWorkbookByName = True
            Exit Function
        End If
    Next Worksheet
End Function

Private Function TryGetWorksheetInCollectionByName(ByVal Collection As Collection, ByVal WorksheetName As String, ByRef OutWorksheet As Worksheet) As Boolean
    Dim Worksheet As Worksheet
    For Each Worksheet In Collection
        If Worksheet.Name Like WorksheetName Then
            Set OutWorksheet = Worksheet
            TryGetWorksheetInCollectionByName = True
            Exit Function
        End If
    Next Worksheet
End Function

Private Function GetAllWorksheetsInApplication(ByVal Application As Application) As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    
    For Each Workbook In Application.Workbooks
        For Each Worksheet In Workbook.Worksheets
            Result.Add Worksheet
        Next Worksheet
    Next Workbook
    
    Set GetAllWorksheetsInApplication = Result
End Function
