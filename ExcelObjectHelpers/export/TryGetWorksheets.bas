Attribute VB_Name = "TryGetWorksheets"
'@Folder "CommonHelpers"
Option Explicit

'@Description "Tries to return a Collection of all the Worksheets with a given name from a Collection or an Application object."
Public Function TryGetWorksheetsByName(ByVal Parent As Object, ByVal WorksheetName As String, ByRef OutWorksheets As Collection) As Boolean
    If TypeOf Parent Is Collection Then
        TryGetWorksheetsByName = TryGetWorksheetsInCollectionByName(Parent, WorksheetName, OutWorksheets)
    ElseIf TypeOf Parent Is Application Then
        TryGetWorksheetsByName = TryGetWorksheetsInCollectionByName(GetAllWorksheetsInApplication(Parent), WorksheetName, OutWorksheets)
    End If
End Function

Private Function TryGetWorksheetsInCollectionByName(ByVal Collection As Collection, ByVal WorksheetName As String, ByRef OutWorksheets As Collection) As Boolean
    Dim Result As Collection
    Set Result = New Collection
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Collection
        If Worksheet.Name Like WorksheetName Then
            Result.Add Worksheet
            TryGetWorksheetsInCollectionByName = True
        End If
    Next Worksheet
    
    Set OutWorksheets = Result
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

