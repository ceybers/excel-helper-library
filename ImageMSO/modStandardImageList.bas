Attribute VB_Name = "StandardImageList"
'@Folder "Helpers.Controls"
Option Explicit

Private Const DEFAULT_ICON_SIZE As Long = 16
Private Const DEFAULT_MSO_KEYS As String = "root,BlogHomePage;wb,FileSaveAsExcelXlsx;" & _
    "ws,HeaderFooterSheetNameInsert;lo,CreateTable;col,TableColumnSelect;activeLo,TableSelect;" & _
    "delete,Delete;AutoSum,AutoSum;MagicWand,QueryBuilder;TraceError,TraceError;" & _
    "Tick,AcceptInvitation;Cross,DeclineInvitation;Key,AdpPrimaryKey;Fx,TableFormulaDialog;" & _
    "Excel,MicrosoftExcel"

'@Description "Returns a new ImageList object pre-populated with a standardised list of default icons."
Public Function GetMSOImageList(Optional ByVal IconSize As Long = DEFAULT_ICON_SIZE) As ImageList
    Dim Result As ImageList
    Set Result = New ImageList
    
    Dim ImageTuple As Variant
    For Each ImageTuple In Split(DEFAULT_MSO_KEYS, ";")
        AddImageToImageList Result, Split(ImageTuple, ",")(0), Split(ImageTuple, ",")(1), IconSize
    Next ImageTuple
    
    Set GetMSOImageList = Result
End Function

Private Sub AddImageToImageList(ByVal ImageList As ImageList, ByVal Key As String, ByVal ImageMso As String, ByVal IconSize As Long)
    Dim Picture As IPictureDisp
    Set Picture = Application.CommandBars.GetImageMso(ImageMso, IconSize, IconSize)
    ImageList.ListImages.Add 1, Key, Picture
End Sub