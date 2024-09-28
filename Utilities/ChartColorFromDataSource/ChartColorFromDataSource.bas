Attribute VB_Name = "ChartColorFromDataSource"
'@Folder("VBAProject")
Option Explicit

Private Const FORMULA_PREFIX As String = "=SERIES("
Private Const DEFAULT_INTERIOR_COLOR As Long = 16777215
Private Const DEFAULT_FONT_COLOR As Long = 0

Private Type SeriesFormula
    Legend As String
    Categories As String
    Values As String
End Type

'@EntryPoint
Public Sub ChartColorFromDataSource()
    If TypeOf Selection Is ChartArea Then
        FormatAllSeriesFromLegend
    ElseIf TypeOf Selection Is Legend Then
        FormatAllSeriesFromLegend
    ElseIf TypeOf Selection Is Series Then
        FormatOneSeriesFromLegend Selection
    ElseIf TypeOf Selection Is Point Then
        FormatOnePointFromValue Selection
    ElseIf TypeOf Selection Is Axis Then
        If Selection.Type = xlCategory Then
            FormatAllSeriesFromCategories
        End If
    ElseIf TypeOf Selection Is AxisTitle Then
        If Selection.Parent.Type = xlCategory Then
            FormatAllSeriesFromCategories
        End If
    End If
End Sub

Private Sub FormatAllSeriesFromLegend()
    Dim Series As Series
    For Each Series In ActiveChart.SeriesCollection
        FormatOneSeriesFromLegend Series
    Next Series
End Sub

Private Sub FormatOneSeriesFromLegend(ByVal Series As Series)
    Dim SeriesSources As SeriesFormula
    SeriesSources = GetSeriesFormulaValues(Series)
    
    Dim Range As Range
    If Not TrySetRange(SeriesSources.Legend, Range) Then Exit Sub
    
    SetFormatFromRange Series.Format, Range
End Sub

Private Sub FormatAllSeriesFromCategories()
    Dim Series As Series
    For Each Series In ActiveChart.SeriesCollection
        FormatOneSeriesFromCategories Series
    Next Series
End Sub

Private Sub FormatOneSeriesFromCategories(ByVal Series As Series)
    Dim SeriesSources As SeriesFormula
    SeriesSources = GetSeriesFormulaValues(Series)
    
    Dim Categories As Range
    If Not TrySetRange(SeriesSources.Categories, Categories) Then Exit Sub
    
    Dim Point As Long
    For Point = 1 To Categories.Cells.Count
        SetFormatFromRange Series.Points.Item(Point).Format, Categories.Cells.Item(Point)
    Next Point
End Sub

Private Sub FormatOnePointFromValue(ByVal Point As Point)
    Dim PointValueRange As Range
    If Not TryGetRangeFromPoint(Point, PointValueRange) Then Exit Sub
    
    SetFormatFromRange Point.Format, PointValueRange
End Sub

Private Sub SetFormatFromRange(ByVal Format As ChartFormat, ByVal Range As Range)
    Dim Color As Long
    If Range.Interior.Color <> DEFAULT_INTERIOR_COLOR Then
        Color = Range.Interior.Color
    ElseIf Range.Font.Color <> DEFAULT_FONT_COLOR Then
        Color = Range.Font.Color
    Else
        Exit Sub
    End If
    
    Format.Fill.ForeColor.RGB = Color
    Format.Line.ForeColor.RGB = Color
End Sub

'@Description "Tries to return the specific cell that a chart Point references its value from."
Private Function TryGetRangeFromPoint(ByVal Point As Point, ByRef OutRange As Range) As Boolean
Attribute TryGetRangeFromPoint.VB_Description = "Tries to return the specific cell that a chart Point references its value from."
    If Point Is Nothing Then Exit Function
    
    Dim PointName As String
    PointName = Point.Name
    
    Dim PointIndex As Long
    On Error Resume Next
    PointIndex = CLng(Mid$(PointName, InStr(PointName, "P") + 1, Len(PointName) - InStr(PointName, "P")))
    On Error GoTo 0
    If PointIndex = 0 Then Exit Function
    
    Dim SeriesSources As SeriesFormula
    SeriesSources = GetSeriesFormulaValues(Point.Parent)
    
    Dim Range As Range
    If Not TrySetRange(SeriesSources.Values, Range) Then Exit Function
    
    If Range.Cells.Count < PointIndex Then Exit Function
    
    Set OutRange = Range.Cells.Item(PointIndex)
    TryGetRangeFromPoint = True
End Function

Private Function TrySetRange(ByVal RangeText As String, ByRef OutRange As Range) As Boolean
    If RangeText = vbNullString Then Exit Function
    If Left$(RangeText, 1) = """" Then Exit Function
    
    On Error Resume Next
    Set OutRange = Application.Range(RangeText)
    On Error GoTo 0
    If OutRange Is Nothing Then Exit Function
    
    TrySetRange = True
End Function

'@Description "Converts a `=SERIES()` formula into a custom Data Type."
Private Function GetSeriesFormulaValues(ByVal Series As Series) As SeriesFormula
Attribute GetSeriesFormulaValues.VB_Description = "Converts a `=SERIES()` formula into a custom Data Type."
    Dim Formula As String
    Formula = Series.Formula
    If Left$(Formula, Len(FORMULA_PREFIX)) <> FORMULA_PREFIX Then Exit Function
    If Right$(Formula, 1) <> ")" Then Exit Function
    Formula = Mid$(Formula, Len(FORMULA_PREFIX) + 1, Len(Formula) - Len(FORMULA_PREFIX) - 1)
    
    Dim FormulaSplit As Variant
    FormulaSplit = Split(Formula, Application.International(xlListSeparator))
    If UBound(FormulaSplit) <> 3 Then Exit Function
    
    With GetSeriesFormulaValues
        .Legend = FormulaSplit(0)
        .Categories = FormulaSplit(1)
        .Values = FormulaSplit(2)
    End With
End Function
