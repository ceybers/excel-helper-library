Option Explicit

'@Description "Removes all headers and footers in a document."
'@EntryPoint
Public Sub RemoveHeadersAndFooters()
    With ActiveDocument.PageSetup
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
    End With
    
    Dim Section As Section
    For Each Section In ActiveDocument.Sections
        Section.Headers(wdHeaderFooterPrimary).Range.Text = vbNullString
        Section.Footers(wdHeaderFooterPrimary).Range.Text = vbNullString
    Next Section
End Sub

'@Description "Resets the page layout to A4 with narrow margins."
'@EntryPoint
Public Sub ResetPageLayout()
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(1.27)
        .BottomMargin = CentimetersToPoints(1.27)
        .LeftMargin = CentimetersToPoints(1.27)
        .RightMargin = CentimetersToPoints(1.27)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
    End With
End Sub