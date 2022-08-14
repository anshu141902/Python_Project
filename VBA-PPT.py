Option Explicit
Sub myPPT()
'Need to defined object for using powerpoint presentation
Dim ppApp As PowerPoint.Application
Dim ppPres As PowerPoint.Presentation
Dim ppSlide As PowerPoint.Slide
Dim ppTextBox As PowerPoint.Shape
Dim ch As Charts

Set ppApp = New PowerPoint.Application
ppApp.Visible = msoCTrue
ppApp.Activate
Set ppPres = ppApp.Presentations.Add

'For adding slides in powerpoint presentation
Set ppSlide = ppPres.Slides.Add(1, ppLayoutTitle)

'for adding shapes in powerpoint presentation
ppSlide.Shapes(1).TextFrame.TextRange.Text = "Movie Presentation"
ppSlide.Shapes(2).TextFrame.TextRange.Text = "By wise owl"

'For adding second slides in powerpoint presentation
Set ppSlide = ppPres.Slides.Add(2, ppLayoutBlank)

'Copy range from Excel to Power Point Presenation
Range("a1").CurrentRegion.Copy
ppSlide.Shapes.PasteSpecial ppPasteBitmap
ppSlide.Shapes(1).Width = ppPres.PageSetup.SlideWidth / 2
ppSlide.Shapes(1).Height = ppPres.PageSetup.SlideHeight / 2
ppSlide.Shapes(1).Left = 0
ppSlide.Shapes(1).Top = (ppPres.PageSetup.SlideHeight / 2) - (ppSlide.Shapes(1).Height / 2)

'for adding textbox in power point presentation
Set ppTextBox = ppSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 20, ppPres.PageSetup.SlideWidth, 60)
With ppTextBox.TextFrame
        .TextRange.Text = "List of Current Films"
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        .TextRange.Font.Size = 26
        .TextRange.Font.Name = "Arial"
        .VerticalAnchor = msoAnchorMiddle
End With

'For adding third slides in powerpoint presentation
Set ppSlide = ppPres.Slides.Add(3, ppLayoutBlank)

'for adding charts from excel to powerpoint presentation
Set ch = ThisWorkbook.Sheets(1).ChartObjects("Chart1")
ch.Copy
ppSlide.Shapes.PasteSpecial ppPastePNG
'Chart alignments
                      ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoCTrue
                      ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoCTrue
End Sub


