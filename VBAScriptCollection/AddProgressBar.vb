' Adds a progress bar at the top of each slide:
' • hidden slides are ignored
' • section heading slides are ignored
' The progress bar has an offset to the left to accomodate for a design
' specific side bar. If that is not required, the variable "Offset" may
' be set to 0.
Sub AddProgressBar()
  On Error Resume Next
  With ActivePresentation
    nonHiddenSlides = 1
    For X = 2 To .Slides.Count ' we don't want to have it on the title slide
      .Slides(X).Shapes("ProgressBar").Delete
      If .Slides(X).SlideShowTransition.Hidden = msoFalse Then
        nonHiddenSlides = nonHiddenSlides + 1
      End If
    Next X:
    
    Offset = 78 ' the width of the talk outline object
    slideCounter = 2
    For X = slideCounter To .Slides.Count
      If .Slides(X).SlideShowTransition.Hidden = msoFalse Then
        ' and neither on the section header slides
        If Not .Slides(X).CustomLayout.Name = "Section Header" Then
          Set s = .Slides(X).Shapes.AddShape(msoShapeRectangle, _
                Offset, 0, _
                slideCounter * (.PageSetup.SlideWidth - Offset) / nonHiddenSlides, 12)
          s.Fill.ForeColor.RGB = RGB(0, 32, 96)
          s.Fill.BackColor.RGB = RGB(0, 32, 96)
          s.Fill.OneColorGradient msoGradientVertical, 1, 0.1
          s.Borders.Enable = False
          s.Line.Visible = False
          s.Name = "ProgressBar"
        End If
        slideCounter = slideCounter + 1
      End If
    Next X:
  End With
End Sub