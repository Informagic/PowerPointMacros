' This VBA skript inserts page numbering into PowerPoint slides:
' • hidden slides are ignored
' • the page numbers have the form "21 of 50"
Sub InsertCorrectPageNumbering()
  On Error Resume Next
  With ActivePresentation
    nonHiddenSlides = 1
    For X = 2 To .Slides.Count ' we don't want to have it on the title slide
      .Slides(X).Shapes("FullSlideNumber").Delete
      If .Slides(X).SlideShowTransition.Hidden = msoFalse Then
        nonHiddenSlides = nonHiddenSlides + 1
      End If
    Next X:
    
    slideCounter = 2
    For X = slideCounter To .Slides.Count
      If .Slides(X).SlideShowTransition.Hidden = msoFalse Then
        Set s = .Slides(X).Shapes.AddTextbox(msoTextOrientationHorizontal, _
              .PageSetup.SlideWidth - 100, _
              .PageSetup.SlideHeight - 30, _
              90, 20)
        With s.TextFrame.TextRange
          .Text = slideCounter & " of " & nonHiddenSlides
          .Font.Name = "Calibri"
          .Font.Size = 12
          .Font.Color.RGB = RGB(137, 137, 137)
          .ParagraphFormat.Alignment = ppAlignRight
        End With
        s.Name = "FullSlideNumber"
        
        slideCounter = slideCounter + 1
      End If
    Next X:
  End With
End Sub