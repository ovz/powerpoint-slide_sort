Attribute VB_Name = "Module13"
Sub sort_back()

    Dim i As Integer
    Dim first As Integer
    Dim last As Integer
    first = 1
    last = ActivePresentation.Slides.Count
    For i = 1 To (ActivePresentation.Slides.Count) / 2
        ActiveWindow.ViewType = ppViewSlideSorter
        ActivePresentation.Slides(last).Copy
        ActivePresentation.Slides.Paste (first)
        ActivePresentation.Slides(last + 1).Delete
        
        ActivePresentation.Slides(first + 1).Copy
        ActivePresentation.Slides.Paste (last + 1)
        ActivePresentation.Slides(first + 1).Delete
        
        first = first + 1
        last = last - 1
        
    Next

End Sub

