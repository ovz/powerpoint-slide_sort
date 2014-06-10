Attribute VB_Name = "Module11"
Sub sort_rand()

    Dim i As Integer
    Dim myvalue As Integer
    Dim islides As Integer
    islides = ActivePresentation.Slides.Count
    For i = 2 To ActivePresentation.Slides.Count - 2
        myvalue = Int((i * Rnd) + 2)
        ActiveWindow.ViewType = ppViewSlideSorter
'        ActivePresentation.Slides(myvalue).Select
'        ActiveWindow.Selection.Cut
        ActivePresentation.Slides(myvalue).Copy
        ActivePresentation.Slides.Paste
        ActivePresentation.Slides(myvalue).Delete
'        ActiveWindow.View.Paste
    Next

End Sub

