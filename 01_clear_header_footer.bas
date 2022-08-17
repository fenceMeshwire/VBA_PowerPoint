Option Explicit

' If you have to deal with the annoying header footer issue... 
' Just run the following script:

Sub clear_header_footer()

Dim slide As slide

For Each slide In ActivePresentation.Slides
  slide.DisplayMasterShapes = True
  slide.HeadersFooters.Clear
Next

End Sub
