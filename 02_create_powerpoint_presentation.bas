Option Explicit

Sub create_powerpoint_presentation()

Dim wkbBook As Excel.Workbook
Dim wksSheet As Excel.Worksheet
' ______________________________________________________________________  
Dim ppApplication As Object
Dim ppSlideShow As Object
Dim ppSlide As Object
Dim ppTable As Object
Dim ppDiagram As Object
Dim strPath As String
' ______________________________________________________________________
strPath = ThisWorkbook.Path & "\" & "presentation.pptx"

Set wkbBook = ThisWorkbook
Set wksSheet = Tabelle1

' Initialize PowerPoint objects
Set ppApplication = CreateObject("PowerPoint.Application")

With ppApplication
  .Visible = msoTrue
  .WindowState = 1
  .Activate
  
  ' Behavior depending on whether the presentation already exists or needs to be created.
  If strPath = "" Then
    Set ppSlideShow = .presentations.Add
  Else
    If Dir(strPath) = "" Then
      Set ppSlideShow = .presentations.Add
    Else
      Set ppSlideShow = .presentations.Open(strPath, msoFalse)
    End If
  End If
  
End With

' Create a new slide
Set ppSlide = ppSlideShow.slides.AddSlide(ppSlideShow.slides.Count + 1, ppSlideShow.SlideMaster.CustomLayouts(6))
ppSlide.Shapes(1).TextFrame.TextRange = wksSheet.Name
ppSlide.Select

' Transfer data
wksSheet.Range("A1:C5").Select
Selection.Copy
ppSlideShow.Application.ActiveWindow.View.Paste

Set ppTable = ppSlide.Shapes(ppSlide.Shapes.Count)
ppTable.Left = 100
ppTable.Top = 150

' Create a new slide
Set ppSlide = ppSlideShow.slides.AddSlide(ppSlideShow.slides.Count + 1, ppSlideShow.SlideMaster.CustomLayouts(6))
ppSlide.Select

' Transfer a chart
ppSlide.Shapes(1).TextFrame.TextRange = wksSheet.Name
wksSheet.ChartObjects(1).Copy
ppSlideShow.Application.ActiveWindow.View.Paste

Set ppDiagram = ppSlide.Shapes(ppSlide.Shapes.Count)
ppDiagram.Left = 100
ppDiagram.Top = 150

' Save PowerPoint presentation
ppSlideShow.SaveAs (strPath)
ppApplication.Quit

Set wkbBook = Nothing
Set wksSheet = Nothing
Set ppSlideShow = Nothing
Set ppApplication = Nothing

End Sub
