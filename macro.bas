Attribute VB_Name = "Module1"
Sub Automate()
'Open the Excel workbook. Change the filename here.
Dim OWB As New Excel.Workbook
Set OWB = Excel.Application.Workbooks.Open(ActivePresentation.Path & "\Automation_Sheet.xlsx")
'Grab the first Worksheet in the Workbook
Dim WS As Excel.Worksheet
Set WS = OWB.Worksheets("Speech")
'Loop through each row until the row is empty'
For i = 2 To WS.Range("A65536").End(xlUp).Row
    'Copy the first slide and paste at the end of the presentation
    ActivePresentation.Slides(1).Copy
    ActivePresentation.Slides.Paste (ActivePresentation.Slides.Count + 1)
 
    'Change the Speaker's name text'
    ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("SpeakerName").TextFrame.TextRange.Text = WS.Cells(i, 1).Value
    'Change the Speaker's job text'
    ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("SpeakerJob").TextFrame.TextRange.Text = WS.Cells(i, 2).Value
    'Change the Speaker's image'
    ActivePresentation.Slides(ActivePresentation.Slides.Count).Shapes("SpeakerImage").Fill.UserPicture (ActivePresentation.Path & "\" & WS.Cells(i, 3).Value)
Next
OWB.Close
End Sub
