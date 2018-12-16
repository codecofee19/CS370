' Based on code by Kelvin Sung
' File: openPpt.vbs
'
' Open a powerpoint file and display the text from each slide.
' The same MSO FileOpenDialog (as used with Excel) is used.

Option Explicit

' specific constants, vbScript cannot see these constants, so 
' find out what they are, and redefine them
const msoFileDialogOpen = 1

dim WshShell
set WshShell = WScript.CreateObject("WScript.Shell")
   ' For getting our operating environment (e.g. Current Working Folder)

' Very similar to word and excel, remember:
'    dim xlAppl
'    set xlAppl = CreateObject("Excel.Application")
dim pptAppl
set pptAppl = CreateObject("Powerpoint.Application")

' OpenFile Dialog uses MSO facility, since MSO is implemented in both
' Word and Excel, use it in exactly the same way. 
dim dlgOpen
set dlgOpen = pptAppl.FileDialog(msoFileDialogOpen)

if (openPPTDocument()) then
   ' PPT uses ActivePresenation whereas
   '    Word  uses ActiveDocument 
   '    Excel uses ActiveWorkbook
   ' slideShow is the entire powerpoint presentation
   dim slideShow, oneSlide
   set slideShow = pptAppl.ActivePresentation.Slides

   MsgBox "This powerpoint file has: " & slideShow.Count & " number of slides"

   dim count
   for count = 1 To slideShow.Count
      set oneSlide = slideShow.Item(count)         ' one of the slides

      dim slideSections, titleName

      ' look at how many input boxes that are defined for this slide
      set slideSections = oneSlide.Shapes

      ' find the title for this slide
      if (slideSections.HasTitle) then
         titleName = slideSections.Title.TextFrame.TextRange.Text
      else
         titleName = "HAS NO TITLE"
      end if

      dim shapeIndex, slideText
      slideText = ""
      for shapeIndex = 1 to slideSections.Count
         ' Rest of the text for this slide
         dim aSection
         set aSection = slideSections.Item(shapeIndex)
         if (aSection.HasTextFrame) then
            slideText = slideText & aSection.TextFrame.TextRange.Text & vbCrLf
         end if
      next
         
      MsgBox "Slide Number Is: " & oneSlide.SlideNumber & vbCrLf & _
             "Title is: " & titleName & vbCrLf & "Texts In This Slide: " & _
             vbCrLf & slideText
   next
      
   pptAppl.ActivePresentation.Close
   pptAppl.Quit
else
   MsgBox("No PowerPoint Document Selected, Bye Bye")
   pptAppl.Quit
end if

'-----------------------------------------------------------------------------
' FUNCTION: openPPTDocument()
'
' Input:   none (uses the global dlgOpen to let user open a .ppt document)
' Returns: none (changes the global pptAppl variable to have an opened document)
' Error:   Checks to make sure input is a number
'
'   Remark: 
FUNCTION openPPTDocument()

   with dlgOpen
      ' Notice when we want to perform multiple operations based on the
      ' same object, we can use the "with" statement. This is much more
      ' efficient (especially when dealing with ActiveX connection)
      '
      .AllowMultiSelect = FALSE
      .InitialFileName = WshShell.CurrentDirectory
      .Filters.Clear
      .Filters.Add "PowerPoint Files", "*.ppt"
      end with

   openPPTDocument = FALSE
   if (dlgOpen.Show() = -1) then
      dim selectedFile                          ' -1 says user clicked "Open"

      selectedFile = dlgOpen.SelectedItems(1)
      pptAppl.Visible = TRUE                    ' first set ppt to visible
      pptAppl.Presentations.Open selectedFile   ' now open the file 
      openPPTDocument = TRUE
   end if
end FUNCTION

