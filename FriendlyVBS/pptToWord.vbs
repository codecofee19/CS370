' Based on code by Kelvin Sung
' File: pptToWord.vbs
'
' Open a powerpoint document and put all the text into a word document.

Option Explicit

' specific constants, vbScript cannot see these constants, so 
' find out what they are, and redefine them
const msoFileDialogOpen = 1

' getting the operating environment (e.g., Current Working Folder)
dim WshShell
set WshShell = WScript.CreateObject("WScript.Shell")

dim pptAppl
set pptAppl = CreateObject("Powerpoint.Application")

dim WdAppl
set WdAppl = CreateObject("Word.Application")

WdAppl.Application.Visible = FALSE

' OpenFile Dialog using MSO facility, since MSO is implemented in both
' Word and Excel, use it in exactly the same way. 
dim DlgOpen
set DlgOpen = pptAppl.FileDialog(msoFileDialogOpen)

if (OpenPPTDocument()) then

   ' Powerpoint uses ActivePresenation just as Word uses ActiveDocument 
   ' and Excel uses ActiveWorkbook.
   '
   ' slideShow is now the entire powerpoint presentation
   dim slideShow, oneSlide
   set slideShow = pptAppl.ActivePresentation.Slides

   MsgBox "This powerpoint file has: " & slideShow.Count & " number of slides"

   SavePptToWord slideShow, pptAppl.ActivePresentation.Name

   ' quit powerpoint
   pptAppl.ActivePresentation.Close
   pptAppl.Quit

   ' quit word
   WdAppl.Application.Quit
else
   MsgBox("No PowerPoint Document Selected, Bye Bye")
   pptAppl.Quit
   WdAppl.Application.Quit
end if

'----------------------------------------------------------------------------
' Function: OpenPPTDocument()
'
' Input:   none (uses the global DlgOpen to let user open a .doc document)
' Returns: none (changes the global pptAppl variable to have an opened document
' Error:   Checks to make sure input is a number
 
FUNCTION OpenPPTDocument()

   ' perform multiple operations based on the same object, so use the
   ' "with" statement. This is cleaner, especially with activeX.
   with DlgOpen
      .AllowMultiSelect = FALSE
      .InitialFileName = WshShell.CurrentDirectory
      .Filters.Clear
      .Filters.Add "PowerPoint Files", "*.ppt"
   end with

   OpenPPTDocument = FALSE
   if (dlgOpen.Show() = -1) then
      dim selectedFile                          ' -1 says user clicked "Open"
      selectedFile = dlgOpen.SelectedItems(1)
      pptAppl.Visible = TRUE                    ' set ppt to visible
      pptAppl.Presentations.Open selectedFile   ' now open the file 
      OpenPPTDocument = TRUE
   end if
end FUNCTION

'----------------------------------------------------------------------------
' SUB SavePptToWord slideShow:
'
' Creates an MS Word Doc document with from ppt slideShow content.
' It is formatted somewhat.
SUB SavePptToWord(ByRef slides, ByVal name)
   const wdToggle = 9999998
   ' create a new MS Word Doc
   dim newWdDoc
   set newWdDoc = WdAppl.Documents.Add

   dim sel
   set sel = WdAppl.Application.Selection

   WdAppl.Application.Visible = TRUE      ' get back reference for editing

   with sel
      .Style = WdAppl.Application.ActiveDocument.Styles("Heading 1")
      .TypeText "Content of the Presenation: " + name
      .TypeParagraph
      .TypeParagraph
   end with

   dim count
   for count = 1 to slides.Count
      Set oneSlide = slides.Item(count)       ' one of the slides

      ' look at how many input boxes are defined for this slide
      dim slideSections, titleName
      set slideSections = oneSlide.Shapes

      ' find the title for this slide
      if (slideSections.HasTitle) then
         titleName = slideSections.Title.TextFrame.TextRange.Text
      else
         titleName = "HAS NO TITLE"
      end if

      with sel
         .Style = WdAppl.Application.ActiveDocument.Styles("Heading 2")
         .TypeText "Slide " & Count & ": " & titleName
         .TypeParagraph
      end with

      dim shapeIndex
      for shapeIndex = 1 to slideSections.Count
         ' Rest of the text for this slide
         dim aSection
         set aSection = slideSections.Item(shapeIndex)
         if (aSection.HasTextFrame) then
            with sel
               .TypeText aSection.TextFrame.TextRange.Text 
               .TypeParagraph
            end with
         end if
      next
   next
end SUB

