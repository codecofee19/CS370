' Based on code by Kelvin Sung
' File: searchWord.vbs
'
' Shows how to open a MS Word document via ActiveX.
' Ask the user for a string and search for it in the document.
'
' Uses the same MSO FileOpenDialog as used via the excel examples.

Option Explicit

' Very similar to opening an excel document.
'   dim XlAppl
'   set XlAppl = CreateObject("Excel.Application")
dim DOCAppl
set DOCAppl = CreateObject("Word.Application")

dim searchString
dim count

if (Open_OFFICE_Document(DOCAppl, "Word Files", "*.doc")) then

   searchString = "Something"
   do while searchString <> ""
      searchString = InputBox("Enter the string to search [enter to quit]: ")
      if (searchString <> "") then
         count =  countStringInDoc(searchString, DOCAppl)
         if (count > 0) Then
            MsgBox "String [" & searchString & "] is found in the document: " _
                   & count & " times"
         else
            MsgBox "String [" & searchString & "] is NOT found in the document"
         end if
      end if
   loop
    
   DOCAppl.Application.Documents.Close
   DOCAppl.Application.Quit
else
   DOCAppl.Application.Quit
end if


'-----------------------------------------------------------------------------
' Function Name:    countStringInDoc(ByVal searchString)
' Description:      Takes in 1 input parameter, and returns count, the
'                   number of times the string is found in the document
'
' Input Parameters: ByRef searchString,  ByRef applObj
'                      1) searchString -- the string you want to find
'                      2)   applObj -- the application object
'
' Returns Value:    INTEGER: number of times string is found

FUNCTION countStringInDoc(ByVal searchString, ByRef applObj)

   ' set up to search the entire document
   dim findRange
   set findRange = applObj.ActiveDocument.Range()

   ' program Word's search facility
   With findRange
      .Find.Text = searchString        ' assign the string to search
      .Find.Forward = TRUE             ' look in the forward direction
   end With

   dim doneSearching, count
   doneSearching = FALSE
   count = 0

   ' search through the doc, counting occurrences of searchString
   do while not doneSearching
      findRange.Find.Execute()         ' start Word searching
      if (findRange.Find.Found) then
         count = count + 1
      else
         doneSearching = TRUE          ' can't find anymore strings
      end if
   loop

   countStringInDoc = count
end FUNCTION


'-----------------------------------------------------------------------------
' Function Name: Open_OFFICE_Document(ByRef applicationObj, ByRef filterDes,
'                                     ByRef filterExt)
' Description:  Takes in 3 input parameter, and returns
'               TRUE or FALSE, whether the file is being opened.
'
'               It uses the WScript's CreateObject method to establish the 
'               connection with the "Operation Enviroment".
'
'               After setting up the "connection", use the Microsoft Office 
'               OpenFileDialog facility to open a file, through the 
'               "connection".
'
' Input Parameters: ByRef applicationObj, ByRef filterDes, ByRef filterExt
'                   1) applicationObj -- The application object
'                   2) filterDes -- file description type you want to open,
'                      pass by reference as String
'                      e.g., filterDes = "MS Word Files" to work with Word 
'                   3) filterExt -- extension of the file type to open,
'                      pass by reference as String, e.g.,
'                      filterExt = "*.ppt" to work with PowerPoint files
'
' Returns Value:    BOOLEAN, whether the file is being opened or not
' Possible Errors:  The passed in applicationObj is not created correctly
'                   The values of filterDes or filterExt are not inside 
'                   double quotations
FUNCTION Open_OFFICE_Document(ByRef applicationObj, ByRef filterDes, _
                              ByRef filterExt)

   ' Constants specific to _OFFICE_ Object. VBScript cannot see these 
   ' constants, so we have to find out what they are, and redefine them.
   const msoFileDialogOpen = 1

   ' Open a "connection" to the "operating environment" .
   ' WshShell contains the environment for which we are operating in.
   ' For example, below it is used to retrieve the current working folder. 
   dim WshShell
   set WshShell = WScript.CreateObject("WScript.Shell")

   ' OpenFile Dialog we use MSO facility, since MSO is implemented in either
   ' Word, Excel, even PowerPoint, we can use it in exactly the same way. 
   dim DlgOpen
   set DlgOpen = applicationObj.FileDialog(msoFileDialogOpen)

   With DlgOpen
      ' Notice when we want to perform multiple operations based on the
      ' same object, we can use the "With" statement. This is much more
      ' efficient (especially when dealing with ActiveX connection)
      .AllowMultiSelect = False
         ' Only allow selection of one file
         ' TRUE:    Allow multiple file selection
         ' FALSE:   Do not allow multiple file selection
      .InitialFileName = WshShell.CurrentDirectory
         ' Start looking from current working folder
         ' CurrentFolder contains the path to the folder of this script 
      .Filters.Clear                           'clear filter

      ' Set up "Filters", so that we only work with certain types of files.
      ' The parameters of Add are:
      ' 1st - description of what is the file type
      ' 2nd - the extension (don't forget the "*")
      ' 3rd - the position for this entry
      .Filters.Add filterDes, filterExt, 1
   end With

   if ( dlgOpen.Show() = -1 ) then
      ' -1 says user clicked on "Open"
      ' Since we do not allow multiple select, the first selected item 
      ' will be our file name
      dim selectedFile
      selectedFile = dlgOpen.SelectedItems(1)

      ' Open the file, can switch the visibility of MSon/off (true/false)
      ' TRUE:    Show Excel is running
      ' FALSE:   Do not show Excel is running
      applicationObj.Visible = TRUE
      
      '--------     ALTERNATIVES     --------
      Select Case filterExt
         case "*.mdb"
               ' For Access files
               applicationObj.OpenCurrentDatabase selectedFile

         case "*.ppt"
               ' For PowerPoint files
               applicationObj.Presentations.Open selectedFile

         case "*.doc"
               ' For Word files
               applicationObj.Application.Documents.Open selectedFile

         case "*.xls"
               ' For Excel files
               applicationObj.Application.Documents.Open selectedFile
      end select

      Open_OFFICE_Document = TRUE             ' set return to TRUE
   else
      'else if user did not click on "Open"
      Open_OFFICE_Document = FALSE            ' set return to FALSE
   end if

end FUNCTION

