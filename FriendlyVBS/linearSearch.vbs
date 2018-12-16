' File: linearSearch.vbs
'
' Open a text file to save names into an array.
' Open another text file to save grades into another array.
' Given a name to search for, display the grade.

Option Explicit

dim names(50)                ' holds names 
dim grades(50)               ' holds grades 

dim count                    ' total number of names, subscript 0 to count-1
dim count2                   ' used for filling array with grades
count = 0
count2 = 0

' Fill the names and grades arrays with data from text files that
' are opened in the subroutine. Display the two arrays in MsgBoxes.
call fillArray(names, count)
call fillArray(grades, count2)
call printOneArray(names, "The names")
call printOneArray(grades, "The grades")

dim targetName                        ' name to search for
dim foundGrade                        ' grade associated with name 

' For each name that is entered, locate the name in the names array,
' and display the associated grade from the grades array.
do 
   targetName = inputBox("Enter a name to search for [enter to quit]")
   if targetName = "" then
      exit do
   end if

   ' search the names array for targetName and return the associated grade
   foundGrade = gradeValue(targetName, names, grades, count)
   if foundGrade <> -1 then
      MsgBox(targetName & " has a grade of " & foundGrade)
   else
      MsgBox(targetName & " was not found in the list of names")
   end if
loop


'------------------------------------------------------------------------------
' gradeValue
' Given an array of names and an array of grades with the count of how 
' many are in the array, return the grade of the entered name.
' Return a -1 if the name is not found.
FUNCTION  gradeValue(ByVal targetName, ByRef names, ByRef grades, ByVal count)
   dim i, found
   i = 0
   found = FALSE                           ' have not found name yet
   gradeValue = -1

   ' Look for targetName as long as we are still within the bounds
   ' of the array and the name has not been found yet.
   do while i <= count and not found
      if names(i) = targetName then
         gradeValue = grades(i)
         found = TRUE
      end if
      i = i + 1
   loop
end FUNCTION   


'------------------------------------------------------------------------------
' fillArray
' Open a text file and read the data into an array.
' The count will be the true size of the array (number of elements in use).
SUB fillArray(ByRef myArray, ByRef count)

   ' ask the scripting runtime environemt for access to files
   dim FS 
   set FS = CreateObject("Scripting.FileSystemObject")

   const FILEFORREADING = 1
   dim dataFile
   dataFile = selectAFileOrFolder("FILE")

   if FS.FileExists(dataFile) then
      ' get the fileHander based on the dataFile
      dim FileHandler
      set FileHandler = FS.GetFile(dataFile)
   
      ' open the file as a inputTextStream so text data can be "streamed"
      dim inputTextStream
      set inputTextStream = FileHandler.OpenAsTextStream(FILEFORREADING)

      dim inputLine                         ' a line in the file

      ' As long as we're not at the end of the file, read a line from
      ' the file and place it into myArray.
      do while not inputTextStream.AtEndOfStream 
         inputLine = inputTextStream.ReadLine
         myArray(count) = inputLine 
         count = count + 1
      loop
   else
      MsgBox "File: " & dataFile & " does not exists."
   end if
end SUB


'----------------------------------------------------------------------------
' printOneArray
' Takes in 2 input parameters, the array to display and a string which 
' is a description of what is displayed to be used as the MsgBox title. 
' It displays the contents of the array in message box titled description.
SUB printOneArray(ByRef myArray, ByRef description)
   dim output, i
   output = ""
   For i = 0 to UBound(myArray)
      output = output & "myArray(" & i & ") = " & myArray(i) & vbNewLine
   next

   MsgBox output, ,description
end SUB


'------------------------------------------------------------------------------
' selectAFileOrFolder
' Takes in one input parameter, "FILE" or "FOLDER", and returns its path
' in a string that contains either the selected file or folder path.
FUNCTION selectAFileOrFolder(ByRef fileOrFolder)

   dim WshShell, currentFolder

   ' Open a "connection" to the "operating environment" .
   ' WshShell contains the environment for which we are operating in.
   ' For example, below we will use it to retrieve the current working folder.
   set WshShell = WScript.CreateObject("WScript.Shell")
   
   ' Get the current folder (where this script is opened from).
   ' currentFolder now contains the path to the folder 
   currentFolder = wshShell.CurrentDirectory
   
   ' XlAppl is our "connection" to all functions MS Excel provides
   dim XlAppl
   set XlAppl = WScript.CreateObject("Excel.Application")
    
   ' don't let user see what is going on,
   ' can switch the visibility of MSon/off (TRUE/FALSE)
   XlAppl.Application.Visible = FALSE

   ' specific constants to OFFICE Object. 
   ' VBScript cannot see these constants, so we have
   ' to find out what they are, and re-define them.
   dim msoFileDialogOpen
   if (fileOrFolder = "FILE") then
      msoFileDialogOpen = 1
   end if
   if (fileOrFolder = "FOLDER") then
      msoFileDialogOpen = 4
   end if
   
   ' OpenFile Dialog we use MSO facility, since MSO is implemented in either
   ' Word, Excel, even PowerPoint. It can be used in exactly the same way. 
   ' Notice "set" is used here too.
   dim DlgOpen   
   set DlgOpen = XlAppl.Application.FileDialog(msoFileDialogOpen)

   ' Only allow selection of one file
   ' TRUE:      Allow multiple file selection
   ' FALSE:   Do not allow multiple file selection
   DlgOpen.AllowMultiSelect = FALSE
   
   ' Set up the "Filters" so that we only work with certain types of files.
   '
   ' DlgOpen.Filters.Clear
   ' DlgOpen.Filters.Add "Text Files", "*.txt", 1
   ' Only want to work with Text files
   ' The parameters of Add are:
   '    1st - description of what is the file type
   '    2nd - the extension (don't forget the "*") 
   '    3rd - the position for this entry

   ' Start looking from current working folder
   ' currentFolder contains the path to the folder where this script is located
   DlgOpen.InitialFileName = currentFolder
   
   SelectAFileOrFolder = ""             ' Set the return value to ""

   if (DlgOpen.Show() = -1) then        ' -1 says user clicked on "Open"
      ' first selected item is the file name
      SelectAFileOrFolder = DlgOpen.SelectedItems(1)
   end if
   
   ' TRUE:    Prompt user to save the work
   ' FALSE:   Do not prompt user to save the work
   XlAppl.Application.DisplayAlerts = FALSE
   
   XlAppl.Application.Quit                  ' quit the application
end FUNCTION

