' Based on code by Kelvin Sung
' File: openTextFile.vbs
'
' Define a Function to use MSO FileOpenDial facility
' to help the user select a text file to open

Option Explicit

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
   dim lineCount                         ' number of lines in the file
   lineCount = 0

   do while not inputTextStream.AtEndOfStream 
      lineCount = lineCount + 1
      inputLine = inputTextStream.ReadLine
      MsgBox "Line " & lineCount & " content:  " & inputLine
   loop

else
   MsgBox "File: " & dataFile & " does not exists."
end if

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

