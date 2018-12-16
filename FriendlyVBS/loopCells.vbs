' Based on code from Kelvin Sung
' File: loopCells.vbs 
' Allow the user to open an existing excel document, and access cells 

Option Explicit           'force all variables to be declared 

' Get information on the "current folder" of this script and open a connection 
' to the "operating environment". 
dim wshShell, currentFolder

' wshShell contains the environment for which we are operating in; 
' use it to retrieve the current work folder
set wshShell = WScript.CreateObject("WScript.Shell")
        
' currentFolder contains the path to the folder where this script is located
currentFolder = wshShell.CurrentDirectory

' open the the msoFileDialog with initial folder setting to the "currentFolder"
dim xlAppl
set xlAppl = CreateObject("Excel.Application")

const msoFileDialogOpen = 1
dim dlgOpen                          ' use MS Office FileDialog to open a file
set dlgOpen = xlAppl.Application.FileDialog(msoFileDialogOpen)

dim selectedFile
dim dlgAnswer

dlgOpen.AllowMultiSelect = false           ' only allow selection of one file
dlgOpen.InitialFileName = currentFolder

if (dlgOpen.Show() = -1) then              ' -1 says user clicked on "Open"
    ' the first selected item will be our file name
    selectedFile = dlgOpen.SelectedItems(1)

    ' now open the file 
    xlAppl.Application.Workbooks.Open(selectedFile)
else
    MsgBox "No document opened!"
end if

' each excel document can have many work sheets, do our work on Sheet1
dim activeSheet
set activeSheet = xlAppl.Worksheets("Sheet1")

'--------------------------------------------------------------------------
' The code to work with the spreadsheet starts here.
dim row                      ' row is row number
dim col                      ' col is column number

' activeSheet.Cells(row, col) is the cell value 
'
' For example, activeSheet.Cells(4, 6) is the cell in row 4, column 6
' The cells are strings, so convert if needed. If you do arithmetic
' with them, they are automatically converted.

' Pause to see the original spreadsheet. 
MsgBox "Notice that there is nothing in column five or six."

' Here is a loop to demonstrate. 
' The loop puts the value of variable row into each cell in rows one
' through ten, column five. It then copies column 5 to column 6.
for row = 1 to 10
   activeSheet.Cells(row, 5) = row 
   activeSheet.Cells(row, 6) = activeSheet.Cells(row, 5) 
next

' To keep the spradsheet open to allow saving 
MsgBox "The End. Click the excel closing box if you wish to save it."

'--------------------------------------------------------------------------
xlAppl.Application.DisplayAlerts = false
xlAppl.Application.Quit

