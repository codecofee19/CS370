' Based on code by Kelvin Sung
' File: openOldExcel.vbs
'
' Use Microsoft Office (mso) ActiveX service to open an existing document
'
' Open a document from user input
 
Option Explicit

' We need to go through MS Excel (or MS Office, Word is fine), the 
' application script editor, to find the value to send to the FileDialog.
' Refer to on-line help page for the details.
'
' In this case, MsoFileDialogType.msoFileDialogOpen with a value of one
' is used to open the file.
const msoFileDialogOpen = 1

dim xlAppl
set xlAppl = CreateObject("Excel.Application")

' use Microsoft Office FileDialog to open a file
dim dlgOpen          
set dlgOpen = xlAppl.Application.FileDialog(msoFileDialogOpen)

dim selectedFile
dim dlgAnswer

' only allow the selection of one file
dlgOpen.AllowMultiSelect = false

' -1 says user clicked on "Open"
if (dlgOpen.Show() = -1 ) then
   ' since we disallow multiple select, the first selected item is our file 
   selectedFile = dlgOpen.SelectedItems(1)

   ' now open the file 
   xlAppl.Application.Workbooks.Open(selectedFile)
   msgbox "selectedFile is: " & selectedFile
else
   MsgBox "No document opened!"
end if

xlAppl.Application.DisplayAlerts = false
xlAppl.Application.Quit

