' File: findLarge.vbs
'
' 6. In a loop, you input one integer at a time. The word "done" 
' terminates the loop. Find and display the largest number inputted.
 
Option Explicit                  ' force declaration of variables before use

dim inputNumber                  ' input from user
dim outputNums                   ' to show all the numbers user inputs
dim largest                      ' always current largest number
dim doneWithInput                ' whether or not user is done entering input
doneWithInput = false
outputNums = ""

' Loop to get one valid number from the user to initialize largest.
' The var doneWithInput will remember if the user is done before
' ever entering a valid number.
do
   inputNumber = InputBox("Please enter one number, enter 'done' to end")
   ' check for validity, exit on entry of word "done"
   if not IsNumeric(inputNumber) then
      if inputNumber = "done" then
         doneWithInput = true                
         exit do
      else
         MsgBox "You must enter a number, try again.", vbOKOnly, "Invalid Input"
      end if
   else
      largest = inputNumber
      outputNums = outputNums & " " & inputNumber
      exit do
   end if
loop

' If user isn't done, continue getting integers to find the largest.
' Continually compare with the current largest to see if it's larger,
' then it becomes the new current largest.
if not doneWithInput then
   do
      inputNumber = InputBox("Please enter one number, enter 'done' to end")

      ' check for validity, exit on entry of word "done"
      if not IsNumeric(inputNumber) then
         if inputNumber = "done" then
            exit do
         else
            MsgBox "You must enter a number, try again.", vbOKOnly, _
                   "Invalid Input"
         end if
      else  
         ' recall that InputBox returns a string, so convert to
         ' make sure they are treated as numbers, not strings
         if Cdbl(inputNumber) > Cdbl(largest) then
            largest = inputNumber
         end if

         outputNums = outputNums & " " & inputNumber
      end if
   loop
end if

if doneWithInput then
   MsgBox "Can't find the largest because you never entered any numbers!!" 
else
   MsgBox "Of all the numbers, " & outputNums & ", the largest is  " & largest
end if

