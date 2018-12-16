' 9. Given some integer number as input, print out a solid square in 
' asterisks the size of that number. E.g., input is 5, output is
' *****
' *****
' *****
' *****
' *****
 
Option Explicit                  ' must declare every variables before use

dim inputNumber                  ' input from user
dim i, j
dim output                           

do
   ' start with no characters in the output string, called the empty string
   output = ""

   inputNumber = InputBox("Please enter one positive number, 'done' to end")

   ' check for validity of input
   if not IsNumeric(inputNumber) then
      if inputNumber = "done" then
         exit do
      else
         MsgBox "You must enter a number. Try Again.", vbOKOnly, "Invalid Input"
      end if

   ' can't have a negative number of asterisks in a box
   elseif inputNumber < 0 then
      MsgBox "You must enter a positive number. Try Again.", vbOKOnly, _
             "Invalid Input"

   else                                     ' have valid inputNumber
      ' i loop yields different lines of asterisks
      for i = 1 to inputNumber

         ' j loop gives all the asterisks in one line
         for j = 1 to inputNumber
            output = output & "*"
         next

         ' after the j loop asterisks are concatenated, need NewLine in output
         output = output & vbNewLine
      next
      MsgBox output
   end if
loop

