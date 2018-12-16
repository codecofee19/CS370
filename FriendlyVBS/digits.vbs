' Based on code by Kelvin Sung
' File:  digits.vbs
'
' example demonstrates mod (modulus, finding remainder) and integer div (\)
' by breaking a five-digit number into individual digits
 
Option Explicit                  ' must declare every variables before use

dim inputNumber                  ' input from user
dim origNumber                   ' save the inputNumber
dim onesDigit, tenDigit, hunDigit, thouDigit, tenThouDigit

do
   inputNumber = InputBox("Please Enter a 5 digit number, zero to end")

   ' check for validity
   if not IsNumeric(inputNumber) then
      MsgBox "You must enter a number. Try Again.", vbOKOnly, "Invalid Input"
   else
      ' check to see if the loop should be terminated
      if inputNumber = 0 then
         exit do
      end if

      ' have obtained a good number, demonstrate different kinds of division
      origNumber = inputNumber
      MsgBox "Using division (/):  " & origNumber & "/100 is  " _
              & origNumber/100
      MsgBox "Using integer division (\):  " & origNumber & "\100 is  " _
              & origNumber\100

      ' extract digits
      tenThouDigit = inputNumber\10000

      inputNumber = inputNumber mod 10000
      thouDigit = inputNumber\1000

      inputNumber = inputNumber mod 1000
      hunDigit = inputNumber\100

      inputNumber = inputNumber mod 100
      tenDigit = inputNumber\10

      onesDigit = inputNumber mod 10

      MsgBox "The digits of  " & origNumber & "  are  " & tenThouDigit & "  " _
              & thouDigit & "  " & hunDigit & "  " & tenDigit & "  " & onesDigit
   end if
loop

MsgBox "All done!"

