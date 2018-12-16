' File:  avg.vbs
'
' count how many data and compute average
 
Option Explicit                  ' must declare every variables before use

dim inputNumber                  ' input from user
dim count                        ' count data items
dim sum                          ' sum data items
dim avg                          ' average of data items

' initialize 
count = 0
sum = 0
avg = 0

' for every valid data item, sum and count it
do
   inputNumber = InputBox("Please enter one number, negative number to end")

   ' check for validity
   if not IsNumeric(inputNumber) then
      MsgBox "You must enter a number. Try Again.", vbOKOnly, "Invalid Input"
   else
      if inputNumber < 0 then
         exit do
      end if
      sum = sum + inputNumber
      count = count + 1
   end if
loop

' computer average and output
if count <> 0 then
   avg = sum/count
else
   avg = 0
end if
MsgBox "The number of items entered:  " & count & "  has an average of  " & avg 

