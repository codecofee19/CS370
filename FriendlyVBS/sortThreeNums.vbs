' File: sortThreeNums.vbs
' Sort three Numbers

   Option Explicit

   dim num1, num2, num3            ' used for input
   dim largeNum                   ' variable for largest number
   dim middleNum                  ' variable for second largest number
   dim smallNum                   ' variable for smallest number
   dim temp

   num1 = InputBox("Please enter the first number")
   num2 = InputBox("Please enter the second number")
   num3 = InputBox("Please enter the third number")

   ' check to make sure all input are "numeric" - integer/float/double
   if IsNumeric(num1) and IsNumeric(num2) and IsNumeric(num3) then
      MsgBox "You have entered: " & num1 & " " & num2 & " " _
             & num3, vbOKOnly, "Entered Values"

      largeNum = CDbl(num1) 
      middleNum = CDbl(num2)
      smallNum = CDbl(num3)

      if largeNum < middleNum then
         temp = largeNum
         largeNum = middleNum
         middleNum = temp
      end if

      if largeNum < smallNum then
         temp = largeNum
         largeNum = smallNum
         smallNum = temp
      end if

      if middleNum < smallNum then
         temp = middleNum
         middleNum = smallNum
         smallNum = temp
      end if

      MsgBox "The numbers sorted: " & smallNum & "  " & middleNum & "  " _
             & largeNum, vbOKOnly, "Sorted Numbers"  
   else
      MsgBox "You must enter three numbers! Try Again", vbOKOnly, _
             "Invalid Input"
   end if

