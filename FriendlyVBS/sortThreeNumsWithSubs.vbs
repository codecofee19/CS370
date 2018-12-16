' File: sortThreeNumsWithSubs.vbs
' Sort three Numbers

   Option Explicit

   dim num1, num2, num3            ' used for input
   dim largeNum                   ' variable for largest number
   dim middleNum                  ' variable for second largest number
   dim smallNum                   ' variable for smallest number

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

      call compareAndSwap(largeNum, middleNum)
      call compareAndSwap(largeNum, smallNum)
      call compareAndSwap(middleNum, smallNum)

      MsgBox "The numbers sorted: " & smallNum & "  " & middleNum & "  " _
             & largeNum, vbOKOnly, "Sorted Numbers"  
   else
      MsgBox "You must enter three numbers! Try Again", vbOKOnly, _
             "Invalid Input"
   end if

   '----------------------------------------------------------------------
   ' compareAndSwap
   ' Compare num1 to num2 and swap if they are out of order.
   ' Result is that at the end of the routine, num1 is always less than num2.
   sub compareAndSwap(num1, num2)
      dim temp
      if num1 < num2 then
         temp = num1
         num1 = num2
         num2 = temp
      end if
   end sub

