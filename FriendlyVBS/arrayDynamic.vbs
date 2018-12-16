' File: arrayDynamic.vbs
' Demonstrate dynamic arrays. 

Option Explicit

dim A(), B()            ' A and B have no memory allocated yet

redim A(3), B(5)        ' allocate 4 elements for A, 6 elements for B

call fillArrayWithMult2(A, 0, UBound(A))
call fillArrayWithMult2(B, 0, UBound(B))
call printOneArray(A, "This is the A Array with 4 values")
call printOneArray(B, "This is the B Array with 6 values")

' redimension A to be of size 11 and display
' notice different memory is allocated and you lose all the values
' of the original array A
redim A(10)
call printOneArray(A, "This is the A Array with 10 values, redimensioned")

' fill again and display
call fillArrayWithMult2(A, 0, UBound(A))
call printOneArray(A, "This is the A Array with 10 values, redimensioned")

' preserve the values in B by using "preserve" keyword
redim preserve B(10)
call printOneArray(B, "This is the B Array with 10 values, preserved")

' can fill some or the rest of B
call fillArrayWithMult2(B, 7, 9)
call printOneArray(B, "This is the B Array with 10 values, preserved and more")


'----------------------------------------------------------------------------
' fillArrayWithMult2:
'
SUB fillArrayWithMult2(ByRef myArray, ByVal beginIndex, ByVal endIndex)
   dim i
   for i = beginIndex to endIndex
      myArray(i) = 2 * i 
   next
end SUB


'----------------------------------------------------------------------------
' printOneArray
' Takes in 2 input parameters, the array to display and a string which 
' is a description of what is displayed to be used as the MsgBox title. 
' It displays the contents of the array in message box titled description.
SUB printOneArray(ByRef myArray, ByRef description)
   dim output, i
   output = ""
   for i = 0 to UBound(myArray)
      output = output & "myArray(" & i & ") = " & myArray(i) & vbNewLine
   next

   ' concatenate lower and upper bound of the array
   output = output & "Array Lower Bound: " & LBound(myArray) & vbNewLine
   output = output & "Array Upper Bound: " & UBound(myArray) & vbNewLine

   MsgBox output, ,description
end SUB

