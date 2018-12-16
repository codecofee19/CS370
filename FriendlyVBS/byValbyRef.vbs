' File: byValbyRef.vbs
' demonstrate difference between byVal parameter and byRef parameter

option explicit

dim a, b, c, d
a = 2
b = 3
c = 2
d = 3

call byValSub(a, b)
call byRefSub(c, d)

MsgBox "a = " & a & "  b = " & b & "  c = " & c & "  d = " & d

'----------------------------------------------------------------------------
' byValSub
' Parameters that are pass by value make a copy of the value passed to
' the parameter and if the parameter is changed, it is only changed
' locally, meaning within the subprogram.
SUB byValSub(ByVal num1, ByVal num2)
   num1 = 10
   num2 = 20
end SUB

'----------------------------------------------------------------------------
' byRefSub
' Parameters that are pass by reference are sent the memory address of the
' sender. If the parameter is changed, it is changed at the sending location.
SUB byRefSub(ByRef num1, ByRef num2)
   num1 = 10
   num2 = 20
end SUB

