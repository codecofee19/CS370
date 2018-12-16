' File:  million.vbs
'
' 4. Suppose you start with $1 and double your money every day. How many days
' does it take to make more than $1000000?
 
Option Explicit                  ' must declare every variables before use

const MAXAMOUNT = 1000000
dim amount                       ' current amount of money
dim sum                          ' total sum of accumulated money
dim days                         ' number of days

' initialize 
amount = 1
sum = 0
days = 0

do until sum > MAXAMOUNT 
   sum = sum + amount
   amount = 2 * amount
   days = days + 1
loop

MsgBox "Start with $1 and double your money everyday. It takes  " _
       & days & "  days to make  " & sum & "  dollars" 

