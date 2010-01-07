Attribute VB_Name = "modLuhn"
Option Explicit
 
'Luhn algorithm overview (from wikipedia.com)
'The formula verifies a number against its included check digit
', which is usually appended to a partial account number to
'generate the full account number.
'This account number must pass the following test:
 
'1. Starting with the rightmost digit (which is the check digit) and moving left,
'   double the value of every second digit. For any digits that thus become 10 or
'   more, add their digits together as if casting out nines. For example, 1111
'   becomes 2121, while 8763 becomes 7733 (from 2×6=12 ? 1+2=3 and 2×8=16 ? 1+6=7)
 
'2. Add all these digits together. For example, if 1111 becomes 2121, then 2+1+2+1
'   is 6; and 8763 becomes 7733, so 7+7+3+3 is 20.
'3. If the total ends in 0 (put another way, if the total modulus 10 is congruent to 0)
'   , then the number is valid according to the Luhn formula; else it is not valid. So,
'   1111 is not valid (as shown above, it comes out to 6), while 8763 is valid (as
'   shown above, it comes out to 20).

'e.g.
''to add a check digit to a number:
'NumberWithCheck = CStr(number) & LuhnCheckDigit(number)
'
''verify a number with check digit included:
'Valid = LuhnValid(NumberWithCheck)
'
Private Function LookupArray() As Integer()
 
'this function is used to provide an array used to lookup for 'doubled' digits
 
Dim aXL(9) As Integer
 
'array values in xL are for the index integer
'0 * 2 = 0 --> 0 = 0
'6*2 = 12 --> 1 + 2 = 3
'having this array available saves us from performing string conversions
'and math operations (just lookup by index)
aXL(0) = 0
aXL(1) = 2
aXL(2) = 4
aXL(3) = 6
aXL(4) = 8
aXL(5) = 1
aXL(6) = 3
aXL(7) = 5
aXL(8) = 7
aXL(9) = 9
 
LookupArray = aXL
 
End Function

Public Function LuhnCheckDigit(ByVal intStr As String) As String
 
'this function is used to return the check digit to be appended to a given number
Dim b() As Byte
Dim x As Integer
Dim xL() As Integer
Dim sD As Integer ' this holds sum of digits (as modified by Luhn algorithm)
Dim lD As Integer ' this is used to store checksum digit (10 - sD Mod 10)
 
'check for numeric input
If Not IsNumeric(intStr & ".0e0") Then
   LuhnCheckDigit = "X"
   Exit Function
End If
 
'create lookup array for loop
xL = LookupArray()
 
sD = 0
ReDim b(Len(intStr))
 
b = StrConv(StrReverse(intStr), vbFromUnicode)
 
'b(x) - 48 == faster way to get integer value from unicode byte value
'first digit (starting from right)is doubled/digits added because once
'check digit is appended this will be the second
For x = LBound(b) To UBound(b)
    If x Mod 2 = 0 Then
        sD = sD + xL(b(x) - 48)
    Else
        sD = sD + (b(x) - 48)
    End If
Next
 
lD = 10 - (sD Mod 10)
 
'we don't want to add 10, if lD calculates to 10 then we really want to add 0
If lD = 10 Then
    lD = 0
End If
 
'return string with check digit appended
LuhnCheckDigit = CStr(lD)
 
End Function

Public Function LuhnValid(ByVal intStr As String) As Boolean

'this function is used to check if a number entered is valid
Dim sD As Integer
Dim bl As Boolean
Dim b() As Byte
Dim x As Integer
Dim xL() As Integer
 
'check for numeric input
If Not IsNumeric(intStr & ".0e0") Then
   LuhnValid = False
   Exit Function
End If
 
'create lookup array for loop
xL = LookupArray()
 
ReDim b(Len(intStr))
 
b = StrConv(intStr, vbFromUnicode)
 
bl = False
sD = 0
 
'start with last digit, work towards first
For x = UBound(b) To LBound(b) Step -1
    If bl Then
        sD = sD + xL(b(x) - 48)
    Else
        sD = sD + b(x) - 48
    End If
    
    bl = Not (bl)
Next
 
LuhnValid = (sD Mod 10 = 0)
 
End Function
