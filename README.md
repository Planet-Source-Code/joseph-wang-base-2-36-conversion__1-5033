<div align="center">

## Base 2\-36 conversion


</div>

### Description

Converts a number from base 2~36 to a number of base 2~36
 
### More Info
 
use all lower case if passed base 10


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joseph Wang](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joseph-wang.md)
**Level**          |Beginner
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joseph-wang-base-2-36-conversion__1-5033/archive/master.zip)





### Source Code

```
Option Explicit
Private Function dec2any(number As Long, convertb As Integer) As String
  On Error Resume Next
  Dim num As Long
  Dim sum As String
  Dim carry As Long
  sum = ""
  num = number
  If convertb > 1 And convertb < 37 Then
    Do
      carry = num Mod convertb
      If carry > 9 Then
        sum = Chr$(carry + 87) + sum
      Else
        sum = carry & sum
      End If
      num = Int(num / convertb)
    Loop Until num = 0
    dec2any = sum
  Else
    dec2any = -1
  End If
End Function
Private Function any2dec(num As String, Optional numbase As Integer = 10) As Long
  On Error Resume Next
  Dim sum As Long
  Dim length As Integer
  Dim count As Integer
  Dim digit As String * 1
  length = Len(num)
  If length > 0 And numbase > 0 And numbase < 37 Then
    For count = 1 To length
      digit = Mid$(num, count, 1)
      If digit <= "9" Then
        sum = sum + digit * numbase ^ (length - count)
      Else
        sum = sum + (Asc(digit) - 87) * numbase ^ (length - count)
      End If
    Next count
    any2dec = sum
  Else
    any2dec = -1
  End If
End Function
Private Function any2any(num1 As String, num1base As Integer, convertbase As Integer) As String
  Dim answer As Long
  If num1base <> convertbase And num1base > 0 And convertbase > 0 _
    And num1base < 37 And convertbase < 37 Then
    answer = any2dec(num1, num1base)
    any2any = dec2any(answer, convertbase)
  Else
    any2any = -1
  End If
End Function
Private Sub Form_Load()
  ' example: converts letter z of base 36 to base 2 (binary)
  Me.Caption = any2any("z", 36, 2)
End Sub
```

