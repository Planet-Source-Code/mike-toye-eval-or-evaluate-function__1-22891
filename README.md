<div align="center">

## Eval or Evaluate function


</div>

### Description

evaluate basic math strings function
 
### More Info
 
math formula, eg; (6+(4.322^6)+(9*7)

result

no bad char validation!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Toye](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-toye.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-toye-eval-or-evaluate-function__1-22891/archive/master.zip)





### Source Code

```
Function Eval(sin As String) As Double
Dim bAreThereBrackets As Boolean
Dim x As Double, y As Double, z As Double
Dim L2R As Integer
Dim sLeft As String, sMid As String, sRight As String
Dim dStack As Double
Dim sPrevOp As String
Dim bInnerFound As Boolean
  sin = IIf(InStr(sin, " ") > 0, RemoveAllSpaces(sin), sin)
  If InStr(sin, "(") Then
  'work from left to right. find the inner most
  'brackets and resolve them into the string, eg;
  '(6+7+(6/3)) becomes (6+7+2)
    L2R = 1
    While InStr(sin, "(") > 0
      'inner loop
      bInnerFound = False
      Do
        x = InStr(L2R, sin, "(")
        y = InStr(x + 1, sin, "(")
        z = InStr(x + 1, sin, ")")
        If y = 0 Then
          L2R = x
          bInnerFound = True
        Else
          If y < z Then
            L2R = y
          Else
            L2R = x
            bInnerFound = True
          End If
        End If
      Loop Until bInnerFound
      x = InStr(L2R, sin, ")")
      sin = Left(sin, L2R - 1) & CStr(Eval(Mid(sin, L2R + 1, x - L2R - 1))) & Mid(sin, x + 1)
      Debug.Print sin
    Wend
    Eval = CDbl(IIf(IsNumeric(sin), sin, Eval(sin)))
  Else
    dStack = 0
    sLeft = ""
    sPrevOp = ""
    For L2R = 1 To Len(sin)
      If Not IsNumeric(Mid(sin, L2R, 1)) And Mid(sin, L2R, 1) <> "." Then
        'we have an operator
        If dStack = 0 Then
          dStack = CDbl(sLeft)
        Else
          dStack = ASMD(dStack, sLeft, sPrevOp)
        End If
        sLeft = ""
        sPrevOp = Mid(sin, L2R, 1)
      Else
        'carry on extracting the current number
        sLeft = sLeft & Mid(sin, L2R, 1)
      End If
    Next L2R
    If sLeft > "" Then
      dStack = ASMD(dStack, sLeft, sPrevOp)
    End If
    Eval = dStack
  End If
End Function
Function RemoveAllSpaces(sin As String) As String
Dim x As Integer
  RemoveAllSpaces = ""
  For x = 1 To Len(sin)
    If Mid(sin, x, 1) <> " " Then
      RemoveAllSpaces = RemoveAllSpaces & Mid(sin, x, 1)
    End If
  Next x
End Function
Function ASMD(dIn As Double, sin As String, sOP As String) As Double
  Select Case sOP
    Case "+"
      ASMD = dIn + CDbl(sin)
    Case "-"
      ASMD = dIn - CDbl(sin)
    Case "*"
      ASMD = dIn * CDbl(sin)
    Case "/"
      ASMD = dIn / CDbl(sin)
    Case "^"
      ASMD = dIn ^ CDbl(sin)
    Case Else
      ASMD = 0
  End Select
End Function
```

