<div align="center">

## RECURSIVE DESCENT PARSER FOR NUMERIC EXPRESSIONS


</div>

### Description

RECURSIVE DESCENT PARSER FOR NUMERIC EXPRESSIONS
 
### More Info
 
A string with a nummeric expression to parse

The result of the parsing as a string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sven\-Erik Dahlrot](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sven-erik-dahlrot.md)
**Level**          |Unknown
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sven-erik-dahlrot-recursive-descent-parser-for-numeric-expressions__1-985/archive/master.zip)





### Source Code

```
Option Explicit
Private Const NON_NUMERIC = 1
Private Const PARENTHESIS_EXPECTED = 2
Private Const NON_NUMERIC_DESCR = "Non numeric value"
Private Const PARENTESIS_DESCR = "Parenthesis expected"
Private Token As Variant   'Current token
'*********************************************************************
'*
'*   RECURSIVE DESCENT PARSER FOR NUMERIC EXPRESSIONS
'*
'* The function parses an string and returns the result.
'* If the string is empty the string "Empty" is returned.
'* If an error occurs the string "Error" is returned.
'*
'* The parser handles numerical expression with parentheses
'* unary operators + and -
'*
'* The following table gives the rules of precedence and associativity
'* for the operators:
'*
'* Operators on the same line have the same precedence and all operators
'* on a given line have higher precedence than those on the line below.
'*
'* -----------------------------------------------------------
'* Operators  Type of operation      Associativity
'*   ( )     Expression         Left to right
'*   + -     Unary            Right to left
'*   * /     Multiplication division   Left to right
'* -----------------------------------------------------------
'*
'* Sven-Erik Dahlrot 100260.1721@compuserve.com
'*
'*********************************************************************
Public Function EvaluateString(expr As String) As String
  Dim result As Variant
  Dim s1 As String
  Dim s2 As String   'White space characters
  Dim s3 As String   'Operators
  s2 = " " & vbTab   'White space characters
  s3 = "+-/*()"    'Operators
  On Error GoTo EvaluateString_Error
  Token = getToken(expr, s2, s3)  'Initialize
   EvalExp result         'Evaluate expression
   EvaluateString = result
Exit Function
EvaluateString_Error:
  EvaluateString = "Error"
End Function
'**** EVALUATE AN EXPRESSION
Private Function EvalExp(ByRef data As Variant)
  If Token <> vbNull And Token <> "" Then
    EvalExp2 data
  Else
    data = "Empty"
  End If
End Function
'* ADD OR SUBTRACT TERMS
Private Function EvalExp2(ByRef data As Variant)
  Dim op As String
  Dim tdata As Variant
  EvalExp3 data
  op = Token
  Do While op = "+" Or op = "-"
    Token = getToken(Null, "", "")
    EvalExp3 tdata
    Select Case op
      Case "+"
        data = Val(data) + Val(tdata)
      Case "-"
        data = Val(data) - Val(tdata)
    End Select
    op = Token
  Loop
End Function
'**** MULTIPLY OR DIVIDE FACTORS
Private Function EvalExp3(ByRef data As Variant)
  Dim op As String
  Dim tdata As Variant
  EvalExp4 data
  op = Token
  Do While op = "*" Or op = "/"
    Token = getToken(Null, "", "")
    EvalExp4 tdata
    Select Case op
      Case "*"
        data = Val(data) * Val(tdata)
      Case "/"
        data = Val(data) / Val(tdata)
    End Select
    op = Token
  Loop
End Function
'**** UNARY EXPRESSION
Private Function EvalExp4(ByRef data As Variant)
  Dim op As String
  If Token = "+" Or Token = "-" Then
    op = Token
    Token = getToken(Null, "", "")
  End If
  EvalExp5 data
  If op = "-" Then data = -Val(data)
 End Function
'**** PROCESS PARENTHESIZED EXPRESSION
Private Function EvalExp5(ByRef data As Variant)
  If Token = "(" Then
    Token = getToken(Null, "", "")
    EvalExp data           'Subexpression
    If Token <> ")" Then
      Err.Raise vbObjectError + PARENTHESIS_EXPECTED, "Expression parser", PARENTESIS_DESCR
    End If
    Token = getToken(Null, "", "")
  Else
    EvalAtom data
  End If
End Function
'* GET VALUE
Private Function EvalAtom(ByRef data As Variant)
  If IsNumeric(Token) Then
    data = Token
  Else
    Err.Raise vbObjectError + NON_NUMERIC, "Expression parser", NON_NUMERIC_DESCR
  End If
  Token = getToken(Null, "", "")
End Function
'****************************************************************
'*
'* Tokenizer function
'*
'*
'* When first time called s1 must contain the string to be tokenized
'* and s2, s3 the delimites and operators, otherwise s1 should be Null
'* and s2,s3 empty strings ""
'*
'* s2 contains delimiters
'* s3 contains operators that act as both delimiters and tokens
'*
'* If no delimiter can be found in s1 the whole local copy is returned
'* If there are no more tokens left, Null is returned
'* If one delimiter follows another, the empty string "" is returned
'*
'* s1 is declared as Variant, because VB doesn't like to assign Null to a string.
'*
'****************************************************************
Public Function getToken(s1 As Variant, s2 As String, s3 As String) As Variant
    Static stmp As Variant
    Static wspace As String
    Static operators As String
    Dim i As Integer
    Dim schr As String
    getToken = Null
    'Initialize first time calling
    If s1 <> "" Then
        stmp = s1
        wspace = s2
        operators = s3
    End If
    'Nothing left to tokenize!
    If VarType(stmp) = vbNull Then
        Exit Function
    End If
    'Loop until we find a delimiter or operator
    For i = 1 To Len(stmp)
      schr = Mid$(stmp, i, 1)
      If InStr(1, wspace, schr, vbTextCompare) = 0 Then    'White space
        If InStr(1, operators, schr, vbTextCompare) Then  'Operator
          getToken = Mid$(stmp, i, 1)            'Get it
          stmp = Mid(stmp, i + 1, Len(stmp))
          Exit Function
        Else                    'It is a numeric value
          getToken = ""
          schr = Mid$(stmp, i, 1)
          Do While (schr >= "0" And schr <= "9") Or schr = "."
            getToken = getToken & schr
            i = i + 1
            schr = Mid$(stmp, i, 1)
          Loop
          If IsNumeric(getToken) Then
            stmp = Mid$(stmp, i, Len(stmp))
            Exit Function
          End If
        End If
      End If
    Next i
    'No tokens was found, return whatever is left in stmp
    getToken = stmp
    stmp = Null
End Function
```

