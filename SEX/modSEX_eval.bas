Attribute VB_Name = "modSEX_Eval"

Public Function Eval(expr As String)

    Dim value As Variant, operand As String
    Dim pos As Integer
    
    pos = 1

    On Error Resume Next
    
    Do Until pos > Len(expr)

    Select Case Mid$(expr, pos, 4)
        Case "like"
            operand = Mid$(expr, pos, 4)
            pos = pos + 4
    End Select
    
    
    Select Case Mid$(expr, pos, 3)
        Case "not", "or ", "and", "xor", "eqv", "imp"
            operand = Mid$(expr, pos, 3)
            pos = pos + 3
    End Select

    
'    MsgBox "expr:~" & expr & "~"
    
    Select Case Mid$(expr, pos, 1)
        Case " "
            pos = pos + 1
        Case "&", "+", "-", "*", "/", "\", "^"
            operand = Mid$(expr, pos, 1)
            pos = pos + 1
        Case ">", "<", "=":
            Select Case Mid$(expr, pos + 1, 1)
                Case "<", ">", "="
                    operand = Mid$(expr, pos, 2)
                    pos = pos + 1
                Case Else
                    operand = Mid$(expr, pos, 1)
            End Select
            pos = pos + 1
        Case Else
            Dim x As String
            'x = value & "~like~" & Token(expr, pos)
            Select Case operand
                Case "": value = Token(expr, pos)
                Case "&":   Eval = Eval & value
                            value = Token(expr, pos)
                Case "+":   Eval = Eval + value
                            value = Token(expr, pos)
                Case "-":
                    If IsNumeric(value) Then
                        Eval = Eval + value
                        value = -Token(expr, pos)
                    Else
                        value = Replace$(value, Token(expr, pos), "")
                    End If
                Case "^":
                    If IsNumeric(value) Then
                        value = value ^ Token(expr, pos)
                    Else
                        value = strrepeat(CStr(value), Val(Token(expr, pos)))
                    End If
                Case "*":   value = value * Token(expr, pos)
                Case "/":   value = value / Token(expr, pos)
                Case "\":   value = value \ Token(expr, pos)
                Case "not": Eval = Eval + value
                            value = Not Token(expr, pos)
                Case "and", "&&": value = value And Token(expr, pos)
                Case "or ", "||": value = value Or Token(expr, pos)
                Case "xor", "^^": value = value Xor Token(expr, pos)
                Case "eqv": value = value Eqv Token(expr, pos)
                Case "imp": value = value Imp Token(expr, pos)
                Case "liek", "like", "matches": value = value Like Token(expr, pos)
                Case "=", "==", "eq":
                    Dim tk
                    tk = Token(expr, pos)
                    value = Left$(CStr(value), Len(CStr(value))) = Left$(CStr(tk), Len(CStr(tk)))
                Case ">": value = value > Token(expr, pos)
                Case "<": value = value < Token(expr, pos)
                Case ">=", "=>", "ge": value = value >= Token(expr, pos)
                Case "<=", "=<", "le": value = value <= Token(expr, pos)
                Case "<>", "!=", "ne": value = value <> Token(expr, pos)
            End Select
        End Select
    Loop
    Eval = Eval + value
End Function


Function Token(expr, pos)
    Dim char As String, value As String, fn As String
    Dim es As Integer, pl As Integer
    Const QUOTE As String = """"
    
    On Error Resume Next

    Do Until pos > Len(expr)
        char = Mid$(expr, pos, 1)
        Select Case char
            Case "&", "+", "-", "/", "\", "*", "^", " ", ">", "<", "=": Exit Do
            Case "("
            pl = 1
            pos = pos + 1
            es = pos
            Do Until pl = 0 Or pos > Len(expr)
                char = Mid$(expr, pos, 1)

                Select Case char
                    Case "(": pl = pl + 1
                    Case ")": pl = pl - 1
                End Select
            pos = pos + 1
        Loop
        value = Mid$(expr, es, pos - es - 1)
        fn = LCase(Token)

        Select Case fn
            Case Else: Token = Eval(value)
        End Select
    Exit Do
    Case QUOTE
    pl = 1
    pos = pos + 1
    es = pos


    Do Until pl = 0 Or pos > Len(expr)
        char = Mid$(expr, pos, 1)
        pos = pos + 1


        If StrComp(char, QUOTE) = True Then
            If StrComp(Mid$(expr, pos, 1), QUOTE) = True Then
                value = value & QUOTE
                pos = pos + 1
            Else
                Exit Do
            End If
        Else
            value = value & char
        End If
    Loop
    Token = value
    Exit Do
    Case Else
    Token = Token & char
    pos = pos + 1
End Select
Loop
    If IsNumeric(Token) Then
        Token = Val(Token)
    ElseIf IsDate(Token) Then
        Token = CDate(Token)
    End If
End Function

