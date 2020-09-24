'* ********************
'*     Alias Module
'* ********************

Public serverID As Integer
Public rootEngine      As clsSSE_Main

'* Constants
Private Const BEGIN_FUNCTION = "["
Private Const END_FUNCTION = "]"
Private Const BEGIN_SGROUP = "("
Private Const END_SGROUP = ")"
Private Const COMMENT_CHAR = ";"
Private Const COMMENT_ML_CHAR = "`"
Private Const PARAM_DELIM = " "
Private Const BEGIN_VARLOCAL = "$"
Private Const BEGIN_VARGLOBAL = "%"
Private Const ESCAPE_CHAR = "\"
Private Const ARRAY_CHAR = "^"


Private currentLine As Integer

'* label stack
Private stack_label(25) As typLabel
Private stack_labelcount    As Integer
Private Type typLabel
    strName As String
    nLine   As Integer
End Type

'* if stack
Private stack_if(25)      As typIf
Private stack_ifcount   As Integer
Private Type typIf
    bTrueYet    As Boolean
    bNoEval     As Boolean
End Type

'* while loop stack
Private stack_while(15)       As typWhileLoop
Private stack_whilecount    As Integer
Private Type typWhileLoop
    bIsFalse    As Boolean
    iReturnLine As Integer
    iLastLine   As Integer
    bInit       As Boolean
End Type

'* loop stack
Private stack_loop(15)    As typLoop
Private stack_loopcount As Integer
Private Type typLoop
    current As Long
    total   As Long
    linenum As Integer
    varname As String
End Type

'* Args
Private args(75)    As String
Private ArgCount    As Integer

'* Var Type
Private variableType As Integer
Private Enum varType
    LOCAL_
    GLOBAL_
End Enum

Public bInComment       As Boolean
Public bGotoNextLine    As Boolean
Public returnValue     As String

'* Alias Information
Private strName         As String
Private AliasType       As enAliasType
Private strExtraParams  As String
Private Enum enAliasType
    at_ALIAS
    at_EVENT
End Enum

'* Local variables
Private variables()     As typVariable
Private varCount        As Integer
Private Enum enumVarType
    xString
    xInteger
    xReal
End Enum
Private Type typVariable
    Name    As String
    value() As String
    type    As enumVarType
End Type

'* Code storage
Private strCode()       As String
Private intCodeLines    As Integer

'* Call stack type
Private Type typAliasCall
    ArgCount    As Integer
    args()      As String
    bQuote      As Boolean
End Type

'* call stack
Private stack_calls(25)       As typAliasCall
Private stack_callcount     As Integer


'* Execution multiplier
Private em_total    As Integer
Private em_current  As Integer

Public Sub AddCodeLine(strCodeLine As String)
    intCodeLines = intCodeLines + 1
    ReDim Preserve strCode(1 To intCodeLines) As String
    strCode(intCodeLines) = strCodeLine
   
    If Left(strCodeLine, 1) = ":" Or Right(strCodeLine, 1) = ":" Then
        Dim strLabel As String
        strLabel = Replace(strCodeLine, ":", "")
        AddLabel strLabel, intCodeLines
    End If
   
End Sub


Private Sub AddLabel(strName As String, nLine As Integer)
    stack_labelcount = stack_labelcount + 1
    If stack_labelcount > 25 Then Exit Sub
   
    stack_label(stack_labelcount).strName = strName
    stack_label(stack_labelcount).nLine = nLine
End Sub


Private Sub AddLocalVar(strVarNameX As String, strValue As String, Optional arrayElement As Integer = 0, Optional tvType As Integer = 0)
    varCount = varCount + 1
    ReDim Preserve variables(1 To varCount) As typVariable
    variables(varCount).Name = strVarNameX
    ReDim variables(varCount).value(arrayElement) As String
    variables(varCount).value(arrayElement) = strValue
    variables(varCount).type = LOCAL_
End Sub

Private Sub CleanUp()
    On Error Resume Next
   
'    ReDim stack_calls(0) As typAliasCall
    stack_callcount = 0
    ReDim stack_calls(0).args(50) As String
    stack_calls(0).ArgCount = 0
   
End Sub



Public Sub CopyAlias(ByRef oldAliasClass As clsSSE_Alias, ByRef newAliasClass As clsSSE_Alias)
    'Set newAliasClass = oldAliasClass
   
    Set newAliasClass = New clsSSE_Alias
    Dim I As Integer
    For I = 1 To intCodeLines
        newAliasClass.AddCodeLine strCode(I)
    Next I
   
    Set newAliasClass.rootEngine = rootEngine
End Sub

Private Sub CopyToArgs(ByRef paramlist() As String)
    If UBound(paramlist) = 0 Then
        ArgCount = 0
        Exit Sub
    End If
   
    ArgCount = UBound(paramlist) + 1
       
    Dim I As Integer
    For I = LBound(paramlist) To UBound(paramlist)
        args(I) = paramlist(I)
    Next I
End Sub




Public Function Execute(ByRef paramlist() As String) As String
   
    Dim bFinished As Boolean
    currentLine = 1
    bGotoNextLine = True
    bInComment = False
   
    CopyToArgs paramlist
    stack_loopcount = 0
    stack_whilecount = 0
    stack_ifcount = 0
    stack_labelcount = 0
    AddLocalVar "me", windowStatus(serverID).strCurNick
   
    Do
        Call CleanUp
        If currentLine = 0 Then
            Execute = ""
            Exit Function
        End If
        returnValue = ExecuteLine(strCode(currentLine))
               
        If bGotoNextLine Then
            currentLine = currentLine + 1
        End If
    Loop Until currentLine > intCodeLines Or bGotoNextLine = False
   
    Execute = returnValue
End Function

Public Function ExecuteLine(ByRef strLine As String)
    '*
    '* This is the heart of SEX
    '*


    Dim strBuffer As String, I As Integer, curChar As String, prevChar As String
    Dim inVariable As Boolean, bEscape As Boolean, strLen As String, bWhiteSpace As Boolean
    strLen = Len(strLine)
   
    '* Clean stacks
    CleanUp
   
    I = 1
   
    If stack_whilecount <= UBound(stack_while) Then
        If stack_while(stack_whilecount).bIsFalse And Left(strLine, 3) <> "end" Then
            Exit Function
        End If
    End If
   
    For I = 1 To strLen
        curChar = Mid(strLine, I, 1)
       
        If prevChar <> ESCAPE_CHAR Then
            If curChar = COMMENT_ML_CHAR Then
                If stack_calls(stack_callcount).bQuote Then
                    '* Else append to other shit..
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & curChar
                   
                    GoTo nextchar
                End If
               
                bWhiteSpace = False
                bInComment = Not bInComment
            End If
            If bInComment Or curChar = COMMENT_ML_CHAR Then GoTo nextchar
       
        End If
       
        '* Append escape char
        If bEscape Then
            bWhiteSpace = False
            'if invar, end var
            If inVariable Then
                '* change this...
                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                strBuffer = ""
                inVariable = False
            End If
            Dim strEscapeChar As String
           
            strEscapeChar = GetEscapeChar(curChar)
           
            stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & strEscapeChar
            bEscape = False
           
        '* Comments, end
        ElseIf curChar = COMMENT_CHAR Then
            If stack_calls(stack_callcount).bQuote Then
                '* Else append to other shit..
                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & curChar
                GoTo nextchar
            End If
       
            bWhiteSpace = False
            '* if invar, end var
            If inVariable Then
                '* change this...
                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetLocalVar(strBuffer)

                strBuffer = ""
                inVariable = False
            End If
           
            GoTo finish
        '* Other stuff..parse
        Else
       
        Select Case curChar
            Case ESCAPE_CHAR
                bWhiteSpace = False
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                    strBuffer = ""
                    inVariable = False
                End If
               
                If bEscape Then
                    '* get var, add escape
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & "\"
               
                End If
               
                bEscape = True
            Case BEGIN_FUNCTION
                bWhiteSpace = False
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                    strBuffer = ""
                    inVariable = False
                End If
           
                stack_callcount = stack_callcount + 1
                'ReDim Preserve stack_calls(stack_callcount) As typAliasCall
                ReDim stack_calls(stack_callcount).args(50)
                stack_calls(stack_callcount).ArgCount = 0
               
            Case END_FUNCTION
                bWhiteSpace = False
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                       
                    strBuffer = ""
                    inVariable = False
                End If
               
                Dim strReturnVal As String
                If stack_callcount = 0 Then GoTo nextchar
               
                If Left(CStr(stack_calls(stack_callcount).args(0)), 1) = "$" Then
                    If stack_calls(stack_callcount).ArgCount < 1 Then
                        strReturnVal = ""
                        GoTo skipFunction
                    End If
                   
                    strReturnVal = GetLocalVar(Mid(CStr(stack_calls(stack_callcount).args(0)), 2), CInt(stack_calls(stack_callcount).args(1)))
                    GoTo skipFunction
                End If
               
                           
                Set rootEngine.cChildAlias = Me
               
                ReDim Preserve stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount)
                strReturnVal = rootEngine.ExecuteAlias(CStr(stack_calls(stack_callcount).args(0)), stack_calls(stack_callcount).args, serverID)
skipFunction:
                stack_calls(stack_callcount).ArgCount = 0
                stack_callcount = stack_callcount - 1
                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & strReturnVal
               
            Case PARAM_DELIM
                If bWhiteSpace = True Then GoTo nextchar
               
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                   
                    strBuffer = ""
                    inVariable = False
                End If
               
                If stack_calls(stack_callcount).bQuote Then
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & PARAM_DELIM
                Else
                    stack_calls(stack_callcount).ArgCount = stack_calls(stack_callcount).ArgCount + 1
                    'ReDim Preserve stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount)
                End If
            Case """"
                bWhiteSpace = False
               
                If (stack_calls(stack_callcount).args(0) = "eval") Or _
                    (stack_calls(stack_callcount).args(0) = "elseif") Or _
                    (stack_calls(stack_callcount).args(0) = "while") Or _
                    (stack_calls(stack_callcount).args(0) = "if") Then
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & curChar
                Else
                    stack_calls(stack_callcount).bQuote = _
                        Not stack_calls(stack_callcount).bQuote
                End If

            Case BEGIN_VARLOCAL
                If I = 1 Then
                    'stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & curChar
                       
                    GoTo nextchar
                End If

           
                bWhiteSpace = False
                If prevChar = BEGIN_FUNCTION Then
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & "$"
                    GoTo nextchar
                End If
               
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                   
                    strBuffer = ""
                    inVariable = False
                End If
                variableType = 1
                inVariable = True
                strBuffer = ""
            Case BEGIN_VARGLOBAL
                bWhiteSpace = False
                variableType = 2
                inVariable = True
                strBuffer = ""
            Case Else
                bWhiteSpace = False
                If inVariable Then
                    '* In variable, append to strbuffer
                    Select Case curChar
                        Case "?", ",", ".", "/", "!", "@", "(", ")", "=", "+", "&", "^", "%", "*", "/"
                            '* change this... no, dont..DONT! PLEASE!
                            stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                           
                            strBuffer = ""
                            inVariable = False

                            stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & curChar
                       
                        Case "}", ">"
                            strBuffer = strBuffer & curChar
                            '* change this...
                            stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                                stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)

                            strBuffer = ""
                            inVariable = False
                        Case Else
                            strBuffer = strBuffer & curChar
                    End Select
                Else
                    '* Else append to other shit..
                   
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & curChar
                   
                End If
        End Select
                       
        End If
nextchar:
   
        prevChar = curChar
       
    Next I
   
    If strBuffer <> "" Then
        If inVariable Then
            '* change this...
            stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) = stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)

            strBuffer = ""
            inVariable = False
        End If
    End If
   
finish:

    Dim strReturn As String
'    MsgBox stack_calls(0).args(0) & "~" & stack_calls(0).args(1) & "~~calls~" & strCode(currentLine)
   
    If stack_calls(0).args(0) = "return" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        '*
        '* return statement
        '*
        ExecuteLine = JoinArray(stack_calls(0).args, " ", 2, stack_calls(0).ArgCount + 1)
        bGotoNextLine = False
        Exit Function
    '*
    '* Goto label
    ElseIf stack_calls(0).args(0) = "goto" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        If UBound(stack_calls(0).args) < 1 Then Exit Function
       
        GotoLabel CStr(stack_calls(0).args(1))
           
    '*
    '* SET 
    ElseIf stack_calls(0).args(0) = "set" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        '*
        '* set a variable
        '*
        If UBound(stack_calls(0).args) < 2 Then Exit Function
       
        If CStr(stack_calls(0).args(1)) Like "*:*" Then
            SetLocalVar CStr(Mid(CStr(stack_calls(0).args(1)), 1, InStr(CStr(stack_calls(0).args(1)), ":") - 1)), JoinArray(stack_calls(0).args, " ", 3, stack_calls(0).ArgCount + 1), Mid(CStr(stack_calls(0).args(1)), InStr(CStr(stack_calls(0).args(1)), ":") + 1)
        Else
            SetLocalVar CStr(stack_calls(0).args(1)), JoinArray(stack_calls(0).args, " ", 3, stack_calls(0).ArgCount + 1)
        End If
   
    '*
    '* set variable
    '*
    ElseIf stack_calls(0).args(1) = "=" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        If UBound(stack_calls(0).args) < 2 Then Exit Function
       
        If CStr(stack_calls(0).args(0)) Like "*:*" Then
            SetLocalVar CStr(Mid(CStr(stack_calls(0).args(0)), 1, InStr(CStr(stack_calls(0).args(0)), ":") - 1)), JoinArray(stack_calls(0).args, " ", 3, stack_calls(0).ArgCount + 1), Mid(CStr(stack_calls(0).args(1)), InStr(CStr(stack_calls(0).args(1)), ":") + 1)
        Else
            SetLocalVar CStr(stack_calls(0).args(0)), JoinArray(stack_calls(0).args, " ", 3, stack_calls(0).ArgCount + 1)
        End If
   
    '*
    '* WHILE LOOP
    '*
    ElseIf stack_calls(0).args(0) = "while" Then
        If stack_whilecount <= UBound(stack_while) Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
       
        If UBound(stack_calls(0).args) < 1 Then Exit Function
       
        stack_whilecount = stack_whilecount + 1
        If stack_whilecount > UBound(stack_while) Then
            'ReDim Preserve stack_while(stack_whilecount) As typWhileLoop
        End If
       
        If stack_while(stack_whilecount).bInit = False Then
            stack_while(stack_whilecount).iReturnLine = currentLine
            stack_while(stack_whilecount).bInit = True
            stack_while(stack_whilecount).bIsFalse = False
        End If
       
        Dim whileStat
        whileStat = Eval(JoinArrayE(stack_calls(0).args, " ", 2, stack_calls(0).ArgCount + 1))
       
        If whileStat <> 0 Then
            stack_while(stack_whilecount).bIsFalse = False
            stack_while(stack_whilecount).iReturnLine = currentLine
        Else
            stack_while(stack_whilecount).bIsFalse = True
            If stack_while(stack_whilecount).iLastLine <> -1 Then
                currentLine = stack_while(stack_whilecount).iLastLine
                Exit Function
            End If
        End If
       
    ElseIf stack_calls(0).args(0) = "loop" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
       
        If UBound(stack_calls(0).args) < 2 Then Exit Function
       
        stack_loopcount = stack_loopcount + 1
        'ReDim Preserve stack_loop(stack_loopcount) As typLoop
   
        stack_loop(stack_loopcount).current = 1
        stack_loop(stack_loopcount).total = stack_calls(0).args(2)
        stack_loop(stack_loopcount).linenum = currentLine '+ 1
        stack_loop(stack_loopcount).varname = stack_calls(0).args(1)
        SetLocalVar stack_loop(stack_loopcount).varname, "1"

    ElseIf stack_calls(0).args(0) = "if" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        If UBound(stack_calls(0).args) < 1 Then Exit Function
       
        stack_ifcount = stack_ifcount + 1
        'ReDim Preserve stack_if(stack_ifcount) As typIf
       
        Dim ifStat
        ifStat = Eval(JoinArrayE(stack_calls(0).args, " ", 2, stack_calls(0).ArgCount + 1))
      
        If ifStat = 0 Then
            stack_if(stack_ifcount).bTrueYet = False
        Else
            stack_if(stack_ifcount).bTrueYet = True
        End If
    ElseIf stack_calls(0).args(0) = "elseif" Then
        '* if in while loop, damnit!
'        MsgBox "BLAH~" & strCode(currentLine)
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
       
        If stack_if(stack_ifcount).bNoEval = True Then Exit Function
        If stack_if(stack_ifcount).bTrueYet = True Then
            stack_if(stack_ifcount).bTrueYet = False
            stack_if(stack_ifcount).bNoEval = True
            Exit Function
        End If
       
        If UBound(stack_calls(0).args) < 1 Then Exit Function
       
        Dim ifStat2
'        MsgBox stack_calls(0).ArgCount & "~argcount~" & strCode(currentLine)
        ifStat2 = Eval(JoinArrayE(stack_calls(0).args, " ", 2, stack_calls(0).ArgCount + 1))
       
        If ifStat2 = 0 Then
            stack_if(stack_ifcount).bTrueYet = False
        Else
            stack_if(stack_ifcount).bTrueYet = True
        End If
    ElseIf stack_calls(0).args(0) = "else" Then
        '* if in while loop, damnit!
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
       
        If stack_if(stack_ifcount).bNoEval = True Then Exit Function
        If stack_if(stack_ifcount).bTrueYet = True Then
            stack_if(stack_ifcount).bTrueYet = False
            stack_if(stack_ifcount).bNoEval = True
            Exit Function
        End If
       
        stack_if(stack_ifcount).bTrueYet = True
    '*
    '* init variables
    '*
    ElseIf stack_calls(0).args(0) = "init" Then
        Dim iLoop As Integer
        For iLoop = 1 To stack_calls(0).ArgCount - 1
            AddLocalVar CStr(stack_calls(0).args(iLoop)), "0"
        Next iLoop
    '*
    '* END
    '*
    ElseIf stack_calls(0).args(0) = "end" Then
   
        If stack_calls(0).ArgCount < 1 Then Exit Function
       
        Select Case stack_calls(0).args(1)
            Case "loop"
                If stack_ifcount > 0 Then
                    If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
                End If

                stack_loop(stack_loopcount).current = stack_loop(stack_loopcount).current + 1
                If stack_loop(stack_loopcount).current > stack_loop(stack_loopcount).total Then
                    stack_loopcount = stack_loopcount - 1
                Else
                    currentLine = stack_loop(stack_loopcount).linenum
                End If
                SetLocalVar stack_loop(stack_loopcount).varname, CStr(stack_loop(stack_loopcount).current)

            Case "if"
                stack_ifcount = stack_ifcount - 1
            Case "while"
                If stack_while(stack_whilecount).bIsFalse Then
                    stack_while(stack_whilecount).bInit = False
                    stack_while(stack_whilecount).bIsFalse = False
                    stack_whilecount = stack_whilecount - 1
                    stack_while(stack_whilecount).iLastLine = -1
                Else
                    stack_whilecount = stack_whilecount - 1
                    currentLine = stack_while(stack_whilecount + 1).iReturnLine - 1
                    stack_while(stack_whilecount + 1).iLastLine = currentLine
                End If
        End Select
    ElseIf Right(stack_calls(0).args(0), 1) = "*" Then
    Else
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
       
        '* if in while loop, damnit!
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
       
        '*
        '* nothing special, call the alias
        '*
        Dim jj As Integer, paramarr() As String
        ReDim paramarr(stack_calls(0).ArgCount)
        For jj = 0 To stack_calls(0).ArgCount
            paramarr(jj) = stack_calls(0).args(jj)
        Next jj
       
        strReturn = rootEngine.ExecuteAlias(CStr(stack_calls(0).args(0)), paramarr, serverID)
    End If
   
    ExecuteLine = ""

End Function


Private Function GetEscapeChar(strChar As String) As String
    Select Case strChar
        Case "n"    'new line
            GetEscapeChar = vbCrLf
        Case "t"    'tab
            GetEscapeChar = Chr(9)
        Case "0"    'char 0
            GetEscapeChar = Chr(0)
        Case "1"    'char 1
            GetEscapeChar = Chr(1)
        Case "c", "k" 'color chars
            GetEscapeChar = Chr(3)
        Case "r"    'reverse char
            GetEscapeChar = Chr(22)
        Case "b"    'bold char
            GetEscapeChar = Chr(2)
        Case "u"    'underline char
            GetEscapeChar = Chr(31)
       
        Case Else
            GetEscapeChar = strChar
    End Select
End Function


Public Function GetLocalVar(strVarName As String, Optional arrayElement As Integer = 0) As String
    Dim I As Integer
    For I = 1 To varCount
        'MsgBox strVarName & "~" & variables(i).Name
        If variables(I).Name = strVarName Then
            If arrayElement > UBound(variables(I).value) Then
                GetLocalVar = ""
            Else
                GetLocalVar = variables(I).value(arrayElement)
            End If
            Exit Function
        End If
    Next I
    GetLocalVar = ""
End Function

Public Function GetName() As String
    GetName = strName
End Function



Public Function GetType() As Integer
    If AliasType = at_ALIAS Then
        GetType = 0
    Else
        GetType = 1
    End If
End Function

Private Function GetVar(strVarName As String) As String
    If (Left(strVarName, 1) = "{" And Right(strVarName, 1) = "}") Or (Left(strVarName, 1) = "<" And Right(strVarName, 1) = ">") Then
        strVarName = Mid(strVarName, 2, Len(strVarName) - 2)
    End If
   
   
    If Right(strVarName, 1) = "-" Then
        If IsNumeric(Left(strVarName, Len(strVarName) - 1)) Then
            Dim I As Integer
            For I = CInt(Left(strVarName, Len(strVarName) - 1)) To Left(strVarName, Val(Len(strVarName) - 1)) + 1
                If I > UBound(args) Then Exit Function
                If args(I) = "" Then
                Else
                    GetVar = GetVar & args(I) & " "
                End If
            Next I
            'GetVar = JoinArray(args, " ", Left(strVarName, Val(Len(strVarName) - 1)) + 1)
            Exit Function
        End If
    End If
   
    If IsNumeric(strVarName) Then
        If Val(strVarName) > ArgCount Then
            GetVar = ""
        Else
            GetVar = args(strVarName)
        End If
        Exit Function
    End If
   
   
    If variableType = 1 Then
        GetVar = GetLocalVar(strVarName)
    Else
        '*get global var
       
        ' finish code
        GetVar = rootEngine.GetGlobalVar(strVarName)
    End If
End Function

Private Sub GotoLabel(strName As String)
    Dim I As Integer
    For I = LBound(stack_label) To UBound(stack_label)
        If stack_label(I).strName = strName Then
            currentLine = stack_label(I).nLine
            Exit Sub
        End If
    Next
End Sub

Public Sub SetInfo(strtheName As String, at As Integer, strtheExtraParams As String)
    strName = strtheName
    AliasType = at
    strExtraParams = strtheExtraParams
End Sub


Public Sub SetLocalVar(strVarName As String, strValue As String, Optional arrayElement As Integer = 0)
    Dim I As Integer
   
    For I = 1 To varCount
       
        If variables(I).Name = strVarName Then
            If arrayElement > UBound(variables(I).value) Then ReDim Preserve variables(I).value(arrayElement) As String
            variables(I).value(arrayElement) = strValue
            Exit Sub
        End If
    Next I
   
    AddLocalVar strVarName, strValue, arrayElement
End Sub