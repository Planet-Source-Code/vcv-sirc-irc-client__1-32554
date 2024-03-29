Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#D:\WINDOWS\System32\stdole2.tlb#OLE Automation
Object={3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0; Richtx32.ocx
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
Object={F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0; COMDLG32.OCX
Reference=*\G{3D5C6BF0-69A3-11D0-B393-00A0C9055D8E}#1.0#0#D:\Program Files\Common Files\designer\MSDERUN.DLL#Microsoft Data Environment Instance 1.0 (SP4)
Reference=*\G{00000205-0000-0010-8000-00AA006D2EA4}#2.5#0#D:\Program Files\Common Files\system\ado\msado25.tlb#Microsoft ActiveX Data Objects 2.5 Library
Form=MDI_Client.frm
Module=modmain; main.bas
Module=modColors; colors.bas
Form=..\..\Program Files\Microsoft Visual Studio\VB98\frm_Status.frm
Module=modLanguage; modLanguage.bas
Module=modWindows; modWindows.bas
Module=modStrManip; modStrManip.bas
Module=modSettings; modSettings.bas
Module=modIAL; modIAL.bas
Form=frmLoadProfile.frm
Form=frmNewProfile.frm
Class=clsSSE_Alias; SEX\clsSEX_Alias.cls
Class=clsSSE_Main; SEX\clsSEX_Main.cls
Class=clsSSE_Script; SEX\clsSEX_Script.cls
Form=SEX\frmDev.frm
Form=SEX\frmFuncList.frm
Module=modSEX_CommonFunctions; SEX\modSEX_CommonFunctions.bas
Module=modSEX_Eval; SEX\modSEX_eval.bas
Form=frmOptions.frm
Form=frm_Channel.frm
Module=modInterpret; modInterpret.bas
Module=modTime; modTime.bas
Object={248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0; Mswinsck.ocx
Class=clsXPMenu; ..\!xpmenu\clsXPMenu.cls
Form=..\!xpmenu\frmXPMenu.frm
Module=modXPMenu; ..\!xpmenu\modXPMenu.bas
Module=modMenus; modMenus.bas
Form=frmDebug.frm
Form=frm_Query.frm
Form=frmAbout.frm
Module=modAPI; modAPI.bas
Module=modHotkeys; modHotkeys.bas
Module=modTreeView; modTreeView.bas
Form=frmDCCs.frm
Form=frmChanCentral.frm
Module=modMisc; modMisc.bas
IconForm="CLIENT"
Startup="Sub Main"
HelpFile=""
Title="sIRC alpha"
ExeName32="sIRC.exe"
Command32=""
Name="sIRC"
HelpContextID="0"
CompatibleMode="0"
MajorVer=0
MinorVer=0
RevisionVer=9
AutoIncrementVer=0
ServerSupportFiles=0
VersionComments="sIRC IRC Client for the Windows 32-bit Environment."
VersionCompanyName="None"
VersionFileDescription="sIRC Executable file."
VersionLegalCopyright="©2001 Matt C."
VersionProductName="sIRC"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1

[MS Transaction Server]
AutoRefresh=1
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               + 1
    ReDim Preserve strCode(1 To intCodeLines) As String
    strCode(intCodeLines) = strCodeLine
    
    If Left$(strCodeLine, 1) = ":" Or Right$(strCodeLine, 1) = ":" Then
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
    Dim i As Integer
    For i = 1 To intCodeLines
        newAliasClass.AddCodeLine strCode(i)
    Next i
    
    Set newAliasClass.rootEngine = rootEngine
End Sub

Private Sub CopyToArgs(ByRef paramlist() As String)
    If UBound(paramlist) = 0 Then
        ArgCount = 0
        Exit Sub
    End If
    
    ArgCount = UBound(paramlist) + 1
        
    Dim i As Integer
    For i = LBound(paramlist) To UBound(paramlist)
        args(i) = paramlist(i)
    Next i
End Sub




Public Function Execute(ByRef paramlist() As String, localVars() As String) As String
    
    Dim i As Integer
    For i = LBound(localVars) To UBound(localVars)
        SetLocalVar LeftOf(localVars(i), ":"), RightOf(localVars(i), ":")
    Next i
    
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
    '* Do not fuck with it...it will hurt you
    '*


    Dim strBuffer As String, i As Integer, curChar As String, prevChar As String
    Dim inVariable As Boolean, bEscape As Boolean, strLen As String, bWhiteSpace As Boolean
    strLen = Len(strLine)
    
    '* Clean stacks
    CleanUp
    
    i = 1
    
    If stack_whilecount <= UBound(stack_while) Then
        If stack_while(stack_whilecount).bIsFalse And Left$(strLine, 3) <> "end" Then
            Exit Function
        End If
    End If
    
    For i = 1 To strLen
        curChar = Mid$(strLine, i, 1)
        
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
                    stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)

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
                
                If Left$(CStr(stack_calls(stack_callcount).args(0)), 1) = "$" Then
                    If stack_calls(stack_callcount).ArgCount < 1 Then
                        strReturnVal = ""
                        GoTo skipFunction
                    End If
                    
                    strReturnVal = GetLocalVar(Mid$(CStr(stack_calls(stack_callcount).args(0)), 2), CInt(stack_calls(stack_callcount).args(1)))
                    GoTo skipFunction
                End If
                
                            
                Set rootEngine.cChildAlias = Me
                
                ReDim Preserve stack_calls(stack_callcount).args(stack_calls(stack_callcount).ArgCount)
                strReturnVal = rootEngine.ExecuteAlias(CStr(stack_calls(stack_callcount).args(0)), stack_calls(stack_callcount).args, serverID, emptyVars)
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
                If i = 1 Then
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
              