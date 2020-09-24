VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSE_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private scripts()        As clsSSE_Script
Private script_count     As Integer

'* Global variables
Private variables()     As typVariable
Private varCount        As Integer
Private Type typVariable
    Name    As String
    value   As String
End Type

Public cChildAlias  As New clsSSE_Alias
Public bLastExec    As Boolean
Function function_ping(ByRef paramlist())
    If UBound(paramlist()) < 1 Then function_ping = -1: Exit Function
    
    Dim theArgs(3)
    theArgs(0) = paramlist(0)
    theArgs(1) = paramlist(1)
    theArgs(2) = "PING"
    theArgs(3) = Timer
    function_ctcp theArgs()
    
    function_ping = ""
End Function

Function function_format(ByRef paramlist())
    If UBound(paramlist()) < 2 Then function_format = -1: Exit Function
    
    function_format = Format(paramlist(1), paramlist(2))
End Function

Function function_queryexists(paramlist())
    If UBound(paramlist) < 1 Then function_queryexists = -1: Exit Function
    
    function_queryexist = QueryExists(CStr(paramlist(1)), CInt(paramlist(0)))
End Function

Public Function AliasCount(scriptID As Integer) As Integer
    
    If ScriptCount = 0 Then AliasCount = 0 Else AliasCount = scripts(scriptID).GetAliasCount()
    
End Function


Function CallByNameVar(obj As Object, strAliasNameX As String, theMethod As VbCallType, XvarArgs() As Variant)
    CallByNameVar = CallByName(obj, strAliasNameX, theMethod, XvarArgs)
End Function


Public Function ExecuteEvent(strEventName As String, ByRef strParamList() As String, serverID As Integer, localVars() As String)

    ' No default events, comment out
    
'    On Error GoTo NotBuiltIn
'    Dim returnValue
'    returnValue = CallByName(Me, "function_" & strEventName, VbMethod, strParamList)
'    ExecuteAlias = CStr(returnValue)
'    Exit Function

'NotBuiltIn:
    'MsgBox strEventName
    Dim i As Integer, j As Integer, returnVal As String
    For i = 1 To script_count
        DoEvents
        CurrentServerID = serverID
        Call scripts(i).ExecuteAlias(strEventName, strParamList, 1, serverID, localVars) '1 = event
        If scripts(i).bExecuted = False Then
            '* **
            '* code to generate error??
            '* **
            bLastExec = False
            ExecuteEvent = ""
            'Exit Function
        Else
            bLastExec = False
            ExecuteEvent = returnValue
        End If
    Next i
End Function

Function function_asctime(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_asctime = 0: Exit Function
    
    On Error GoTo errhand
    function_asctime = AscTime(CLng(paramlist(1)))
    Exit Function
errhand:
    function_asctime = ""
End Function

Public Function ExecuteCTCP(strCtcpName As String, ByRef strParamList() As String, serverID As Integer, localVars() As String)
    Dim i As Integer, j As Integer, returnVal As String
    For i = 1 To script_count
        CurrentServerID = serverID
        returnValue = scripts(i).ExecuteAlias(strCtcpName, strParamList, 2, serverID, localVars) '1 = ctcp
        If scripts(i).bExecuted = False Then
            '* **
            '* code to generate error??
            '* **
        
            ExecuteCTCP = ""
            'Exit Function
        Else
            ExecuteCTCP = returnValue
        End If
    Next i
End Function


Public Function ExecuteHotkey(strVirtKey As String, strKey As String, ByRef strParamList() As String, serverID As Integer, localVars() As String) As Boolean
    Dim i As Integer, j As Integer, returnVal As String
    For i = 1 To script_count
        CurrentServerID = serverID
        returnValue = scripts(i).ExecuteAlias(strVirtKey, strParamList, 3, serverID, localVars, strKey) '1 = ctcp
        If scripts(i).bExecuted = False Then
            '* **
            '* code to generate error??
            '* **
        
            ExecuteHotkey = False
            'Exit Function
        Else
            ExecuteHotkey = True
        End If
    Next i
End Function
Function function_channelcentral(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_channelcentral = 0: Exit Function
    
    Dim theChan As CHANNEL
    Set theChan = GetFormByName(CStr(paramlist(1)), CInt(paramlist(0)))
    
    If theChan Is Nothing Then
    Else
        With frmChanCentral
            .strChanModes = windowStatus(theChan.serverID).strChannelModes
            .strSetChanModes = theChan.strModes
            .strKey = theChan.strKey
            .strLimit = theChan.strLimit
            .SetInfo
            
            Set .refWindow = theChan
            .LoadTopics
            
            .Show
        End With
    End If
End Function
Function function_calc(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_calc = -1: Exit Function
    
    Dim strEval As String
    strEval = Eval(JoinArrayV(paramlist, " ", 2))
    
    function_calc = strEval
End Function

Function function_ctcpreply(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_ctcpreply = -1: Exit Function
    
    Dim strReply As String
    strReply = JoinArrayV(paramlist, " ", 3)
    
    windowStatus(CurrentServerID).SendData "NOTICE " & paramlist(1) & " :" & Chr(1) & strReply & Chr(1)
    'Dim paramListX() As String, vars(3) As String
    'vars(0) = "target:" & paramlist(1)
    'vars(1) = "nick:" & paramlist(1)
    'vars(2) = "command:" & paramlist(2)
    'vars(3) = "data:" & JoinArrayV(paramlist, " ", 4)
    'paramListX = Split(JoinArrayV(paramlist, " ", 2), " ")
    'ExecuteEvent "ctcpreply", paramListX, CInt(paramlist(0)), vars
End Function

Function function_ctcp(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_ctcp = -1: Exit Function
    
    Dim strReply As String
    paramlist(2) = UCase(paramlist(2))
    strReply = JoinArrayV(paramlist, " ", 3)
    
    windowStatus(CurrentServerID).SendData "PRIVMSG " & paramlist(1) & " :" & Chr(1) & strReply & Chr(1)
    Dim paramListX() As String, vars(3) As String
    vars(0) = "target:" & paramlist(1)
    vars(1) = "nick:" & paramlist(1)
    vars(2) = "command:" & paramlist(2)
    vars(3) = "data:" & JoinArrayV(paramlist, " ", 4)
    paramListX = Split(JoinArrayV(paramlist, " ", 2), " ")
    ExecuteEvent "ctcpsend", paramListX, CInt(paramlist(0)), vars
End Function
Function function_chancount(ByRef paramlist())
    Dim i As Integer
    
    For i = 1 To MAX_CHANS
        If ChanInUse(i) Then
            If Channels(i).serverID = paramlist(0) Then 'paramlist(0) = serverid
                function_chancount = function_chancount + 1
            End If
        End If
    Next i
    
    Exit Function
End Function


Function function_channel(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_channel = -1: Exit Function
    
    Dim i As Integer, chanNum As Integer
    
    If Val(paramlist(1)) = 0 Then
        function_channel = function_chancount(paramlist)
        Exit Function
    End If
    
    For i = 1 To MAX_CHANS
        If ChanInUse(i) Then
            If Channels(i).serverID = paramlist(0) Then 'paramlist(0)=serverid
                chanNum = chanNum + 1
                If chanNum = Val(paramlist(1)) Then
                    function_channel = Channels(i).strChanName
                    Exit Function
                End If
            End If
        End If
    Next i
    
    function_channel = ""
End Function

Function function_charcount(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_charcount = 0: Exit Function
    
    If Len(paramlist(1)) Then
        function_charcount = UBound(Split(paramlist(1), paramlist(2)))
    Else
        function_charcount = 0
    End If

End Function

Function function_dec(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_dec = -1: Exit Function
    
    'MsgBox Val(paramlist(1)) & "~" & Val(paramlist(2))
    If UBound(paramlist) >= 2 Then
        function_dec = Val(paramlist(1)) - Val(paramlist(2))
    Else
        function_dec = Val(paramlist(1)) - 1
    End If
End Function
Function function_doevents(ByRef paramlist())
    DoEvents
End Function

Function function_beep(ByRef paramlist())
    Beep
End Function
Function function_echo(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_echo = -1: Exit Function
    
    If paramlist(0) = 0 Then Exit Function
    
    Dim frmLoop As Form, i As Integer
    For Each frmLoop In Forms
        On Error GoTo nextthing
        If TypeOf frmLoop Is MDIForm Then GoTo nextthing
        On Error GoTo nextthing
        If frmLoop.serverID = paramlist(0) Then
            If Left$(paramlist(1), 1) = "-" Then
                For i = 2 To Len(paramlist(1))
                    If Mid$(paramlist(1), i, 1) = "c" Then
                        'nothing yet
                    End If
                    Select Case Mid$(paramlist(1), i, 1)
                        Case "s"
                            If frmLoop.WinType() = "Status" Then
                                PutText frmLoop.rt_Output, JoinArrayV(paramlist, " ", 3)
                                Exit Function
                            End If
                        Case "a"
                            On Error Resume Next
                            If CLIENT.ActiveForm.serverID = paramlist(0) Then
                                PutText CLIENT.ActiveForm.rt_Output, JoinArrayV(paramlist, " ", 3)
                            Else
                                PutText windowStatus(CInt(paramlist(0))).rt_Output, JoinArrayV(paramlist, " ", 3)
                            End If
                            Exit Function
                        Case "q"
                            If QueryExists(CStr(paramlist(2)), CInt(paramlist(0))) Then
                                If frmLoop.GetTitle = CStr(paramlist(2)) Then
                                    PutText frmLoop.rt_Output, JoinArrayV(paramlist, " ", 4)
                                    Exit Function
                                End If
                            Else
                                If frmLoop.WinType() = "Status" Then
                                    PutText frmLoop.rt_Output, JoinArrayV(paramlist, " ", 4)
                                    Exit Function
                                End If
                            End If
                    End Select
                Next i
            Else
                If LCase(paramlist(1)) = LCase(frmLoop.strTitle) Then
                    PutText frmLoop.rt_Output, JoinArrayV(paramlist, " ", 3)
                    Exit Function
                End If
            End If
        End If
nextthing:
    Next
End Function
Function function_eval(ByRef paramlist())
    'MsgBox "WEE"
    If UBound(paramlist) < 1 Then function_eval = -1: Exit Function
    
    Dim strEval, i As Integer, strExpr As String
    strExpr = JoinArrayV(paramlist, " ", 2)
    
    strEval = Eval(strExpr)
    
    function_eval = CVar(strEval)
End Function

Function function_inc(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_inc = -1: Exit Function
    
    If UBound(paramlist) >= 2 Then
        function_inc = Val(paramlist(1)) + Val(paramlist(2))
    Else
        function_inc = Val(paramlist(1)) + 1
    End If
End Function

Function function_inchannel(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_inchannel = -1: Exit Function
    
    Dim i As Integer
    
    For i = 1 To MAX_CHANS
        If ChanInUse(i) Then
            If Channels(i).serverID = paramlist(0) Then
                If LCase(Channels(i).strChanName) = LCase(paramlist(1)) Then
                    If Channels(i).InChannel(CStr(paramlist(2))) Then
                        function_inchannel = 1
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    
    function_inchannel = 0
End Function

Function function_instr(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_instr = -1: Exit Function
    
    If UBound(paramlist) = 2 Then
        function_instr = InStr(1, paramlist(1), paramlist(2))
    Else
        function_instr = InStr(Val(paramlist(3)), paramlist(1), paramlist(2))
    End If
End Function

Function function_join(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_join = -1: Exit Function
    
    windowStatus(paramlist(0)).SendData "JOIN " & JoinArrayV(paramlist, " ", 2)
End Function

Function function_left(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_left = -1: Exit Function
    
    If paramlist(1) = "" Or Val(paramlist(2)) < 0 Then Exit Function
    function_left = Left$(paramlist(1), Val(paramlist(2)))
End Function
Function function_len(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_len = -1: Exit Function
    
    function_len = Len(paramlist(1))
End Function

Function function_me(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_me = -1: Exit Function
    
    windowStatus(CurrentServerID).SendData "PRIVMSG " & paramlist(1) & " :" & Chr$(ACTION) & "ACTION " & paramlist(2) & Chr$(ACTION)
    Dim paramListX(3) As String, vars(3) As String
    vars(0) = "target:" & paramlist(1)
    vars(1) = "nick:" & windowStatus(paramlist(0)).strCurNick
    vars(2) = "message:" & paramlist(2)
    vars(3) = "action:" & paramlist(2)
    paramListX(1) = paramlist(1)
    paramListX(2) = windowStatus(paramlist(0)).strCurNick
    paramListX(3) = paramlist(2)
    ExecuteEvent "action", paramListX, CInt(paramlist(0)), vars

End Function

Function function_mid(ByRef paramlist())
    
    If UBound(paramlist) < 2 Then function_mid = -1:  Exit Function
    
    If UBound(paramlist) = 2 Then
        function_mid = Mid$(paramlist(1), Val(paramlist(2)))
    Else
        function_mid = Mid$(paramlist(1), Val(paramlist(2)), Val(paramlist(3)))
    End If

End Function

Function function_date(ByRef paramlist())
    
    If UBound(paramlist) < 1 Then
        function_date = Date
    Else
        function_date = Format(Now, paramlist(1))
    End If

End Function
Function function_time(ByRef paramlist())
    
    If UBound(paramlist) = 1 Then
        function_time = Time
    Else
        function_time = Format(Time, paramlist(2))
    End If

End Function
Function function_msg(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_msg = -1: Exit Function
    
    Dim msg As String
    strMsg = JoinArrayV(paramlist, " ", 3)
    
    windowStatus(CurrentServerID).SendData "PRIVMSG " & paramlist(1) & " :" & strMsg
    Dim paramListX(3) As String, vars(4) As String
    vars(0) = "target:" & paramlist(1)
    vars(1) = "channel:" & paramlist(1)
    vars(2) = "chan:" & paramlist(1)
    vars(3) = "nick:" & windowStatus(paramlist(0)).strCurNick
    vars(4) = "message:" & strMsg
    
    paramListX(1) = paramlist(1)
    paramListX(2) = windowStatus(paramlist(0)).strCurNick
    paramListX(3) = strMsg
    ExecuteEvent "text", paramListX, CInt(paramlist(0)), vars
End Function
Function function_notice(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_notice = -1: Exit Function
    
    Dim msg As String
    strMsg = JoinArrayV(paramlist, " ", 3)
    
    windowStatus(CurrentServerID).SendData "NOTICE " & paramlist(1) & " :" & strMsg
    Dim paramListX(3) As String, vars(4) As String
    vars(0) = "target:" & paramlist(1)
    vars(1) = "channel:" & paramlist(1)
    vars(2) = "chan:" & paramlist(1)
    vars(3) = "nick:" & windowStatus(paramlist(0)).strCurNick
    vars(4) = "message:" & strMsg
    
    paramListX(1) = paramlist(1)
    paramListX(2) = windowStatus(paramlist(0)).strCurNick
    paramListX(3) = strMsg
    ExecuteEvent "noticesend", paramListX, CInt(paramlist(0)), vars
End Function
Function function_opnotice(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_opnotice = -1: Exit Function
    
    Dim msg As String
    strMsg = JoinArrayV(paramlist, " ", 3)
    
    windowStatus(CurrentServerID).SendData "NOTICE @" & paramlist(1) & " :" & strMsg
    Dim paramListX(3) As String, vars(4) As String
    vars(0) = "target:" & paramlist(1)
    vars(1) = "channel:" & paramlist(1)
    vars(2) = "chan:" & paramlist(1)
    vars(3) = "nick:" & windowStatus(paramlist(0)).strCurNick
    vars(4) = "message:" & strMsg
    
    paramListX(1) = paramlist(1)
    paramListX(2) = windowStatus(paramlist(0)).strCurNick
    paramListX(3) = strMsg
    ExecuteEvent "opnoticesend", paramListX, CInt(paramlist(0)), vars
End Function
Function function_query(ByRef paramlist())
    If UBound(paramlist()) < 1 Then function_query = -1: Exit Function
    
    If QueryExists(CStr(paramlist(1)), CInt(paramlist(0))) Then
        Dim theQuery As Form
        Set theQuery = GetFormByName(CStr(paramlist(1)), CInt(paramlist(0)))
        
        If theQuery Is Nothing Then Exit Function
        theQuery.setFocus
    Else
        newQuery CStr(paramlist(1)), CInt(paramlist(0))
    End If
            
    If UBound(paramlist()) >= 2 Then
        Dim paramListX(3) As String, vars(2) As String, strMessage As String
        strMessage = JoinArrayV(paramlist, " ", 3)
        
        windowStatus(CurrentServerID).SendData "PRIVMSG " & paramlist(1) & " :" & strMessage
        
        vars(0) = "target:" & paramlist(1)
        vars(1) = "nick:" & windowStatus(paramlist(0)).strCurNick
        vars(2) = "message:" & strMessage
        paramListX(1) = paramlist(1)
        paramListX(2) = windowStatus(paramlist(0)).strCurNick
        paramListX(3) = strMessage
        ExecuteEvent "text", paramListX, CInt(paramlist(0)), vars
    End If
End Function

Function function_topic(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_topic = -1: Exit Function
    
    windowStatus(paramlist(0)).SendData "TOPIC " & paramlist(1) & " :" & paramlist(2)
End Function

Function function_kick(ByRef paramlist())
    If UBound(paramlist) < 3 Then function_kick = -1: Exit Function
    
    windowStatus(paramlist(0)).SendData "KICK " & paramlist(1) & " " & paramlist(2) & " :" & JoinArrayV(paramlist, " ", 4)
End Function

Function function_part(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_part = -1: Exit Function
    
    If UBound(paramlist) >= 2 Then
        windowStatus(paramlist(0)).SendData "PART " & paramlist(1) & " :" & JoinArrayV(paramlist, " ", 3)
    Else
        windowStatus(paramlist(0)).SendData "PART " & paramlist(1)
    End If
End Function


Function function_quit(ByRef paramlist())
    If UBound(paramlist) < 1 Then
        windowStatus(paramlist(0)).SendData "QUIT :Leaving - " & strVersion
    Else
        windowStatus(paramlist(0)).SendData "QUIT :" & JoinArrayV(paramlist, " ", 2)
    End If
End Function


Function function_right(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_right = -1: Exit Function
    
    If paramlist(1) = "" Or Val(paramlist(2)) < 0 Then Exit Function
    function_right = Right$(paramlist(1), Val(paramlist(2)))
    
End Function


Function function_senddata(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_senddata = -1: Exit Function
    
    windowStatus(paramlist(0)).SendData JoinArrayV(paramlist, " ", 2) & vbCrLf
End Function

Function function_set(ByRef paramlist())
    If UBound(paramlist) < 2 Then function_set = -1: Exit Function
    
    'MsgBox paramlist(1) & "~fucker"
    If paramlist(1) Like "*:*" Then
        cChildAlias.SetLocalVar CStr(Mid$(paramlist(1), 1, InStr(paramlist(1), ":") - 1)), JoinArrayV(paramlist, " ", 3), Mid$(paramlist(1), InStr(paramlist(1), ":") + 1)
    End If
    
    cChildAlias.SetLocalVar CStr(paramlist(1)), JoinArrayV(paramlist, " ", 3)
End Function

Function function_ticks(ByRef paramlist())
    
    function_ticks = Timer
    
End Function

Function function_getctime(ByRef paramlist())
    
    function_getctime = CTime
    
End Function

Function function_ctime(ByRef paramlist())
    
    function_ctime = CTime
    
End Function
Public Function GetAlias(scriptID As Integer, AliasNum As Integer) As String
    GetAlias = scripts(scriptID).GetAlias(AliasNum)
End Function

Public Function GetGlobalVar(strName As String) As String
    Dim i As Integer
    For i = 1 To varCount
        If StrComp(variables(i).Name, strName) = True Then
            GetGlobalVar = variables(i).value
            Exit Function
        End If
    Next i
    GetGlobalVar = ""
End Function


Private Sub AddGlobalVar(strName As String, strValue As String, Optional tvType As Integer = 0)
    varCount = varCount + 1
    ReDim Preserve variables(1 To varCount) As typVariable
    variables(varCount).Name = strName
    variables(varCount).value = strValue
    'variables(varCount).type = tvType
End Sub
Public Sub KillAllScripts()
    script_count = 0
    ReDim Preserve scripts(1 To 1) As clsSSE_Script

End Sub

Public Sub SetGlobalVar(strName As String, strValue As String)
    Dim i As Integer
    
    For i = 1 To varCount
        If StrComp(variables(i).Name, strName) = True Then
            variables(i).value = strValue
            Exit Sub
        End If
    Next i
    AddGlobalVar strName, strValue
End Sub



    


Public Function ExecuteAlias(strAliasNameX As String, ByRef strParamList() As String, serverID As Integer, localVars() As String)

    On Error GoTo NotBuiltIn
    Dim returnValue, varArgsX(), i As Integer
    ReDim varArgsX(UBound(strParamList))
    
    For i = 0 To UBound(strParamList)
        'MsgBox strParamList(i)
        varArgsX(i) = CVar(strParamList(i))
    Next i
    
    CurrentServerID = serverID
    varArgsX(0) = serverID
    returnValue = CallByName(Me, "function_" & strAliasNameX, VbMethod, varArgsX)
    ExecuteAlias = CStr(returnValue)
    bLastExec = True
    Exit Function

NotBuiltIn:
    'If Err Then MsgBox Err.Number & " -> " & Err.Description & " (on " & Err.Source & ")"
    
    Dim j As Integer, returnVal As String
    For i = 1 To script_count
        
        CurrentServerID = serverID
        returnValue = scripts(i).ExecuteAlias(strAliasNameX, strParamList, 0, serverID, localVars)
        
        If scripts(i).bExecuted = False Then
            '* **
            '* code to generate error??
            '* **
            bLastExec = False
            ExecuteAlias = ""
            Exit Function
        Else
            bLastExec = True
            ExecuteAlias = returnValue
            Exit Function
        End If
    Next i
    
End Function

Function function_input(ByRef paramlist())
    If UBound(paramlist) < 1 Then function_input = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        function_input = InputBox(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        function_input = InputBox(paramlist(1), paramlist(2))
    Else
        function_input = InputBox(paramlist(1), paramlist(2), paramlist(3))
    End If
End Function

Function function_msgbox(ByRef paramlist())
    On Error Resume Next
    If UBound(paramlist) < 1 Then function_msgbox = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        function_msgbox = MsgBox(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        function_msgbox = MsgBox(paramlist(1), 0, paramlist(2))
    Else
        function_msgbox = MsgBox(paramlist(1), paramlist(3), paramlist(2))
    End If
    
End Function

Function function_lcase(ByRef paramlist())
    On Error Resume Next
    If UBound(paramlist) < 1 Then function_lcase = -1: Exit Function
    
    function_lcase = LCase(JoinArrayV(paramlist, " ", 2))
    
End Function
Function function_ucase(ByRef paramlist())
    On Error Resume Next
    If UBound(paramlist) < 1 Then function_ucase = -1: Exit Function
    
    function_ucase = UCase(JoinArrayV(paramlist, " ", 2))
    
End Function

Function function_server(ByRef paramlist())
    On Error Resume Next
    If UBound(paramlist) < 1 Then function_server = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        windowStatus(paramlist(0)).SetServerInfo CStr(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        windowStatus(paramlist(0)).SetServerInfo CStr(paramlist(1)), CInt(paramlist(2))
    Else
        windowStatus(paramlist(0)).SetServerInfo CStr(paramlist(1)), CInt(paramlist(2)), CStr(paramlist(3))
    End If
    
    windowStatus(paramlist(0)).Disconnect
    windowStatus(paramlist(0)).Connect
    
End Function
Public Sub LoadScript(script_id As Integer, strFileName As String)
    scripts(script_id).LoadScript strFileName, Me
    
End Sub

Public Function NewScript()
    script_count = script_count + 1
    ReDim Preserve scripts(1 To script_count) As clsSSE_Script
    Set scripts(script_count) = New clsSSE_Script
    Set scripts(script_count).parent = Me
    NewScript = script_count
End Function

Public Function ScriptCount() As Integer
    ScriptCount = script_count
End Function

