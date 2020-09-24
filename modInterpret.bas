Attribute VB_Name = "modInterpret"
Option Explicit

Public varArgs()
Public iLoop As Long
Public Sub AddChannelList(strChannel As String, strUserCount As String, strTopic As String, serverID As Integer)
    Dim li As ListItem
    On Error Resume Next
    With windowStatus(serverID).lvChannels
        Set li = .ListItems.Add(, , strChannel)
        li.SubItems(1) = strUserCount
        li.SubItems(2) = strTopic
    End With
    
    'script stuff..add it
End Sub


Sub ChangeNick(strOldNick As String, strNewNick As String, serverID As Integer)
    Dim i As Integer
    On Error Resume Next
    
    For i = 1 To MAX_CHANS
        If ChanInUse(i) Then
            If Channels(i).serverID = serverID Then Channels(i).ChangeUserNick strOldNick, strNewNick
        End If
    Next i
    
    If QueryExists(strOldNick, serverID) Then
        Dim theQuery As Form
        Set theQuery = GetFormByName(strOldNick, serverID)
        
        If theQuery Is Nothing Then Exit Sub
        
        XPM_ServerMenu(serverID).SetText theQuery.queryNum + MAX_CHANS + 3, strNewNick
        theQuery.strNick = strNewNick
        theQuery.UpdateCaption
        
        CLIENT.DrawTaskbarAllServers
    End If
    
    If strOldNick = windowStatus(serverID).strCurNick Then
        windowStatus(serverID).strCurNick = strNewNick
    End If
    
    Dim argsX(2) As String, vars(1) As String
    argsX(1) = strOldNick
    argsX(2) = strNewNick
    vars(0) = "oldnick:" & strOldNick
    vars(1) = "newnick:" & strNewNick
    scriptEngine.ExecuteEvent "nick", argsX, serverID, vars
End Sub


Sub Event001(strWelcomeMsg As String, serverID As Integer)
    Dim vars(1) As String, args(1) As String
    vars(1) = "message:" & strWelcomeMsg
    args(1) = strWelcomeMsg
    
    scriptEngine.ExecuteEvent "001", args, serverID, vars
End Sub

Sub Event255(strClients As String, serverID As Integer)
    Dim vars(3) As String, args(1) As String
    vars(1) = "message:" & strClients
    vars(2) = "clients:" & strClients
    vars(3) = "clientcount:" & strClients
    args(1) = strClients
    
    scriptEngine.ExecuteEvent "255", args, serverID, vars
End Sub

Sub BeginMOTD(strText As String, serverID As Integer)
    Dim vars(2) As String, args(1) As String
    vars(1) = "message:" & strText
    vars(2) = "text:" & strText
    args(1) = strText
    
    scriptEngine.ExecuteEvent "375", args, serverID, vars
End Sub
Sub MOTD(strText As String, serverID As Integer)
    Dim vars(2) As String, args(1) As String
    vars(1) = "message:" & strText
    vars(2) = "text:" & strText
    args(1) = strText
    
    scriptEngine.ExecuteEvent "372", args, serverID, vars
End Sub
Sub EndMOTD(strText As String, serverID As Integer)
    Dim vars(2) As String, args(1) As String
    vars(1) = "message:" & strText
    vars(2) = "text:" & strText
    args(1) = strText
    
    On Error Resume Next
    scriptEngine.ExecuteEvent "376", args, serverID, vars
End Sub
Sub Event251(strUserCount As String, serverID As Integer)
    Dim vars(1) As String, args(1) As String
    vars(1) = "message:" & strUserCount
    args(1) = strUserCount
    
    scriptEngine.ExecuteEvent "251", args, serverID, vars
End Sub
Sub OpersOn(strOpersOn As String, strText As String, serverID As Integer)
    Dim vars(4) As String, args(2) As String
    vars(1) = "opers:" & strOpersOn
    vars(2) = "operson:" & strOpersOn
    vars(3) = "message:" & strText
    vars(4) = "text:" & strText
    args(1) = strOpersOn
    args(2) = strText
    
    scriptEngine.ExecuteEvent "252", args, serverID, vars
End Sub

Sub ChannelsFormed(ByRef strChannelsFormed As String, ByRef strText As String, serverID As Integer)
    Dim vars(4) As String, args(2) As String
    vars(1) = "chans:" & strChannelsFormed
    vars(2) = "chansformed:" & strChannelsFormed
    vars(3) = "message:" & strText
    vars(4) = "text:" & strText
    args(1) = strChannelsFormed
    args(2) = strText
    
    scriptEngine.ExecuteEvent "254", args, serverID, vars
End Sub
Sub Event002(strYourServer As String, serverID As Integer)
    Dim vars(1) As String, args(1) As String
    vars(1) = "message:" & strYourServer
    args(1) = strYourServer
    
    scriptEngine.ExecuteEvent "002", args, serverID, vars
End Sub
Sub Event003(strServerCreated As String, serverID As Integer)
    Dim vars(3) As String, args(1) As String
    vars(1) = "created:" & strServerCreated
    vars(2) = "servercreated:" & strServerCreated
    vars(3) = "message:" & strServerCreated
    args(1) = strServerCreated
    
    scriptEngine.ExecuteEvent "003", args, serverID, vars
End Sub


Sub Event004(strServer As String, strVersion As String, strUserModes As String, strChannelModes As String, serverID As Integer)
    Dim vars(7) As String, args(4) As String
    vars(1) = "server:" & strServer
    vars(2) = "serverv:" & strVersion
    vars(3) = "serverver:" & strVersion
    vars(4) = "serverversion:" & strVersion
    vars(5) = "version:" & strVersion
    vars(6) = "usermodes:" & strUserModes
    vars(7) = "chanmodes:" & strChannelModes
    args(1) = strServer
    args(2) = strVersion
    args(3) = strUserModes
    args(4) = strChannelModes
    
    scriptEngine.ExecuteEvent "004", args, serverID, vars
End Sub
Sub Event005(strSupported() As String, serverID As Integer)
    Dim vars(2) As String, args(2) As String
    vars(1) = "supported:" & Join(strSupported, " ")
    
    'scriptEngine.ExecuteEvent "005", args, serverID, Vars
End Sub
Sub ParseChannelModes(strChannel As String, strData As String, serverID As Integer)
    Dim strModes() As String, strChar As String, theChannel As Form
    Dim i As Integer, intParam As Integer
    Dim bAdd As Boolean
    Set theChannel = GetFormByName(strChannel, serverID)
    
    On Error Resume Next
    
    bAdd = True
    strModes = Split(strData, " ")
    For i = 1 To Len(strModes(0))
        strChar = Mid$(strModes(0), i, 1)
        
        Select Case strChar
            Case "+"
                bAdd = True
            Case "-"
                bAdd = False
            Case "k", "l"
                intParam = intParam + 1
                theChannel.DoMode bAdd, strChar, strModes(intParam)
            Case Else
                theChannel.DoMode bAdd, strChar, ""
        End Select
    Next i
        
    If Left$(strChannel, 1) = "#" Or Left$(strChannel, 1) = "&" Then
        Dim argsX(2) As String, vars(3) As String
        argsX(1) = strChannel
        argsX(2) = strData
        vars(0) = "chan:" & strChannel
        vars(1) = "channel:" & strChannel
        vars(2) = "data:" & strData
        vars(3) = "modes:" & strData
        
        scriptEngine.ExecuteEvent "324", argsX, serverID, vars
    End If
End Sub



Sub DoNotice(strFrom As String, strTo As String, strMessage As String, bFromServer As Boolean, serverID As Integer)
    Dim argsX(3) As String, vars(2) As String
    argsX(1) = strFrom
    argsX(2) = strTo
    argsX(3) = strMessage
    vars(0) = "target:" & strTo
    vars(1) = "nick:" & strFrom
    vars(2) = "message:" & strMessage
    
    If Left$(strMessage, 1) = Chr(ACTION) Then
        Dim CTCPcom As String, CTCPArgs() As String, CTCPAllArgs As String, Trash As String
        CTCPcom = Right$(strMessage, Len(strMessage) - 1)
        CTCPcom = Left$(CTCPcom, Len(CTCPcom) - 1)
        
        Trash = CTCPcom
        CTCPcom = LeftOf(Trash, " ")
        CTCPAllArgs = RightOf(Trash, " ")
        
        CTCPArgs = Split(CTCPAllArgs, " ")
        
        HandleCTCP strTo, strFrom, CTCPcom, CTCPArgs, serverID, False
    Else
        scriptEngine.ExecuteEvent "notice", argsX, serverID, vars
    End If

End Sub

Sub DoServerNotice(strServer As String, strMessage As String, serverID As Integer)
    Dim argsX(3) As String, vars(2) As String
    argsX(1) = strServer
    argsX(2) = strMessage
    vars(0) = "server:" & strServer
    vars(1) = "message:" & strMessage
        
    scriptEngine.ExecuteEvent "snotice", argsX, serverID, vars
End Sub
Sub DoPrivMSG(strNickx As String, theArgs() As String, serverID As Integer, Optional strIdent As String = "", Optional strHost As String = "")
    Dim params(3) As String, vars(2) As String
    params(1) = theArgs(0)
    params(2) = strNickx
    params(3) = theArgs(1)
    vars(0) = "target:" & theArgs(0)
    vars(1) = "nick:" & strNickx
    vars(2) = "message:" & theArgs(1)
    
    If Left$(theArgs(1), 1) = Chr(ACTION) Then
        Dim CTCPcom As String, CTCPArgs() As String, CTCPAllArgs As String, Trash As String
        CTCPcom = Right$(theArgs(1), Len(theArgs(1)) - 1)
        CTCPcom = Left$(CTCPcom, Len(CTCPcom) - 1)
        
        Trash = CTCPcom
        CTCPcom = LeftOf(Trash, " ")
        CTCPAllArgs = RightOf(Trash, " ")
        
        CTCPArgs = Split(CTCPAllArgs, " ")
        
        HandleCTCP params(1), params(2), CTCPcom, CTCPArgs, serverID
    Else
        If LCase(params(1)) = LCase(windowStatus(serverID).strCurNick) Then
            params(1) = strNickx        ' change target from yourself to user
            vars(0) = "target:" & strNickx
            
            Dim strFullHost As String
            If strIdent <> "" And strHost <> "" Then
                strFullHost = strIdent & "@" & strHost
            ElseIf strIdent = "" And strHost <> "" Then
                strFullHost = strHost
            End If
            
            If QueryExists(strNickx, serverID) Then
            Else
                newQuery strNickx, serverID, strFullHost
            End If
        Else
        End If
        
        Dim theForm As Form
        Set theForm = GetFormByName(params(1), serverID)
        
        If theForm Is Nothing Then
            Exit Sub
        Else
            If theForm.hwnd <> CLIENT.ActiveForm.hwnd Then
                Dim nInd1 As Integer, nInd2 As Integer
                nInd1 = treeview_GetChannelIndex(CLIENT.tvServers, params(1), serverID)
                nInd2 = treeview_GetQueryIndex(CLIENT.tvServers, params(1), serverID)
                
                If nInd1 <> -1 Then
                    CLIENT.tvServers.Nodes.item(nInd1).ForeColor = vbRed
                ElseIf nInd2 <> -1 Then
                    CLIENT.tvServers.Nodes.item(nInd2).ForeColor = vbRed
                End If
                
                theForm.bNewData = True
                CLIENT.DrawTaskbarAllServers
            End If
        End If
        
        Dim theParams() As String
        theParams = Split(theArgs(1), " ")
        scriptEngine.ExecuteEvent "text", theParams, serverID, vars
    End If

End Sub
Sub HandleCTCP(strToWhere As String, strFromWho As String, strCommand As String, theArgs() As String, serverID As Integer, Optional bReply As Boolean = True)
    Dim args() As String, vars(3) As String
    
    Select Case LCase(strCommand)
        Case "action"
            ReDim args(3) As String
            args(1) = strToWhere
            args(2) = strFromWho
            args(3) = JoinArray(theArgs, " ", 1)
            vars(0) = "target:" & strToWhere
            vars(1) = "nick:" & strFromWho
            vars(2) = "message:" & args(3)
            vars(3) = "action:" & args(3)
            
            If bReply Then scriptEngine.ExecuteEvent "action", args, serverID, vars
            Exit Sub
        Case "version"
            '* automatic version reply, cannot be changed
            If bReply Then windowStatus(serverID).SendData "PRIVMSG " & strFromWho & " :" & Chr(1) & "VERSION " & strVersion & Chr(1)
    End Select
    
    Dim theParams() As String
    If bReply Then
        theParams = Split(JoinArray(theArgs, " ", 1), " ")
        vars(0) = "target:" & strToWhere
        vars(1) = "nick:" & strFromWho
        vars(2) = "command:" & strCommand
        vars(3) = "data:" & JoinArray(theArgs, " ", 1)
        scriptEngine.ExecuteCTCP strCommand, theParams, serverID, vars
    Else
        theParams = Split(JoinArray(theArgs, " ", 1), " ")
        vars(0) = "target:" & strToWhere
        vars(1) = "nick:" & strFromWho
        vars(2) = "command:" & strCommand
        vars(3) = "data:" & JoinArray(theArgs, " ", 1)
        scriptEngine.ExecuteEvent "ctcpreply", theParams, serverID, vars
    End If
End Sub


Sub HandleError(strError As String, serverID As Integer)
    Dim xArgs(1) As String, vars(1) As String
    xArgs(1) = strError
    vars(0) = "error:" & strError
    vars(1) = "message:" & strError
    
    scriptEngine.ExecuteEvent "error", xArgs, serverID, vars
End Sub

Sub JoinChannel(strNickx As String, strChannelX As String, strHost As String, strIdent As String, serverID As Integer)
    Dim params(3) As String, vars(2) As String
    vars(0) = "channel:" & strChannelX
    vars(1) = "chan:" & strChannelX
    vars(2) = "nick:" & strNickx
    
    If strNickx = windowStatus(serverID).strCurNick Then
        newChannel strChannelX, serverID
        windowStatus(serverID).SendData "MODE " & strChannelX
    Else
        Dim theForm As Form
        Set theForm = GetFormByName(strChannelX, serverID)
        
        If theForm Is Nothing Then Exit Sub
        
        theForm.SetNick strNickx, strHost, strIdent
        
        params(1) = strChannelX
        params(2) = strNickx
        scriptEngine.ExecuteEvent "join", params, serverID, vars
    End If
    
End Sub




Sub KickUser(strKickedNick As String, strKicker As String, strChannelX As String, strReason As String, serverID As Integer)
    Dim params(4) As String, vars(6) As String
    
    Dim frmChan As Form
    Set frmChan = GetFormByName(strChannelX, serverID)
    
    If strKickedNick = windowStatus(serverID).strCurNick Then
        If bCloseOnKick Then
            frmChan.Tag = "KICKED"
            Unload frmChan
        End If
    End If
    
    If frmChan Is Nothing Then
    Else
        frmChan.RemoveNick strKickedNick
    End If
    
    params(1) = strChannelX
    params(2) = strKicker
    params(3) = strKickedNick
    params(4) = strReason
    vars(0) = "channel:" & strChannelX
    vars(1) = "knick:" & strKickedNick
    vars(2) = "kickednick:" & strKickedNick
    vars(3) = "nick:" & strKicker
    vars(4) = "message:" & strReason
    vars(5) = "reason:" & strReason
    vars(6) = "target:" & strChannelX
    scriptEngine.ExecuteEvent "kick", params, serverID, vars
    
End Sub


Sub ParseChannelNicks(ByRef strChannel As String, ByRef strNicks As String, serverID As Integer)
    Dim strNicksArray() As String, i As Integer, theChannel As Form
    strNicksArray = Split(strNicks, " ")
    Set theChannel = GetFormByName(strChannel, serverID)
    
    If theChannel Is Nothing Then Exit Sub
        
    For i = LBound(strNicksArray) To UBound(strNicksArray)
        If strNicksArray(i) <> "" Then
            Select Case Left$(strNicksArray(i), 1)
                Case "@"
                    theChannel.SetNick Right$(strNicksArray(i), Len(strNicksArray(i)) - 1), "", "", True
                Case "%"
                    theChannel.SetNick Right$(strNicksArray(i), Len(strNicksArray(i)) - 1), "", "", False, True
                Case "+"
                    theChannel.SetNick Right$(strNicksArray(i), Len(strNicksArray(i)) - 1), "", "", False, False, True
                Case Else
                    theChannel.SetNick strNicksArray(i), "", ""
            End Select
        End If
    Next i
    theChannel.UpdateNickList
    'theChannel.tvNickList.Sorted = True

End Sub

Sub PartChannel(strNickx As String, strChannelX As String, strReason As String, serverID As Integer)
    Dim params(3) As String, vars(5) As String
    
    If strNickx = windowStatus(serverID).strCurNick Then
    
    Else
        Dim theChan As Form
        Set theChan = GetFormByName(strChannelX, serverID)
        theChan.RemoveNick strNickx
    
        params(1) = strChannelX
        params(2) = strNickx
        params(3) = strReason
        vars(0) = "channel:" & strChannelX
        vars(1) = "chan:" & strChannelX
        vars(2) = "target: " & strChannelX
        vars(3) = "nick:" & strNickx
        vars(4) = "reason:" & strReason
        vars(5) = "message:" & strReason
        scriptEngine.ExecuteEvent "part", params, serverID, vars
    End If
    
End Sub


Public Sub SetTopic(strNick As String, strHost As String, strChannel As String, strNewTopic As String, serverID As Integer)
    Dim theChannel As Form
    Set theChannel = GetFormByName(strChannel, serverID)
    
    If theChannel Is Nothing Then Exit Sub
    
    PutText_Reset theChannel.txtTopic, strNewTopic
    theChannel.SetNewTopic strNewTopic, strNick
    
    Dim argsX(3) As String, vars(4) As String
    argsX(1) = strChannel
    argsX(2) = strNick
    argsX(3) = strNewTopic
    vars(0) = "target:" & strChannel: vars(1) = "channel:" & strChannel
    vars(2) = "newtopic:" & strNewTopic
    vars(3) = "topic:" & strNewTopic
    vars(4) = "nick:" & strNick
    
    scriptEngine.ExecuteEvent "topic", argsX, serverID, vars
    
End Sub

Sub topicOnJoin(args() As String, serverID As Integer)

    Dim Chan As Form, i As Integer
    Set Chan = GetFormByName(args(1), serverID)
    
    If Chan Is Nothing Then
        Exit Sub
    End If
    'Chan.strTopic = JoinArray(args, " ", 3)
    Chan.SetNewTopic JoinArray(args, " ", 3), ""
    
    Dim vars(2) As String
    vars(0) = "target:" & args(1)
    vars(1) = "channel:" & args(1)
    vars(2) = "topic:" & JoinArray(args, " ", 3)
    
    PutText_Reset Chan.txtTopic, JoinArray(args, " ", 3)
    scriptEngine.ExecuteEvent "332", args, serverID, vars
    
End Sub


Sub topicSetBy(args() As String, serverID As Integer)

    Dim Chan As Form, i As Integer
    Set Chan = GetFormByName(args(1), serverID)
    
    If Chan Is Nothing Then
        Exit Sub
    End If
    
    Chan.SetFirstTopicInfo args(2), args(3)

    Dim vars(4) As String
    vars(0) = "target:" & args(1)
    vars(1) = "channel:" & args(1)
    vars(2) = "nick:" & args(2)
    vars(3) = "when:" & args(3)
    vars(4) = "unixtime:" & args(3)

    scriptEngine.ExecuteEvent "333", args, serverID, vars
    
End Sub


Sub UserQuit(strNickx As String, strHost As String, strIdent As String, strReason As String, serverID As Integer)
    
    Dim argsX(2) As String, vars(2) As String
    argsX(1) = strNickx
    argsX(2) = strReason
    vars(0) = "nick:" & strNickx
    vars(1) = "reason:" & strReason
    vars(2) = "message:" & strReason
    scriptEngine.ExecuteEvent "quit", argsX, serverID, vars
    
    'moo
    Dim i As Integer
    For i = 1 To MAX_CHANS
        If ChanInUse(i) Then
            If Channels(i).serverID = serverID Then Channels(i).RemoveNick strNickx
        End If
    Next i
    
End Sub


Sub WhoIs_311(strNick As String, strUsername As String, strAddress As String, strInfo As String, serverID As Integer)

    Dim varsx(5) As String, argsX(4) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "username:" & strUsername
    varsx(5) = "name:" & strUsername
    varsx(2) = "address:" & strAddress
    varsx(3) = "info:" & strInfo
    varsx(4) = "email:" & strInfo
    argsX(1) = strNick
    argsX(2) = strUsername
    argsX(3) = strAddress
    argsX(4) = strInfo

    scriptEngine.ExecuteEvent "311", argsX, serverID, varsx
    

End Sub

Sub WhoIs_319(strNick As String, strChannels As String, serverID As Integer)

    Dim varsx(1) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "channels:" & strChannels
    argsX(1) = strNick
    argsX(2) = strChannels

    scriptEngine.ExecuteEvent "319", argsX, serverID, varsx

End Sub
Sub WhoIs_318(strNick As String, strMessage As String, serverID As Integer)

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "message:" & strMessage
    varsx(1) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "318", argsX, serverID, varsx

End Sub

Sub WhoIs_307(strNick As String, strMessage As String, serverID As Integer)

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "message:" & strMessage
    varsx(2) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "307", argsX, serverID, varsx

End Sub
Sub WhoIs_310(strNick As String, strMessage As String, serverID As Integer)
    ' 310: available for help

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "message:" & strMessage
    varsx(2) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "310", argsX, serverID, varsx

End Sub
Sub WhoIs_301(strNick As String, strReason As String, serverID As Integer)

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "reason:" & strReason
    varsx(2) = "message:" & strReason
    argsX(1) = strNick
    argsX(2) = strReason

    scriptEngine.ExecuteEvent "301", argsX, serverID, varsx

End Sub

Sub Away_305(strNick As String, strMessage As String, serverID As Integer)
    ' 305: back from away

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "info:" & strMessage
    varsx(2) = "message:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "305", argsX, serverID, varsx

End Sub

Sub Away_306(strNick As String, strMessage As String, serverID As Integer)
    ' 306: now away

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "info:" & strMessage
    varsx(2) = "message:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "306", argsX, serverID, varsx

End Sub


Sub WhoIs_313(strNick As String, strMessage As String, serverID As Integer)
    '   313 : NICK :is an ircop - net admin

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "message:" & strMessage
    varsx(2) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "313", argsX, serverID, varsx

End Sub
Sub Error_433(strNick As String, strMessage As String, serverID As Integer)
    '   433 : NICK :is already in use

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "message:" & strMessage
    varsx(2) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "433", argsX, serverID, varsx

End Sub
Sub Error_401(strNick As String, strMessage As String, serverID As Integer)
    '   401 : NICK :No such nick

    Dim varsx(2) As String, argsX(2) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "message:" & strMessage
    varsx(2) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strMessage

    scriptEngine.ExecuteEvent "401", argsX, serverID, varsx

End Sub
Sub Error_471(strNick As String, strChannel As String, strMessage As String, serverID As Integer)
    '   433 : NICK :is already in use

    Dim varsx(4) As String, argsX(3) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "chan:" & strChannel
    varsx(2) = "channel:" & strChannel
    varsx(3) = "message:" & strMessage
    varsx(4) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strChannel
    argsX(3) = strMessage

    scriptEngine.ExecuteEvent "471", argsX, serverID, varsx

End Sub
Sub Error_438(strNick As String, strNewNick As String, strMessage As String, serverID As Integer)
    '   433 : NICK :is already in use

    Dim varsx(3) As String, argsX(3) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "newnick:" & strNewNick
    varsx(2) = "message:" & strMessage
    varsx(3) = "info:" & strMessage
    argsX(1) = strNick
    argsX(2) = strNewNick
    argsX(3) = strMessage

    scriptEngine.ExecuteEvent "438", argsX, serverID, varsx

End Sub

Sub WhoIs_320(strMessage, serverID As Integer)

    Dim varsx(1) As String, argsX(1) As String
    varsx(0) = "message:" & strMessage
    varsx(1) = "info:" & strMessage
    argsX(1) = strMessage

    scriptEngine.ExecuteEvent "320", argsX, serverID, varsx

End Sub
Sub WhoIs_312(strNick As String, strServer As String, strServerDesc As String, serverID As Integer)

    Dim varsx(3) As String, argsX(3) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "server:" & strServer
    varsx(2) = "desc:" & strServerDesc
    varsx(3) = "info:" & strServerDesc
    argsX(1) = strNick
    argsX(2) = strServer
    argsX(3) = strServerDesc

    scriptEngine.ExecuteEvent "312", argsX, serverID, varsx

End Sub
Sub WhoIs_317(strNick As String, strIdle As String, strSignOnTime As String, strMessage As String, serverID As Integer)

    Dim varsx(4) As String, argsX(4) As String
    varsx(0) = "nick:" & strNick
    varsx(1) = "idle:" & strIdle
    varsx(2) = "signon:" & strSignOnTime
    varsx(3) = "info:" & strMessage
    varsx(4) = "message:" & strMessage
    argsX(1) = strNick
    argsX(2) = strIdle
    argsX(3) = strSignOnTime
    argsX(4) = strMessage

    scriptEngine.ExecuteEvent "317", argsX, serverID, varsx

End Sub
