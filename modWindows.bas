Attribute VB_Name = "modWindows"
Option Explicit

Global Const MAX_CHANS = 25
Global Const MAX_CONNS = 5
Global Const MAX_QUERIES = 50

Public Channels(1 To MAX_CHANS)         As CHANNEL
Public ChanInUse(1 To MAX_CHANS)        As Boolean
Public Queries(1 To MAX_QUERIES)        As QUERY
Public QueriesInUse(1 To MAX_QUERIES)   As Boolean
Public windowStatus(1 To MAX_CONNS)     As STATUS
Public Connections(1 To MAX_CONNS)      As Boolean

Public Type typTopic
    SetBy   As String
    SetWhen As String
    Topic   As String
End Type

Public Type typSwitchButton
    serverID    As Integer
    strTitle    As String
    strText     As String
    hIcon       As Long
    bNewData    As Boolean
    NewLines    As Integer
    lRect       As Rect
    hwnd        As Long
End Type
Public SwitchWindows(100)   As typSwitchButton
Public SwitchWindowCount    As Integer
Public SwitchServers(10)    As typSwitchButton
Public SwitchServerCount    As Integer
Public Const SWITCH_WINDOWS = 0
Public Const SWITCH_SERVERS = 1
Public Sub FillSwitchbar(whichBar As Integer, Optional strWinType As String = "", Optional serverID As Integer = 0, Optional bShowServer As Boolean = True)
    If whichBar = 0 Then
        SwitchWindowCount = 0
    Else
        SwitchServerCount = 0
    End If
    
    Dim child As Form
    For Each child In Forms
        If TypeOf child Is MDIForm Then
        Else
            If child.bShowInTaskbar Then
                If (child.serverID = serverID Or serverID = 0) And child.serverID <> 0 Then
                    If strWinType = child.WinType() Or strWinType = "" Then
                        If bShowServer = False And child.WinType = "Status" Then
                        Else
                            If whichBar = 0 Then
                                SwitchWindowCount = SwitchWindowCount + 1
                                SwitchWindows(SwitchWindowCount).serverID = child.serverID
                                SwitchWindows(SwitchWindowCount).strTitle = child.strTitle
                                SwitchWindows(SwitchWindowCount).strText = child.GetTitle
                                SwitchWindows(SwitchWindowCount).hIcon = child.Icon
                                SwitchWindows(SwitchWindowCount).bNewData = child.bNewData
                                SwitchWindows(SwitchWindowCount).NewLines = child.intNewLines
                                SwitchWindows(SwitchWindowCount).hwnd = child.hwnd
                            Else
                                SwitchServerCount = SwitchServerCount + 1
                                SwitchServers(SwitchServerCount).serverID = child.serverID
                                SwitchServers(SwitchServerCount).strTitle = child.strTitle
                                SwitchServers(SwitchServerCount).strText = child.GetTitle
                                SwitchServers(SwitchServerCount).hIcon = child.Icon
                                SwitchServers(SwitchServerCount).bNewData = child.bNewData
                                SwitchServers(SwitchServerCount).NewLines = child.intNewLines
                                SwitchServers(SwitchServerCount).hwnd = child.hwnd
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub


Public Function LVWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim NM As NMHDR, statNum As Long
    statNum = LV_GetOldProcAddrByHwnd(hwnd)
    If statNum = -1 Then Exit Function
    
    Select Case uMsg
        Case WM_NOTIFY
            CopyMemory NM, ByVal lParam, 12&
            Debug.Print "lParam: " & lParam & "; wParam: " & wParam & "; hWnd: " & NM.hwndFrom
            Select Case NM.code
                Case 1242872
                    Exit Function
            End Select
    End Select
    
    LVWndProc = CallWindowProc(windowStatus(statNum).oldLVProcAddr, hwnd, uMsg, wParam, lParam)
End Function

Public Sub RESIZECLIENT()
    MoveWindow CLIENT.hwnd, CLIENT.Left \ 15, CLIENT.Top \ 15, CLIENT.Width \ 15, CLIENT.Height \ 15, True
    CLIENT.MDIForm_Resize
End Sub


Private Function S_GetOldProcAddrByHwnd(hwnd As Long) As Long
    Dim i As Integer
    For i = 1 To MAX_CONNS
        If windowStatus(i).hwnd = hwnd Then
            S_GetOldProcAddrByHwnd = i
            Exit Function
        End If
    Next i
    S_GetOldProcAddrByHwnd = -1
End Function

Public Function StatWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim statNum As Long, DI As DRAWITEMSTRUCT, NM As NMHDR
    statNum = S_GetOldProcAddrByHwnd(hwnd)
    If statNum = -1 Then Exit Function
    
    Select Case uMsg
        Case WM_ACTIVATE
            If wParam = WA_ACTIVE Then
                windowStatus(statNum).Resize_event
            End If
        Case WM_SIZE
            If SIZE_MAXIMIZED = wParam Then
                'SetMenu CLIENT.hwnd, 0
                StatWndProc = 0
                windowStatus(statNum).Resize_event
                Exit Function
            End If
        Case WM_NOTIFY
            CopyMemory NM, ByVal lParam, 12&
            Debug.Print "lParam: " & lParam & "; wParam: " & wParam & "; hWnd: " & NM.hwndFrom
    End Select
    
    If statNum = -1 Then
        StatWndProc = 0
    Else
        StatWndProc = CallWindowProc(windowStatus(statNum).oldProcAddr, hwnd, uMsg, wParam, lParam)
    End If
End Function
Private Function C_GetOldProcAddrByHwnd(hwnd As Long) As Long
    Dim i As Integer
    For i = 1 To MAX_CHANS
        If Channels(i).hwnd = hwnd Then
            C_GetOldProcAddrByHwnd = i
            Exit Function
        End If
    Next i
    C_GetOldProcAddrByHwnd = -1
End Function


Private Function LV_GetOldProcAddrByHwnd(hwnd As Long) As Long
    Dim i As Integer
    For i = 1 To MAX_CONNS
        If windowStatus(i).lvChannels.hwnd = hwnd Then
            LV_GetOldProcAddrByHwnd = i
            Exit Function
        End If
    Next i
    LV_GetOldProcAddrByHwnd = -1
End Function



Private Function Q_GetOldProcAddrByHwnd(hwnd As Long) As Long
    Dim i As Integer
    For i = 1 To MAX_CHANS
        If Queries(i).hwnd = hwnd Then
            Q_GetOldProcAddrByHwnd = i
            Exit Function
        End If
    Next i
    Q_GetOldProcAddrByHwnd = -1
End Function
Public Function ChanWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim chanNum As Long
    chanNum = C_GetOldProcAddrByHwnd(hwnd)
    If chanNum = -1 Then Exit Function
    
    Select Case uMsg
        Case WM_ACTIVATE
            If wParam = WA_ACTIVE Then
                Channels(chanNum).Resize_event
            End If
        Case WM_SIZE
            If SIZE_MAXIMIZED = wParam Then
                'SetMenu CLIENT.hwnd, 0
                ChanWndProc = 0
                Channels(chanNum).Resize_event
                Exit Function
            End If
    End Select
    
    If chanNum = -1 Then
        ChanWndProc = 0
    Else
        ChanWndProc = CallWindowProc(Channels(chanNum).oldProcAddr, hwnd, uMsg, wParam, lParam)
    End If
End Function

Public Function MDIWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_MDISETMENU
            Debug.Print hwnd & "~" & uMsg & "~" & wParam
            MDIWndProc = -1
            Exit Function
    End Select
    
    MDIWndProc = CallWindowProc(CLIENT.oldProcAddr, hwnd, uMsg, wParam, lParam)
End Function


Public Function QueryWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim queryNum As Long
    queryNum = Q_GetOldProcAddrByHwnd(hwnd)
    If queryNum = -1 Then Exit Function
    
    Select Case uMsg
        Case WM_ACTIVATE
            If wParam = WA_ACTIVE Then
                Channels(queryNum).Resize_event
            End If
        Case WM_SIZE
            If SIZE_MAXIMIZED = wParam Then
                'SetMenu CLIENT.hwnd, 0
                QueryWndProc = 0
                Queries(queryNum).Resize_event
                Exit Function
            End If
    End Select
    
    If queryNum = -1 Then
        QueryWndProc = 0
    Else
        QueryWndProc = CallWindowProc(Queries(queryNum).oldProcAddr, hwnd, uMsg, wParam, lParam)
    End If
End Function


Sub ButtonizeForm(frm As Form)
    Dim child As Control
    
    For Each child In frm
        If TypeOf child Is CommandButton Then cppButton child
    Next child
End Sub





Sub Center(frm As Form)
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
End Sub

Sub CenterDialog(frmName As Form)
    frmName.Move (Screen.Width - frmName.Width) / 2, (Screen.Height - frmName.Height) / 2
End Sub




Sub cppButton(Btn As CommandButton)
    SendMessage Btn.hwnd, &HF4&, &H0&, 0&
End Sub

Function GetFormByName(strName As String, serverIDx As Integer) As Form
    On Error Resume Next
    
    Dim retForm As Form
    For Each retForm In Forms
    
        If TypeOf retForm Is MDIForm Then GoTo skipnext
        
        If LCase(strName) = LCase(retForm.strTitle) Then
            If serverIDx = retForm.serverID Then
                Set GetFormByName = retForm
                Exit Function
            End If
        End If
        GoTo skipnext
nextOne:
skipnext:
    Next retForm
End Function

Function GetTaskText(strText As String, givenWidth As Integer) As String
    If CLIENT.picTask.textWidth(strText) < givenWidth Then
        GetTaskText = strText
    Else
        Dim i As Integer
        For i = 1 To Len(strText)
            If CLIENT.picTask.textWidth(Left$(strText, i) & "...") > givenWidth - 5 Then
                GetTaskText = Left$(strText, i - 1) & "..."
                Exit Function
            End If
        Next i
    End If
End Function

Function GetWindowCount() As Integer
    On Error Resume Next
    Dim winCount As Integer
    winCount = 0
    Dim frm As Form
    For Each frm In Forms
        On Error GoTo skipit
        If frm.bShowInTaskbar And Not TypeOf frm Is MDIForm Then winCount = winCount + 1
skipit:
    Next
    GetWindowCount = winCount
    Exit Function
End Function

Sub newChannel(strChannelName As String, serverID As Integer)
    CLIENT.DrawTaskbarAllServers
    Dim i As Integer
    For i = 1 To MAX_CHANS
        If ChanInUse(i) Then
            If LCase(Channels(i).strChanName) = LCase(strChannelName) And Channels(i).serverID = serverID Then
                Channels(i).chanNum = i
                Channels(i).strTitle = strChannelName
                Channels(i).bWindowInUse = True
                Channels(i).bShowInTaskbar = True
                Channels(i).SetCaption
                Channels(i).NickCount = 0
                Channels(i).realNickCnt = 0
                Channels(i).tvNickList.Clear
                CLIENT.SetActive strChannelName, serverID
                Channels(i).setFocus
                
                '* update taskbar
                CLIENT.DrawTaskbarAllServers
                
                '* update dynamic menu
                XPM_ServerMenu(serverID).SetVisible i + 2, True
                XPM_ServerMenu(serverID).SetText i + 2, strChannelName
                
                treeview_AddChannel CLIENT.tvServers, strChannelName, serverID
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To MAX_CHANS
        If ChanInUse(i) = False Then
            Set Channels(i) = New CHANNEL
            Channels(i).chanNum = i
            Channels(i).serverID = serverID
            Channels(i).strChanName = strChannelName
            Channels(i).strTitle = strChannelName
            Channels(i).bWindowInUse = True
            Channels(i).bShowInTaskbar = True
            ChanInUse(i) = True
            CLIENT.SetActive strChannelName, serverID
            Channels(i).SetCaption
            If XPM_Window_Auto.GetCheck(1) Then
                ShowWindow Channels(i).hwnd, SW_SHOWMAXIMIZED
            Else
                Channels(i).Visible = True
            End If
            
            treeview_AddChannel CLIENT.tvServers, strChannelName, serverID
            
            '* update taskbar
            CLIENT.DrawTaskbarAllServers
            
            '* update dynamic menu
            XPM_ServerMenu(serverID).SetVisible i + 2, True
            XPM_ServerMenu(serverID).SetText i + 2, strChannelName
        
            Exit Sub
        End If
    Next i
    
    MsgBox "You can only have a maximum of " & MAX_CHANS & " Channels.  If you wish to join another, please close one first, or complain to the author.", vbCritical

End Sub

Sub newQuery(strNick As String, serverID As Integer, Optional strHost As String = "")
    Dim i As Integer
    For i = 1 To MAX_QUERIES
        If QueriesInUse(i) Then
            If LCase(Queries(i).strNick) = LCase(strNick) And Queries(i).serverID = serverID Then
                Queries(i).queryNum = i
                Queries(i).strTitle = strNick
                Queries(i).bWindowInUse = True
                Queries(i).bShowInTaskbar = True
                Queries(i).UpdateCaption
                Queries(i).setFocus
                
                treeview_AddQuery CLIENT.tvServers, strNick, serverID
                
                '* update taskbar
                CLIENT.DrawTaskbarAllServers
                
                '* update dynamic menu
                XPM_ServerMenu(serverID).SetVisible i + MAX_CHANS + 3, True
                XPM_ServerMenu(serverID).SetText i + MAX_CHANS + 3, strNick
                Exit Sub
            End If
        End If
    Next i
    
    For i = 1 To MAX_QUERIES
        If QueriesInUse(i) = False Then
            Set Queries(i) = New QUERY
            Queries(i).queryNum = i
            Queries(i).serverID = serverID
            Queries(i).strNick = strNick
            Queries(i).strHost = strHost
            Queries(i).strTitle = strNick
            Queries(i).bWindowInUse = True
            Queries(i).bShowInTaskbar = True
            QueriesInUse(i) = True
            Queries(i).UpdateCaption
            If XPM_Window_Auto.GetCheck(1) Then
                ShowWindow Queries(i).hwnd, SW_SHOWMAXIMIZED
            Else
                Queries(i).Visible = True
            End If
            
            treeview_AddQuery CLIENT.tvServers, strNick, serverID
            
            '* update taskbar
            CLIENT.DrawTaskbarAllServers
            
            '* update dynamic menu
            XPM_ServerMenu(serverID).SetVisible i + MAX_CHANS + 3, True
            XPM_ServerMenu(serverID).SetText i + MAX_CHANS + 3, strNick
            Exit Sub
        End If
    Next i
    
    'MsgBox "You can only have a maximum of " & MAX_CHANS & " Channels.  If you wish to join another, please close one first, or complain to the author.", vbCritical

End Sub
Sub NewConnection(Optional strServer As String = "", Optional intport As Integer = 6667, Optional strPass As String = "")
    
    Dim i As Integer
    For i = 1 To MAX_CONNS
        If Connections(i) = False Then
            Set windowStatus(i) = New STATUS
            Connections(i) = True
            windowStatus(i).serverID = i
            windowStatus(i).strTitle = "Status"
            windowStatus(i).bWindowInUse = True
            windowStatus(i).SetServerInfo ServerAddr, CLng(ServerPort), strPass
            If XPM_Window_Auto.GetCheck(1) Then
                ShowWindow windowStatus(i).hwnd, SW_SHOWMAXIMIZED
            Else
                windowStatus(i).Visible = True
            End If
            CLIENT.DrawTaskbarAllServers
            
            XPM_Window.SetVisible (15 + i), True
            XPM_Window.SetText (15 + i), ServerAddr
        
            
            '* New treeview
            treeview_NewServer CLIENT.tvServers, ServerAddr, i
            treeview_SetActive CLIENT.tvServers, "Status", i
            Exit Sub
        End If
    Next i
    
    MsgBox "You can only have a maximum of " & MAX_CONNS & " connections.  If you wish to start a new one, you must first close an existing one.", vbCritical
End Sub


Function QueryExists(strNick As String, serverID As Integer) As Boolean
    Dim i As Integer
    For i = 1 To MAX_QUERIES
        If QueriesInUse(i) Then
            If Queries(i).strNick = strNick And Queries(i).serverID = serverID Then
                QueryExists = True
                Exit Function
            End If
        End If
    Next i
    QueryExists = False
End Function

Public Sub setFocus(hwnd As Long)
    SendMessage hwnd, WM_SETFOCUS, 0&, vbNullString
End Sub

Public Sub StayOnTop(frmForm As Form, Optional fOnTop As Boolean = True)
   
    'Const HWND_TOPMOST = -1
    'Const HWND_NOTOPMOST = -2
   
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
   
    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hwnd, lState, iLeft - 100, iTop - 100, iWidth, iHeight, 0)
End Sub

Sub WP_RememberWindow(frm As Form)
    On Error Resume Next
    
    Dim strTitle As String, strPos As String
    With frm
        strTitle = .GetTitle() & "(" & .serverID & ")"
        strPos = .Left & "," & .Top & "," & .Width & "," & .Height
        If Left$(strTitle, 1) = "#" Then strPos = strPos & "," & .picNickList.Width
        If strTitle = "" Then Exit Sub
    End With
    
    'MsgBox winINI
    PutINI winINI, "pos", strTitle, strPos
End Sub
Sub WP_ForgetWindow(frm As Form)
    On Error Resume Next
    
    Dim strTitle As String
    With frm
        strTitle = .GetTitle() & "(" & .serverID & ")"
        If strTitle = "" Then Exit Sub
    End With
    
    'MsgBox winINI
    PutINI winINI, "pos", strTitle, ""
End Sub
Sub WP_RememberClient()
    On Error Resume Next
    
    Dim strTitle As String, strPos As String
    With CLIENT
        strTitle = "CLIENT"
        strPos = .Left & "," & .Top & "," & .Width & "," & .Height
    End With
    
    PutINI winINI, "clientpos", strTitle, strPos
End Sub
Sub WP_ForgetClient()
    On Error Resume Next
    
    Dim strTitle As String, strPos As String
    With CLIENT
        strTitle = "CLIENT"
    End With
    
    PutINI winINI, "clientpos", strTitle, ""
End Sub
Sub WP_RememberAll()
    On Error Resume Next
    
    Dim strTitle As String, strPos As String, frm As Form
    
    For Each frm In Forms
        With frm
            strTitle = ""
            strTitle = .GetTitle() & "(" & .serverID & ")"
            strPos = .Left & "," & .Top & "," & .Width & "," & .Height
            If strTitle <> "" Then
                PutINI winINI, "pos", strTitle, strPos
            End If
        End With
    Next
    
End Sub
Sub WP_ForgetAll()
    On Error Resume Next
    
    Dim strTitle As String, strPos As String, frm As Form
    
    For Each frm In Forms
        With frm
            strTitle = ""
            strTitle = .GetTitle() & "(" & .serverID & ")"
            If strTitle <> "" Then
                PutINI winINI, "pos", strTitle, ""
            End If
        End With
    Next
    
End Sub
Sub WP_ResetWindow(frm As Form)
    On Error Resume Next
    
    Dim strTitle As String, strPos As String
    With frm
        strTitle = .GetTitle() & "(" & .serverID & ")"
        strPos = GetINI(winINI, "pos", strTitle, "")
        If strTitle = "" Or strPos = "" Then Exit Sub
    End With
    
    Dim intPos() As String
    intPos = Split(strPos, ",")
    
    If UBound(intPos) = 4 Then
        frm.picNickList.Width = intPos(4)
    End If
    frm.Move CInt(intPos(0)), CInt(intPos(1)), CInt(intPos(2)), CInt(intPos(3))
    
End Sub


Sub WP_ResetAll()
    On Error Resume Next
    
    Dim intPos() As String, strTitle As String, strPos As String, frm As Form
    
    For Each frm In Forms
        With frm
            strTitle = .GetTitle() & "(" & .serverID & ")"
            strPos = GetINI(winINI, "pos", strTitle, "")
            If strTitle <> "" Then
                intPos = Split(strPos, ",")
                If UBound(intPos) = 4 Then
                    frm.picNickList.Width = intPos(4)
                End If
                frm.Move CInt(intPos(0)), CInt(intPos(1)), CInt(intPos(2)), CInt(intPos(3))
            End If
        End With
    Next
    
End Sub
Sub WP_MAXALL()
    On Error Resume Next
    
    Dim intPos() As String, strTitle As String, strPos As String, frm As Form
    
    For Each frm In Forms
        If Not TypeOf frm Is MDIForm Then
            If frm.WindowState = vbMaximized Then
                If frm.bShowInTaskbar Then
                    WP_SetMax frm
                End If
            End If
        End If
    Next
    
End Sub
Sub WP_ResetClient()
    On Error Resume Next
    
    Dim strTitle As String, strPos As String
    With CLIENT
        strTitle = "CLIENT"
        strPos = GetINI(winINI, "clientpos", strTitle, "")
        If strTitle = "" Or strPos = "" Then Exit Sub
    End With
    
    Dim intPos() As String
    intPos = Split(strPos, ",")
    
    CLIENT.Move CInt(intPos(0)), CInt(intPos(1)), CInt(intPos(2)), CInt(intPos(3))
    
End Sub
Sub WP_SetMax(whichForm As Form)

    ShowWindow whichForm.hwnd, SW_MAXIMIZE
        
End Sub


