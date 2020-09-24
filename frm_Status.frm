VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form STATUS 
   Caption         =   "Status"
   ClientHeight    =   5055
   ClientLeft      =   3660
   ClientTop       =   3525
   ClientWidth     =   8580
   FillColor       =   &H80000015&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Status.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picChannelList 
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   0
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   12
      Top             =   405
      Visible         =   0   'False
      Width           =   6450
      Begin MSComctlLib.ListView lvChannels 
         Height          =   4650
         Left            =   0
         TabIndex        =   13
         Top             =   -15
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   8202
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Channel"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Users"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Topic"
            Object.Width           =   7937
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   7530
      Top             =   2685
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":0CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":1058
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":13F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":1D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":20C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Status.frx":245A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Connect"
            Object.ToolTipText     =   "Connect to the server"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Disconnect"
            Object.ToolTipText     =   "Disconnect from server"
            ImageIndex      =   5
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Seperator"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Server Information"
            Object.ToolTipText     =   "Status window"
            ImageIndex      =   6
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Channel List"
            Object.ToolTipText     =   "Channel List"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "DCC Transfers"
            Object.ToolTipText     =   "DCC Transfers"
            ImageIndex      =   9
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      Height          =   4725
      Left            =   15
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   482
      TabIndex        =   0
      Top             =   405
      Width           =   7230
      Begin VB.PictureBox picServerTool 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   477
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   7155
         Begin VB.TextBox txtPassword 
            BorderStyle     =   0  'None
            Height          =   255
            IMEMode         =   3  'DISABLE
            Left            =   4995
            PasswordChar    =   "*"
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   30
            Width           =   1110
         End
         Begin VB.TextBox txtPort 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4440
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "6667"
            Top             =   30
            Width           =   465
         End
         Begin VB.TextBox txtServer 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   75
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "irc.othersideirc.net"
            Top             =   30
            Width           =   4215
         End
         Begin VB.Label lblPassword 
            Height          =   195
            Left            =   5145
            TabIndex        =   8
            Top             =   30
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lblPort 
            Alignment       =   1  'Right Justify
            Caption         =   "6667"
            Height          =   225
            Left            =   4440
            TabIndex        =   7
            Top             =   30
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Shape shpServer2 
            BorderColor     =   &H80000005&
            Height          =   270
            Left            =   0
            Top             =   15
            Width           =   4365
         End
         Begin VB.Label lblServer 
            Caption         =   "irc.othersideirc.net"
            Height          =   240
            Left            =   75
            TabIndex        =   6
            Top             =   30
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.Shape shpServer 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   285
            Left            =   0
            Top             =   15
            Width           =   4365
         End
         Begin VB.Shape shpPort 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   285
            Left            =   4380
            Top             =   15
            Width           =   570
         End
         Begin VB.Shape shpPort2 
            BorderColor     =   &H80000005&
            Height          =   270
            Left            =   4395
            Top             =   15
            Width           =   555
         End
         Begin VB.Shape shpPassword 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   285
            Left            =   4950
            Top             =   15
            Width           =   1320
         End
         Begin VB.Shape shpPassword2 
            BorderColor     =   &H80000005&
            Height          =   270
            Left            =   4995
            Top             =   15
            Width           =   1305
         End
         Begin VB.Label lblHide 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   11.25
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   6570
            TabIndex        =   5
            Top             =   15
            Width           =   255
         End
         Begin VB.Shape shpServerInfo 
            BackColor       =   &H8000000F&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   6555
         End
         Begin VB.Shape shpHide 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   270
            Left            =   6570
            Top             =   15
            Width           =   270
         End
      End
      Begin RichTextLib.RichTextBox rt_Input 
         Height          =   270
         Left            =   30
         TabIndex        =   9
         Top             =   4485
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   476
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         MultiLine       =   0   'False
         MaxLength       =   512
         Appearance      =   0
         OLEDropMode     =   1
         TextRTF         =   $"frm_Status.frx":27F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rt_Output 
         Height          =   4065
         Left            =   45
         TabIndex        =   10
         Top             =   330
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   7170
         _Version        =   393217
         BorderStyle     =   0
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         Appearance      =   0
         OLEDropMode     =   1
         TextRTF         =   $"frm_Status.frx":2870
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line ln_Sep 
         BorderColor     =   &H8000000F&
         X1              =   0
         X2              =   434
         Y1              =   295
         Y2              =   295
      End
   End
   Begin MSWinsockLib.Winsock socket 
      Left            =   8070
      Top             =   165
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrCommandQueue 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7650
      Top             =   165
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7230
      Top             =   165
   End
End
Attribute VB_Name = "STATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oldProcAddr     As Long
Public oldLVProcAddr   As Long
Private curWS   As Integer

Public strTopic As String
Public strChannelModes  As String
    
Private ShiftDown       As Boolean
Private ShiftKey        As Integer

Public bConnected       As Boolean
Public bShowServerInfo  As Boolean

Public strServerName    As String
Public intServerPort    As Long
Public serServerPass    As String

Public bWindowInUse     As Boolean
Public serverID         As Integer

Public strDataBuffer    As String
Public strTitle         As String

Public bShowInTaskbar   As Boolean
Public currentNick      As Integer

Public strCurNick       As String
Public strNetwork       As String

Private CommandQueue As New Collection

'* new data?
Public bNewData As Boolean
Public intNewLines As Integer

Private startX As Long, startY As Long

Private textHistory As New Collection
Private intCurHist  As Integer

Private whichButton As Long

Function AddHistory(strText As String) As Boolean
    'On Error Resume Next
    If intCurHist <> 0 And intCurHist <= textHistory.Count Then
        If textHistory.item(intCurHist) = strText Then
            AddHistory = False
            Exit Function
        End If
    End If

    If textHistory.Count > MAX_TEXT_HISTORY Then
        textHistory.Remove 1
    End If
    textHistory.Add strText
    intCurHist = textHistory.Count + 1
    AddHistory = True
End Function

'End Of DECS
Sub Connect()
    If socket.State = sckConnected Then
        PutText rt_Output, "2* You are already connected"
    Else
        On Error GoTo ErrorHandling:
        socket.Close
        'socket.RemoteHost = strServerName
        'socket.RemotePort = intServerPort
        socket.Connect strServerName, intServerPort
        PutText rt_Output, "2* Connecting to " & strServerName & " (port: " & intServerPort & ")"
        Exit Sub
    End If
    
ErrorHandling:
    PutText rt_Output, "4* Socket Error: (" & Err & ") " & Error & ""
End Sub


Sub Disconnect()
    If socket.State = sckConnected Then
        socket.Close
        PutText rt_Output, "02* Connection to server closed"
    Else
        PutText rt_Output, "02* Not connected to server"
    End If
End Sub


Public Function GetTitle() As String
    If bConnected Then
        GetTitle = strNetwork & " [" & strCurNick & "]"
    Else
        GetTitle = txtServer.Text & " [Offline]"
    End If
End Function

Sub HideAllServerInfo()
    
    'txtServer.Visible = False
    'shpServer.Visible = False
    'txtPort.Visible = False
    'shpPort.Visible = False
    'shpPassword.Visible = False
    'txtPassword.Visible = False
End Sub


Private Sub HideTabs()
    picStatus.Visible = False
    picChannelList.Visible = False
End Sub

Sub interpret(strDataX As String)
    
    Dim parsed As ParsedData, inttemp As Integer
    Dim i As Integer, StrChan As String, strTemp As String
    Dim args() As String, vars() As String
    
    strData = Replace$(strDataX, Chr(10), "")
    ParseData strDataX, parsed
    If Len(parsed.strCommand) = 0 Then Exit Sub
    
    ReDim args(UBound(parsed.strParams)) As String
    i = 1
    Do Until i > parsed.intParams
        
        args(i - 1) = Replace$(parsed.strParams(i), Chr(13), "")
        args(i - 1) = Replace$(args(i - 1), Chr(10), "")
        i = i + 1
    Loop
    
    If bDebug Then
        With frmDebug.txtDebugIn
            .selStart = Len(.Text)
            .selLength = 0
            .seltext = vbCrLf & strDataX
        End With
    End If
    
    Select Case UCase$(parsed.strCommand)
        Case "001"      ' Welcome
            strCurNick = args(0)
            Event001 args(1), serverID
            bConnected = True
            
            Dim strSplits() As String
            strSplits = Split(args(1), " ")
            strNetwork = strSplits(3)
            
            '* set dynamic menu info
            XPM_Window.SetText 15 + serverID, strNetwork
            
            '* update taskbar
            CLIENT.DrawTaskbarAllServers
        Case "002"
            Event002 args(1), serverID
        Case "003"
            Event003 args(1), serverID
        Case "004"
            strChannelModes = args(4)
            Event004 args(1), args(2), args(3), args(4), serverID
        Case "251"
            Event251 args(1), serverID
        Case "252"
            OpersOn args(1), args(2), serverID
        Case "254"
            ChannelsFormed args(1), args(2), serverID
        Case "301"      ' whois nick (is away) :reason
            WhoIs_301 args(1), args(2), serverID
        Case "305"      ' whois nick :no longer away
            Away_305 args(0), args(1), serverID
        Case "306"      ' whois nick :you have been marked as being away
            Away_306 args(0), args(1), serverID
        Case "307"      ' whois nick :has identified
            WhoIs_307 args(1), args(2), serverID
        Case "310"      ' whois nick :is available for help
            WhoIs_310 args(1), args(2), serverID
        Case "311"      ' whois nick username address :info
            WhoIs_311 args(1), args(2), args(3), args(5), serverID
        Case "312"      ' whois nick server :desc
            WhoIs_312 args(1), args(2), args(3), serverID
        Case "313"      ' whois nick :is an irc operator
            WhoIs_313 args(1), args(2), serverID
        Case "317"      ' whois nick secondsidle signontime :message
            WhoIs_317 args(1), args(2), args(3), args(4), serverID
        Case "318"      ' whois nick :end of whois
            WhoIs_318 args(1), args(2), serverID
        Case "319"      ' whois nick :channels
            WhoIs_319 args(1), args(2), serverID
        Case "320"      ' whois :nick is registered nickname
            WhoIs_320 args(1), serverID
        Case "321"      ' begin channel list
            lvChannels.ListItems.Clear
            Toolbar.Buttons.item(4).value = tbrUnpressed
            Toolbar.Buttons.item(5).value = tbrPressed
            Toolbar.Buttons.item(6).value = tbrUnpressed
            HideTabs
            picChannelList.Visible = True
        Case "322"      ' add channel  to channel list
            AddChannelList args(1), args(2), JoinArray(args, " ", 4), serverID
        Case "323"      ' end channel list
            'end.
        Case "324"      ' channel modes
            ParseChannelModes args(1), JoinArray(args, " ", 3), serverID
        Case "332"      ' topic (passed when you join a channel)
            topicOnJoin args, CInt(serverID)
        Case "333"      ' topic set by, 1=chan, 2=nick, 3=time
            topicSetBy args, CInt(serverID)
        Case "353"      ' channel names
            ParseChannelNicks args(2), args(3), serverID
        Case "372"      ' motd text
            MOTD args(1), serverID
        Case "375"      ' begin motd
            BeginMOTD args(1), serverID
        Case "376"      ' end motd
            EndMOTD args(1), serverID
        Case "401"      ' NICK :no such nick
            Error_401 args(1), args(2), serverID
        Case "433"      ' NICK :is already in use
            If bConnected = False Then
                If currentNick >= UBound(strNicks) Then
                    '* do something
                    rt_Input.Text = "/nick "
                    rt_Input.setFocus
                    rt_Input.selStart = Len(rt_Input.Text)
                Else
                    SendData "NICK " & strNicks(currentNick)
                    strCurNick = strNicks(currentNick)
                    currentNick = currentNick + 1
                End If
            Else
                Error_433 args(1), args(2), serverID
            End If
        Case "438"  ' nick -> newnick :nick change too fast, wait ...
            Error_438 args(1), args(2), args(3), serverID
        Case "471"  ' nick #chan :cannot join chan (...)
            Error_471 args(0), args(1), args(2), serverID
        Case "ERROR"
            HandleError JoinArray(args, " ", 1), serverID
        Case "JOIN"
            JoinChannel parsed.strNick, CStr(args(0)), parsed.strHost, parsed.strIdent, serverID
        Case "KICK"
            KickUser args(1), parsed.strNick, args(0), args(2), serverID
        Case "MODE"
            ParseMode parsed.strNick, args(0), JoinArray(args, " ", 2), serverID
        Case "NICK"
            ChangeNick parsed.strNick, args(0), serverID
        Case "NOTICE"
            If args(0) = "AUTH" And parsed.bFromServer Then
                DoServerNotice socket.RemoteHost, args(1), serverID
                Exit Sub
            End If
            
            If parsed.bFromServer Then
                DoNotice parsed.strFullHost, args(0), args(1), parsed.bFromServer, serverID
            Else
                DoNotice parsed.strNick, args(0), args(1), parsed.bFromServer, serverID
            End If
        Case "PART"
            PartChannel parsed.strNick, CStr(args(0)), CStr(args(1)), serverID
        Case "PING"
            ReDim vars(0) As String
            vars(0) = "value:" & args(0)
        
            SendData "PONG " & parsed.AllParams
            scriptEngine.ExecuteEvent "ping", args, serverID, vars
        Case "PRIVMSG"
            DoPrivMSG parsed.strNick, args, serverID, parsed.strIdent, parsed.strHost
        Case "QUIT"
            UserQuit parsed.strNick, parsed.strHost, parsed.strIdent, JoinArray(args, " ", 1), serverID
            If parsed.strNick = strCurNick Then
                bConnected = False
            End If
        Case "TOPIC"
            SetTopic parsed.strNick, parsed.strHost, args(0), args(1), serverID
        Case Else
            
    End Select

End Sub

Public Sub Resize_event()
    On Error Resume Next
    
    If Me.WindowState <> curWS Then
        If Me.WindowState = vbMaximized Then
            CLIENT.DrawMenu
        ElseIf Me.WindowState = vbNormal Then
            
        End If
        curWS = Me.WindowState
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub

    picStatus.Move 0, picStatus.Top, Me.ScaleWidth + 1, Me.ScaleHeight - picStatus.Top
    picChannelList.Move 0, picChannelList.Top, Me.ScaleWidth + 1, Me.ScaleHeight - picChannelList.Top
    
    ResizeToolbar
    DoEvents

    '* Resize Status Stuff
    Dim intTextHeight As Integer
    picStatus.FontName = rt_Input.Font.Name
    picStatus.FontSize = rt_Input.Font.Size
    intTextHeight = picStatus.textHeight("Ab_¯") + 4     '# 4 for buffer (2 on top, 2 bottom)
    rt_Output.Width = picStatus.ScaleWidth - 5 'buffer
    rt_Output.Height = picStatus.ScaleHeight - (intTextHeight + 7) - (rt_Output.Top - 1) '# 7 for buffer
    rt_Input.Top = picStatus.ScaleHeight - intTextHeight
    rt_Input.Height = intTextHeight + 4
    rt_Input.Width = picStatus.ScaleWidth - 4
    ln_Sep.Y1 = picStatus.ScaleHeight - (intTextHeight + 4)   '# 4 for buffer again
    ln_Sep.Y2 = ln_Sep.Y1
    ln_Sep.X2 = picStatus.ScaleWidth
    
    '* Resize Channel List
    lvChannels.Move lvChannels.Left, lvChannels.Top, picChannelList.ScaleWidth - (lvChannels.Left * 2) - 1, picChannelList.ScaleHeight - (lvChannels.Top * 2) - 1
    lvChannels.ColumnHeaders.item(3).Width = lvChannels.Width - lvChannels.ColumnHeaders.item(1).Width - lvChannels.ColumnHeaders.item(2).Width - 25
    
End Sub

Sub ResizeToolbar()
    On Error Resume Next
    If bShowServerInfo Then
        picServerTool.Width = picStatus.ScaleWidth
        shpServerInfo.Width = picStatus.Width

        shpServer.Width = picStatus.ScaleWidth - 144
        shpServer2.Width = shpServer.Width
        lblServer.Width = shpServer.Width - 6
        txtServer.Width = lblServer.Width
        
        shpPort.Left = picStatus.ScaleWidth - 143
        shpPort2.Left = shpPort.Left
        txtPort.Left = shpPort.Left + 4
        lblPort.Left = txtPort.Left
        
        shpPassword.Left = picStatus.ScaleWidth - 105
        shpPassword2.Left = shpPassword.Left
        txtPassword.Left = shpPassword.Left + 4
        lblPassword.Left = txtPassword.Left
        lblPassword.Width = txtPassword.Width
        
        shpHide.Left = picStatus.ScaleWidth - 17
        lblHide.Left = shpHide.Left
        picServerTool.Visible = True
        rt_Output.Top = 22
    Else
        picServerTool.Visible = False
        rt_Output.Top = 2
    End If
    
End Sub


Sub SendData(strData As String)
    If strData = "" Then Exit Sub
    
    If socket.State = sckConnected Then
        socket.SendData strData & vbCrLf
        If bDebug Then
            With frmDebug.txtDebugOut
                .selStart = Len(.Text)
                .selLength = 0
                .seltext = vbCrLf & strData
            End With
        End If
    Else
        PutText rt_Output, "2* Not connected to server"
    End If
End Sub

Public Sub SetServerInfo(strServer As String, Optional intport As Integer = 6667, Optional strPassword As String = "")
    strServerName = strServer
    If strServer = "" Then strServerName = txtServer.Text
    intServerPort = intport
    strserverpass = strPassword
    
    txtServer = strServer
    txtPort = intport
    txtPassword = strPassword
    
    '* set dynamic menu info
    XPM_Window.SetText 15 + serverID, strServer
    
    Toolbar.Buttons(1).ToolTipText = "Connect to the server " & strServer
    
    Dim nInd As Integer
    nInd = treeview_GetStatusIndex(CLIENT.tvServers, serverID)
    If nInd <> -1 Then CLIENT.tvServers.Nodes(nInd).Text = serverID & ": " & strServer
End Sub

Function WinType() As String
    WinType = "Status"
End Function






Private Sub Form_Activate()
    
    On Error Resume Next
    treeview_SetActive CLIENT.tvServers, "Status", serverID
    
    bNewData = False
    CurrentServerID = serverID
    CLIENT.SetActive strTitle, serverID
    rt_Input.setFocus
    
    CLIENT.DrawMenu
    
    XPM_View.SetText 4, "Server Bar"
    XPM_View.SetCheck 4, picServerTool.Visible
    XPM_View.SetDisable 4, False
End Sub

Private Sub Form_GotFocus()
    Resize_event
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If HandleHotkey(ShiftKey, KeyCode, serverID) Then
        KeyCode = 0
    End If
End Sub




Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ShiftKey = 0
End Sub

Private Sub Form_Load()

    oldProcAddr = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf StatWndProc)
    'oldLVProcAddr = SetWindowLong(lvChannels.hwnd, GWL_WNDPROC, AddressOf LVWndProc)
    'SetWindowLong lvChannels.hwnd, GWL_STYLE, GetWindowLong(lvChannels.hwnd, GWL_STYLE) Or LVS_OWNERDRAWFIXED

    strTitle = "Status"

    '* reset window pos
    SetServerInfo ServerAddr, CLng(ServerPort), ""
    WP_ResetWindow Me
    
    bShowInTaskbar = True
    'bShowServerInfo = True
    
    ScaleMode = 3
    PutText rt_Output, "02Welcome to sIRC.  You are using alpha v0." & String(2 - Len(CStr(App.Revision)), Asc("0")) & App.Revision & "." & lngBuild
    
    Timer1.Interval = GetCaretBlinkTime()
    strServerName = ServerAddr
    intServerPort = ServerPort
    lblServer = ServerAddr
    txtServer = ServerAddr
    lblPort = ServerPort
    txtPort = ServerPort
        
    CLIENT.DrawTaskbarAllServers
    
    
End Sub


Private Sub Form_Resize()
    Resize_event
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    treeview_RemoveServer CLIENT.tvServers, serverID
    
    SendData "QUIT :sIRC alpha v0." & String(2 - Len(CStr(App.Revision)), Asc("0")) & App.Revision & "." & lngBuild
    bShowInTaskbar = False
    
    On Error Resume Next
    
    Dim blah As Form
    For Each blah In Forms
        If TypeOf blah Is MDIForm Then GoTo nextt
        On Error GoTo nextt
        If blah.serverID = serverID Then
            SetWindowLong Me.hwnd, GWL_WNDPROC, oldProcAddr
            SetWindowLong lvChannels.hwnd, GWL_WNDPROC, oldLVProcAddr
            Unload blah
        End If
nextt:
    On Error Resume Next
    Next blah
    
    serverID = 0
    
    XPM_Window.SetVisible 15 + serverID, False
    Connections(serverID) = False
    CLIENT.DrawTaskbarAllServers
    
    Dim i As Integer
    For i = 1 To 50
        DoEvents
    Next i
    socket.Close
End Sub

Private Sub lblHide_Click()
    bShowServerInfo = False
    ResizeToolbar
    Form_Resize
    XPM_View.SetCheck 3, False
End Sub

Private Sub lblPort_DblClick()
    'txtPort.Visible = True
    'txtPort.setFocus
    'txtPort.Tag = ""
    'shpPort.Visible = True
    'txtServer.Visible = False
    'shpServer.Visible = False
    'shpPassword.Visible = False
    'txtPassword.Visible = False
End Sub




Private Sub lvChannels_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvChannels.SortKey = ColumnHeader.Index - 1
    If lvChannels.SortOrder = lvwAscending Then
        lvChannels.SortOrder = lvwDescending
    Else
        lvChannels.SortOrder = lvwAscending
    End If
    lvChannels.Sorted = True
End Sub


Private Sub rt_Input_Change()
    rt_Input.Font.Name = strFontName
    rt_Input.Font.Size = intFontSize
End Sub

Private Sub rt_Input_GotFocus()
    HideAllServerInfo
End Sub

Private Sub rt_Input_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If Shift <> 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    
    If KeyCode = 38 Then    'UP KEY!
        If intCurHist <= 1 Then Beep: Exit Sub
        
        If intCurHist >= textHistory.Count And rt_Input.Text <> "" Then
            If AddHistory(rt_Input.Text) Then
                intCurHist = intCurHist - 1
            End If
        End If
        
        intCurHist = intCurHist - 1
        rt_Input.Text = textHistory.item(intCurHist)
        rt_Input.selStart = Len(rt_Input.Text)
        KeyCode = 0
    ElseIf KeyCode = 40 Then    'down key!
        If intCurHist >= textHistory.Count Then
            If rt_Input.Text <> "" Then
                If AddHistory(rt_Input.Text) = False Then
                    intCurHist = intCurHist + 1
                End If
                rt_Input.Text = ""
                KeyCode = 0
            Else
                Beep
            End If
            Exit Sub
        End If
        
        intCurHist = intCurHist + 1
        rt_Input.Text = textHistory.item(intCurHist)
        rt_Input.selStart = Len(rt_Input.Text)
        KeyCode = 0
    End If

End Sub

Private Sub rt_Input_KeyPress(KeyAscii As Integer)
    Dim params(3) As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If rt_Input.Text = "" Then Exit Sub
        
        AddHistory rt_Input.Text
        
        If Left$(rt_Input.Text, Len(COMMANDCHAR)) = COMMANDCHAR Then
        
            If Len(rt_Input.Text) = 1 Then Exit Sub
            Dim strData As String
            strData = Right$(rt_Input.Text, Len(rt_Input.Text) - 1)
            Dim argsX() As String
            argsX = Split(strData, " ")
            
            If DoCommandLine(argsX, "Status", serverID) = False Then
                windowStatus(serverID).SendData strData
            End If
        Else
            If socket.State <> sckConnected Then Exit Sub
            
            windowStatus(serverID).SendData rt_Input.Text
        End If
        
        rt_Input.Text = ""
    End If
End Sub


Private Sub rt_Input_KeyUp(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If Shift <> 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
End Sub



Private Sub rt_Input_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim point As POINTAPI
        GetCursorPos point
        XPM_Edit.ShowMenu point.x * 15, point.y * 15
    End If
End Sub

Private Sub rt_Input_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Data.GetFormat(1) = True Then rt_Input.seltext = Data.GetData(1)
End Sub

Private Sub rt_Input_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    rt_Input.setFocus
End Sub




Private Sub rt_Output_Click()
    If whichButton = 2 Then
        Dim point As POINTAPI
        GetCursorPos point
        bMenuShown = False
        bPopupmenu = False
        XPM_Edit.ShowMenu point.x * 15, point.y * 15
        Exit Sub
    End If

End Sub

Private Sub rt_Output_GotFocus()
    HideAllServerInfo
End Sub

Private Sub rt_Output_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> 9 Then
        KeyCode = 0
    End If
    
End Sub

Private Sub rt_Output_KeyPress(KeyAscii As Integer)
    rt_Input.setFocus
    rt_Input.seltext = Chr(KeyAscii)
    KeyAscii = 0
End Sub


Private Sub rt_Output_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HideCaret rt_Output.hwnd
    startX = x
    startY = y
End Sub


Private Sub rt_Output_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then HideCaret rt_Output.hwnd
End Sub


Private Sub rt_Output_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    whichButton = Button
   
    If startX = x And startY = y Then
        rt_Input.setFocus
    End If
    
End Sub

Private Sub socket_Close()
'    PutText rt_Output, "2* Connection with server has been closed."
    
    '* Testing
    bConnected = False
    Dim vars(0) As String
    vars(0) = ":"
    scriptEngine.ExecuteEvent "disconnect", strEmptyParams, serverID, vars
    Toolbar.Buttons(2).value = tbrPressed
    Toolbar.Buttons(1).value = tbrUnpressed
End Sub

Private Sub socket_Connect()
'    PutText rt_Output, "2* Connected to server"
    
    Dim localVars(3) As String
    localVars(0) = "server:" & socket.RemoteHost
    localVars(1) = "serverip:" & socket.RemoteHostIP
    localVars(2) = "port:" & socket.RemotePort
    localVars(3) = "localport:" & socket.LocalPort
    
    scriptEngine.ExecuteEvent "connect", strEmptyParams, serverID, localVars
    
    currentNick = 0
    
    SendData "USER " & strNicks(currentNick) & " irc local :" & strEmail
    SendData "NICK " & strNicks(currentNick)
    strCurNick = strNicks(currentNick)
    
    currentNick = currentNick + 1
    
    CLIENT.DrawTaskbarAllServers
    
    Toolbar.Buttons(1).value = tbrPressed
    Toolbar.Buttons(2).value = tbrUnpressed
    
End Sub

Private Sub socket_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, AllParams As String
    Dim strData() As String, i As Integer
    
    'lngPingReply = GetTickCount
    
    On Error Resume Next
    If socket.State = sckConnected Then socket.GetData dat, vbString
    
    'If DebugWin.Visible Then
    '    If Len(DebugWin.txtDataIn.Text) + Len(dat) > 60000 Then DebugWin.txtDataIn.Text = Right$(DebugWin.txtDataIn.Text, 10000)
    '    On Error Resume Next
    '    If InStr(dat, "303 " & strMyNick & " :") Then Else DebugWin.txtDataIn = DebugWin.txtDataIn & "<< INCOMING DATA << " & vbCrLf
    '    DebugWin.txtDataIn.selStart = Len(DebugWin.txtDataIn)
    'End If
    
    strDataBuffer = strDataBuffer & dat
    
    If Right$(strDataBuffer, 1) <> Chr(13) And Right$(strDataBuffer, 1) <> Chr(10) Then
        Exit Sub
    Else
        dat = strDataBuffer
        strDataBuffer = ""
    End If
    
    strData = Split(dat, Chr(10))
    
    For i = LBound(strData) To UBound(strData)
        'interpret strData(i)
        CommandQueue.Add strData(i)
        If tmrCommandQueue.Enabled = False Then tmrCommandQueue.Enabled = True
        'PutText rt_Output, strData(i)
    Next i


End Sub

Private Sub socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'PutText rt_Output, "4* Socket Error (" & Number & "): " & Description & ""
    
    Dim XParams(2) As String, vars(2) As String
    XParams(1) = Number
    XParams(2) = Description
    vars(0) = "number:" & Number
    vars(1) = "desc:" & Description
    vars(2) = "source:" & Source
    scriptEngine.ExecuteEvent "connecterror", XParams, serverID, vars
    'scriptEngine.ExecuteEvent "connect", strEmptyParams
End Sub

Private Sub Timer1_Timer()
    HideCaret rt_Output.hwnd
End Sub


Private Sub tmrCommandQueue_Timer()
    interpret CommandQueue.item(1)
    CommandQueue.Remove 1
    
    If CommandQueue.Count = 0 Then tmrCommandQueue.Enabled = False
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Me.Connect
            Toolbar.Buttons(1).value = tbrPressed
            Toolbar.Buttons(2).value = tbrUnpressed
        Case 2
            Me.Disconnect
            Toolbar.Buttons(2).value = tbrPressed
            Toolbar.Buttons(1).value = tbrUnpressed
        Case 4
            HideTabs
            picStatus.Visible = True
        Case 5
            HideTabs
            picChannelList.Visible = True
    End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.Tag = "'"
        'txtPassword.Visible = False
        'shpPassword.Visible = False
        lblPassword.Caption = String(Len(txtPassword.Text), txtPassword.PasswordChar)
        KeyAscii = 0
        SetServerInfo txtServer.Text, txtPort.Text, txtPassword.Text
    End If
End Sub


Private Sub txtPassword_LostFocus()
'    txtPassword.Visible = False
'    shpPassword.Visible = False
    lblPassword.Caption = String(Len(txtPassword.Text), txtPassword.PasswordChar)
End Sub


Private Sub txtPassword_Validate(Cancel As Boolean)
    If txtPassword.Tag = "" Then Cancel = True
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPort.Tag = "'"
        'txtPort.Visible = False
        'shpPort.Visible = False
        lblPort = txtPort.Text
        KeyAscii = 0
        SetServerInfo txtServer.Text, txtPort.Text, txtPassword.Text
    End If
End Sub


Private Sub txtPort_LostFocus()
'    txtPort.Visible = False
'    shpPort.Visible = False
    lblPort = txtPort.Text
End Sub


Private Sub txtPort_Validate(Cancel As Boolean)
    If txtPort.Tag = "" Then Cancel = True
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtServer.Tag = "'"
        'txtServer.Visible = False
        'shpServer.Visible = False
        lblServer.Caption = txtServer.Text
        KeyAscii = 0
        SetServerInfo txtServer.Text, txtPort.Text, txtPassword.Text
        CLIENT.DrawTaskbarAllServers
        rt_Input.setFocus
    End If
End Sub

Private Sub txtServer_LostFocus()
    'txtServer.Visible = False
    'shpServer.Visible = False
    lblServer.Caption = txtServer.Text
End Sub


Private Sub txtServer_Validate(Cancel As Boolean)
    If txtServer.Tag = "" Then Cancel = True
End Sub


