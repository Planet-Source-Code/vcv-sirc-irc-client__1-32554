VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form CHANNEL 
   BackColor       =   &H00FFFFFF&
   Caption         =   "#channel [#][+]: [topic]"
   ClientHeight    =   4770
   ClientLeft      =   4560
   ClientTop       =   5040
   ClientWidth     =   8025
   Icon            =   "frm_Channel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   Begin VB.PictureBox picSep2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   4110
      Left            =   6345
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   8
      Top             =   315
      Visible         =   0   'False
      Width           =   45
   End
   Begin RichTextLib.RichTextBox rt_Output 
      Height          =   4065
      Left            =   45
      TabIndex        =   4
      Top             =   330
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   7170
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   1
      TextRTF         =   $"frm_Channel.frx":038A
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
   Begin VB.PictureBox picNickSep 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   6390
      MousePointer    =   9  'Size W E
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   7
      Top             =   300
      Width           =   105
      Begin VB.Line ln_Sep2 
         BorderColor     =   &H8000000F&
         X1              =   3
         X2              =   3
         Y1              =   0
         Y2              =   277
      End
   End
   Begin VB.PictureBox picNickList 
      BorderStyle     =   0  'None
      Height          =   4080
      Left            =   6570
      ScaleHeight     =   272
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   5
      Top             =   330
      Width           =   1395
      Begin VB.ListBox tvNickList 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4185
         IntegralHeight  =   0   'False
         Left            =   -15
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   -15
         Width           =   1425
      End
   End
   Begin VB.Timer tmrCaret 
      Interval        =   500
      Left            =   2055
      Top             =   1695
   End
   Begin VB.PictureBox picServerTool 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7140
      Begin RichTextLib.RichTextBox txtTopic 
         Height          =   240
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   423
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frm_Channel.frx":0406
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   2
         Top             =   15
         Width           =   255
      End
      Begin VB.Shape shpTopic2 
         BorderColor     =   &H80000005&
         Height          =   270
         Left            =   0
         Top             =   15
         Width           =   3210
      End
      Begin VB.Label lblServer 
         Caption         =   "irc.othersideirc.net"
         Height          =   240
         Left            =   75
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Shape shpTopic 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   0
         Top             =   15
         Width           =   4365
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
         Width           =   255
      End
   End
   Begin RichTextLib.RichTextBox rt_Input 
      Height          =   270
      Left            =   30
      TabIndex        =   3
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
      TextRTF         =   $"frm_Channel.frx":0483
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
   Begin VB.Line ln_Sep 
      BorderColor     =   &H8000000F&
      X1              =   0
      X2              =   520
      Y1              =   295
      Y2              =   295
   End
End
Attribute VB_Name = "CHANNEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* TOPIC VARIABLES
Private CurrentTopic     As typTopic
Public topicCount       As Integer
Private topicHistory()   As typTopic

Public oldProcAddr      As Long
Private curWS           As Integer

Private intShift As Integer

Public bNewData As Boolean
Public intNewLines As Integer

Private intMB As Integer

Private startX As Long, startY As Long, OrigLeft As Long, OrigTop As Long

Const MF_STRING = &H0&
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public pOldProc As Long  ' pointer to Form's previous window procedure
Public ontop As Boolean  ' identifies if Form is always on top or not

Public bShowChannelInfo As Boolean
Const strType = "channel"
Public strChanName      As String
Public strTitle         As String

Public strServer        As String
Public serverID         As Integer
Public bWindowInUse     As Boolean

Public bShowInTaskbar   As Boolean
Public chanNum          As Integer

Private Nicks()  As typeNick
Private Type typeNick
    strNick     As String
    strHost     As String
    strIdent    As String
    bOP         As Boolean
    bHalfOP     As Boolean
    bVoice      As Boolean
    nListIndex  As Long
End Type
Public NickCount   As Long
Public realNickCnt As Long

Public strModes As String
Public strKey   As String
Public strLimit As String

Private textHistory As New Collection
Private intCurHist  As Integer
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


Public Function GetCurTopic() As String
    GetCurTopic = CurrentTopic.Topic
End Function

Public Function GetCurTopicSetBy() As String
    GetCurTopicSetBy = CurrentTopic.SetBy
End Function

Public Function GetCurTopicSetWhen() As String
    GetCurTopicSetWhen = CurrentTopic.SetWhen
End Function

Function GetModeString() As String
    Dim strModeX As String
    strModeX = strModes
    
    If strLimit <> "" Then
        strModeX = strModeX & "l"
    End If
    If strKey <> "" Then
        strModeX = strModeX & "k"
    End If
    If strLimit <> "" Then
        strModeX = strModeX & " " & strLimit
    End If
    If strKey <> "" Then
        strModeX = strModeX & " " & strKey
    End If
    
    If strModeX = "" Then
        GetModeString = ""
    Else
        GetModeString = "[" & strModeX & "]"
    End If
End Function

Public Sub AddNick(strNick As String, strHost As String, strIdent As String)
    Dim i As Integer
    
    For i = 1 To NickCount
        If Nicks(i).strNick = "" Then
            Nicks(i).strNick = strNick
            Nicks(i).bOP = False
            Nicks(i).bHalfOP = False
            Nicks(i).bVoice = False
            Nicks(i).strHost = strHost
            Nicks(i).strIdent = strIdent
            Exit Sub
        End If
    Next i
    
    NickCount = NickCount + 1
    realNickCnt = realNickCnt + 1
    ReDim Preserve Nicks(NickCount) As typeNick
    Nicks(NickCount).strNick = strNick
    Nicks(NickCount).bOP = False
    Nicks(NickCount).bHalfOP = False
    Nicks(NickCount).bVoice = False
    Nicks(NickCount).strHost = strHost
    Nicks(NickCount).strIdent = strIdent
    UpdateNickList
End Sub


Public Sub ChangeUserNick(strOldNick As String, strNewNick As String)
    
    Dim lngFV As Long
    
    Dim nickIndex As Long
    nickIndex = GetNickIndex(strOldNick)
    
    If strOldNick <> "" Then
        lngFV = SendMessage(tvNickList.hwnd, LB_FINDSTRINGEXACT, -1, ByVal GetDisplayNick(CInt(nickIndex)))
    End If
    
    If nickIndex <> -1 Then
        Nicks(nickIndex).strNick = strNewNick
    End If
    
    If lngFV > 0 Then
        tvNickList.List(lngFV) = GetDisplayNick(CInt(nickIndex))
    End If
    
    UpdateNickList
    
End Sub

Public Function GetTitle() As String
    GetTitle = strChanName

End Function
Public Sub DoMode(bAdd As Boolean, strChar As String, strMode As String)
    
    'On Error GoTo errhandler
    
    Dim lngFV, nickIndex As Long
    nickIndex = -1
    lngFV = -1
    
    If strMode <> "" Then
        nickIndex = GetNickIndex(strMode)
    
        If nickIndex = -1 Then Exit Sub
    
        If strMode <> "" And strChar <> "k" And strChar <> "l" Then
            lngFV = SendMessage(tvNickList.hwnd, LB_FINDSTRINGEXACT, -1, ByVal GetDisplayNick(CInt(nickIndex)))
            If lngFV = -1 Then Exit Sub
        End If
    End If
    
    Select Case strChar
    
        Case "o"
            If strMode = windowStatus(serverID).strCurNick Then
                txtTopic.Locked = Not bAdd
            End If
            
            Nicks(nickIndex).bOP = bAdd
        Case "v"
            Nicks(nickIndex).bVoice = bAdd
        Case "h"
            Nicks(nickIndex).bHalfOP = bAdd
        Case "a", "q"
        Case "k"
            If bAdd Then
                strKey = strMode
            Else
                strKey = ""
            End If
        Case "l"
            If bAdd Then
                strLimit = strMode
            Else
                strLimit = ""
            End If
        Case Else
            If bAdd Then
                If InStr(strModes, strChar) Then
                Else
                    strModes = strChar & strModes
                End If
            Else
                strModes = Replace$(strModes, strChar, "")
            End If
    End Select
        
    UpdateCaption
    
    If lngFV <> -1 Then
        tvNickList.RemoveItem lngFV
        tvNickList.AddItem GetDisplayNick(CInt(nickIndex))
        'tvNickList.List(lngFV) = GetDisplayNick(CInt(nickIndex))
    End If
    
    Exit Sub
errhandler:
    MsgBox Err & ": " & Error & ", " & Err.Source
End Sub

Function GetDisplayNick(nickIndex As Integer) As String
    If Nicks(nickIndex).bOP Then
        GetDisplayNick = "@" & Nicks(nickIndex).strNick
    ElseIf Nicks(nickIndex).bHalfOP Then
        GetDisplayNick = "%" & Nicks(nickIndex).strNick
    ElseIf Nicks(nickIndex).bVoice Then
        GetDisplayNick = "+" & Nicks(nickIndex).strNick
    Else
        GetDisplayNick = Nicks(nickIndex).strNick
    End If
End Function


Public Function GetNickIndex(strNickx As String) As Long
    Dim i As Integer
    For i = 0 To NickCount
        If Nicks(i).strNick = strNickx Then
            GetNickIndex = i
            Exit Function
        End If
    Next i
    
    GetNickIndex = -1
End Function

Public Function GetTopic(which As Integer) As String
    If which > topicCount Then GetTopic = "": Exit Function
    GetTopic = topicHistory(which).Topic
End Function

Public Function GetTopicSetBy(which As Integer) As String
    If which > topicCount Then GetTopicSetBy = "": Exit Function
    GetTopicSetBy = topicHistory(which).SetBy
End Function
Public Function GetTopicSetWhen(which As Integer) As String
    If which > topicCount Then GetTopicSetWhen = "": Exit Function
    GetTopicSetWhen = topicHistory(which).SetWhen
End Function
Public Function InChannel(strNick As String) As Boolean
    Dim nickIndex As Integer
    
    nickIndex = GetNickIndex(strNick)
    
'    MsgBox "InChannel function call" & vbCrLf & "nick:" & strNick & vbCrLf & "index:" & nickIndex
    
    If nickIndex = -1 Then
        InChannel = False
    Else
        InChannel = True
    End If
End Function

Public Sub RemoveNick(strNick As String)
    
    Dim lngFV As Long
    
    Dim nickIndex As Long
    nickIndex = GetNickIndex(strNick)
    
    If nickIndex = -1 Then Exit Sub
    
    If strNick <> "" Then
        lngFV = SendMessage(tvNickList.hwnd, LB_FINDSTRINGEXACT, -1, ByVal GetDisplayNick(CInt(nickIndex)))
    End If
            
    If lngFV > -1 Then
        tvNickList.RemoveItem lngFV
    End If
    
    If nickIndex <> -1 Then
        Nicks(nickIndex).strNick = ""
        realNickCnt = realNickCnt - 1
    End If
    
    UpdateCaption
    UpdateNickList
End Sub

Public Sub Resize_event()

    If Me.WindowState <> curWS Then
        If Me.WindowState = vbMaximized Then
            CLIENT.DrawMenu
        ElseIf Me.WindowState = vbNormal Then
            
        End If
        curWS = Me.WindowState
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ResizeToolbar
    DoEvents
    
    On Error Resume Next
    Dim intTextHeight As Integer
    Me.FontName = rt_Input.Font.Name
    Me.FontSize = rt_Input.Font.Size
    intTextHeight = Me.textHeight("Ab_¯") + 4     '# 4 for buffer (2 on top, 2 bottom)
        
    rt_Output.Width = Me.ScaleWidth - picNickList.ScaleWidth - 11 'buffer
    rt_Output.Height = Me.ScaleHeight - (intTextHeight + 7) - (rt_Output.Top - 1) '# 7 for buffer
    
    picNickList.Top = rt_Output.Top
    picNickList.Left = Me.ScaleWidth - picNickList.ScaleWidth - 2
    picNickList.Height = rt_Output.Height
    
    tvNickList.Width = picNickList.Width + 2
    tvNickList.Height = picNickList.Height + 2
    'tvNickList.ColumnHeaders(1).Width = tvNickList.Width - 4

    rt_Input.Top = Me.ScaleHeight - intTextHeight
    rt_Input.Height = intTextHeight + 4
    rt_Input.Width = Me.ScaleWidth - 4
    ln_Sep.Y1 = Me.ScaleHeight - (intTextHeight + 4)  '# 4 for buffer again
    ln_Sep.Y2 = ln_Sep.Y1
    ln_Sep.X2 = Me.ScaleWidth
    ln_Sep2.Y1 = 0
    ln_Sep2.Y2 = 10000
    picNickSep.Left = Me.ScaleWidth - picNickList.ScaleWidth - 9
    picNickSep.Top = picNickList.Top - 2
    picNickSep.Height = picNickList.Height + 4

End Sub

Public Sub SetFirstTopicInfo(strWho As String, strWhen As String)
    CurrentTopic.SetBy = strWho
    CurrentTopic.SetWhen = strWhen
End Sub

Sub SetNewTopic(strTopic As String, strWho As String)
    
    If topicCount = 0 Then
        CurrentTopic.SetBy = strWho
        CurrentTopic.Topic = strTopic
        CurrentTopic.SetWhen = CTime()
    End If
    
    topicCount = topicCount + 1
    ReDim Preserve topicHistory(1 To topicCount) As typTopic
    topicHistory(topicCount).SetBy = CurrentTopic.SetBy
    topicHistory(topicCount).SetWhen = CurrentTopic.SetWhen
    topicHistory(topicCount).Topic = CurrentTopic.Topic
    
    CurrentTopic.SetBy = strWho
    CurrentTopic.Topic = strTopic
    CurrentTopic.SetWhen = CTime()
End Sub

Public Sub UpdateCaption()
    Me.Caption = strChanName & " [" & realNickCnt & "] " & GetModeString
End Sub


Public Sub UpdateNickList()
    'tvNickList.Sorted = True
End Sub





Public Sub SetNick(strNickx As String, strHost As String, strIdent As String, Optional bOP As Boolean = False, Optional bHalfOP As Boolean = False, Optional bVoice As Boolean = False)
    Dim i As Integer, nickIndex As Integer
    
    If NickCount = 0 Then GoTo NoNicks
    nickIndex = NickCount + 1
    
    For i = LBound(Nicks) To NickCount
        If Nicks(i).strNick = strNickx Then
            Nicks(i).strHost = strHost
            Nicks(i).strIdent = strIdent
            realNickCnt = realNickCnt + 1
            UpdateCaption
            Exit Sub
        End If
        If Nicks(i).strNick = "" Then
            nickIndex = i
        End If
    Next i
NoNicks:
    
    realNickCnt = realNickCnt + 1
    If nickIndex >= NickCount Then
        NickCount = NickCount + 1
        ReDim Preserve Nicks(NickCount) As typeNick
        nickIndex = NickCount
    End If
    
    Nicks(nickIndex).strNick = strNickx
    Nicks(nickIndex).strHost = strHost
    Nicks(nickIndex).strIdent = strIdent
    Nicks(nickIndex).bHalfOP = bHalfOP
    Nicks(nickIndex).bOP = bOP
    Nicks(nickIndex).bVoice = bVoice
    UpdateCaption
    
    Call tvNickList.AddItem(GetDisplayNick(CInt(nickIndex)))
End Sub


Sub ResizeToolbar()
    On Error Resume Next
    If bShowChannelInfo Then
        picServerTool.Width = Me.ScaleWidth
        shpServerInfo.Width = Me.Width

        shpTopic.Width = Me.ScaleWidth - 18
        shpTopic2.Width = shpTopic.Width
        txtTopic.Width = shpTopic.Width - 3
        
        shpHide.Left = Me.ScaleWidth - 17
        lblHide.Left = shpHide.Left
        
        picServerTool.Visible = True
        rt_Output.Top = 22
    Else
        picServerTool.Visible = False
        rt_Output.Top = 2
    End If
    
End Sub


Sub SetCaption()
    Me.Caption = strChanName
End Sub

Function WinType() As String
    WinType = "Channel"
End Function

Private Sub Form_Activate()
    
    bNewData = False
    CurrentServerID = serverID
    CLIENT.SetActive strTitle, serverID
    rt_Input.setFocus
    
    Dim nIndex As Integer
    nIndex = treeview_GetChannelIndex(CLIENT.tvServers, strChanName, serverID)
    If nIndex <> -1 Then
        Set CLIENT.tvServers.selectedItem = CLIENT.tvServers.Nodes.item(nIndex)
        CLIENT.tvServers.Nodes.item(nIndex).ForeColor = vbBlack
    End If
    
    XPM_View.SetText 4, "Topic Bar"
    XPM_View.SetCheck 4, picServerTool.Visible
    XPM_View.SetDisable 4, False
    
    RESIZECLIENT
End Sub

Private Sub Form_GotFocus()

    CLIENT.SetActive strTitle, serverID
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    intShift = Shift
    If HandleHotkey(intShift, KeyCode, serverID) Then
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    intShift = 0
End Sub

Private Sub Form_Load()
    
    oldProcAddr = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf ChanWndProc)
    
    txtTopic = ""
    bShowChannelInfo = True
    
    WP_ResetWindow Me
    
    '* Let's properly draw the nick seperator thing
    Dim x As Integer, y As Integer, i As Long
    i = 1
    For x = 0 To picSep2.ScaleWidth
        For y = 0 To picSep2.ScaleHeight
            If i Mod 2 Then
                picSep2.ForeColor = vbBlack
            Else
                picSep2.ForeColor = vbWhite
            End If
            picSep2.PSet (x, y)
            i = i + 1
        Next
    Next
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    rt_Input.setFocus
End Sub

Private Sub Form_Resize()
    Resize_event
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Me.Tag <> "KICKED" Then
        SetWindowLong Me.hwnd, GWL_WNDPROC, oldProcAddr
        treeview_RemoveChannel CLIENT.tvServers, strChanName, serverID
    
        If serverID <> 0 Then
            windowStatus(serverID).SendData "PART " & strChanName
        End If
        strChanName = ""
        serverID = 0
        
        '* update dynamic menu
        XPM_ServerMenu(serverID).SetVisible chanNum + 2, False
        XPM_ServerMenu(serverID).SetText chanNum + 2, ""
        
        ChanInUse(chanNum) = False
    End If
    
    bWindowInUse = False
    bShowInTaskbar = False
    strTitle = ""
End Sub

Private Sub lblHide_Click()
    bShowChannelInfo = False
    ResizeToolbar
    Form_Resize
    XPM_View.SetCheck 3, False
End Sub


Private Sub picNickSep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    startX = x
    startY = y
    
    picSep2.Top = picNickSep.Top
    picSep2.Left = picNickSep.Left + 2
    picSep2.Height = picNickSep.Height
    
    picNickSep.Visible = False
    picSep2.Visible = True
End Sub


Private Sub picNickSep_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        picSep2.Left = picSep2.Left - (startX - x)
        startX = x
    End If
End Sub


Private Sub picNickSep_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    picNickSep.Left = picSep2.Left - 2
    
    picSep2.Visible = False
    picNickSep.Visible = True
    
    picNickList.Left = picNickSep.Left + 7
    picNickList.Width = Me.ScaleWidth - picNickSep.Left - 9
    tvNickList.Width = picNickList.Width + 2
    
    rt_Output.Width = picNickSep.Left - 2
End Sub

Private Sub rt_Input_Change()
    rt_Input.Font.Name = strFontName
    rt_Input.Font.Size = intFontSize
End Sub

Private Sub rt_Input_KeyDown(KeyCode As Integer, Shift As Integer)
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
        
        If Me.Tag = "KICKED" Then Exit Sub
        If rt_Input.Text = "" Then Exit Sub
        
        AddHistory rt_Input.Text
        
        If Left$(rt_Input.Text, Len(COMMANDCHAR)) = COMMANDCHAR Then
        
            If Len(rt_Input.Text) = 1 Then Exit Sub
            Dim strData As String
            strData = Right$(rt_Input.Text, Len(rt_Input.Text) - 1)
            Dim argsX() As String
            argsX = Split(strData, " ")
            
            If DoCommandLine(argsX, strChanName, serverID) = False Then
                windowStatus(serverID).SendData strData
            End If
        Else
            windowStatus(serverID).SendData "PRIVMSG " & strTitle & " :" & rt_Input.Text
            params(1) = strChanName
            params(2) = windowStatus(serverID).strCurNick
            params(3) = rt_Input.Text
            Dim vars(2) As String
            vars(0) = "target:" & strChanName
            vars(1) = "nick:" & windowStatus(serverID).strCurNick
            vars(2) = "message:" & rt_Input.Text
            scriptEngine.ExecuteEvent "text", params, serverID, vars
        End If
        
        rt_Input.Text = ""
    End If
End Sub



Private Sub rt_Input_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim point As POINTAPI
        GetCursorPos point
        XPM_Edit.ShowMenu point.x * 15, point.y * 15
    End If
End Sub

Private Sub rt_Output_Click()
    HideCaret rt_Output.hwnd
End Sub

Private Sub rt_Output_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HideCaret rt_Output.hwnd
    startX = x
    startY = y
End Sub

Private Sub rt_Output_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim point As POINTAPI
        GetCursorPos point
        XPM_Edit.ShowMenu point.x * 15, point.y * 15
    End If
    
    If startX = x And startY = y Then
        rt_Input.setFocus
    End If
    
End Sub



Private Sub txtTopic_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        windowStatus(serverID).SendData "TOPIC " & strChanName & " :" & txtTopic.Text
        rt_Input.setFocus
    End If
End Sub

Private Sub txtTopic_Validate(Cancel As Boolean)
    PutText_Reset txtTopic, CurrentTopic.Topic
End Sub


