VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "sIRC - New Profile"
   ClientHeight    =   4050
   ClientLeft      =   2370
   ClientTop       =   4575
   ClientWidth     =   6150
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   0
      Picture         =   "frmNewProfile.frx":038A
      ScaleHeight     =   3510
      ScaleWidth      =   1545
      TabIndex        =   7
      ToolTipText     =   "text art by rory (raw@the-flipside.co.uk)"
      Top             =   0
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3630
      Width           =   1065
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<  &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3630
      Width           =   1065
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next  >"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3630
      Width           =   1065
   End
   Begin VB.PictureBox picForm 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   2
      Left            =   1680
      ScaleHeight     =   3015
      ScaleWidth      =   4470
      TabIndex        =   22
      Top             =   390
      Visible         =   0   'False
      Width           =   4470
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   3690
         TabIndex        =   26
         Top             =   2610
         Width           =   645
      End
      Begin VB.TextBox txtServerName 
         Height          =   315
         Left            =   840
         TabIndex        =   24
         Top             =   2610
         Width           =   2790
      End
      Begin MSComctlLib.TreeView tvServers 
         Height          =   2460
         Left            =   120
         TabIndex        =   23
         Top             =   60
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   4339
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   6
         Appearance      =   1
      End
      Begin VB.Label lblServerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server&:"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   2640
         Width           =   540
      End
   End
   Begin VB.PictureBox picForm 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   1680
      ScaleHeight     =   3015
      ScaleWidth      =   4395
      TabIndex        =   8
      Top             =   390
      Width           =   4395
      Begin VB.TextBox txtProfileName 
         Height          =   315
         Left            =   1005
         MaxLength       =   35
         TabIndex        =   10
         Text            =   "New User"
         Top             =   1350
         Width           =   1950
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Profile Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1005
         TabIndex        =   9
         Top             =   1080
         Width           =   1125
      End
   End
   Begin VB.PictureBox picForm 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   1
      Left            =   1680
      ScaleHeight     =   3015
      ScaleWidth      =   4395
      TabIndex        =   14
      Top             =   390
      Visible         =   0   'False
      Width           =   4395
      Begin VB.CommandButton cmdNicksDOWN 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1755
         Width           =   240
      End
      Begin VB.CommandButton cmdNicksUP 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1455
         Width           =   240
      End
      Begin VB.TextBox txtNickToAdd 
         Height          =   285
         Left            =   2145
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1110
         Width           =   2220
      End
      Begin VB.CommandButton cmdNicksClear 
         Caption         =   "c"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2655
         Width           =   240
      End
      Begin VB.CommandButton cmdNicksRemove 
         Caption         =   "-"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2355
         Width           =   240
      End
      Begin VB.CommandButton cmdNicksAdd 
         Caption         =   "+"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2055
         Width           =   240
      End
      Begin VB.ListBox lstNicks 
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   2145
         TabIndex        =   18
         Top             =   1470
         Width           =   1905
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   2130
         MaxLength       =   35
         TabIndex        =   1
         Top             =   420
         Width           =   2220
      End
      Begin VB.TextBox txtYourName 
         Height          =   285
         Left            =   2130
         MaxLength       =   25
         TabIndex        =   0
         Top             =   45
         Width           =   2220
      End
      Begin VB.Label lblNickInfo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNewProfile.frx":1630
         Height          =   1440
         Left            =   195
         TabIndex        =   27
         Top             =   1425
         Width           =   1845
      End
      Begin VB.Label lblNicks 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Nicks&:"
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Top             =   1125
         Width           =   945
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   285
         X2              =   4300
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   285
         X2              =   4300
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Email&:"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   450
         Width           =   1305
      End
      Begin VB.Label lblYourName 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Name&:"
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   75
         Width           =   1305
      End
   End
   Begin VB.Label lblStep 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   210
      Left            =   5610
      TabIndex        =   13
      Top             =   90
      Width           =   180
   End
   Begin VB.Label lblSteps 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   210
      Left            =   5940
      TabIndex        =   12
      Top             =   90
      Width           =   180
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Step    of"
      Height          =   210
      Left            =   5325
      TabIndex        =   11
      Top             =   90
      Width           =   825
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   105
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   180
      X2              =   347
      Y1              =   13
      Y2              =   13
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   180
      X2              =   347
      Y1              =   14
      Y2              =   14
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   -2
      X2              =   498
      Y1              =   234
      Y2              =   234
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   500
      Y1              =   235
      Y2              =   235
   End
End
Attribute VB_Name = "frmNewProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curForm As Integer
Const maxForm = 3

Public bShowInTaskbar As Boolean
Public serverID     As Integer
Public strTitle As String
Sub CreateProfile()
    Dim strServerINI As String

    If FileExists(PATH & txtProfileName.Text & "-settings.ini") Then
        MsgBox "A Profile with that name already exists, please choose another name", vbCritical
        Exit Sub
    End If
    
    strINI = PATH & txtProfileName.Text & "-settings.ini"
    strServerINI = PATH & txtProfileName.Text & "-servers.ini"
    
    '* create settings ini
    Open strINI For Output As #1
        Print #1, ""
    Close #1
        
    '* write server info
    Dim strNicks As String, i As Integer
    For i = 0 To lstNicks.ListCount - 1
        strNicks = strNicks & lstNicks.List(i) & ","
    Next i
    
    Open strServerINI For Output As #1
        Print #1, ""
    Close #1
    
    PutINI strServerINI, "All", "name", txtYourName.Text
    PutINI strServerINI, "All", "email", txtEmail.Text
    PutINI strServerINI, "All", "nicks", strNicks
    PutINI strServerINI, "All", "address", txtServerName.Text
    PutINI strServerINI, "All", "port", txtPort.Text
End Sub


Sub Enable_Step2()
    If txtYourName.Text = "" Or _
        txtEmail.Text = "" Or _
        lstNicks.ListCount = 0 Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
End Sub


Sub Enable_Step3()
    If txtServerName.Text = "" Or _
        txtPort.Text = "" Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
End Sub


Sub SetTitle(intStep As Integer)
    Select Case intStep
        Case 1:
            lblTitle.Caption = "Profile Name"
        Case 2:
            lblTitle.Caption = "Your Information"
        Case 3:
            lblTitle.Caption = "Server"
        Case 4:
            lblTitle.Caption = "IRC Settings"
    End Select
    Line1.X1 = lblTitle.Width + 118
    Line2.X1 = lblTitle.Width + 118
End Sub



Private Sub cmdBack_Click()
    curForm = curForm - 1
    SetTitle curForm
    picForm(curForm - 1).Visible = True
    picForm(curForm).Visible = False
    lblStep = curForm
    If curForm <> maxForm Then
        cmdNext.Caption = "&Next >>"
    End If
    If curForm = 1 Then
        If txtProfileName.Text <> "" Then
            cmdNext.Enabled = True
        End If
        cmdBack.Enabled = False
    ElseIf curForm = 2 Then
        Enable_Step2
    End If
End Sub

Private Sub cmdNext_Click()
    
    If curForm = 1 Then
        If FileExists(PATH & txtProfileName.Text & "-settings.ini") Then
            MsgBox "A Profile with that name already exists, please choose another name", vbCritical
            Exit Sub
        End If
    End If

    If cmdNext.Caption = "&Finish" Then
        Call CreateProfile      'Create the profile
        strProfile = txtProfileName.Text
        Unload Me
        Unload frmLoadProfile
        CLIENT.Show
    Else ' Caption = Next
        cmdBack.Enabled = True
        picForm(curForm - 1).Visible = False
        picForm(curForm).Visible = True
        curForm = curForm + 1
        SetTitle curForm
        lblStep = curForm
        If curForm = maxForm Then
            cmdNext.Caption = "&Finish"
        End If
        cmdNext.Enabled = False
    End If
    
    If curForm = 2 Then
        Call Enable_Step2
        txtYourName.setFocus
    ElseIf curForm = 3 Then
        Call Enable_Step3
        On Error Resume Next
        tvServers.setFocus
    End If
End Sub

Private Sub cmdNicksAdd_Click()
    Dim i As Integer
    For i = 0 To lstNicks.ListCount - 1
        If LCase(lstNicks.List(i)) = LCase(txtNickToAdd.Text) Then Exit Sub
    Next i
    lstNicks.AddItem txtNickToAdd.Text
    txtNickToAdd.Text = ""
    cmdNicksClear.Enabled = True
    txtNickToAdd.setFocus
    DoEvents
    Enable_Step2
    txtNickToAdd.setFocus
    
    If lstNicks.ListIndex = -1 Then
        cmdNicksUP.Enabled = False
        cmdNicksDOWN.Enabled = False
    End If
    
    If lstNicks.ListIndex <> 0 Then
        cmdNicksUP.Enabled = True
    Else
        cmdNicksUP.Enabled = False
    End If
    If lstNicks.ListIndex <> lstNicks.ListCount - 1 Then
        cmdNicksDOWN.Enabled = True
    Else
        cmdNicksDOWN.Enabled = False
    End If
End Sub


Private Sub cmdNicksClear_Click()
    Dim intYesNo As Integer
    intYesNo = MsgBox("You are about to clear the list of nicks you created, are you sure you would like to do this?", vbYesNo Or vbQuestion)
    If intYesNo = vbYes Then
        lstNicks.Clear
        cmdNicksClear.Enabled = False
        cmdNicksRemove.Enabled = False
    End If
    Enable_Step2
End Sub

Private Sub cmdNicksDOWN_Click()
    Dim strTmp As String
    strTmp = lstNicks.List(lstNicks.ListIndex)
    lstNicks.List(lstNicks.ListIndex) = lstNicks.List(lstNicks.ListIndex + 1)
    lstNicks.List(lstNicks.ListIndex + 1) = strTmp
    lstNicks.ListIndex = lstNicks.ListIndex + 1
End Sub

Private Sub cmdNicksRemove_Click()
    Dim i As Integer, a As Integer
    For i = 0 To lstNicks.ListCount - 1
        If lstNicks.Selected(a) = True Then
            lstNicks.RemoveItem a
        Else
            a = a + 1
        End If
    Next i
    If lstNicks.ListIndex = -1 Then
        cmdNicksRemove.Enabled = False
    End If
    If lstNicks.ListCount = 0 Then
        cmdNicksClear.Enabled = False
    End If
    
    If lstNicks.ListIndex = -1 Then
        cmdNicksUP.Enabled = False
        cmdNicksDOWN.Enabled = False
    End If
    
    
    Enable_Step2
End Sub

Private Sub cmdNicksUP_Click()
    Dim strTmp As String
    strTmp = lstNicks.List(lstNicks.ListIndex)
    lstNicks.List(lstNicks.ListIndex) = lstNicks.List(lstNicks.ListIndex - 1)
    lstNicks.List(lstNicks.ListIndex - 1) = strTmp
    lstNicks.ListIndex = lstNicks.ListIndex - 1
End Sub

Private Sub Command2_Click()
    Unload Me
    frmLoadProfile.Show
End Sub

Private Sub Form_Load()
    LoadServers tvServers
    
    lblSteps.Caption = maxForm
    curForm = 1
    SetTitle curForm
    Center Me
    SendMessage txtProfileName.hwnd, WM_SETFOCUS, 0&, vbNullString
    
    ButtonizeForm Me
End Sub



Private Sub lstNicks_Click()
    If lstNicks.ListIndex <> -1 Then
        cmdNicksRemove.Enabled = True
        If lstNicks.ListIndex <> 0 Then
            cmdNicksUP.Enabled = True
        Else
            cmdNicksUP.Enabled = False
        End If
        If lstNicks.ListIndex <> lstNicks.ListCount - 1 Then
            cmdNicksDOWN.Enabled = True
        Else
            cmdNicksDOWN.Enabled = False
        End If
    End If
End Sub

Private Sub picForm_Click(Index As Integer)
    SendMessage txtProfileName.hwnd, WM_SETFOCUS, 0&, vbNullString
End Sub

Private Sub tvServers_Click()
    Dim strDat As String
    On Error Resume Next
    If tvServers.selectedItem.parent Then DoEvents
    strDat = tvServers.selectedItem.Text
    If InStr(strDat, ".") = 0 Then
        
        Exit Sub
    End If
    If InStr(strDat, " (") Then
        txtServerName = LeftOf(strDat, " (")
        txtPort = RightOf(strDat, " (")
        txtPort = Left$(txtPort, Len(txtPort) - 1)
        If Left$(txtPort, 1) = "(" Then txtPort = Right$(txtPort, Len(txtPort) - 1)
        
    Else
        txtServerName = strDat
        txtPort = "6667"
    End If
    
    
End Sub

Private Sub txtEmail_Change()
    Enable_Step2
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEmail <> "" Then
            KeyAscii = 0
            txtNickToAdd.setFocus
        End If
    End If
End Sub


Private Sub txtNickToAdd_Change()
    If txtNickToAdd = "" Then
        cmdNicksAdd.Enabled = False
    Else
        cmdNicksAdd.Enabled = True
    End If
End Sub

Private Sub txtNickToAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtNickToAdd = "" Then Exit Sub
        Dim i As Integer
        For i = 0 To lstNicks.ListCount - 1
            If LCase(lstNicks.List(i)) = LCase(txtNickToAdd.Text) Then Exit Sub
        Next i
        lstNicks.AddItem txtNickToAdd.Text
        txtNickToAdd.Text = ""
        cmdNicksClear.Enabled = True
        Enable_Step2
    End If
End Sub


Private Sub txtPort_Change()
    Enable_Step3
End Sub

Private Sub txtProfileName_Change()
    If txtProfileName.Text <> "" Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
End Sub

Private Sub txtProfileName_GotFocus()
    If txtProfileName.Tag = "" Then
        txtProfileName.Tag = "z"
        txtProfileName.Text = ""
    End If
End Sub


Private Sub txtProfileName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtProfileName = "" Then
        Else
            KeyAscii = 0
            cmdNext_Click
        End If
    End If
End Sub


Private Sub txtServerName_Change()
    Enable_Step3
End Sub

Private Sub txtYourName_Change()
    Enable_Step2
End Sub


Private Sub txtYourName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtYourName <> "" Then
            KeyAscii = 0
            txtEmail.setFocus
        End If
    End If
End Sub


