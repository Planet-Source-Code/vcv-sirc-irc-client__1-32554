VERSION 5.00
Begin VB.Form frmLoadProfile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "sIRC - Load Profile"
   ClientHeight    =   3060
   ClientLeft      =   2220
   ClientTop       =   3375
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoadProfile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4440
   Begin VB.CheckBox chkLoadAuto 
      Caption         =   " Load this profile automatically from now on. "
      Enabled         =   0   'False
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   3990
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   45
      Picture         =   "frmLoadProfile.frx":038A
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   120
      Width           =   225
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1065
      Width           =   900
   End
   Begin VB.CommandButton cmdLoadProfile 
      Caption         =   "&Load Profile"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2595
      Width           =   1320
   End
   Begin VB.CommandButton cmdNewProfile 
      Caption         =   "&New Profile..."
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2595
      Width           =   1320
   End
   Begin VB.ListBox lstProfiles 
      Enabled         =   0   'False
      Height          =   1005
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   1
      Top             =   1065
      Width           =   3105
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   75
      X2              =   4400
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   75
      X2              =   4400
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   1550
      X2              =   4400
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   1550
      X2              =   4400
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Profiles"
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
      Left            =   360
      TabIndex        =   4
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLoadProfile.frx":05EC
      Height          =   675
      Left            =   255
      TabIndex        =   0
      Top             =   360
      Width           =   4305
   End
End
Attribute VB_Name = "frmLoadProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bShowInTaskbar As Boolean
Public serverID As Integer
Public strTitle As String
Private Sub cmdDelete_Click()
    Dim strSel As String
    strSel = lstProfiles.List(lstProfiles.ListIndex)
    
    If lstProfiles.ListIndex = -1 Then
    Else
        Dim intReturn As Integer
        intReturn = MsgBox("Are you sure you would like to delete the user profile named """ & strSel & """?  Doing so will delete the profile altogether, which is irreversible.", vbYesNo Or vbQuestion)
        
        If intReturn = vbYes Then
            On Error GoTo errhandler
            lstProfiles.RemoveItem lstProfiles.ListIndex
            Kill PATH & strSel & "-settings.ini"
            Kill PATH & strSel & "-windows.ini"
            Kill PATH & strSel & "-servers.ini"
            Exit Sub
errhandler:
            If Err.Number = 53 Then Exit Sub
            MsgBox "While trying to delete the user profile """ & strSel & """, an error occured: " & vbCrLf & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Source: " & Err.Source & vbCrLf & "Please email the author with this error.", vbCritical
        End If
    End If

End Sub

Private Sub cmdLoadProfile_Click()
    Dim strSel As String
    strSel = lstProfiles.List(lstProfiles.ListIndex)
    
    If lstProfiles.ListIndex = -1 Then
    Else
        If chkLoadAuto.value = 1 Then
            SaveSetting "sIRC", "options", "loadautoname", strSel
            SaveSetting "sIRC", "options", "loadauto", "true"
        Else
            SaveSetting "sIRC", "options", "loadauto", "false"
        End If
    
        strINI = PATH & strSel & "-settings.ini"
        strProfile = strSel
        winINI = PATH & strSel & "-windows.ini"
        Unload Me
        CLIENT.Show
    End If
End Sub

Private Sub cmdNewProfile_Click()
    frmNewProfile.Show
    Me.Hide
End Sub


Private Sub Form_Load()
    
    strTitle = "sIRC - Load Profile"
    
    Dim slash As String
    If Right$(App.PATH, 1) <> "\" Then slash$ = "\"
    PATH = App.PATH & slash$
    
    strGlobalINI = PATH & "sIRC.ini"
    
    cppButton cmdNewProfile
    cppButton cmdLoadProfile
    cppButton cmdDelete
    Center Me
        
    ' Load profiles..
    lstProfiles.Enabled = True
    Dim strFile As String, vDir As String
    strFile = Dir(PATH & "*-settings.ini", vbDirectory)
    Do While strFile <> ""
        lstProfiles.AddItem Left$(strFile, InStr(strFile, "-") - 1)
        strFile = Dir
    Loop
        
    If lstProfiles.ListCount = 0 Then
        SendMessage cmdNewProfile.hwnd, WM_SETFOCUS, 0, vbNullString
    End If
    
    '* Auto Load Profile
    Dim strAL As String, strName As String
    strAL = GetSetting("sIRC", "options", "loadauto")
    strName = GetSetting("sIRC", "options", "loadautoname")
    If strAL = "true" Then
        chkLoadAuto.Enabled = True
        chkLoadAuto.value = 1
        Dim i As Integer
        For i = 0 To lstProfiles.ListCount - 1
            If lstProfiles.List(i) = strName Then
                lstProfiles.ListIndex = i
                Exit Sub
            End If
        Next i
    End If
End Sub


Private Sub lstProfiles_Click()
    If lstProfiles.ListIndex <> -1 Then
        cmdDelete.Enabled = True
        chkLoadAuto.Enabled = True
        cmdLoadProfile.Enabled = True
    End If
        
End Sub


Private Sub lstProfiles_DblClick()
    Call cmdLoadProfile_Click
End Sub


