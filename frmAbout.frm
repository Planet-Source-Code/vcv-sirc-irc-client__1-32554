VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About sIRC"
   ClientHeight    =   4035
   ClientLeft      =   3795
   ClientTop       =   4410
   ClientWidth     =   5610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Ok"
      Height          =   390
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3510
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      Picture         =   "frmAbout.frx":038A
      ScaleHeight     =   1215
      ScaleWidth      =   5610
      TabIndex        =   0
      Top             =   0
      Width           =   5610
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "v0.11.3146"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4110
         TabIndex        =   2
         Top             =   780
         Width           =   1455
      End
   End
   Begin VB.Label lblWebpage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://sirc.ath.cx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   780
      MouseIcon       =   "frmAbout.frx":16770
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3585
      Width           =   1245
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":16A7A
      Height          =   1065
      Left            =   780
      TabIndex        =   6
      Top             =   2310
      Width           =   4440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001-2002"
      Height          =   195
      Left            =   780
      TabIndex        =   5
      Top             =   1815
      Width           =   1725
   End
   Begin VB.Label lblVerInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.08 (Build 1350 for 9x/ME/NT/2k/XP)"
      Height          =   195
      Left            =   780
      TabIndex        =   4
      Top             =   1575
      Width           =   3270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sIRC IRC Client"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   3
      Top             =   1350
      Width           =   1260
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bShowInTaskbar   As Boolean
Public serverID         As Integer

Public strEEBuffer As String
Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    strEEBuffer = strEEBuffer & Chr(KeyAscii)
    
    If InStr(strEEBuffer, "moo") Then
        PlayWaveRes "MOO"
        strEEBuffer = ""
        Exit Sub
    End If
    
    If Len(strEEBuffer) > 50 Then strEEBuffer = ""
End Sub

Private Sub Form_Load()

    lblVerInfo.Caption = "Version 0." & String(2 - Len(CStr(App.Revision)), Asc("0")) & App.Revision & " (Build " & lngBuild & " for 9x/ME/NT/2k/XP)"

    lblVersion.Caption = "v0." & String(2 - Len(CStr(App.Revision)), Asc("0")) & App.Revision & "." & lngBuild

    cppButton cmdOkay
    Center Me
    
End Sub


Private Sub TabStrip1_Click()

End Sub


Private Sub tsAbout_Click()
    Select Case tsAbout.selectedItem.Index
        Case 1
            If FileExists(PATH & "about.rtf") Then
                rtbAbout.LoadFile PATH & "about.rtf"
            End If
    End Select
End Sub




Private Sub lblWebpage_Click()
    ShellExecute 0, "open", "http://sirc.ath.cx", "", "", 0
End Sub


