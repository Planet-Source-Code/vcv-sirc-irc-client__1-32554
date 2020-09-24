VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug Window"
   ClientHeight    =   3375
   ClientLeft      =   10800
   ClientTop       =   8625
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDebugIn 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   390
      Width           =   6315
   End
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Incoming Data"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outgoing Data"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin VB.TextBox txtDebugOut 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   390
      Width           =   6315
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intPrevTop As Integer

Public strTitle As String
Public serverID As Integer

Private Sub Form_Resize()
    On Error Resume Next
    tabStrip.Width = Me.ScaleWidth
    txtDebugIn.Move 0, txtDebugIn.Top, Me.ScaleWidth, Me.ScaleHeight - 26
    txtDebugOut.Move 0, txtDebugOut.Top, Me.ScaleWidth, Me.ScaleHeight - 26
    cmdClose.Move Me.ScaleWidth - 18
    
End Sub


Private Sub tabStrip_Click()
    Select Case tabStrip.selectedItem.Index
        Case 1
            txtDebugIn.Visible = True
            txtDebugOut.Visible = False
        Case 2
            txtDebugIn.Visible = False
            txtDebugOut.Visible = True
    End Select
End Sub


Private Sub txtDebugIn_Change()
    txtDebugIn.selStart = Len(txtDebugIn)
End Sub


Private Sub txtDebugOut_Change()
    txtDebugOut.selStart = Len(txtDebugOut)
End Sub


