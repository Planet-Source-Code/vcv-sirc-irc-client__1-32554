VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChanCentral 
   Caption         =   "Channel Properties"
   ClientHeight    =   4680
   ClientLeft      =   7590
   ClientTop       =   2745
   ClientWidth     =   6060
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   Begin VB.PictureBox picTopic 
      BorderStyle     =   0  'None
      Height          =   3585
      Left            =   135
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   386
      TabIndex        =   10
      Top             =   435
      Visible         =   0   'False
      Width           =   5790
      Begin VB.ListBox lstTopicHistory 
         Height          =   2445
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   12
         Top             =   450
         Width           =   5655
      End
      Begin RichTextLib.RichTextBox rtbTopic 
         Height          =   345
         Left            =   45
         TabIndex        =   11
         Top             =   75
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   609
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmChanCentral.frx":0000
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thu Dec 20 12:12:12 2001"
         Height          =   195
         Left            =   675
         TabIndex        =   16
         Top             =   3285
         Width           =   1920
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick!ident@hostname.extension"
         Height          =   195
         Left            =   675
         TabIndex        =   15
         Top             =   3015
         Width           =   2325
      End
      Begin VB.Label lblSetDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   3285
         Width           =   450
      End
      Begin VB.Label lblSetUser 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   3015
         Width           =   435
      End
   End
   Begin VB.PictureBox picModes 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   105
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   4
      Top             =   420
      Width           =   5835
      Begin VB.TextBox txtLimit 
         Height          =   315
         Left            =   2730
         TabIndex        =   9
         Top             =   3225
         Width           =   450
      End
      Begin VB.TextBox txtKey 
         Height          =   315
         Left            =   540
         TabIndex        =   7
         Top             =   3225
         Width           =   1470
      End
      Begin MSComctlLib.ListView lvModes 
         Height          =   3000
         Left            =   90
         TabIndex        =   5
         Top             =   105
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mode"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lblLimit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Limit:"
         Height          =   195
         Left            =   2220
         TabIndex        =   8
         Top             =   3270
         Width           =   375
      End
      Begin VB.Label lblKey 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Key:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   3270
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2850
      TabIndex        =   3
      Top             =   4245
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3885
      TabIndex        =   2
      Top             =   4245
      Width           =   990
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   4245
      Width           =   990
   End
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   4080
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   7197
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Modes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Topic"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ban List"
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
End
Attribute VB_Name = "frmChanCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strChanModes As String
Public Sub AddModeType(strMode As String)
    Dim newItem As ListItem
    Set newItem = lvModes.ListItems.Add(, , strMode)
    newItem.SubItems(1) = GetINI(PATH & "modes.inf", "cmdesc", strMode, "")
    
End Sub


Private Sub AddModeTypes()
    Dim i As Integer
    For i = 1 To Len(strChanModes)
        Select Case Mid(strChanModes, i, 1)
            Case "v", "h", "o", "b", "q", "k", "l"
            Case Else
                AddModeType Mid(strChanModes, i, 1)
        End Select
    Next i
End Sub

Private Sub HideTabs()
    picModes.Visible = False
    picTopic.Visible = False
End Sub









Private Sub Form_Load()
    '* Change
    strChanModes = "bHiklmnopqrsStv"
    
    AddModeTypes
    lvModes.Sorted = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    '* buttons, tabstrip, etc
    tabStrip.Move tabStrip.Left, tabStrip.Top, Me.ScaleWidth - 7, Me.ScaleHeight - 40
    cmdOk.Move Me.ScaleWidth - 214, Me.ScaleHeight - 29
    cmdCancel.Move Me.ScaleWidth - 145, Me.ScaleHeight - 29
    cmdApply.Move Me.ScaleWidth - 76, Me.ScaleHeight - 29
    picModes.Width = tabStrip.Width - 10
    picModes.Height = tabStrip.Height - 27
    picTopic.Width = tabStrip.Width - 10
    picTopic.Height = tabStrip.Height - 27
    
    '* picmodes stuff
    lvModes.ColumnHeaders.item(2).Width = lvModes.Width - 60
    lvModes.Move lvModes.Left, lvModes.Top, picModes.ScaleWidth - 10, picModes.ScaleHeight - 45
    lblKey.Top = picModes.ScaleHeight - 26
    lblLimit.Top = picModes.ScaleHeight - 26
    txtKey.Top = picModes.ScaleHeight - 29
    txtLimit.Top = picModes.ScaleHeight - 29
    
    '* picTopic stuff
    rtbTopic.Width = picTopic.ScaleWidth - 8
    lstTopicHistory.Move lstTopicHistory.Left, lstTopicHistory.Top, picTopic.ScaleWidth - 10, picTopic.ScaleHeight - 80
    lblSetUser.Top = picTopic.ScaleHeight - 44
    lblUser.Top = picTopic.ScaleHeight - 44
    lblSetDate.Top = picTopic.ScaleHeight - 26
    lblDate.Top = picTopic.ScaleHeight - 26
End Sub


Private Sub lvModes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvSort lvModes, ColumnHeader.Index
End Sub


Private Sub tabStrip_Click()
    HideTabs
    Select Case tabStrip.selectedItem.Index
        Case 1
            picModes.Visible = True
        Case 2
            picTopic.Visible = True
    End Select

End Sub


