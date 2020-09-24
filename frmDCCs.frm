VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDCCs 
   Caption         =   "DCC Transfers"
   ClientHeight    =   3450
   ClientLeft      =   6300
   ClientTop       =   5265
   ClientWidth     =   5625
   Icon            =   "frmDCCs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5625
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8343
            MinWidth        =   882
            Text            =   "0 Active Transfers / 0 Total Transfers"
            TextSave        =   "0 Active Transfers / 0 Total Transfers"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0%"
            TextSave        =   "0%"
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
   Begin MSComctlLib.ListView lvTransfers 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Progress"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User"
         Object.Width           =   3175
      EndProperty
   End
End
Attribute VB_Name = "frmDCCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub AddDCC(strUser As String, strFile As String)
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    lvTransfers.Move lvTransfers.Left, lvTransfers.Top, Me.ScaleWidth, Me.ScaleHeight - sbInfo.Height
    lvTransfers.ColumnHeaders.item(1).Width = Me.Width - 2900
End Sub


