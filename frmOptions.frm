VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4950
   ClientLeft      =   4470
   ClientTop       =   1305
   ClientWidth     =   6315
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
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   1995
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   405
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4470
      Width           =   810
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4470
      Width           =   810
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   405
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4470
      Width           =   810
   End
   Begin MSComctlLib.TreeView tvOptions 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   8281
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   450
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
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
   End
   Begin VB.PictureBox picDisplayMenu 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   1650
      ScaleHeight     =   3885
      ScaleWidth      =   4485
      TabIndex        =   23
      Top             =   450
      Visible         =   0   'False
      Width           =   4485
      Begin VB.CommandButton cmdMenuChangeFont 
         Caption         =   "&Change ..."
         Height          =   345
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2490
         Width           =   1020
      End
      Begin VB.PictureBox pic_MD_FontSample 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   210
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   183
         TabIndex        =   62
         Top             =   2460
         Width           =   2805
      End
      Begin VB.CheckBox chkMenuSink 
         Caption         =   " Sink when down"
         Height          =   255
         Left            =   555
         TabIndex        =   44
         Top             =   3315
         Width           =   1575
      End
      Begin VB.Label lbl_MD_Font 
         AutoSize        =   -1  'True
         Caption         =   "  Font  "
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
         Left            =   240
         TabIndex        =   61
         Top             =   2145
         Width           =   555
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000014&
         X1              =   45
         X2              =   4440
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   45
         X2              =   4440
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label lblTextDisabled 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disabled"
         Height          =   195
         Left            =   1755
         TabIndex        =   60
         Top             =   390
         Width           =   600
      End
      Begin VB.Label lbl_mc_TextDisabled 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1350
         TabIndex        =   59
         Top             =   375
         Width           =   345
      End
      Begin VB.Label lbl_mc_SepShadow 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2685
         TabIndex        =   58
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label lbl_mc_SepHilight 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1650
         TabIndex        =   57
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label lbl_mc_FaceDown 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   3120
         TabIndex        =   56
         Top             =   1350
         Width           =   345
      End
      Begin VB.Label lbl_mc_FaceOver 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2190
         TabIndex        =   55
         Top             =   1350
         Width           =   345
      End
      Begin VB.Label lbl_mc_FaceOff 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1350
         TabIndex        =   54
         Top             =   1350
         Width           =   345
      End
      Begin VB.Label lbl_mc_ShadowDown 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   3120
         TabIndex        =   53
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label lbl_mc_ShadowOver 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2190
         TabIndex        =   52
         Top             =   1035
         Width           =   345
      End
      Begin VB.Label lbl_mc_shadowOff 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1350
         TabIndex        =   51
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label lbl_mc_HilightDown 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   3120
         TabIndex        =   50
         Top             =   690
         Width           =   345
      End
      Begin VB.Label lbl_mc_hilightOver 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2190
         TabIndex        =   49
         Top             =   690
         Width           =   345
      End
      Begin VB.Label lbl_mc_hilightOff 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1350
         TabIndex        =   48
         Top             =   690
         Width           =   345
      End
      Begin VB.Label lbl_mc_TextDown 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   3120
         TabIndex        =   47
         Top             =   60
         Width           =   345
      End
      Begin VB.Label lbl_mc_TextOver 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2190
         TabIndex        =   46
         Top             =   60
         Width           =   345
      End
      Begin VB.Label lbl_mc_textoff 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   1350
         TabIndex        =   45
         Top             =   60
         Width           =   345
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "  Behaviour  "
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
         Left            =   240
         TabIndex        =   43
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000010&
         X1              =   45
         X2              =   4440
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000014&
         X1              =   45
         X2              =   4440
         Y1              =   3105
         Y2              =   3105
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seperator:"
         Height          =   195
         Left            =   405
         TabIndex        =   42
         Top             =   1755
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow"
         Height          =   195
         Left            =   3105
         TabIndex        =   41
         Top             =   1785
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hilight"
         Height          =   195
         Left            =   2055
         TabIndex        =   40
         Top             =   1785
         Width           =   435
      End
      Begin VB.Label lblFaceColors 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Face:"
         Height          =   195
         Left            =   780
         TabIndex        =   39
         Top             =   1335
         Width           =   405
      End
      Begin VB.Label lblFaceDown 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down"
         Height          =   195
         Left            =   3540
         TabIndex        =   38
         Top             =   1365
         Width           =   405
      End
      Begin VB.Label lblFaceOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over"
         Height          =   195
         Left            =   2610
         TabIndex        =   37
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label lblFaceOff 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Off"
         Height          =   195
         Left            =   1755
         TabIndex        =   36
         Top             =   1365
         Width           =   240
      End
      Begin VB.Label lblShadowColors 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow:"
         Height          =   195
         Left            =   555
         TabIndex        =   35
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblShadowDown 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down"
         Height          =   195
         Left            =   3540
         TabIndex        =   34
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label lblShadowOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over"
         Height          =   195
         Left            =   2610
         TabIndex        =   33
         Top             =   1035
         Width           =   360
      End
      Begin VB.Label lblShadowOff 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Off"
         Height          =   195
         Left            =   1755
         TabIndex        =   32
         Top             =   1035
         Width           =   240
      End
      Begin VB.Label lblHilightColors 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hilight:"
         Height          =   195
         Left            =   690
         TabIndex        =   31
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lblHilightDown 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down"
         Height          =   195
         Left            =   3540
         TabIndex        =   30
         Top             =   705
         Width           =   405
      End
      Begin VB.Label lblHilightOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over"
         Height          =   195
         Left            =   2610
         TabIndex        =   29
         Top             =   705
         Width           =   360
      End
      Begin VB.Label lblHilightOff 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Off"
         Height          =   195
         Left            =   1755
         TabIndex        =   28
         Top             =   705
         Width           =   240
      End
      Begin VB.Label lblTextcolors 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         Height          =   195
         Left            =   795
         TabIndex        =   27
         Top             =   45
         Width           =   390
      End
      Begin VB.Label lblTextDown 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down"
         Height          =   195
         Left            =   3540
         TabIndex        =   26
         Top             =   75
         Width           =   405
      End
      Begin VB.Label lblTextOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over"
         Height          =   195
         Left            =   2610
         TabIndex        =   25
         Top             =   75
         Width           =   360
      End
      Begin VB.Label lblTextOff 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Off"
         Height          =   195
         Left            =   1755
         TabIndex        =   24
         Top             =   75
         Width           =   240
      End
   End
   Begin VB.PictureBox picUserInfo 
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
      Height          =   3435
      Left            =   1710
      ScaleHeight     =   3435
      ScaleWidth      =   4440
      TabIndex        =   4
      Top             =   930
      Visible         =   0   'False
      Width           =   4440
      Begin VB.ListBox lstNicks 
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   1455
         TabIndex        =   15
         Top             =   1740
         Width           =   1905
      End
      Begin VB.CommandButton cmdNicksAdd 
         Caption         =   "+"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2325
         Width           =   240
      End
      Begin VB.CommandButton cmdNicksRemove 
         Caption         =   "-"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2625
         Width           =   240
      End
      Begin VB.CommandButton cmdNicksClear 
         Caption         =   "c"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2925
         Width           =   240
      End
      Begin VB.TextBox txtNickToAdd 
         Height          =   285
         Left            =   1455
         MaxLength       =   25
         TabIndex        =   11
         Top             =   1380
         Width           =   2220
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
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1725
         Width           =   240
      End
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
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2025
         Width           =   240
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   1455
         TabIndex        =   8
         Top             =   810
         Width           =   2220
      End
      Begin VB.TextBox txtRealName 
         Height          =   315
         Left            =   1455
         TabIndex        =   6
         Top             =   435
         Width           =   2220
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "  User Information  "
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
         Left            =   195
         TabIndex        =   22
         Top             =   0
         Width           =   1635
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   4395
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   4395
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nicks:"
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   1410
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Address:"
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   855
         Width           =   1110
      End
      Begin VB.Label lblConnect01 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Real Name:"
         Height          =   195
         Left            =   555
         TabIndex        =   5
         Top             =   480
         Width           =   825
      End
   End
   Begin VB.PictureBox picServers 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1590
      ScaleHeight     =   450
      ScaleWidth      =   4725
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   4725
      Begin VB.CommandButton cmdDelServer 
         Caption         =   "&Delete"
         Height          =   330
         Left            =   3825
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   15
         Width           =   705
      End
      Begin VB.CommandButton cmdEditServer 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   3165
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   15
         Width           =   630
      End
      Begin VB.CommandButton cmdAddServer 
         Caption         =   "&Add"
         Height          =   330
         Left            =   2505
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   15
         Width           =   630
      End
      Begin VB.ComboBox cmbServers 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   15
         Width           =   2340
      End
   End
   Begin VB.PictureBox picScripting 
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
      Height          =   3975
      Left            =   1665
      ScaleHeight     =   3975
      ScaleWidth      =   4500
      TabIndex        =   64
      Top             =   435
      Visible         =   0   'False
      Width           =   4500
      Begin VB.PictureBox picEE1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2520
         Picture         =   "frmOptions.frx":038A
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   72
         Top             =   3210
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CommandButton cmdScriptDown 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3165
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3150
         Width           =   330
      End
      Begin VB.CommandButton cmdScriptUp 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3150
         Width           =   330
      End
      Begin VB.CommandButton cmdRemoveScript 
         Caption         =   "Remove"
         Height          =   345
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3150
         Width           =   855
      End
      Begin VB.CommandButton cmdAddScript 
         Caption         =   "Add"
         Height          =   345
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3150
         Width           =   660
      End
      Begin VB.ListBox lstScripts 
         Height          =   2010
         IntegralHeight  =   0   'False
         Left            =   945
         TabIndex        =   67
         Top             =   1095
         Width           =   2535
      End
      Begin VB.CommandButton cmdScriptDefFolder 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox txtDefScriptFolder 
         Height          =   315
         Left            =   945
         TabIndex        =   65
         Text            =   "C:\apps\sIRC\"
         Top             =   330
         Width           =   2565
      End
      Begin VB.Label lblEE1 
         BackStyle       =   0  'Transparent
         Height          =   345
         Left            =   2355
         TabIndex        =   75
         Top             =   3255
         Width           =   345
      End
      Begin VB.Label lblSc 
         BackStyle       =   0  'Transparent
         Caption         =   "Loaded Scripts&:"
         Height          =   315
         Left            =   855
         TabIndex        =   74
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblSF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Script Folder&:"
         Height          =   195
         Left            =   855
         TabIndex        =   73
         Top             =   75
         Width           =   1530
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "  Servers  "
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
      Left            =   1935
      TabIndex        =   76
      Top             =   180
      Width           =   840
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000014&
      X1              =   116
      X2              =   409
      Y1              =   19
      Y2              =   19
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000010&
      X1              =   116
      X2              =   409
      Y1              =   18
      Y2              =   18
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bShowInTaskbar As Boolean
Public serverID As Integer
Public strTitle As String

Sub DrawMenuFontSample()
    Dim strText As String
    strText = strMenuFontName & ", " & intMenuFontSize & "pt"
    With pic_MD_FontSample
        .FontName = strMenuFontName
        .FontSize = intMenuFontSize
        .Cls
        .CurrentX = (.ScaleWidth - .textWidth(strText)) \ 2
        .CurrentY = (.ScaleHeight - .textHeight(strText)) \ 2
        pic_MD_FontSample.Print strText
    End With
End Sub

Sub FillFields()

    Dim i As Integer
    
    '* Scripting
    txtDefScriptFolder.Text = strDefScriptFolder
    For i = LBound(strScripts) To UBound(strScripts)
        If strScripts(i) <> "" Then lstScripts.AddItem strScripts(i)
    Next i
    
    '* Connected
    txtRealName.Text = strFName
    txtEmail.Text = strEmail
    lstNicks.Clear
    For i = LBound(strNicks) To UBound(strNicks)
        If strNicks(i) <> "" Then lstNicks.AddItem strNicks(i)
    Next i
    
    FillFields_DisplayMenu
    
End Sub
Sub FillFields_DisplayMenu()
    lbl_mc_textoff.BackColor = mc_TextOff
    lbl_mc_TextOver.BackColor = mc_TextOver
    lbl_mc_TextDown.BackColor = mc_TextDown
    lbl_mc_TextDisabled.BackColor = mc_TextDisabled
    lbl_mc_shadowOff.BackColor = mc_BShadowOff
    lbl_mc_ShadowOver.BackColor = mc_BShadowOver
    lbl_mc_ShadowDown.BackColor = mc_BShadowDown
    lbl_mc_hilightOff.BackColor = mc_BHilightOff
    lbl_mc_hilightOver.BackColor = mc_BHilightOver
    lbl_mc_HilightDown.BackColor = mc_BHilightDown
    lbl_mc_FaceOff.BackColor = mc_HilightOff
    lbl_mc_FaceOver.BackColor = mc_HilightOver
    lbl_mc_FaceDown.BackColor = mc_HilightDown
    lbl_mc_SepHilight.BackColor = mc_hilight
    lbl_mc_SepShadow.BackColor = mc_shadow
    
    chkMenuSink.value = IIf(mv_bSink, 1, 0)
    
    pic_MD_FontSample.FontName = strMenuFontName
    pic_MD_FontSample.FontSize = intMenuFontSize
    DrawMenuFontSample
End Sub

Sub HideAll()
    picScripting.Visible = False
    picServers.Visible = False
    picUserInfo.Visible = False
    picDisplayMenu.Visible = False
End Sub


Sub LoadOptions()

    'this procedure WAS used to load buddies from the ini to treeview.
    'taken from Chad Cox's AIM example, thanks Chad. (ass)
    
    Dim strBuffer As String * 600, lngSize As Long, arrBuddies() As String, lngDo As Long
    Dim nod() As Node, intGroup As Integer, strOptions As String
    
    
    strOptions = "g Connect" & Chr(1) & _
                 "b Options" & Chr(1) & _
                 "b Identd" & Chr(1) & _
                 "b Firewall" & Chr(1) & _
                 "g Scripting" & Chr(1) & _
                 "g Display" & Chr(1) & _
                 "b Menus" & Chr(1) & _
                 "b Taskbar" & Chr(1) & _
                 "b Windows" & Chr(1)
                 'g = main item
                 'b = sub item
                 
    With tvOptions
        arrBuddies$ = Split(strOptions, Chr(1))
        .Nodes.Clear
        For lngDo& = LBound(arrBuddies$) To UBound(arrBuddies$)
            ReDim Preserve nod(1 To .Nodes.Count + 1)
            If arrBuddies$(lngDo&) <> "" Then
                If Left$(arrBuddies$(lngDo&), 1) = "g" Then
                    Set nod(.Nodes.Count) = .Nodes.Add(, , , Right$(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2))
                    intGroup% = .Nodes.Count
                Else
                    If .Nodes.Count > 0 Then
                        Set nod(.Nodes.Count) = .Nodes.Add(nod(intGroup%), tvwChild, , Right$(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2))
                        nod(.Nodes.Count).EnsureVisible
                    End If
                End If
            End If
        Next
    End With


    FillFields
End Sub

Sub SaveOptions()
    Me.MousePointer = 11
    Call SaveOptions_Scripting
    Call SaveOptions_Connect
    Call SaveOptions_DisplayMenu
    
    ReLoadScripts
    Me.MousePointer = 1
End Sub

Sub SaveOptions_Connect()
    strFName = txtRealName.Text
    WriteINI "user", "name", strFName
    strEmail = txtEmail.Text
    WriteINI "user", "email", strEmail
    
    Dim i As Integer, strAllNicks As String
    ReDim strNicks(lstNicks.ListCount - 1) As String
    For i = 0 To lstNicks.ListCount - 1
        strNicks(i) = lstNicks.List(i)
        strAllNicks = strAllNicks & strNicks(i) & ","
    Next i
    WriteINI "user", "nicks", strAllNicks
End Sub


Sub SaveOptions_DisplayMenu()
    mc_TextOff = lbl_mc_textoff.BackColor
    WriteINI "display", "m_textoff", CStr(mc_TextOff)
    mc_TextOver = lbl_mc_TextOver.BackColor
    WriteINI "display", "m_textover", CStr(mc_TextOver)
    mc_TextDown = lbl_mc_TextDown.BackColor
    WriteINI "display", "m_textdown", CStr(mc_TextDown)
    mc_TextDisabled = lbl_mc_TextDisabled.BackColor
    WriteINI "display", "m_textdisabled", CStr(mc_TextDisabled)
    
    mc_BHilightOff = lbl_mc_hilightOff.BackColor
    WriteINI "display", "m_bhilightoff", CStr(mc_BHilightOff)
    mc_BHilightOver = lbl_mc_hilightOver.BackColor
    WriteINI "display", "m_bhilightover", CStr(mc_BHilightOver)
    mc_BHilightDown = lbl_mc_HilightDown.BackColor
    WriteINI "display", "m_bhilightdown", CStr(mc_BHilightDown)
    
    mc_BShadowOff = lbl_mc_shadowOff.BackColor
    WriteINI "display", "m_bshadowoff", CStr(mc_BShadowOff)
    mc_BShadowOver = lbl_mc_ShadowOver.BackColor
    WriteINI "display", "m_bshadowover", CStr(mc_BShadowOver)
    mc_BShadowDown = lbl_mc_ShadowDown.BackColor
    WriteINI "display", "m_bshadowdown", CStr(mc_BShadowDown)
    
    mc_HilightOff = lbl_mc_FaceOff.BackColor
    WriteINI "display", "m_hilightoff", CStr(mc_HilightOff)
    mc_HilightOver = lbl_mc_FaceOver.BackColor
    WriteINI "display", "m_hilightover", CStr(mc_HilightOver)
    mc_HilightDown = lbl_mc_FaceDown.BackColor
    WriteINI "display", "m_hilightdown", CStr(mc_HilightDown)
    
    mv_bSink = IIf(chkMenuSink.value = 1, True, False)
    WriteINI "display", "m_bsink", CStr(mv_bSink)
    
    WriteINI "display", "m_fontname", strMenuFontName
    intMenuFontHeight = GetCharHeight(intMenuFontSize)
    WriteINI "display", "m_fontsize", CStr(intMenuFontSize)
End Sub

Sub SaveOptions_Scripting()
    strDefScriptFolder = txtDefScriptFolder.Text
    WriteINI "scripting", "def_sf", strDefScriptFolder
    
    Dim i As Long, strScripts As String
    For i = 0 To lstScripts.ListCount - 1
        strScripts = strScripts & lstScripts.List(i) & ","
    Next i
    WriteINI "scripting", "scripts", strScripts
End Sub


Sub SetTitle(strTitle As String)
    lblTitle.Caption = "  " & strTitle & "  "
End Sub

Sub ShowOptions(strWhich As String)
    '* Hide All
    HideAll
    
    '* Show right options
    Select Case strWhich
        Case "Scripting"
            SetTitle "Scripting"
            picScripting.Visible = True
        Case "Connect"
            SetTitle "Servers"
            picUserInfo.Visible = True
            picServers.Visible = True
        Case "Menus"
            SetTitle "Colors"
            picDisplayMenu.Visible = True
        Case Else
            SetTitle "Options..."
    End Select
End Sub



Private Sub cmdAddScript_Click()
    On Error GoTo errhandler
    
    cmDialog.DefaultExt = "*.sex"
    cmDialog.DialogTitle = "Add a script"
    cmDialog.Filter = "sIRC Scripts (*.sex)|*.sex|"
    cmDialog.ShowOpen
    
    If FileExists(cmDialog.FileName) Then
        If Left$(cmDialog.FileName, Len(PATH)) = PATH Then
            lstScripts.AddItem Right$(cmDialog.FileName, Len(cmDialog.FileName) - Len(PATH))
        Else
            lstScripts.AddItem cmDialog.FileName
        End If
    End If
    
    Exit Sub
errhandler:
End Sub

Private Sub cmdAddServer_Click()
    Dim strNewServer As String
    strNewServer = InputBox("Enter the name (not address) of the new server:", "Add server")
    
    If IsNull(strNewServer) Then Exit Sub
    
    Dim i As Integer
    For i = 0 To cmbServers.ListCount
        If LCase(cmbServers.List(i)) = LCase(strNewServer) Then
            MsgBox "Server already exists in list.", vbCritical
            Exit Sub
        End If
    Next i
    
    'cmbServers.ListIndex = cmbServers.AddItem(strNewServer)
End Sub

Private Sub cmdApply_Click()
    SaveOptions
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdDelServer_Click()
    If cmbServers.ListIndex < 0 Then
        MsgBox "You must first select a server before editing it.", vbCritical
        Exit Sub
    End If
    
    Dim intResult As Integer
    intResult = MsgBox("Are you sure you would like to delete the server """ & cmbServers.List(cmbServers.ListIndex) & """?", vbCritical)
    If intResult = vbYes Then
        cmbServers.RemoveItem cmbServers.ListIndex
    End If
End Sub

Private Sub cmdEditServer_Click()
    If cmbServers.ListIndex < 0 Then
        MsgBox "You must first select a server before editing it.", vbCritical
        Exit Sub
    End If

    Dim strNewServer As String
    strNewServer = InputBox("Enter the NEW name (not address) of the server:", "Add server", cmbServers.List(cmbServers.ListIndex))
    
    If IsNull(strNewServer) Then Exit Sub
        
    cmbServers.List(cmbServers.ListIndex) = strNewServer
End Sub

Private Sub cmdMenuChangeFont_Click()
    On Error GoTo errHandle
    cmDialog.Flags = cdlCFScreenFonts
    cmDialog.FontName = strMenuFontName
    cmDialog.FontSize = intMenuFontSize
    cmDialog.ShowFont
    strMenuFontName = cmDialog.FontName
    intMenuFontSize = cmDialog.FontSize
    DrawMenuFontSample
    Exit Sub
errHandle:
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
  
End Sub

Private Sub cmdNicksUP_Click()
    Dim strTmp As String
    strTmp = lstNicks.List(lstNicks.ListIndex)
    lstNicks.List(lstNicks.ListIndex) = lstNicks.List(lstNicks.ListIndex - 1)
    lstNicks.List(lstNicks.ListIndex - 1) = strTmp
    lstNicks.ListIndex = lstNicks.ListIndex - 1
End Sub

Private Sub cmdOk_Click()
    SaveOptions
    Unload Me
End Sub

Private Sub cmdRemoveScript_Click()
    If lstScripts.ListIndex = -1 Then
    Else
        lstScripts.RemoveItem lstScripts.ListIndex
    End If
End Sub


Private Sub cmdScriptDown_Click()
    Dim intIndex As Integer, tmpStr As String
    intIndex = lstScripts.ListIndex
    
    If intIndex = lstScripts.ListCount - 1 Or intIndex = -1 Then Exit Sub
    
    With lstScripts
        tmpStr = .List(intIndex + 1)
        .List(intIndex + 1) = .List(intIndex)
        .List(intIndex) = tmpStr
        .ListIndex = .ListIndex + 1
    End With
End Sub


Private Sub cmdScriptUp_Click()
    Dim intIndex As Integer, tmpStr As String
    intIndex = lstScripts.ListIndex
    
    If intIndex < 1 Then Exit Sub
    
    With lstScripts
        tmpStr = .List(intIndex - 1)
        .List(intIndex - 1) = .List(intIndex)
        .List(intIndex) = tmpStr
        .ListIndex = .ListIndex - 1
    End With
    
End Sub



Private Sub Form_Load()
    
    LoadOptions
    ButtonizeForm Me
    
    picUserInfo.Visible = True
    picServers.Visible = True
        
End Sub




Private Sub lbl_mc_FaceDown_Click()
    lbl_mc_FaceDown.BackColor = ColorPicker.GetColor(lbl_mc_FaceDown.BackColor)
End Sub

Private Sub lbl_mc_FaceOff_Click()
    lbl_mc_FaceOff.BackColor = ColorPicker.GetColor(lbl_mc_FaceOff.BackColor)
End Sub

Private Sub lbl_mc_FaceOver_Click()
    lbl_mc_FaceOver.BackColor = ColorPicker.GetColor(lbl_mc_FaceOver.BackColor)
End Sub


Private Sub lbl_mc_HilightDown_Click()
    lbl_mc_HilightDown.BackColor = ColorPicker.GetColor(lbl_mc_HilightDown.BackColor)
End Sub

Private Sub lbl_mc_hilightOff_Click()
    lbl_mc_hilightOff.BackColor = ColorPicker.GetColor(lbl_mc_hilightOff.BackColor)
End Sub

Private Sub lbl_mc_hilightOver_Click()
    lbl_mc_hilightOver.BackColor = ColorPicker.GetColor(lbl_mc_hilightOver.BackColor)
End Sub


Private Sub lbl_mc_SepHilight_Click()
    lbl_mc_SepHilight.BackColor = ColorPicker.GetColor(lbl_mc_SepHilight.BackColor)
End Sub

Private Sub lbl_mc_SepShadow_Click()
    lbl_mc_SepShadow.BackColor = ColorPicker.GetColor(lbl_mc_SepShadow.BackColor)
End Sub


Private Sub lbl_mc_ShadowDown_Click()
    lbl_mc_ShadowDown.BackColor = ColorPicker.GetColor(lbl_mc_ShadowDown.BackColor)
End Sub

Private Sub lbl_mc_shadowOff_Click()
    lbl_mc_shadowOff.BackColor = ColorPicker.GetColor(lbl_mc_shadowOff.BackColor)
End Sub

Private Sub lbl_mc_ShadowOver_Click()
    lbl_mc_ShadowOver.BackColor = ColorPicker.GetColor(lbl_mc_ShadowOver.BackColor)
End Sub



Private Sub lbl_mc_TextDown_Click()
    lbl_mc_TextDown.BackColor = ColorPicker.GetColor(lbl_mc_TextDown.BackColor)
End Sub

Private Sub lbl_mc_textoff_Click()
    lbl_mc_textoff.BackColor = ColorPicker.GetColor(lbl_mc_textoff.BackColor)
End Sub

Private Sub lbl_mc_TextOver_Click()
    lbl_mc_TextOver.BackColor = ColorPicker.GetColor(lbl_mc_TextOver.BackColor)
End Sub


Private Sub lblEE1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picEE1.Visible = True
End Sub





Private Sub lblEE1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picEE1.Visible = False
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

Private Sub lstNicks_DblClick()
    Call cmdNicksRemove_Click
End Sub


Private Sub tvOptions_Click()
    ShowOptions tvOptions.selectedItem.Text
End Sub




Private Sub tvOptions_NodeClick(ByVal Node As MSComctlLib.Node)
    ShowOptions tvOptions.selectedItem.Text
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
    End If
End Sub





