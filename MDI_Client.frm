VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.MDIForm CLIENT 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "sIRC Alpha v0.#.#"
   ClientHeight    =   4965
   ClientLeft      =   3645
   ClientTop       =   2565
   ClientWidth     =   8310
   Icon            =   "MDI_Client.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Picture         =   "MDI_Client.frx":038A
   ScrollBars      =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmrTask2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3900
      Top             =   1365
   End
   Begin VB.PictureBox picTask2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   5
      Top             =   4245
      Visible         =   0   'False
      Width           =   8310
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   6690
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   6
         Top             =   135
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ilTreeView 
      Left            =   3480
      Top             =   1965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":0714
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":0E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":11E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":157C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":1916
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picServerList 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3900
      ScaleWidth      =   1830
      TabIndex        =   3
      Top             =   345
      Visible         =   0   'False
      Width           =   1830
      Begin MSComctlLib.TreeView tvServers 
         Height          =   4620
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   8149
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   423
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "ilTreeView"
         Appearance      =   0
         MousePointer    =   1
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
      Begin VB.Shape shpSL 
         BorderColor     =   &H8000000F&
         BorderWidth     =   2
         Height          =   4710
         Left            =   0
         Top             =   0
         Width           =   1755
      End
   End
   Begin VB.Timer tmrMenu 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1920
      Top             =   1980
   End
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   2910
      Top             =   1965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":1EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":224A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":27E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":2B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":3118
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":34B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":3A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":3FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":4580
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":4E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":522C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":57C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":5B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Client.frx":5EFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMenu 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      FillColor       =   &H80000010&
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   1
      Top             =   0
      Width           =   8310
   End
   Begin VB.Timer tmrTool 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   705
      Top             =   1560
   End
   Begin VB.Timer tmrTask 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3435
      Top             =   1335
   End
   Begin VB.PictureBox picTask 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   554
      TabIndex        =   0
      Top             =   4605
      Width           =   8310
      Begin VB.PictureBox picTaskIcon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   6690
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   2
         Top             =   135
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSWinsockLib.Winsock IDENT 
      Left            =   2175
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   113
   End
   Begin VB.Menu mnu_Connect 
      Caption         =   "&Connect"
      Visible         =   0   'False
      Begin VB.Menu mnu_Connect_NewServer 
         Caption         =   "&New Server"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_Connect_lb01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Connect_Connect 
         Caption         =   "&Connect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_Connect_Disconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnu_Tools_Options 
         Caption         =   "&Options..."
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnu_Tools_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tools_Scripts 
         Caption         =   "&Scripts..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Tools_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Tools_ChangeProfile 
         Caption         =   "&Change Profile"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnu_Edit_Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_Edit_LB0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_Cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_Edit_copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_Edit_Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_Edit_Delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu_Edit_LB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Format 
      Caption         =   "F&ormat"
      Visible         =   0   'False
      Begin VB.Menu mnu_Format_Cancel 
         Caption         =   "&Cancel"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_Format_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Format_Bold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu_Format_Color 
         Caption         =   "&Color"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnu_Format_Reverse 
         Caption         =   "&Reverse"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_Format_Underline 
         Caption         =   "&Underline"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnu_Commands 
      Caption         =   "&Commands"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Window_Close 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnu_Window_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Window_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_Window_TileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnu_Window_TileV 
         Caption         =   "Tile &Veritcally"
      End
      Begin VB.Menu mnu_Window_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Window_Auto 
         Caption         =   "&Auto"
         Begin VB.Menu mnu_Window_AutoMax 
            Caption         =   "Maximize"
         End
         Begin VB.Menu mnu_Window_AutoTileH 
            Caption         =   "Tile Horizontally"
         End
         Begin VB.Menu mnu_Window_AutoTileV 
            Caption         =   "Tile Veritically"
         End
      End
      Begin VB.Menu mnu_Window_LB03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Window_Remember 
         Caption         =   "&Remember"
         Begin VB.Menu mnu_Window_Remember_CurrentWindow 
            Caption         =   "&Current Window"
         End
         Begin VB.Menu mnu_Window_Remember_Client 
            Caption         =   "Client (Main Window)"
         End
         Begin VB.Menu mnu_Window_Remember_AllWindows 
            Caption         =   "&All Windows"
         End
      End
      Begin VB.Menu mnu_Window_Forget 
         Caption         =   "&Forget"
         Begin VB.Menu mnu_Window_Forget_CurrentWindow 
            Caption         =   "&Current Window"
         End
         Begin VB.Menu mnu_Window_Forget_Client 
            Caption         =   "Client (Main Window)"
         End
         Begin VB.Menu mnu_Window_Forget_AllWindows 
            Caption         =   "&All Windows"
         End
      End
      Begin VB.Menu mnu_Window_Reset 
         Caption         =   "&Reset"
         Begin VB.Menu mnu_Window_Reset_CurrentWindow 
            Caption         =   "&Current Window"
         End
         Begin VB.Menu mnu_Window_Reset_Client 
            Caption         =   "Client (Main Window)"
         End
         Begin VB.Menu mnu_Window_Reset_AllWindows 
            Caption         =   "&All Windows"
         End
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "CLIENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intPrevY As Single
Private startX As Single, startY As Single
Private ShiftKey As Integer

Public oldProcAddr As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Const MF_STRING = &H0&

Public curSysMenuHWND As Long
Private WithEvents mParent  As Form
Attribute mParent.VB_VarHelpID = -1
Public bShowInTaskbar As Boolean

'* MENU!
Private mrcnt As Integer
Public bMenuShown As Boolean
Private menuActive      As clsXPMenu
Private mnuActiveWhich  As Integer
Public mnuOverWhich     As Integer
Private mnuDownWhich    As Integer
Private bMenuDrew       As Boolean
Private MenuArray()     As String
Private MenuYPos()      As Integer
Private Const XM_Buffer As Integer = 5
Private Const YM_Buffer As Integer = 2

Private tbActiveWhich As Integer
Private tbOverWhich   As Integer
Private tbDownWhich   As Integer
Private btbDrew       As Boolean

Private ActiveWhich As Integer
Private OverWhich   As Integer
Private DownWhich   As Integer
Private bDown       As Boolean
Private bLastDown   As Boolean
Private bDrew       As Boolean

Private DActiveWhich As Integer
Private DOverWhich As Integer
Private DDownWhich As Integer
Private bDDrew As Boolean

'* These will be changed to global variables..yaddy yadda
Const clr_LeftMargin As Long = &HD1D8D8

Const HilightColorOff = &H8000000F
Const ShadowColorOff = &H8000000F
Const buttonColorOff = &H8000000F

Const HilightColorOver = &H8000000D
Const ShadowColorOver = &H8000000D
Const ButtonColorOver = 13811126

Const HilightColorDown = &H8000000D
Const ShadowColorDown = &H8000000D
Const ButtonColorDown = 11899525

Const HilightColorActive = &H8000000D
Const ShadowColorActive = &H8000000D
Const ButtonColorActive = 14210516

Const YBuffer = 0
Const XBuffer = 1
Const YStart = 1
Const XStart = 12
Const TBButtons = 4

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020
Private Const WM_NCLBUTTONDOWN  As Long = &HA1&
Private Const HTCLIENT          As Long = 1
Private Const HTCAPTION         As Long = &H2&


Sub DisplayMenu(strName As String)
    Dim lpRect As Rect, YPos As Long
    GetWindowRect picMenu.hwnd, lpRect
    YPos = lpRect.Bottom * 15 - 15

    Select Case strName
        Case "Connect"
            Set menuActive = XPM_Connect
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(0) * 15), YPos, MenuYPos(1) - MenuYPos(0) - XBuffer * 2 - 1
        Case "Tools"
            Set menuActive = XPM_Tools
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(1) * 15), YPos, MenuYPos(2) - MenuYPos(1) - XBuffer * 2 - 1
        Case "Edit"
            Set menuActive = XPM_Edit
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(2) * 15), YPos, MenuYPos(3) - MenuYPos(2) - XBuffer * 2 - 1
        Case "View"
            Set menuActive = XPM_View
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(3) * 15), YPos, MenuYPos(4) - MenuYPos(3) - XBuffer * 2 - 1
        Case "Format"
            Set menuActive = XPM_Format
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(4) * 15), YPos, MenuYPos(5) - MenuYPos(4) - XBuffer * 2 - 1
        Case "Commands"
            Set menuActive = XPM_Commands
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(5) * 15), YPos, MenuYPos(6) - MenuYPos(5) - XBuffer * 2 - 1
        Case "Window"
            Set menuActive = XPM_Window
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(6) * 15), YPos, MenuYPos(7) - MenuYPos(6) - XBuffer * 2 - 1
        Case "Help"
            Set menuActive = XPM_Help
            menuActive.ShowMenu CLIENT.Left + (MenuYPos(7) * 15), YPos, 29
    End Select
    
End Sub


Public Sub DrawMenuOriginal()
    Dim offSet As Integer
    offSet = 2
    
    picMenuBuffer.Cls
    Dim ly As Integer, lx As Integer
    lx = 3
    For ly = 5 To 17 Step 2
        picMenuBuffer.Line (4, ly)-(4 + lx, ly), COLOR_DGRAY
    Next ly
    
    Dim intMenus As Integer, CenX As Integer, j As Integer, realWidth As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, intBegin As Integer
    Dim intEnd As Integer, strDrawText As String, temp As Integer
    Dim intStartY As Integer, textWidth As Integer, textHeight As Integer
    Dim drawbevel As Boolean, curX As Integer, intHeight As Integer
    Dim newmsgs As String
    intStartY = 1
    'bReDraw = True
    
    picMenuBuffer.Width = picMenu.Width
    
    intMenus = UBound(MenuArray) + 1
    If intMenus <= 0 Then GoTo finishit
    
    i = 0
    temp = 0
        
    textHeight = picMenuBuffer.textWidth("gW")
    intHeight = textHeight + (YM_Buffer * 2)
    
    curX = XStart
    
    On Error Resume Next
    For j = 1 To intMenus
        textWidth = picMenuBuffer.textWidth(MenuArray(j - 1))
        intWidth = textWidth + (XM_Buffer * 2)
    
        MenuYPos(j - 1) = curX + 4
        If mnuOverWhich = j Then
            'If menuSelected = j Then
                ' draw menu to popup
                'picMenuBuffer.Line (curX, YStart)-(curX + intWidth, YStart + intHeight), clr_LeftMargin, BF
            'Else
                picMenuBuffer.Line (curX, YStart)-(curX + intWidth, YStart + intHeight), ButtonColorOver, BF
            'End If
            
            picMenuBuffer.Line (curX, YStart)-(curX + intWidth, YStart), HilightColorOver
            picMenuBuffer.Line (curX, YStart)-(curX, YStart + intHeight), HilightColorOver
        
            picMenuBuffer.Line (curX + intWidth, YStart)-(curX + intWidth, YStart + intHeight), ShadowColorOver
            picMenuBuffer.Line (curX, YStart + intHeight)-(curX + intWidth + 1, YStart + intHeight), ShadowColorOver
        Else
            picMenuBuffer.Line (curX, YStart)-(curX + intWidth, YStart + intHeight), buttonColorOff, BF
            
            picMenuBuffer.Line (curX, YStart)-(curX + intWidth, YStart), HilightColorOff
            picMenuBuffer.Line (curX, YStart)-(curX, YStart + intHeight), HilightColorOff
        
            picMenuBuffer.Line (curX + intWidth, YStart)-(curX + intWidth, YStart + intHeight), ShadowColorOff
            picMenuBuffer.Line (curX, YStart + intHeight)-(curX + intWidth, YStart + intHeight), ShadowColorOff
        End If
        
        TextOut picMenuBuffer.hdc, curX + XM_Buffer, (picMenu.ScaleHeight - textHeight) \ 2 + 1, MenuArray(j - 1), Len(MenuArray(j - 1))
        
        curX = curX + intWidth + 1
        
        i = i + 1
    Next
    
finishit:
    picMenu.Picture = picMenuBuffer.Image
    bMenuDrew = True


End Sub

Public Sub DrawMenu()


    Dim intMenus As Integer, j As Integer, intWidth As Integer, i As Integer
    Dim textWidth As Long, textHeight As Long, curX As Long, intHeight As Long
    Dim tBrush As Long, BMP As BitmapStruc, hFont As Long, theSize As Size, tBrush2 As Long
    Dim oldObj As Long, oldObj2 As Long, oldFont As Long, tPen As Long, hFontSym As Long
    Dim offSet As Integer, ly As Integer, lx As Integer, bDebugMode As Boolean, intDown As Integer
    offSet = 2
    
    '* DISABLE IF NOT DRAWING
    bDebugMode = True
    
    Static bDrawing As Boolean
    
    If bDrawing Then Exit Sub
    bDrawing = True
    
    '* Create the fonts to be used
    hFont = CreateFont(intMenuFontHeight, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, strMenuFontName)
    If hFont = 0 Then Exit Sub
    hFontSym = CreateFont(14, 13, 0, 0, FW_LIGHT, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, DEFAULT_PITCH, "Marlett")
    If hFontSym = 0 Then DeleteObject hFont: Exit Sub
    
    '* Set the Area
    BMP.Area.Left = 0
    BMP.Area.Top = 0
    BMP.Area.Right = picMenu.ScaleWidth
    BMP.Area.Bottom = picMenu.ScaleHeight
    
    '* Create bitmap
    BMP.hDcMemory = CreateCompatibleDC(picMenu.hdc)
    BMP.hDcBitmap = CreateCompatibleBitmap(picMenu.hdc, picMenu.ScaleWidth, picMenu.ScaleHeight)
    BMP.hDcPointer = SelectObject(BMP.hDcMemory, BMP.hDcBitmap)
            
    If BMP.hDcMemory = 0 Or BMP.hDcBitmap = 0 Then
        DeleteObject BMP.hDcBitmap
        DeleteDC BMP.hDcMemory
        DeleteObject hFont
        Exit Sub
    End If
    
    '* Save
    SaveDC BMP.hDcMemory

    '* Copy the background of picMenu into the DC
    tBrush = CreateSolidBrush(GetSysColor(COLOR_3DFACE))
    oldObj = SelectObject(BMP.hDcMemory, tBrush)
    tPen = CreatePen(0, 0, GetSysColor(COLOR_3DFACE))
    oldObj2 = SelectObject(BMP.hDcMemory, tPen)
    Rectangle BMP.hDcMemory, 0, 0, picMenu.ScaleWidth + 1, picMenu.ScaleHeight + 1
    SelectObject BMP.hDcMemory, oldObj
    SelectObject BMP.hDcMemory, oldObj2
    DeleteObject tBrush
    DeleteObject tPen
    
    '* Draw the uh..thing on the left
    tBrush = CreateSolidBrush(clrLines)
    oldObj = SelectObject(BMP.hDcMemory, tBrush)
    tPen = CreatePen(0, 1, clrLines)
    oldObj2 = SelectObject(BMP.hDcMemory, tPen)
    lx = 3
    For ly = 5 To 17 Step 2
        Rectangle BMP.hDcMemory, 4, ly, 4 + lx, ly + 1
    Next ly
    SelectObject BMP.hDcMemory, oldObj
    SelectObject BMP.hDcMemory, oldObj2
    DeleteObject tBrush
    DeleteObject tPen
    
    '* Set The Font
    oldFont = SelectObject(BMP.hDcMemory, hFont)
        
    '* background of text transparent
    SetBkMode BMP.hDcMemory, 0
    
    intMenus = UBound(MenuArray) + 1
    If intMenus <= 0 Then GoTo finishit
    
    textHeight = picMenu.textWidth("gW")
    intHeight = textHeight + (YM_Buffer * 2)
    curX = XStart
    i = 0
    
    On Error Resume Next
    For j = 1 To intMenus
        Call GetTextExtentPoint32(BMP.hDcMemory, MenuArray(j - 1), Len(MenuArray(j - 1)), theSize)
        textWidth = theSize.cx
        intWidth = textWidth + (XM_Buffer * 2)
        
        SetTextColor BMP.hDcMemory, vbBlack
        intDown = 0
        MenuYPos(j - 1) = curX + 4
        If bDebugMode = True Then
            If mnuOverWhich = j Then
                If bMenuShown Then
                    SetTextColor BMP.hDcMemory, mc_TextDown
                    DrawButton BMP.hDcMemory, mc_HilightDown, mc_BHilightDown, mc_BShadowDown, curX, YStart + 1, intWidth - 1, YStart + intHeight - 2
                    If mv_bSink Then intDown = 1
                Else
                    SetTextColor BMP.hDcMemory, mc_TextDown
                    DrawButton BMP.hDcMemory, mc_HilightDown, mc_BShadowDown, mc_BHilightDown, curX, YStart + 1, intWidth - 1, YStart + intHeight - 3
                End If
            Else
                SetTextColor BMP.hDcMemory, mc_TextOff
                DrawButton BMP.hDcMemory, mc_HilightOff, mc_BShadowOff, mc_BHilightOff, curX, YStart + 1, intWidth - 1, YStart + intHeight - 3
            End If
        End If
        
        If bDebugMode = True Then
            TextOut BMP.hDcMemory, curX + XM_Buffer + intDown, intDown + (picMenu.ScaleHeight - textHeight) \ 2 + 1, MenuArray(j - 1), Len(MenuArray(j - 1))
        End If
        curX = curX + intWidth + 1
        i = i + 1
    Next
   
    If CLIENT.ActiveForm Is Nothing Then
    Else
        If CLIENT.ActiveForm.WindowState = 2 Then
            oldObj = SelectObject(BMP.hDcMemory, hFontSym)
            If mnuOverWhich = 9999 Then
                'SetTextColor BMP.hDcMemory, clrTextOver
                DrawButton BMP.hDcMemory, mc_HilightOff, mc_BHilightOver, mc_BShadowOver, BMP.Area.Right - 20, 1, 20, 20
            End If
            TextOut BMP.hDcMemory, BMP.Area.Right - 17, 4, "2", 1
            SelectObject BMP.hDcMemory, oldObj
        End If
    End If
        
finishit:
    'picMenu.Picture = picMenuBuffer.Image
    BitBlt picMenu.hdc, BMP.Area.Left, BMP.Area.Top, BMP.Area.Right, BMP.Area.Bottom, BMP.hDcMemory, 0, 0, SRCCOPY
    
    SelectObject BMP.hDcMemory, BMP.hDcPointer
    SelectObject BMP.hDcMemory, oldFont
    
    '* RestoreDC State
    RestoreDC BMP.hDcMemory, -1
    
    DeleteObject tBrush
    DeleteObject tPen
    DeleteObject oldObj
    DeleteObject oldObj2
    DeleteObject hFont
    DeleteObject hFontSym
    DeleteObject BMP.hDcBitmap
    DeleteDC BMP.hDcMemory
    
    bMenuDrew = True
    bDrawing = False

End Sub


Sub DrawTaskbarAllServers()
    
    '* DrawTaskbar - redone 12/9/01 with API
    '* redone again on 2/17/02 with sorting/filtering abilities
    
    Dim intSeps As Integer, j As Integer, realWidth As Long, iconBuffer As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, iconY  As Long, strText As String
    Dim intStartY As Integer, intDown As Integer, TextX As Long, TextY As Long
    Dim ly As Integer, lx As Integer, curX As Long, hFont2 As Long, TextWidth2 As Long
    Dim tBrush As Long, BMP As BitmapStruc, hFont As Long, theSize As Size  '* stuff for DC (buffer)
    
    On Error Resume Next
    
    If b_DualSwitch Then
        FillSwitchbar SWITCH_WINDOWS, "", CLIENT.ActiveForm.serverID, False
        DrawTaskbarServers
    Else
        FillSwitchbar SWITCH_WINDOWS
    End If
    
    '* Get number of "seperators" (not used anymore)
    intSeps = SwitchWindowCount
    
    If intSeps = 0 Then
        picTask.Cls
        Exit Sub
    End If
    
    '* Nothing to draw...go to end (so we can blit what we did draw)
    If bStretchButtons Then
        realWidth = (picTask.ScaleWidth) - XStart - 1
        minSize = (intSeps) * (ICON_SIZE + 40)
        If realWidth < minSize Then realWidth = minSize Else minSize = realWidth
    Else
        '* remove.. once you do options
        intButtonWidth = 125
        
        realWidth = (intSeps) * intButtonWidth - XStart
        minSize = realWidth
    End If
    If realWidth + XStart + 2 >= picTask.ScaleWidth Then realWidth = picTask.ScaleWidth - XStart - 2
    
    '* Create the fonts to be used
    hFont = CreateFont(13, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    If hFont = 0 Then Exit Sub
    hFont2 = CreateFont(13, 6, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    If hFont2 = 0 Then Exit Sub
    
    '* Set the Area
    BMP.Area.Left = 0
    BMP.Area.Top = 0
    BMP.Area.Right = realWidth + XStart - 2
    BMP.Area.Bottom = picTask.ScaleHeight
    
    '* Create bitmap
    BMP.hDcMemory = CreateCompatibleDC(picTask.hdc)
    BMP.hDcBitmap = CreateCompatibleBitmap(picTask.hdc, picTask.ScaleWidth, picTask.ScaleHeight)
    BMP.hDcPointer = SelectObject(BMP.hDcMemory, BMP.hDcBitmap)
            
    If BMP.hDcMemory = 0 Or BMP.hDcBitmap = 0 Then
        DeleteObject BMP.hDcBitmap
        DeleteDC BMP.hDcMemory
        DeleteObject hFont
        Exit Sub
    End If
    
    '* SAVE!
    SaveDC BMP.hDcMemory
    
    '* Copy the background of picMenu into the DC
    tBrush = CreateSolidBrush(clrBackground)
    oldObj = SelectObject(BMP.hDcMemory, tBrush)
    tPen = CreatePen(0, 0, clrBackground)
    oldObj2 = SelectObject(BMP.hDcMemory, tPen)
    Rectangle BMP.hDcMemory, 0, 0, picTask.ScaleWidth + 1, picTask.ScaleHeight + 1
    SelectObject BMP.hDcMemory, oldObj
    SelectObject BMP.hDcMemory, oldObj2
    DeleteObject tBrush
    DeleteObject tPen
   
    '* Set The Font
    Call SelectObject(BMP.hDcMemory, hFont)
    
    '* Draw the uh..thing on the left
    tBrush = CreateSolidBrush(clrLines)
    oldObj = SelectObject(BMP.hDcMemory, tBrush)
    tPen = CreatePen(0, 1, clrLines)
    oldObj2 = SelectObject(BMP.hDcMemory, tPen)
    lx = 3
    For ly = 6 To 18 Step 2
        Rectangle BMP.hDcMemory, 4, ly, 4 + lx, ly + 1
    Next ly
    SelectObject BMP.hDcMemory, oldObj
    SelectObject BMP.hDcMemory, oldObj2
    DeleteObject tBrush
    DeleteObject tPen
    
    '* background of text transparent
    SetBkMode BMP.hDcMemory, 0
    
    intStartY = 1
    bReDraw = True
    
    If intSeps <= 0 Then GoTo finishit
    
    '* Set some variables
    intWidth = Int((realWidth / (intSeps)) - 0.5)
    GetTextExtentPoint32 BMP.hDcMemory, "wYz", 3, theSize
    textHeight = theSize.cy
    TextY = (picTask.ScaleHeight - textHeight) \ 2
    iconY = (picTask.ScaleHeight - 16) \ 2
    iconBuffer = ICON_SIZE + (XBuffer * 2) + 2
    
    curX = XStart
    
    For i = 1 To SwitchWindowCount
        If ActiveWhich = i Then
            SetTextColor BMP.hDcMemory, tb_TextActive
            DrawButton BMP.hDcMemory, tb_ButtonActive, tb_HilightActive, tb_ShadowActive, curX, (YStart + 1), intWidth - 1, picTask.ScaleHeight - (YStart * 2) ' - 1
            
            picTaskIcon.BackColor = tb_ButtonActive
        ElseIf DownWhich = i Then
            SetTextColor BMP.hDcMemory, tb_TextDown
            DrawButton BMP.hDcMemory, tb_ButtonDown, tb_HilightDown, tb_ShadowDown, curX, (YStart + 1), (intWidth - 1), picTask.ScaleHeight - (YStart * 2) '- 1
            
            picTaskIcon.BackColor = tb_ButtonDown
        ElseIf OverWhich = i Then
            SetTextColor BMP.hDcMemory, tb_TextOver
            DrawButton BMP.hDcMemory, tb_ButtonOver, tb_HilightOver, tb_ShadowOver, curX, (YStart + 1), (intWidth - 1), picTask.ScaleHeight - (YStart * 2) ' - 1
            
            picTask.ToolTipText = " " & SwitchWindows(i).strText & " "
            picTaskIcon.BackColor = tb_ButtonOver
        Else
            SetTextColor BMP.hDcMemory, tb_TextOff
            DrawButton BMP.hDcMemory, tb_ButtonOff, tb_HilightOff, tb_ShadowOff, curX, (YStart + 1), (intWidth - 1), picTask.ScaleHeight - (YStart * 2) '- 1
            
            picTaskIcon.BackColor = tb_ButtonOff
        End If
        
        '* set dimensions
        With SwitchWindows(i).lRect
            .Left = curX
            .Top = YStart + 1
            .Right = intWidth + curX - 1
            .Bottom = picTask.ScaleHeight - (YStart * 2)
        End With
        
        If (DownWhich = i) Or (ActiveWhich = i) Then
            intDown = 1
        Else
            intDown = 0
        End If
        
        'BitBlt BMP.hDcMemory, curX + XBuffer * 2 + intDown + 1, iconY + intDown, ICON_SIZE, ICON_SIZE, picTaskIcon.hdc, 0, 0, SRCCOPY
        'DrawIcon BMP.hDcMemory, curX + XBuffer * 2 + intDown + 1, iconY + intDown, SwitchWindows(i).hIcon
        DrawIconEx BMP.hDcMemory, curX + XBuffer * 2 + intDown + 1, iconY + intDown, SwitchWindows(i).hIcon, ICON_SIZE, ICON_SIZE, 0, 0, DI_NORMAL
        
        strText = SwitchWindows(i).strText
        If SwitchWindows(i).bNewData Then
            strText = strText & " (" & SwitchWindows(i).NewLines & ")"
            SetTextColor BMP.hDcMemory, vbRed
        End If
        
        '* draw actual text
        strText = GetTaskText(strText, intWidth - ICON_SIZE - (XBuffer * 2) - 6)
        TextOut BMP.hDcMemory, curX + XBuffer * 2 + intDown + TextWidth2 + iconBuffer + 1, TextY + intDown, strText, Len(strText)
                
        '* increment variables
        curX = curX + intWidth
        
    Next i
    
finishit:
    BitBlt picTask.hdc, BMP.Area.Left, BMP.Area.Top, BMP.Area.Right, BMP.Area.Bottom, BMP.hDcMemory, 0, 0, SRCCOPY
    
    '* Restore!
    RestoreDC BMP.hDcMemory, -1
    
    DeleteObject tBrush
    DeleteObject tPen
    DeleteObject hFont2
    DeleteObject hFont
    DeleteObject BMP.hDcBitmap
    DeleteDC BMP.hDcMemory
    bDrew = True
End Sub



Function GetMenuItem(x As Single, y As Single) As Integer
    
    If x < picMenu.ScaleWidth - 2 And x > picMenu.ScaleWidth - 18 Then
        GetMenuItem = 9999
    End If
    
    Dim intMenus As Integer, j As Integer
    Dim intWidth As Integer, i As Integer
    Dim intStartY As Integer, textWidth As Integer, textHeight As Integer
    Dim curX As Integer, intHeight As Integer
    
    intMenus = UBound(MenuArray) + 1
    If intMenus <= 0 Then GoTo finishit
    
    i = 0
        
    textHeight = picMenu.textWidth("gW")
    intHeight = textHeight + (YM_Buffer * 2)
    
    curX = XStart
    
    For j = 1 To intMenus
        textWidth = picMenu.textWidth(MenuArray(j - 1))
        intWidth = textWidth + (XM_Buffer * 2)
    
        If y >= YStart And y <= YStart + intHeight And _
           x >= curX And x <= curX + intWidth Then
            GetMenuItem = j
            Exit Function
        End If
                
        curX = curX + intWidth
        i = i + 1
    Next
    
    Exit Function
finishit:
    GetMenuItem = 0

End Function

Sub DrawTaskbarServers()
    
    '* DrawTaskbarServers - for dual switchbars
    '* done on 2/17/02 with sorting/filtering abilities
    
    Dim intSeps As Integer, j As Integer, realWidth As Long, iconBuffer As Long
    Dim strTitle As String, intWidth As Integer, i As Integer, iconY  As Long, strText As String
    Dim intStartY As Integer, intDown As Integer, TextX As Long, TextY As Long, activeServer As Integer
    Dim ly As Integer, lx As Integer, curX As Long, hFont2 As Long, TextWidth2 As Long
    Dim tBrush As Long, BMP As BitmapStruc, hFont As Long, theSize As Size  '* stuff for DC (buffer)
    
    On Error Resume Next
    
    FillSwitchbar SWITCH_SERVERS, "Status", 0, True
    activeServer = CLIENT.ActiveForm.serverID
    
    '* Get number of "seperators" (not used anymore)
    intSeps = SwitchServerCount
    
    If intSeps = 0 Then
        picTask2.Cls
        Exit Sub
    End If
    
    '* Nothing to draw...go to end (so we can blit what we did draw)
    If bStretchButtons Then
        realWidth = (picTask2.ScaleWidth) - XStart - 1
        minSize = (intSeps) * (ICON_SIZE + 40)
        If realWidth < minSize Then realWidth = minSize Else minSize = realWidth
    Else
        '* remove.. once you do options
        intButtonWidth = 125
        
        realWidth = (intSeps) * intButtonWidth - XStart
        minSize = realWidth
    End If
    If realWidth + XStart + 2 >= picTask2.ScaleWidth Then realWidth = picTask2.ScaleWidth - XStart - 2
    
    '* Create the fonts to be used
    hFont = CreateFont(13, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    If hFont = 0 Then Exit Sub
    hFont2 = CreateFont(13, 6, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
    If hFont2 = 0 Then Exit Sub
    
    '* Set the Area
    BMP.Area.Left = 0
    BMP.Area.Top = 0
    BMP.Area.Right = realWidth + XStart - 2
    BMP.Area.Bottom = picTask2.ScaleHeight
    
    '* Create bitmap
    BMP.hDcMemory = CreateCompatibleDC(picTask2.hdc)
    BMP.hDcBitmap = CreateCompatibleBitmap(picTask2.hdc, picTask2.ScaleWidth, picTask2.ScaleHeight)
    BMP.hDcPointer = SelectObject(BMP.hDcMemory, BMP.hDcBitmap)
            
    If BMP.hDcMemory = 0 Or BMP.hDcBitmap = 0 Then
        DeleteObject BMP.hDcBitmap
        DeleteDC BMP.hDcMemory
        DeleteObject hFont
        Exit Sub
    End If
    
    '* SAVE!
    SaveDC BMP.hDcMemory
    
    '* Copy the background of picMenu into the DC
    tBrush = CreateSolidBrush(clrBackground)
    oldObj = SelectObject(BMP.hDcMemory, tBrush)
    tPen = CreatePen(0, 0, clrBackground)
    oldObj2 = SelectObject(BMP.hDcMemory, tPen)
    Rectangle BMP.hDcMemory, 0, 0, picTask.ScaleWidth + 1, picTask.ScaleHeight + 1
    SelectObject BMP.hDcMemory, oldObj
    SelectObject BMP.hDcMemory, oldObj2
    DeleteObject tBrush
    DeleteObject tPen
   
    '* Set The Font
    Call SelectObject(BMP.hDcMemory, hFont)
    
    '* Draw the uh..thing on the left
    tBrush = CreateSolidBrush(clrLines)
    oldObj = SelectObject(BMP.hDcMemory, tBrush)
    tPen = CreatePen(0, 1, clrLines)
    oldObj2 = SelectObject(BMP.hDcMemory, tPen)
    lx = 3
    For ly = 6 To 18 Step 2
        Rectangle BMP.hDcMemory, 4, ly, 4 + lx, ly + 1
    Next ly
    SelectObject BMP.hDcMemory, oldObj
    SelectObject BMP.hDcMemory, oldObj2
    DeleteObject tBrush
    DeleteObject tPen
    
    '* background of text transparent
    SetBkMode BMP.hDcMemory, 0
    
    intStartY = 1
    bReDraw = True
    
    If intSeps <= 0 Then GoTo finishit
    
    '* Set some variables
    intWidth = Int((realWidth / (intSeps)) - 0.5)
    GetTextExtentPoint32 BMP.hDcMemory, "wYz", 3, theSize
    textHeight = theSize.cy
    TextY = (picTask.ScaleHeight - textHeight) \ 2
    iconY = (picTask.ScaleHeight - 16) \ 2
    iconBuffer = ICON_SIZE + (XBuffer * 2) + 2
    
    curX = XStart
    
    For i = 1 To SwitchServerCount
        If activeServer = SwitchServers(i).serverID Then
            SetTextColor BMP.hDcMemory, tb_TextActive
            DrawButton BMP.hDcMemory, tb_ButtonActive, tb_HilightActive, tb_ShadowActive, curX, (YStart + 1), intWidth - 1, picTask.ScaleHeight - (YStart * 2) ' - 1
            
            picTaskIcon.BackColor = tb_ButtonActive
            intDown = 1
        ElseIf DDownWhich = i Then
            SetTextColor BMP.hDcMemory, tb_TextDown
            DrawButton BMP.hDcMemory, tb_ButtonDown, tb_HilightDown, tb_ShadowDown, curX, (YStart + 1), (intWidth - 1), picTask.ScaleHeight - (YStart * 2) '- 1
            
            picTaskIcon.BackColor = tb_ButtonDown
            intDown = 1
        ElseIf DOverWhich = i Then
            SetTextColor BMP.hDcMemory, tb_TextOver
            DrawButton BMP.hDcMemory, tb_ButtonOver, tb_HilightOver, tb_ShadowOver, curX, (YStart + 1), (intWidth - 1), picTask.ScaleHeight - (YStart * 2) ' - 1
            
            picTask.ToolTipText = " " & SwitchServers(i).strText & " "
            picTaskIcon.BackColor = tb_ButtonOver
        Else
            SetTextColor BMP.hDcMemory, tb_TextOff
            DrawButton BMP.hDcMemory, tb_ButtonOff, tb_HilightOff, tb_ShadowOff, curX, (YStart + 1), (intWidth - 1), picTask.ScaleHeight - (YStart * 2) '- 1
            
            picTaskIcon.BackColor = tb_ButtonOff
        End If
        
        '* set dimensions
        With SwitchServers(i).lRect
            .Left = curX
            .Top = YStart + 1
            .Right = intWidth + curX - 1
            .Bottom = picTask.ScaleHeight - (YStart * 2)
        End With
        
        DrawIconEx BMP.hDcMemory, curX + XBuffer * 2 + intDown + 1, iconY + intDown, SwitchServers(i).hIcon, ICON_SIZE, ICON_SIZE, 0, 0, DI_NORMAL
        
        strText = SwitchServers(i).strText
        If SwitchServers(i).bNewData Then
            'strText = strText & " (" & SwitchServers(i).NewLines & ")"
            SetTextColor BMP.hDcMemory, vbRed
        End If
        
        '* draw actual text
        strText = GetTaskText(strText, intWidth - ICON_SIZE - (XBuffer * 2) - 8)
        TextOut BMP.hDcMemory, curX + XBuffer * 2 + intDown + TextWidth2 + iconBuffer + 1, TextY + intDown, strText, Len(strText)
                
        '* increment variables
        curX = curX + intWidth
        
    Next i
    
finishit:
    BitBlt picTask2.hdc, BMP.Area.Left, BMP.Area.Top, BMP.Area.Right, BMP.Area.Bottom, BMP.hDcMemory, 0, 0, SRCCOPY
    
    '* Restore!
    RestoreDC BMP.hDcMemory, -1
    
    DeleteObject tBrush
    DeleteObject tPen
    DeleteObject hFont2
    DeleteObject hFont
    DeleteObject BMP.hDcBitmap
    DeleteDC BMP.hDcMemory
    bDrew = True
End Sub




Function GetToolTip(nIndex As Integer) As String
    Select Case nIndex
        Case 0
            GetToolTip = "Connect to the server specified in the status window of the active server"
        Case 1
            GetToolTip = "Disconnect from the server specified in the status window of the active server"
        Case 2
            GetToolTip = "Open the script editor"
        Case 3
            GetToolTip = "Change options for active user profile (" & strProfile & ")"
    End Select
End Function

Sub SetActive(strTitleX As String, serverIDx As Integer)
    Dim i As Integer
    For i = 1 To SwitchWindowCount
        Me.Caption = SwitchWindows(i).serverID & " = " & serverIDx & " And " & SwitchWindows(i).strTitle & " = " & strTitleX
        If SwitchWindows(i).serverID = serverIDx And SwitchWindows(i).strTitle = strTitleX Then
            ActiveWhich = i
            DrawTaskbarAllServers
            Exit Sub
        End If
    Next i
End Sub




Private Sub IDENT_ConnectionRequest(ByVal requestID As Long)
    
    IDENT.Close
    IDENT.Accept requestID
        
End Sub

Private Sub IDENT_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    IDENT.GetData dat, vbString
        
    If dat Like "*, *" Then
        dat = LeftR(dat, 2)
        dat = dat & " : USERID : UNIX : " & strFName
        IDENT.SendData dat
    End If
End Sub




Private Sub MDIForm_Load()
    
    '* Debugging
    If bDebug Then
        frmDebug.Show
    End If
    
    mrcnt = 0
        
    On Error Resume Next
        
    '* Center menu
    Center Me
    
    '* Menu Items
    MenuArray = Split("Connect,Tools,Edit,View,Format,Commands,Window,Help", ",")
    ReDim MenuYPos(UBound(MenuArray)) As Integer
    
    bStretchButtons = True
    
    OverWhich = -1
    
    '* version
    strVersion = "sIRC alpha v0." & String(2 - Len(CStr(App.Revision)), Asc("0")) & App.Revision & "." & lngBuild
   
    LoadUserOptions
    
    CLIENT.Picture = LoadPicture("")
    lngForeColor = vbBlack
    lngBackColor = &H80000005
    strFontName = "Verdana"
    intFontSize = 8
    strBold = Chr(BOLD)
    strUnderline = Chr(UNDERLINE)
    strColor = Chr(Color)
    strReverse = Chr(REVERSE)
    strAction = Chr(ACTION)
    'DoEvents
    
    ' will be changed with settings
    'bKeepChildrenInBounds = True
    bTimestamp = True
    strTimeFormat = "hh:mm"
    intIndent = 270
    
    
    If b_DualSwitch Then
        picTask2.Visible = True
    End If
    
    '* Load scripts
    Dim i As Integer
    scriptEngine.KillAllScripts
    For i = LBound(strScripts) To UBound(strScripts)
        If FileExists(strScripts(i)) Then
            scriptEngine.LoadScript scriptEngine.NewScript(), strScripts(i)
        ElseIf FileExists(strDefScriptFolder & strScripts(i)) Then
            scriptEngine.LoadScript scriptEngine.NewScript(), strDefScriptFolder & strScripts(i)
        End If
    Next i

    '* Show Window
    WP_ResetClient

    '* draw stuff
    DrawTaskbarAllServers
    DrawMenu
    
    '* load time..nevermind :\
    CLIENT.Caption = "sIRC - Alpha v0." & String(2 - Len(CStr(App.Revision)), Asc("0")) & App.Revision & "." & lngBuild
        
    'DropShadow Me.hWnd
    Me.Visible = True
    
    '* subclass
    'oldProcAddr = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MDIWndProc)


End Sub

Public Sub MDIForm_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then
        If menuActive Is Nothing Then
        Else
            menuActive.KillPopupMenus: menuActive.UnloadMenu
            mnuOverWhich = 0
            DrawMenu
        End If
        Exit Sub
    End If
    
    DrawMenu
    DrawTaskbarAllServers
    bDrew = False
    
    lblRestore.Left = picMenu.ScaleWidth - 400
    
    '* resize treeview
    tvServers.Height = CLIENT.ScaleHeight - 60
    tvServers.Width = picServerList.Width - 90
    shpSL.Height = CLIENT.ScaleHeight + 75
    shpSL.Width = picServerList.Width
    
    
    If XPM_Window_Auto.GetCheck(3) Then
        CLIENT.Arrange vbTileVertical
        Exit Sub
    ElseIf XPM_Window_Auto.GetCheck(2) Then
        CLIENT.Arrange vbTileHorizontal
        Exit Sub
    End If
    
    ' Keep children in bounds
    If bKeepChildrenInBounds Then
        Dim child As Form, newWidth As Integer, newHeight As Integer
        For Each child In Forms
            If TypeOf child Is MDIForm Or UCase$(child.Name) <> child.Name Then GoTo nextCHILD
            On Error Resume Next
            newHeight = child.Height: newWidth = child.Width
            If Me.ScaleWidth - child.Left < child.Width Then newWidth = Me.ScaleWidth - child.Left
            If Me.ScaleHeight - child.Top < child.Height Then newHeight = Me.ScaleHeight - child.Top
            child.Move child.Left, child.Top, newWidth, newHeight
nextCHILD:
        Next child
    End If
    
    WP_MAXALL
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'    SetWindowLong Me.hwnd, GWL_WNDPROC, oldProcAddr
    
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm
    SetParent frmDebug.hwnd, 0
    
    If CLIENT.Tag <> "NO" Then
        '* unload scripts
        scriptEngine.KillAllScripts
        End
    End If
        
End Sub

Public Sub mnu_Connect_Click()

End Sub

Public Sub mnu_Connect_Connect_Click()
    On Error Resume Next
    windowStatus(CLIENT.ActiveForm.serverID).currentNick = 0
    windowStatus(CLIENT.ActiveForm.serverID).Connect
    windowStatus(CLIENT.ActiveForm.serverID).setFocus
    windowStatus(CLIENT.ActiveForm.serverID).rt_Input.setFocus
    
    If CLIENT.IDENT.State <> sckListening Then
        CLIENT.IDENT.Close
        CLIENT.IDENT.Listen
    End If
End Sub

Public Sub mnu_Connect_Disconnect_Click()
    On Error Resume Next
    windowStatus(CLIENT.ActiveForm.serverID).Disconnect
End Sub


Public Sub mnu_Connect_NewServer_Click()
    NewConnection CStr(strServerName), intServerPort \ 1
End Sub

Private Sub mnu_Edit_Click()
    On Error Resume Next
    If CLIENT.ActiveForm.ActiveControl.Name = "rt_Output" Then
        CLIENT.mnu_Edit_Undo.Enabled = False
        CLIENT.mnu_Edit_Cut.Enabled = False
        CLIENT.mnu_Edit_Delete.Enabled = False
        CLIENT.mnu_Edit_Paste.Enabled = False
        
    Else
        CLIENT.mnu_Edit_Undo.Enabled = True
        CLIENT.mnu_Edit_Cut.Enabled = True
        CLIENT.mnu_Edit_Delete.Enabled = True
        CLIENT.mnu_Edit_Paste.Enabled = True
    End If
End Sub

Public Sub mnu_Edit_Copy_Click()

    On Error Resume Next
    Clipboard.SetText CLIENT.ActiveForm.ActiveControl.seltext, 1
    

End Sub

Public Sub mnu_Edit_Cut_Click()

    On Error Resume Next
    Clipboard.SetText CLIENT.ActiveForm.ActiveControl.seltext, 1
    CLIENT.ActiveForm.ActiveControl.seltext = ""
    
End Sub

Public Sub mnu_Edit_Delete_Click()

    On Error Resume Next
    CLIENT.ActiveForm.ActiveControl.seltext = ""
    

End Sub

Public Sub mnu_Edit_Paste_Click()

    On Error Resume Next
    CLIENT.ActiveForm.ActiveControl.seltext = Clipboard.GetText(1)
    

End Sub

Public Sub mnu_Edit_SelectAll_Click()

    On Error Resume Next
    CLIENT.ActiveForm.ActiveControl.selStart = 0
    CLIENT.ActiveForm.ActiveControl.selLength = Len(CLIENT.ActiveForm.ActiveControl.Text)
    

End Sub














Public Sub mnu_Edit_Undo_Click()
    'ehem
End Sub

Public Sub mnu_Format_Bold_Click()
    On Error Resume Next
    If TypeOf CLIENT.ActiveForm.ActiveControl Is RichTextBox Then HandleKeypress CLIENT.ActiveForm.ActiveControl, strBold
End Sub


Public Sub mnu_Format_Cancel_Click()
    On Error Resume Next
    If TypeOf CLIENT.ActiveForm.ActiveControl Is RichTextBox Then HandleKeypress CLIENT.ActiveForm.ActiveControl, Chr(15)
End Sub

Private Sub mnu_Format_Click()

    On Error GoTo errorHandler
    If TypeOf CLIENT.ActiveForm.ActiveControl Is RichTextBox And _
        CLIENT.ActiveForm.ActiveControl.Name <> "rt_Output" Then
        mnu_Format_Bold.Enabled = True
        mnu_Format_Color.Enabled = True
        mnu_Format_Reverse.Enabled = True
        mnu_Format_Underline.Enabled = True
    Else
        mnu_Format_Bold.Enabled = False
        mnu_Format_Color.Enabled = False
        mnu_Format_Reverse.Enabled = False
        mnu_Format_Underline.Enabled = False
    End If
    
    Exit Sub
    
errorHandler:
    mnu_Format_Bold.Enabled = False
    mnu_Format_Color.Enabled = False
    mnu_Format_Reverse.Enabled = False
    mnu_Format_Underline.Enabled = False
End Sub

Public Sub mnu_Format_Color_Click()
    If TypeOf CLIENT.ActiveForm.ActiveControl Is RichTextBox Then HandleKeypress CLIENT.ActiveForm.ActiveControl, strColor
End Sub


Public Sub mnu_Format_Reverse_Click()
    If TypeOf CLIENT.ActiveForm.ActiveControl Is RichTextBox Then
        HandleKeypress CLIENT.ActiveForm.ActiveControl, strReverse
        CLIENT.ActiveForm.ActiveControl.SelAlignment = vbleft
    End If
    
End Sub


Public Sub mnu_Format_Underline_Click()
    If TypeOf CLIENT.ActiveForm.ActiveControl Is RichTextBox Then HandleKeypress CLIENT.ActiveForm.ActiveControl, strUnderline
End Sub



Private Sub mnu_Tools_Options_Click()
    frmOptions.Show
End Sub

Private Sub mnu_Tools_Scripts_Click()
    frmSexIDE.Show
End Sub

Private Sub mnu_Window_AutoMax_Click()
    mnu_Window_AutoMax.Checked = Not mnu_Window_AutoMax.Checked
    On Error Resume Next
    
    If mnu_Window_AutoMax.Checked Then
        Dim ass As Form
        For Each ass In Forms
            If TypeOf ass Is MDIForm Then Else ass.WindowState = vbMaximized
        Next ass
    End If
        
End Sub

Private Sub mnu_Window_AutoTileH_Click()
    mnu_Window_AutoTileH.Checked = Not mnu_Window_AutoTileH.Checked
        
    If mnu_Window_AutoTileH.Checked Then
        CLIENT.Arrange vbTileHorizontal
        mnu_Window_AutoTileV.Checked = False
    End If
End Sub

Private Sub mnu_Window_AutoTileV_Click()
    mnu_Window_AutoTileV.Checked = Not mnu_Window_AutoTileV.Checked
    
    If mnu_Window_AutoTileV.Checked Then
        CLIENT.Arrange vbTileVertical
        mnu_Window_AutoTileH.Checked = False
    End If
End Sub


Private Sub mnu_Window_Cascade_Click()
    CLIENT.Arrange vbCascade
End Sub


Private Sub mnu_Window_Close_Click()
    On Error Resume Next
    Unload CLIENT.ActiveForm
End Sub

Private Sub mnu_Window_TileH_Click()
    CLIENT.Arrange vbTileHorizontal
    
End Sub


Private Sub mnu_Window_TileV_Click()
    CLIENT.Arrange vbTileVertical
End Sub





Private Sub mParent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim lCoords As Long
    Dim ptStart As POINTAPI
    Dim ptMove  As POINTAPI

    If (Button = vbLeftButton) Then
        'And ((meAutoDrag = Always) Or _
      (meAutoDrag = WhenTransparent And mbTransparent)) Then
        ptStart.x = mParent.Left
        ptStart.y = mParent.Top
        Call ReleaseCapture
        Call SendMessage(mParent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, &H0&)
        'Send the MouseUp.
        Call SetCapture(mParent.hwnd)
        lCoords = mParent.ScaleX(x, mParent.ScaleMode, vbPixels) _
          + (mParent.ScaleY(y, mParent.ScaleMode, vbPixels) * &H10000)
        Call PostMessage(mParent.hwnd, WM_LBUTTONUP, HTCLIENT, lCoords)
    End If
    
End Sub


Private Sub picMenu_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If HandleHotkey(ShiftKey, KeyCode, 0) Then
        KeyCode = 0
    End If

End Sub

Private Sub picMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itemHilight As Integer
    itemHilight = GetMenuItem(x, y)
    
    If itemHilight = 0 Then
        mnuOverWhich = 0
        'bPopupShown = False
        'bMenuShown = False
        If Not menuActive Is Nothing Then menuActive.UnloadMenu
        DrawMenu
        Exit Sub
    End If

    bPopupShown = True
    bMenuShown = True
    
    mnuOverWhich = itemHilight
    
    If itemHilight = 9999 Then
        On Error Resume Next
        mnuOverWhich = -1
        bPopupShown = False
        bMenuShown = False
        bMenuDrew = False
        If CLIENT.ActiveForm.WindowState = 2 Then CLIENT.ActiveForm.WindowState = 0
    Else
        DisplayMenu MenuArray(itemHilight - 1)
    End If
    DrawMenu
    
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itemHilight As Integer, startTime As Long
    itemHilight = GetMenuItem(x, y)
    
    If itemHilight = mnuOverWhich And bMenuDrew = False Then
        DrawMenu
        
        If MenuVisible() Or bMenuShown Then
            menuActive.KillPopupMenus
            Set menuActive = Nothing
        End If
        
    ElseIf itemHilight <> mnuOverWhich And bMenuDrew = True Then
        If bMenuShown = False Then
            mnuOverWhich = itemHilight
            DrawMenu
        Else
            If itemHilight = 0 Then
                Exit Sub
            End If
            
            Dim tempMenu As clsXPMenu
            Set tempMenu = menuActive
                                    
            mnuOverWhich = itemHilight
            
            If mnuOverWhich <> 9999 Then
                If itemHilight = mnuOverWhich Then
                    DisplayMenu MenuArray(mnuOverWhich - 1)
                End If
            End If
            
            DrawMenu
            
            tempMenu.KillPopupMenus
            tempMenu.UnloadMenu
        End If
    End If
    'mnuOverWhich = itemHilight
    tmrMenu.Enabled = True
End Sub


Private Sub picMenu_Paint()
    DrawMenu
End Sub







Private Sub picServerList_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If HandleHotkey(ShiftKey, KeyCode, 0) Then
        KeyCode = 0
    End If
End Sub


Private Sub picServerList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    startX = x
    startY = y
End Sub


Private Sub picServerList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        picServerList.Width = picServerList.Width - (startX - x)
        tvServers.Width = picServerList.Width - 90
        shpSL.Width = picServerList.Width
    End If
    startX = x
End Sub


Private Sub picTask_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If HandleHotkey(ShiftKey, KeyCode, 0) Then
        KeyCode = 0
    End If
End Sub


Private Sub picTask_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim realWidth As Integer, minSize As Integer, which As Integer, intWidth As Integer, wID As Integer
    
    If Button <> 1 Then
    
        Exit Sub
    End If
    
    If SwitchWindowCount = 0 Then Exit Sub
    
    which = -1
    Dim i As Integer
    For i = 1 To SwitchWindowCount
        With SwitchWindows(i)
            If x >= .lRect.Left And x <= .lRect.Right And _
               y >= .lRect.Top And y <= .lRect.Bottom Then
               which = i
               Exit For
            End If
        End With
    Next i
    
    If which = DownWhich And bDrew = False Then
        DrawTaskbarAllServers
    ElseIf which <> DownWhich Or bDrew = False Then
        DownWhich = which
        DrawTaskbarAllServers
    End If
    'Me.Caption = which
    DownWhich = which
   
End Sub

Private Sub picTask_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim realWidth As Integer, minSize As Integer, which As Integer, intWidth As Integer, wID As Integer
    
    If Button <> 1 Then
    
        'Exit Sub
    End If
    
    If SwitchWindowCount = 0 Then Exit Sub
    
    which = -1
    Dim i As Integer
    For i = 1 To SwitchWindowCount
        With SwitchWindows(i)
            If x >= .lRect.Left And x <= .lRect.Right And _
               y >= .lRect.Top And y <= .lRect.Bottom Then
               which = i
               Exit For
            End If
        End With
    Next i
    
    If which = OverWhich And bDrew = False Then
        DrawTaskbarAllServers
    ElseIf which <> OverWhich Or bDrew = False Then
        OverWhich = which
        DrawTaskbarAllServers
    End If
    OverWhich = which
    tmrTask.Enabled = True
End Sub


Private Sub picTask_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then
    
        'Exit Sub
    End If
    
    Dim whichForm As Form
    
    'Me.Caption = SwitchWindows(ActiveWhich).strText
    If ActiveWhich = DownWhich Then
        If ActiveWhich < 1 Then Exit Sub
        Set whichForm = GetFormByName(SwitchWindows(ActiveWhich).strTitle, SwitchWindows(ActiveWhich).serverID)
        ActiveWhich = -1
        DownWhich = -1
        DrawTaskbarAllServers
        whichForm.Hide
        Exit Sub
        whichForm.Resize_event
    End If

    ActiveWhich = DownWhich
    DownWhich = -1
    DrawTaskbarAllServers
    
    If ActiveWhich < 1 Then Exit Sub
    
    Set whichForm = GetFormByName(SwitchWindows(ActiveWhich).strTitle, SwitchWindows(ActiveWhich).serverID)
    On Error Resume Next
    If whichForm.Visible = False Then whichForm.Visible = True
    whichForm.setFocus
    
End Sub

Private Sub picTask_Paint()
    DrawTaskbarAllServers
End Sub

Private Sub picToolBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    
    Dim realWidth As Integer, minSize As Integer, which As Integer, intWidth As Integer, wID As Integer
        
    '* remove..
    intButtonWidth = 25
    
    realWidth = TBButtons * intButtonWidth - XStart
    minSize = realWidth
    intWidth = realWidth / (intSeps + 1)
    
    wID = realWidth \ TBButtons
    which = Int(((x - (XStart * 2) + 7) \ wID) + 0.5) + 1
            
    If which > TBButtons Then which = -1
    
    If which = tbDownWhich And btbDrew = False Then
'        DrawToolbar
    ElseIf which <> tbDownWhich Or btbDrew = False Then
        tbDownWhich = which
'        DrawToolbar
    End If
    'Me.Caption = which
    tbDownWhich = which
   
End Sub

Private Sub picToolBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim realWidth As Integer, minSize As Integer, which As Integer, intWidth As Integer, wID As Integer
        
    '* remove..
    intButtonWidth = 25
    
    realWidth = TBButtons * intButtonWidth - XStart
    minSize = realWidth
    intWidth = realWidth / (intSeps + 1)
    
    wID = realWidth \ TBButtons
    which = Int(((x - (XStart * 2) + 7) \ wID) + 0.5) + 1
            
    If which > TBButtons Then which = -1
        
    If which = tbOverWhich And btbDrew = False Then
'        DrawToolbar
        'Me.Caption = "draw" & tbOverWhich
    ElseIf which <> tbOverWhich Or btbDrew = False Then
        tbOverWhich = which
'        DrawToolbar
        'Me.Caption = "whichdraw" & tbOverWhich
    End If
    'Me.Caption = which
    tbOverWhich = which
    tmrTool.Enabled = True

End Sub


Private Sub picToolBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case tbDownWhich
        Case 1
            Call mnu_Connect_Connect_Click
        Case 2
            Call mnu_Connect_Disconnect_Click
        Case 3
            Call mnu_Tools_Scripts_Click
        Case 4
            Call mnu_Tools_Options_Click
    End Select

    tbDownWhich = -1
'    DrawToolbar
End Sub

Private Sub picTask2_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If HandleHotkey(ShiftKey, KeyCode, 0) Then
        KeyCode = 0
    End If
End Sub


Private Sub picTask2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim realWidth As Integer, minSize As Integer, which As Integer, intWidth As Integer, wID As Integer
    
    If Button <> 1 Then
    
        Exit Sub
    End If
    
    If SwitchServerCount = 0 Then Exit Sub
    
    which = -1
    Dim i As Integer
    For i = 1 To SwitchServerCount
        With SwitchServers(i)
            If x >= .lRect.Left And x <= .lRect.Right And _
               y >= .lRect.Top And y <= .lRect.Bottom Then
               which = i
               Exit For
            End If
        End With
    Next i
    
    If which = DDownWhich And bDDrew = False Then
        DrawTaskbarServers
    ElseIf which <> DownWhich Or bDDrew = False Then
        DDownWhich = which
        DrawTaskbarServers
    End If
    'Me.Caption = which
    DDownWhich = which
   

End Sub


Private Sub picTask2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim realWidth As Integer, minSize As Integer, which As Integer, intWidth As Integer, wID As Integer
    
    If Button <> 1 Then
    
        'Exit Sub
    End If
    
    If SwitchServerCount = 0 Then Exit Sub
    
    which = -1
    Dim i As Integer
    For i = 1 To SwitchServerCount
        With SwitchServers(i)
            If x >= .lRect.Left And x <= .lRect.Right And _
               y >= .lRect.Top And y <= .lRect.Bottom Then
               which = i
               Exit For
            End If
        End With
    Next i
    
    If which = DOverWhich And bDDrew = False Then
        DrawTaskbarServers
    ElseIf which <> OverWhich Or bDDrew = False Then
        DOverWhich = which
        DrawTaskbarServers
    End If
    DOverWhich = which
    tmrTask2.Enabled = True
End Sub

Private Sub picTask2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then
    
        'Exit Sub
    End If
    
    Dim whichForm As Form
    
    'Me.Caption = SwitchWindows(ActiveWhich).strText
    If DActiveWhich = DDownWhich Then
        If DActiveWhich < 1 Then Exit Sub
        Set whichForm = GetFormByName(SwitchServers(DActiveWhich).strTitle, SwitchServers(DActiveWhich).serverID)
        DActiveWhich = -1
        DDownWhich = -1
        DrawTaskbarServers
        whichForm.Hide
        Exit Sub
    End If

    DActiveWhich = DDownWhich
    DDownWhich = -1
    DrawTaskbarServers
    
    If DActiveWhich < 1 Then Exit Sub
    
    Set whichForm = GetFormByName(SwitchServers(DActiveWhich).strTitle, SwitchServers(DActiveWhich).serverID)
    On Error Resume Next
    If whichForm.Visible = False Then whichForm.Visible = True
    whichForm.setFocus

End Sub

Private Sub tmrMenu_Timer()
    Dim pt As POINTAPI, lngRet As Long, hwnd As Long
    lngRet = GetCursorPos(pt)
    
    hwnd = WindowFromPoint(pt.x, pt.y)
    If hwnd <> picMenu.hwnd And bMenuShown = False Then
        mnuOverWhich = -1
        DrawMenu
        tmrMenu.Enabled = False
    End If
End Sub

Private Sub tmrTask_Timer()
    Dim pt As POINTAPI, lngRet As Long, hwnd As Long
    lngRet = GetCursorPos(pt)
    
    hwnd = WindowFromPoint(pt.x, pt.y)
    If hwnd <> picTask.hwnd Then
        OverWhich = -1
        DrawTaskbarAllServers
        tmrTask.Enabled = False
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnu_Connect_Connect_Click
        Case 2
            mnu_Connect_Disconnect_Click
    End Select
End Sub


Private Sub tmrTask2_Timer()
    Dim pt As POINTAPI, lngRet As Long, hwnd As Long
    lngRet = GetCursorPos(pt)
    
    hwnd = WindowFromPoint(pt.x, pt.y)
    If hwnd <> picTask2.hwnd Then
        OverWhich = -1
        DrawTaskbarServers
        tmrTask2.Enabled = False
    End If

End Sub

Private Sub tmrTool_Timer()
    Dim pt As POINTAPI, lngRet As Long, hwnd As Long
    lngRet = GetCursorPos(pt)
    
    hwnd = WindowFromPoint(pt.x, pt.y)
    If hwnd <> picToolBar.hwnd Then
        tbOverWhich = -1
        'DrawToolbar
        tmrTool.Enabled = False
    End If
End Sub


Private Sub tvServers_Click()
    Dim si As Node, temp As String, serverID As Integer, whichForm As Form
    Set si = tvServers.selectedItem
    
    If si Is Nothing Then Exit Sub
    If si.Text = "Channels" Or si.Text = "Queries" Then Exit Sub
    
    si.ForeColor = vbBlack
    
    If si.parent Is Nothing Then
        windowStatus(LeftOf(si.Text, ":") \ 1).setFocus
        Exit Sub
    End If
    
    serverID = LeftOf(si.parent.parent.Text, ":") \ 1
    Set whichForm = GetFormByName(si.Text, serverID)
    
    If whichForm Is Nothing Then
    Else
        whichForm.setFocus
    End If
End Sub


Private Sub tvServers_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftKey = Shift
    If HandleHotkey(ShiftKey, KeyCode, 0) Then
        KeyCode = 0
    End If

End Sub


Private Sub tvServers_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim temp As String, serverID As Integer, whichForm As Form
    Set si = Node
    
    If si Is Nothing Then Exit Sub
    If si.parent Is Nothing Then Exit Sub
    If si.Text = "Channels" Or si.Text = "Queries" Then Exit Sub
    
    si.ForeColor = vbBlack
    
    If si.Text = "Status" Then
        windowStatus(CInt(LeftOf(si.parent.Text, ":"))).setFocus
        Exit Sub
    End If
    
    serverID = CInt(LeftOf(si.parent.parent.Text, ":"))
    Set whichForm = GetFormByName(si.Text, serverID)
    
    If whichForm Is Nothing Then
    Else
    End If
End Sub
