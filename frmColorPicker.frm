VERSION 5.00
Begin VB.Form ColorPicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Picker"
   ClientHeight    =   4770
   ClientLeft      =   5850
   ClientTop       =   2190
   ClientWidth     =   6840
   ClipControls    =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5955
      TabIndex        =   17
      Top             =   795
      Width           =   780
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   5955
      TabIndex        =   16
      Top             =   390
      Width           =   780
   End
   Begin VB.PictureBox picTheColor 
      AutoRedraw      =   -1  'True
      Height          =   1050
      Left            =   4815
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   15
      Top             =   390
      Width           =   930
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   " System Colors "
      Height          =   1485
      Left            =   4755
      TabIndex        =   12
      Top             =   3180
      Width           =   1950
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   28
         Left            =   30
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   52
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   29
         Left            =   315
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   51
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   30
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   50
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   31
         Left            =   855
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   49
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   32
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   48
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   33
         Left            =   1395
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   47
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   34
         Left            =   1665
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   46
         Top             =   1290
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   6
         Left            =   1665
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   45
         Top             =   330
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   13
         Left            =   1665
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   44
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   20
         Left            =   1665
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   43
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   27
         Left            =   1665
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   42
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   5
         Left            =   1395
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   41
         Top             =   330
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   12
         Left            =   1395
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   40
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   19
         Left            =   1395
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   39
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   26
         Left            =   1395
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   38
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   25
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   37
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   18
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   36
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   11
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   35
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   1125
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   34
         Top             =   330
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   24
         Left            =   855
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   33
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   23
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   32
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   22
         Left            =   315
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   31
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   21
         Left            =   30
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   30
         Top             =   1050
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   17
         Left            =   855
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   29
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   16
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   28
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   15
         Left            =   315
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   27
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   14
         Left            =   30
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   26
         Top             =   810
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   10
         Left            =   855
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   25
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   9
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   24
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   8
         Left            =   315
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   23
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   7
         Left            =   30
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   22
         Top             =   570
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   3
         Left            =   855
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   21
         Top             =   330
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   585
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   20
         Top             =   330
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   315
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   19
         Top             =   330
         Width           =   225
      End
      Begin VB.PictureBox picSysColor 
         AutoRedraw      =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   30
         ScaleHeight     =   135
         ScaleWidth      =   165
         TabIndex        =   18
         Top             =   330
         Width           =   225
      End
      Begin VB.Label lblBlah 
         AutoSize        =   -1  'True
         Caption         =   "  System Colors  "
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   75
         Width           =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   1920
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   1920
         Y1              =   165
         Y2              =   165
      End
   End
   Begin VB.CheckBox chkWebSafe 
      Caption         =   " Show only Web-safe Colors "
      Height          =   225
      Left            =   150
      TabIndex        =   10
      Top             =   4440
      Width           =   2430
   End
   Begin VB.PictureBox picColorSlider 
      Height          =   3900
      Left            =   4260
      Picture         =   "frmColorPicker.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   9
      Top             =   390
      Width           =   315
   End
   Begin VB.PictureBox picColorData 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   4815
      ScaleHeight     =   1395
      ScaleWidth      =   1845
      TabIndex        =   2
      Top             =   1785
      Width           =   1845
      Begin VB.TextBox txtHEX 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "FF0000"
         Top             =   1095
         Width           =   735
      End
      Begin VB.TextBox txtBp 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "0"
         ToolTipText     =   "Red"
         Top             =   690
         Width           =   405
      End
      Begin VB.TextBox txtGp 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "0"
         ToolTipText     =   "Red"
         Top             =   345
         Width           =   405
      End
      Begin VB.TextBox txtRp 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "100"
         ToolTipText     =   "Red"
         Top             =   0
         Width           =   405
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   225
         TabIndex        =   5
         Text            =   "255"
         ToolTipText     =   "Red"
         Top             =   0
         Width           =   420
      End
      Begin VB.TextBox txtG 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   225
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Green"
         Top             =   345
         Width           =   420
      End
      Begin VB.TextBox txtB 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   225
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Blue"
         Top             =   690
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hex:"
         Height          =   195
         Left            =   315
         TabIndex        =   56
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label lblB 
         AutoSize        =   -1  'True
         Caption         =   "B:          /255 (          )%"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         ToolTipText     =   "Blue"
         Top             =   720
         Width           =   1710
      End
      Begin VB.Label lblG 
         AutoSize        =   -1  'True
         Caption         =   "G:          /255 (          )%"
         Height          =   195
         Left            =   45
         TabIndex        =   7
         ToolTipText     =   "Green"
         Top             =   375
         Width           =   1725
      End
      Begin VB.Label lblR 
         AutoSize        =   -1  'True
         Caption         =   "R:          /255 (          )%"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         ToolTipText     =   "Red"
         Top             =   30
         Width           =   1725
      End
   End
   Begin VB.PictureBox picColors 
      Height          =   3900
      Left            =   150
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   1
      Top             =   390
      Width           =   3900
      Begin VB.Shape shpPick 
         Height          =   105
         Left            =   3795
         Shape           =   2  'Oval
         Top             =   -60
         Width           =   105
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "  Color Information  "
      Height          =   195
      Left            =   4920
      TabIndex        =   14
      Top             =   1530
      Width           =   1440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   317
      X2              =   445
      Y1              =   109
      Y2              =   109
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   317
      X2              =   445
      Y1              =   108
      Y2              =   108
   End
   Begin VB.Label lblSlider 
      BackStyle       =   0  'Transparent
      Caption         =   "4   3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4110
      TabIndex        =   11
      Top             =   390
      Width           =   645
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select color:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   885
   End
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long


Private rV As Integer, gV As Integer, bV As Integer
Private pX As Integer, pY As Integer

Private bmpHDC As Long, memHDC As Long, oldDC As Long

Private oldColor As Long, bSetColor As Boolean
Private Sub CreateBitmap()
    memHDC = CreateCompatibleDC(picColors.hdc)
    bmpHDC = CreateCompatibleBitmap(picColors.hdc, picColors.ScaleWidth, picColors.ScaleHeight)
    oldDC = SelectObject(memHDC, bmpHDC)
End Sub

Private Sub DestroyBitmap()
    SelectObject picColors.hdc, oldDC
    DeleteDC memHDC
    DeleteObject bmpHDC
End Sub


Public Sub DrawColors(ByVal r As Integer, ByVal g As Integer, ByVal B As Integer)
    Dim rDiff As Double, gDiff As Double, bDiff As Double, posY As Integer, _
        rC As Double, gC As Double, bC As Double
    rDiff = (0 - r) / 255
    gDiff = (0 - g) / 255
    bDiff = (0 - B) / 255
    rC = r
    gC = g
    bC = B
    
    Me.Caption = "Color Picker - Generating colors..."
    posY = 0
    For posY = 255 To 0 Step -1
        DrawFade 255 - posY, posY, posY, posY, CInt(r), CInt(g), CInt(B)  '(255 - posY), (255 - posY), (255 - posY)
        'r, g, b
        rC = rC + rDiff
        gC = gC + rDiff
        bC = bC + rDiff
        DoEvents
    Next posY
    
    BitBlt picColors.hdc, 0, 0, picColors.ScaleWidth, picColors.ScaleHeight, memHDC, 0, 0, vbSrcCopy
    picColors.DrawMode = 6
    picColors.Circle (pX, pY), 3
    picColors.DrawMode = 13
    UpdateRGB
    UpdatePicker
    picTheColor.Line (0, 0)-(picTheColor.ScaleWidth, 33), picColors.point(pX, pY), BF
    
    Me.Caption = "Color Picker"
End Sub

Sub DrawFade(YPos As Integer, ByVal r1 As Integer, ByVal g1 As Integer, ByVal b1 As Integer, r2 As Integer, g2 As Integer, b2 As Integer)
    Dim r_Diff As Double, g_Diff As Double, b_Diff As Double, posX As Integer
    Dim rC As Double, gC As Double, bC As Double, theColor As Long
    r_Diff = (r2 - r1) / 255
    g_Diff = (g2 - g1) / 255
    b_Diff = (b2 - b1) / 255
    rC = r1
    gC = g1
    bC = b1
    
    
    If chkWebSafe.value = 1 Then
        For posX = 0 To 255
            theColor = RGB(WebSafe(Abs(rC)), WebSafe(Abs(gC)), WebSafe(Abs(bC)))
            SetPixelV memHDC, posX, YPos, theColor
            
            rC = rC + r_Diff
            gC = gC + g_Diff
            bC = bC + b_Diff
        Next posX
    Else
        For posX = 0 To 255
            theColor = RGB(Abs(rC), Abs(gC), Abs(bC))
            SetPixelV memHDC, posX, YPos, theColor
            
            rC = rC + r_Diff
            gC = gC + g_Diff
            bC = bC + b_Diff
        Next posX
    End If
    
    
End Sub


Public Function GetColor(Optional lngOldColor As Long = 0) As Long
    oldColor = lngOldColor
    bSetColor = False
    
    Dim i As Integer
    For i = 0 To 34
        picSysColor(i).BackColor = GetSysColor(i)
    Next i
    Me.Visible = True
    DoEvents
    pX = 255
    pY = 0
    rV = 0: gV = 0: bV = 0
    
    Me.Visible = True
    
    CreateBitmap
    DrawColors 255, 0, 0
    
    Do
        DoEvents
        GetColor = picColors.point(pX, pY) 'RGB(rV, gV, bV)
    Loop Until bSetColor = True
    
    Unload Me
    Exit Function
End Function

Private Function GetHex(intVal As Integer) As String
    Dim strHex As String
    strHex = Hex(intVal)
    If Len(strHex) = 1 Then strHex = "0" & strHex
    GetHex = strHex
End Function

Private Sub UpdatePicker()
    If pX < 0 Then pX = 0
    If pX > 255 Then pX = 255
    If pY < 0 Then pY = 0
    If pY > 255 Then pY = 255
    
    BitBlt picColors.hdc, 0, 0, picColors.ScaleWidth, picColors.ScaleHeight, memHDC, 0, 0, vbSrcCopy
    picColors.DrawMode = 6
    picColors.Circle (pX, pY), 3
    picColors.DrawMode = 13
    UpdateRGB
    picTheColor.Line (0, 0)-(picTheColor.ScaleWidth, 33), picColors.point(pX, pY), BF
    picTheColor.Line (0, 33)-(picTheColor.ScaleWidth, 66), oldColor, BF
End Sub

Private Sub UpdateRGB()
    Dim r As Integer, g As Integer, B As Integer
    GetRGB picColors.point(pX, pY), r, g, B
    txtR = r
    txtG = g
    txtB = B
    txtRp = CInt(txtR / 2.55)
    txtGp = CInt(txtG / 2.55)
    txtBp = CInt(txtB / 2.55)
    txtHEX = GetHex(r) & GetHex(g) & GetHex(B)
End Sub

Sub GetRGB(ByRef cl As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
    Dim C As Long
    C = cl
    Red = C Mod &H100
    C = C \ &H100
    Green = C Mod &H100
    C = C \ &H100
    Blue = C Mod &H100
End Sub
Function WebSafe(intVal As Integer) As Integer
    Select Case intVal
        Case 0, 51, 102, 153, 204, 255
            WebSafe = intVal
        Case Else
            If intVal <= 26 Then
                WebSafe = 0: Exit Function
            ElseIf intVal > 26 And intVal <= 76 Then
                WebSafe = 51: Exit Function
            ElseIf intVal > 76 And intVal <= 127 Then
                WebSafe = 102: Exit Function
            ElseIf intVal > 127 And intVal <= 178 Then
                WebSafe = 153: Exit Function
            ElseIf intVal > 178 And intVal <= 229 Then
                WebSafe = 204: Exit Function
            ElseIf intVal > 229 Then
                WebSafe = 255: Exit Function
            End If
    End Select
End Function

Private Sub chkWebSafe_Click()
    DrawColors rV, gV, bV
End Sub









Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    bSetColor = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyBitmap
End Sub


Private Sub picColors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    pX = x
    pY = y
    UpdatePicker
    
End Sub


Private Sub picColors_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        pX = x
        pY = y
        UpdatePicker
    End If
End Sub


Private Sub picColors_Paint()
    DrawColors rV, gV, bV
End Sub

Private Sub picColorSlider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblSlider.Top = picColorSlider.Top + y - 2
End Sub


Private Sub picColorSlider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If y > 0 And y < picColorSlider.ScaleHeight Then
            lblSlider.Top = picColorSlider.Top + y - 2
        End If
    End If
End Sub


Private Sub picColorSlider_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GetRGB picColorSlider.point(0, y), rV, gV, bV
    
    DrawColors rV, gV, bV
    'picColors.BackColor = RGB(Abs(r), Abs(g), Abs(b))
End Sub


Private Sub picSysColor_Click(Index As Integer)
    'Dim r As Integer, g As Integer, b As Integer
    GetRGB picSysColor(Index).BackColor, rV, gV, bV
    pX = 255
    pY = 0
    DrawColors rV, gV, bV
    
End Sub


