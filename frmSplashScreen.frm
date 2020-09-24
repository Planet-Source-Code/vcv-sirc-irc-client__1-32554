VERSION 5.00
Begin VB.Form SplashScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   2025
   ClientTop       =   2190
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_PERCENT = 5
Private CurPer As Integer
Private PrevText As String

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Declare Function GetTickCount Lib "kernel32" () As Long




Sub DrawPercent(pic As PictureBox, lngPercent As Long, lngMax As Long)
    Dim I As Integer, startPer As Integer, endPer As Integer, j As Integer
    startPer = (pic.ScaleWidth / 100) * (CurPer / lngMax * 100)
    endPer = (pic.ScaleWidth / 100) * (lngPercent / lngMax * 100)
    pic.Cls
    For I = startPer To endPer Step 5
        pic.ForeColor = RGB(0, 0, ((I / pic.ScaleWidth) * 100) + 64)
        pic.Line (0, 0)- _
             (I, pic.ScaleHeight), pic.ForeColor, BF
        DoEvents
    Next I
    CurPer = lngPercent
End Sub
Sub SetStat(strText As String)
    DoEvents
    Dim y As Integer, nHeight As Integer, I As Integer
    
    If PrevText = "" Then
        picBuffer.Cls
        picBuffer.Print strText
        picDisplay.Picture = picBuffer.Image
        DoEvents
        PrevText = strText
    Else
        nHeight = picBuffer.textHeight(strText)
        picDisplay.Height = nHeight
        For y = 0 To nHeight Step 3
            picBuffer.Cls
            picBuffer.CurrentY = 0 - y
            picBuffer.Print PrevText
            picBuffer.Print strText
            picDisplay.Picture = picBuffer.Image
            
                DoEvents
            
        Next y
        PrevText = strText
    End If
End Sub

Private Sub Form_Load()
    CenterDialog Me
    Me.Visible = True
    DoEvents
    
    If Right(App.PATH, 1) <> "/" Then slash$ = "/"
    PATH = App.PATH & slash$
    
    strGlobalINI = PATH & "sIRC.ini"
    
    lblVersion.Caption = "0." & App.Minor & App.Revision
        
    
    SetStat "Drawing toolbar..."
    DrawPercent picP, 1, MAX_PERCENT
    CLIENT.DrawToolbar
    SetStat "Initializing language packs..."
    DrawPercent picP, 2, MAX_PERCENT
    Language_Init
    SetStat "Loading language ""English""..."
    DrawPercent picP, 3, MAX_PERCENT
    LoadLanguage "english"
    SetStat "Setting Language..."
    DrawPercent picP, 4, MAX_PERCENT
    SetLang_Menu
    SetStat "Load color information..."
    DrawPercent picP, 5, MAX_PERCENT
    LoadColors
    DrawPercent picP, 6, MAX_PERCENT
    
    DoEvents
    
    Unload Me
    Load CLIENT
    CLIENT.Show
End Sub


