VERSION 5.00
Begin VB.Form frmXPMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2760
   ClientLeft      =   4935
   ClientTop       =   1920
   ClientWidth     =   1875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000011&
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
   Moveable        =   0   'False
   ScaleHeight     =   184
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   125
   ShowInTaskbar   =   0   'False
   Tag             =   "XPMenu"
   Visible         =   0   'False
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   825
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   465
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrKillPopup 
      Enabled         =   0   'False
      Interval        =   290
      Left            =   705
      Top             =   840
   End
   Begin VB.Timer tmrHover 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrActive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   2115
   End
End
Attribute VB_Name = "frmXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XPMenuClass As clsXPMenu
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public upY As Single

Public bShowInTaskbar As Boolean
Public serverID As Integer
Public strTitle As String
Public bVisible As Boolean
Public Function ShowWindow() As Boolean
    'ShowWindow = AnimateWindow(Me.hwnd, 500, 0)
    'Me.Visible = True
End Function


Private Sub Form_Click()
    Dim selectedItem As Long
    selectedItem = XPMenuClass.GetHilightedItem(upY)
    
    If selectedItem = 0 Then Exit Sub
    If XPMenuClass.IsTextItem(CInt(selectedItem)) And XPMenuClass.GetDisable(CInt(selectedItem)) = False Then
        XPMenuClass.KillAllMenus
        
        HandleClick XPMenuClass.GetMenuName(), CInt(selectedItem), XPMenuClass.GetItemText(CInt(selectedItem))
    End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then    'UP!
        XPMenuClass.SetPreviousHilightItem
    ElseIf KeyCode = 40 Then    'DOWN!
        XPMenuClass.SetNextHilightItem
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If XPMenuClass.GetHilightNum < 1 Then
            XPMenuClass.KillAllMenus
        Else
            Dim selectedItem As Long
            selectedItem = XPMenuClass.GetHilightNum
            If selectedItem < 1 Then Exit Sub
            If XPMenuClass.IsTextItem(CInt(selectedItem)) And XPMenuClass.GetDisable(CInt(selectedItem)) = False Then
                XPMenuClass.KillAllMenus
                HandleClick XPMenuClass.GetMenuName(), CInt(selectedItem), XPMenuClass.GetItemText(CInt(selectedItem))
            Else
                Beep
            End If
        End If
    End If
End Sub


Private Sub Form_Load()
    StayOnTop Me, True
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim getHilight As Long
    getHilight = XPMenuClass.GetHilightedItem(y)
    
    If getHilight = XPMenuClass.GetHilightNum Then Exit Sub
    XPMenuClass.setHilightedItem CInt(getHilight)

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    upY = y
End Sub


Private Sub Form_Paint()
    On Error Resume Next
    XPMenuClass.DrawMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bVisible = False
End Sub

Private Sub tmrActive_Timer()
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Tag = "XPMenu" And GetActiveWindow() = frm.hwnd Then Exit Sub
    Next frm
    
    tmrActive.Enabled = False
    
    XPMenuClass.KillAllMenus
    XPMenuClass.UnloadMenu
    
    DoEvents
    CLIENT.bMenuShown = False
    CLIENT.mnuOverWhich = 0
    CLIENT.DrawMenu
    
End Sub


Private Sub tmrHover_Timer()
    Dim pt As POINTAPI
    GetCursorPos pt
    
    Dim hw As Long
    hw = WindowFromPoint(pt.x, pt.y)
    
    If hw <> Me.hwnd Then
        If XPMenuClass.PopupShown() = False Then
            XPMenuClass.setHilightedItem -1
            XPMenuClass.DrawMenu
            tmrHover.Enabled = False
        End If
    End If
End Sub


Private Sub tmrKillPopup_Timer()
    On Error Resume Next
    'If XPMenuClass.MenuItems(tmrKillPopup.Tag).bPopupmenu Then
    XPMenuClass.KillSpecPopup tmrKillPopup.Tag
    tmrKillPopup.Enabled = False
    'End If
End Sub


