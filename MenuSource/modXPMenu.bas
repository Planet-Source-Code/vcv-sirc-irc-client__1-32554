Attribute VB_Name = "modXPMenu"

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect, ByVal bErase As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const DFC_MENU = 2
Public Const DFCS_MENUBULLET = 2
Public Const DFCS_MENUCHECK = 1
Public Const FW_LIGHT = 300
Public Const DEFAULT_CHARSET = 1
Public Const DEFAULT_PITCH = 0
Sub HandleClick(menuName As String, itemNum As Integer, strItemText As String)
    On Error Resume Next
    setFocus CLIENT.ActiveForm.hwnd

    Select Case menuName
        '* *******
        '* Connect
        '* *******
        Case "Connect"
            Select Case itemNum
                Case 1
                    CLIENT.mnu_Connect_NewServer_Click
                Case 3
                    CLIENT.mnu_Connect_Connect_Click
                Case 4
                    CLIENT.mnu_Connect_Disconnect_Click
                Case 6
                    Unload CLIENT
            End Select
        '* *****
        '* Tools
        '* *****
        Case "Tools"
            Select Case itemNum
                Case 1
                    Load frmOptions
                    frmOptions.Show
                Case 2
                    Load frmSexIDE
                Case 4
                    CLIENT.Tag = "NO"
                    Unload CLIENT
                    frmLoadProfile.Show
            End Select
        '* ****
        '* Edit
        '* ****
        Case "Edit"
            Select Case itemNum
                Case 1
                    SendMessage CLIENT.ActiveForm.ActiveControl.hwnd, EM_UNDO, 0, 0
                Case 3
                    CLIENT.mnu_Edit_Cut_Click
                Case 4
                    CLIENT.mnu_Edit_Copy_Click
                Case 5
                    CLIENT.mnu_Edit_Paste_Click
                Case 6
                    CLIENT.mnu_Edit_Delete_Click
                Case 8
                    CLIENT.mnu_Edit_SelectAll_Click
            End Select
        '* ****
        '* View
        '* ****
        Case "View"
            Select Case itemNum
                Case 1  'debug window
                    'nothing
                Case 2  'uhmmmuhmm whats it called!..oh yeah..treeview
                    XPM_View.SetCheck 2, Not XPM_View.GetCheck(2)
                    CLIENT.picServerList.Visible = XPM_View.GetCheck(2)
                Case 4
                    Select Case XPM_View.GetItemText(4)
                        Case "Server Bar"
                            XPM_View.SetCheck 4, Not XPM_View.GetCheck(4)
                            CLIENT.ActiveForm.bShowServerInfo = XPM_View.GetCheck(4)
                            CLIENT.ActiveForm.ResizeToolbar
                            CLIENT.ActiveForm.Resize_event
                        Case "Topic Bar"
                            XPM_View.SetCheck 4, Not XPM_View.GetCheck(4)
                            CLIENT.ActiveForm.bShowChannelInfo = XPM_View.GetCheck(4)
                            CLIENT.ActiveForm.ResizeToolbar
                            CLIENT.ActiveForm.Resize_event
                    End Select
            End Select
        '* ******
        '* Window
        '* ******
        Case "Window"
            Select Case itemNum
                Case 1  'close
                    Unload CLIENT.ActiveForm
                Case 3  'cascade
                    CLIENT.Arrange vbCascade
                Case 4  'tile horizontal
                    CLIENT.Arrange vbTileHorizontal
                Case 5
                    CLIENT.Arrange vbTileVertical
                Case 7  ' maximize
                    WP_SetMax CLIENT.ActiveForm
                Case 8  ' maximize all
                    Dim frm As Form
                    For Each frm In Forms
                        If frm.bShowInTaskbar Then WP_SetMax frm
                    Next
            End Select
        '* ******
        '* Window _ Auto
        '* ******
        Case "Window_Auto"
            With XPM_Window_Auto
                Select Case itemNum
                    Case 1
                        .SetCheck 1, Not .GetCheck(1)
                        .SetCheck 2, False
                        .SetCheck 3, False
                    Case 2
                        .SetCheck 2, Not .GetCheck(2)
                        .SetCheck 1, False
                        .SetCheck 3, False
                        
                        CLIENT.Arrange vbTileHorizontal
                    Case 3
                        .SetCheck 3, Not .GetCheck(3)
                        .SetCheck 2, False
                        .SetCheck 1, False
                        
                        CLIENT.Arrange vbTileVertical
                End Select
            End With
        '* ******
        '* Window _ Remember
        '* ******
        Case "Window_Remember"
            Select Case itemNum
                Case 1      ' remember current window
                    Call WP_RememberWindow(CLIENT.ActiveForm)
                Case 2      ' remember client
                    Call WP_RememberClient
                Case 3
                    Call WP_RememberAll
            End Select
        '* ******
        '* Window _ Forget
        '* ******
        Case "Window_Forget"
            Select Case itemNum
                Case 1      ' remember current window
                    Call WP_ForgetWindow(CLIENT.ActiveForm)
                Case 2      ' remember client
                    Call WP_ForgetClient
                Case 3
                    Call WP_ForgetAll
            End Select
        '* ******
        '* Window - Reset
        '* ******
        Case "Window_Reset"
            Select Case itemNum
                Case 1
                    Call WP_ResetWindow(CLIENT.ActiveForm)
                Case 2
                    Call WP_ResetClient
                Case 3
                    Call WP_ResetAll
            End Select
        '* ****
        '* Help
        '* ****
        Case "Help"
            Select Case itemNum
                Case 1
                    If FileExists(PATH & "help\sirc.hlp") Then
                        ShellExecute 0, "open", PATH & "help\sirc.hlp", "", "", 10
                    Else
                        MsgBox "Help file could not be found!", vbCritical
                    End If
                Case 3
                    Load frmAbout
                    frmAbout.Show
                    frmAbout.Visible = True
                    StayOnTop frmAbout, True
                    SendMessage frmAbout.hwnd, WM_SETFOCUS, 0, vbNullString
            End Select
    End Select
End Sub


