VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form QUERY 
   BackColor       =   &H00FFFFFF&
   Caption         =   "<nick> (user@host)"
   ClientHeight    =   4575
   ClientLeft      =   4860
   ClientTop       =   4095
   ClientWidth     =   6660
   Icon            =   "frm_Query.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   Begin RichTextLib.RichTextBox rt_Input 
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   4155
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   476
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      MultiLine       =   0   'False
      MaxLength       =   512
      Appearance      =   0
      OLEDropMode     =   1
      TextRTF         =   $"frm_Query.frx":038A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rt_Output 
      Height          =   4065
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   7170
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   1
      TextRTF         =   $"frm_Query.frx":0406
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line ln_Sep 
      BorderColor     =   &H8000000F&
      X1              =   0
      X2              =   434
      Y1              =   273
      Y2              =   273
   End
End
Attribute VB_Name = "QUERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oldProcAddr      As Long
Private curWS           As Integer
Public intShift         As Integer
Public bWindowInUse     As Boolean
Public bShowInTaskbar   As Boolean
Public serverID         As Integer
Public strNick          As String
Public strHost          As String
Public strTitle         As String
Public queryNum         As Long

Public bNewData         As Boolean
Public intNewLines      As Integer

Private startX As Long, startY As Long

Private textHistory As New Collection
Private intCurHist  As Integer

Function AddHistory(strText As String) As Boolean
    'On Error Resume Next
    If intCurHist <> 0 And intCurHist <= textHistory.Count Then
        If textHistory.item(intCurHist) = strText Then
            AddHistory = False
            Exit Function
        End If
    End If

    If textHistory.Count > MAX_TEXT_HISTORY Then
        textHistory.Remove 1
    End If
    textHistory.Add strText
    intCurHist = textHistory.Count + 1
    AddHistory = True
End Function


Public Sub Resize_event()

    If Me.WindowState <> curWS Then
        If Me.WindowState = vbMaximized Then
            CLIENT.DrawMenu
        ElseIf Me.WindowState = vbNormal Then
            
        End If
        curWS = Me.WindowState
    End If
    

    If Me.WindowState = vbMinimized Then Exit Sub
    
    On Error Resume Next
    Dim intTextHeight As Integer
    Me.FontName = rt_Input.Font.Name
    Me.FontSize = rt_Input.Font.Size
    intTextHeight = Me.textHeight("Ab_¯") + 4     '# 4 for buffer (2 on top, 2 bottom)
        
    rt_Output.Top = 1
    rt_Output.Width = Me.ScaleWidth - 5 'buffer
    rt_Output.Height = Me.ScaleHeight - (intTextHeight + 7) - (rt_Output.Top - 1) '# 7 for buffer
    rt_Input.Top = Me.ScaleHeight - intTextHeight
    rt_Input.Height = intTextHeight + 4
    rt_Input.Width = Me.ScaleWidth - 4
    ln_Sep.Y1 = Me.ScaleHeight - (intTextHeight + 4)  '# 4 for buffer again
    ln_Sep.Y2 = ln_Sep.Y1
    ln_Sep.X2 = Me.ScaleWidth

End Sub


Function WinType() As String
    WinType = "Query"
End Function

Public Function GetTitle() As String
    GetTitle = strNick
End Function


Public Sub UpdateCaption()
    strTitle = strNick
    If strHost = "" Then
        Me.Caption = strNick
    Else
        Me.Caption = strNick & " (" & strHost & ")"
    End If
End Sub

Private Sub Form_Activate()
    
    RESIZECLIENT
        
    bNewData = False
    CurrentServerID = serverID
    CLIENT.SetActive strTitle, serverID
    rt_Input.setFocus
    
    Dim nIndex As Integer
    nIndex = treeview_GetQueryIndex(CLIENT.tvServers, strNick, serverID)
    If nIndex <> -1 Then
        Set CLIENT.tvServers.selectedItem = CLIENT.tvServers.Nodes.item(nIndex)
        CLIENT.tvServers.Nodes.item(nIndex).ForeColor = vbBlack
    End If
    
    XPM_View.SetText 3, "No window bar present"
    XPM_View.SetCheck 3, False
    XPM_View.SetDisable 3, True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    intShift = Shift
    If HandleHotkey(intShift, KeyCode, serverID) Then
        KeyCode = 0
    End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    intShift = 0
End Sub


Private Sub Form_Load()

    oldProcAddr = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf QueryWndProc)

    bShowInTaskbar = True
    WP_ResetWindow Me
    
    CLIENT.DrawTaskbarAllServers
    
End Sub

Private Sub Form_Resize()
    Resize_event
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong Me.hwnd, GWL_WNDPROC, oldProcAddr
    treeview_RemoveQuery CLIENT.tvServers, strNick, serverID
    
    '* update dynamic menu
    XPM_ServerMenu(serverID).SetVisible queryNum + MAX_CHANS + 3, False
    XPM_ServerMenu(serverID).SetText queryNum + MAX_CHANS + 3, strNick
    
    strNick = ""
    serverID = 0
    QueriesInUse(queryNum) = False
    bWindowInUse = False
    bShowInTaskbar = False
    strTitle = ""
    
End Sub

Private Sub rt_Input_Change()
    rt_Input.Font.Name = strFontName
    rt_Input.Font.Size = intFontSize
End Sub


Private Sub rt_Input_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then    'UP KEY!
        If intCurHist <= 1 Then Beep: Exit Sub
        
        If intCurHist >= textHistory.Count And rt_Input.Text <> "" Then
            If AddHistory(rt_Input.Text) Then
                intCurHist = intCurHist - 1
            End If
        End If
        
        intCurHist = intCurHist - 1
        rt_Input.Text = textHistory.item(intCurHist)
        rt_Input.selStart = Len(rt_Input.Text)
        KeyCode = 0
    ElseIf KeyCode = 40 Then    'down key!
        If intCurHist >= textHistory.Count Then
            If rt_Input.Text <> "" Then
                If AddHistory(rt_Input.Text) = False Then
                    intCurHist = intCurHist + 1
                End If
                rt_Input.Text = ""
                KeyCode = 0
            Else
                Beep
            End If
            Exit Sub
        End If
        
        intCurHist = intCurHist + 1
        rt_Input.Text = textHistory.item(intCurHist)
        rt_Input.selStart = Len(rt_Input.Text)
        KeyCode = 0
    End If

End Sub


Private Sub rt_Input_KeyPress(KeyAscii As Integer)
    Dim params(3) As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If rt_Input.Text = "" Then Exit Sub
        
        If Left$(rt_Input.Text, Len(COMMANDCHAR)) = COMMANDCHAR Then
        
            If Len(rt_Input.Text) = 1 Then Exit Sub
            Dim strData As String
            strData = Right$(rt_Input.Text, Len(rt_Input.Text) - 1)
            Dim argsX() As String
            argsX = Split(strData, " ")
            
            If DoCommandLine(argsX, strNick, serverID) = False Then
                windowStatus(serverID).SendData strData
            End If
        Else
            windowStatus(serverID).SendData "PRIVMSG " & strNick & " :" & rt_Input.Text
            params(1) = strNick
            params(2) = windowStatus(serverID).strCurNick
            params(3) = rt_Input.Text
            Dim vars(2) As String
            vars(0) = "target:" & strNick
            vars(1) = "nick:" & windowStatus(serverID).strCurNick
            vars(2) = "message:" & rt_Input.Text
            
            scriptEngine.ExecuteEvent "text", params, serverID, vars
        End If
        
        rt_Input.Text = ""
    End If

End Sub


Private Sub rt_Input_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim point As POINTAPI
        GetCursorPos point
        XPM_Edit.ShowMenu point.x * 15, point.y * 15
    End If
End Sub

Private Sub rt_Output_KeyPress(KeyAscii As Integer)
    rt_Input.setFocus
    rt_Input.seltext = Chr(KeyAscii)
    KeyAscii = 0
End Sub


Private Sub rt_Output_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HideCaret rt_Output.hwnd
    startX = x
    startY = y
End Sub


Private Sub rt_Output_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then HideCaret rt_Output.hwnd

End Sub


Private Sub rt_Output_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim point As POINTAPI
        GetCursorPos point
        XPM_Edit.ShowMenu point.x * 15, point.y * 15
    End If
    
    If startX = x And startY = y Then
        rt_Input.setFocus
    End If
End Sub


