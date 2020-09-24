Attribute VB_Name = "modHotkeys"
Private Enum VIRTKEYS
    VIRT_SHIFT = 1
    VIRT_CTRL
    VIRT_CTRLSHIFT
    VIRT_ALT
    VIRT_ALTSHIFT
    VIRT_ALTCTRL
    VIRT_ALTCTRLSHIFT
End Enum

Private Enum FKEYS
    KEY_F1 = 112
    KEY_F2
    KEY_F3
    KEY_F4
    KEY_F5
    KEY_F6
    KEY_F7
    KEY_F8
    KEY_F9
    KEY_F10
    KEY_F11
    KEY_F12
End Enum

Private Enum LKEYS
    KEY_A = 65
    KEY_B
    KEY_C
    KEY_D
    KEY_E
    KEY_F
    KEY_G
    KEY_H
    KEY_I
    KEY_J
    KEY_K
    KEY_L
    KEY_M
    KEY_N
    KEY_O
    KEY_P
    KEY_Q
    KEY_R
    KEY_S
    KEY_T
    KEY_U
    KEY_V
    KEY_W
    KEY_X
    KEY_Y
    KEY_Z
End Enum
    
Private Function GetKeyString(intValue As Integer) As String

    Select Case intValue
        Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123
            GetKeyString = "f" & CStr(intValue - 111)
        Case 9
            GetKeyString = "tab"
        Case 13
            GetKeyString = "enter"
        Case Else
            GetKeyString = Chr(intValue)
    End Select
End Function

Public Function HandleHotkey(intShift As Integer, intCharX As Integer, Optional serverID As Integer = 0) As Boolean
    If intShift = 0 Then Exit Function

    Dim intChar As Integer, KeyPressed As String
    intChar = intCharX
    KeyPressed = GetKeyString(intChar)

    If intChar = 16 Or intChar = 17 Or intChar = 18 Then    'alt, ctrl, shift : ignore
        HandleHotkey = False
        Exit Function
    End If
    
    Dim params(1) As String, localVars(2) As String, res As Integer
    params(0) = "blah"
    
    Select Case intShift
    Case VIRT_CTRL   ' Control key...not so obviously
        Select Case KeyPressed
            Case "B"
                CLIENT.mnu_Format_Bold_Click
                HandleHotkey = True
            Case "U"
                CLIENT.mnu_Format_Underline_Click
                HandleHotkey = True
            Case "R"
                CLIENT.mnu_Format_Reverse_Click
                HandleHotkey = True
            Case "O"
                CLIENT.mnu_Format_Cancel_Click
                HandleHotkey = True
            Case "K"
                CLIENT.mnu_Format_Color_Click
                HandleHotkey = True
            Case "N"
                CLIENT.mnu_Connect_NewServer_Click
                HandleHotkey = True
            Case "S"
                frmSexIDE.Show
            Case Else
                localVars(1) = "virtkey:ctrl"
                localVars(2) = "key:" & KeyPressed
                If scriptEngine.ExecuteHotkey("ctrl", KeyPressed, params(), serverID, localVars()) Then
                   HandleHotkey = True
                Else
                   HandleHotkey = False
                End If
        End Select
    Case VIRT_SHIFT
        localVars(1) = "virtkey:shift"
        localVars(2) = "key:" & KeyPressed
        If scriptEngine.ExecuteHotkey("shift", KeyPressed, params(), serverID, localVars()) Then
            HandleHotkey = True
        Else
            HandleHotkey = False
        End If
    Case VIRT_CTRLSHIFT
        localVars(1) = "virtkey:ctrl+shift"
        localVars(2) = "key:" & KeyPressed
        If scriptEngine.ExecuteHotkey("shift", KeyPressed, params(), serverID, localVars()) Then
            HandleHotkey = True
        Else
            HandleHotkey = False
        End If
    Case VIRT_ALT
        localVars(1) = "virtkey:alt"
        localVars(2) = "key:" & KeyPressed
        If scriptEngine.ExecuteHotkey("shift", KeyPressed, params(), serverID, localVars()) Then
            HandleHotkey = True
        Else
            HandleHotkey = False
        End If
    Case VIRT_ALTSHIFT
        localVars(1) = "virtkey:alt+shift"
        localVars(2) = "key:" & KeyPressed
        If scriptEngine.ExecuteHotkey("shift", KeyPressed, params(), serverID, localVars()) Then
            HandleHotkey = True
        Else
            HandleHotkey = False
        End If
    Case VIRT_ALTCTRL
        localVars(1) = "virtkey:alt+ctrl"
        localVars(2) = "key:" & KeyPressed
        If scriptEngine.ExecuteHotkey("shift", KeyPressed, params(), serverID, localVars()) Then
            HandleHotkey = True
        Else
            HandleHotkey = False
        End If
    Case VIRT_ALTCTRLSHIFT
        localVars(1) = "virtkey:alt+ctrl+shift"
        localVars(2) = "key:" & KeyPressed
        If scriptEngine.ExecuteHotkey("shift", KeyPressed, params(), serverID, localVars()) Then
            HandleHotkey = True
        Else
            HandleHotkey = False
        End If
    
    Case Else
        HandleHotkey = False
    End Select
        
End Function


