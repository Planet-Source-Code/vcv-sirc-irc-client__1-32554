Attribute VB_Name = "modMenus"
Public XPM_Connect As New clsXPMenu
Public XPM_Tools As New clsXPMenu
Public XPM_Edit As New clsXPMenu
Public XPM_View As New clsXPMenu
Public XPM_Format As New clsXPMenu
Public XPM_Commands As New clsXPMenu
Public XPM_Window As New clsXPMenu
Public XPM_Window_Auto As New clsXPMenu
Public XPM_Window_Remember As New clsXPMenu
Public XPM_Window_Forget As New clsXPMenu
Public XPM_Window_Reset As New clsXPMenu
Public XPM_Help As New clsXPMenu
Public XPM_ServerMenu(5) As New clsXPMenu
Sub CreateMenu_Tools()
    With XPM_Tools
        .Init "Tools", CLIENT.ilMenu
        .AddItem 3, "Options...", False, False
        '.AddItem 0, "", False, True
        .AddItem 4, "Scripts...", False, False
        .AddItem 0, "", False, True
        .AddItem 5, "Change Profile...", False, False
    End With
    
End Sub

Sub CreateMenus()
    CreateMenu_Connect
    CreateMenu_Tools
    CreateMenu_Edit
    CreateMenu_View
    CreateMenu_Format
    CreateMenu_Commands
    CreateMenu_Window
    CreateMenu_Help
End Sub

Sub CreateMenu_Connect()
    With XPM_Connect
        .Init "Connect", CLIENT.ilMenu
        .AddItem 11, "New Server...", False, False
        .AddItem 0, "", False, True
        .AddItem 1, "Connect", False, False
        .AddItem 2, "Disconnect", False, False
        .AddItem 0, "", False, True
        .AddItem 0, "Exit", False, False
    End With
        
End Sub

Sub CreateMenu_View()
    With XPM_View
        .Init "View", CLIENT.ilMenu
        .AddItem 0, "Debug Window", False, False
        If bDebug Then
            .SetCheck 1, True
        Else
            .SetDisable 1, True
        End If
        .AddItem 0, "Treeview display", False, False
        .AddItem 0, "", False, True
        .AddItem 0, "No window bar present", False, False
        .SetDisable 3, True
        
    End With
        
End Sub

Sub CreateMenu_Format()
    With XPM_Format
        .Init "Format", CLIENT.ilMenu
        .AddItem 0, "Menu not done", False, False
        .SetDisable 1, True
        .AddItem 0, "Try again soon :)", False, False
        .SetDisable 2, True
    End With
        
End Sub

Sub CreateMenu_Commands()
    With XPM_Commands
        .Init "Commands", CLIENT.ilMenu
        .AddItem 0, "Menu not done", False, False
        .SetDisable 1, True
        .AddItem 0, "Try again soon :)", False, False
        .SetDisable 2, True
    End With
        
End Sub
Sub CreateMenu_Help()
    With XPM_Help
        .Init "Help", CLIENT.ilMenu
        .AddItem 12, "Contents", False, False
        .AddItem 0, "", False, True
        .AddItem 5, "About sIRC...", False, False
        
    End With
        
End Sub
Sub CreateMenu_Window()
    With XPM_Window_Auto
        .Init "Window_Auto"
        .AddItem 0, "Maximize", False, False
        .AddItem 0, "Tile Horizontally", False, False
        .AddItem 0, "Tile Vertically", False, False
    End With
    
    With XPM_Window_Remember
        .Init "Window_Remember"
        .AddItem 0, "Current Window", False, False
        .AddItem 0, "Client (Main Window)", False, False
        .AddItem 0, "All Windows", False, False
    End With
    
    With XPM_Window_Forget
        .Init "Window_Forget"
        .AddItem 0, "Current Window", False, False
        .AddItem 0, "Client (Main Window)", False, False
        .AddItem 0, "All Windows", False, False
    End With
    
    With XPM_Window_Reset
        .Init "Window_Reset"
        .AddItem 0, "Current Window", False, False
        .AddItem 0, "Client (Main Window)", False, False
        .AddItem 0, "All Windows", False, False
    End With
    
    With XPM_Window
        .Init "Window", CLIENT.ilMenu
        .AddItem 0, "Close", False, False
        .AddItem 0, "", False, True
        .AddItem 13, "Cascade", False, False
        .AddItem 0, "Tile Horizontally", False, False
        .AddItem 0, "Tile Vertically", False, False
        .AddItem 0, "", False, True
        .AddItem 0, "Maximize", False, False
        .AddItem 0, "Maximize All", False, False
        .AddItem 0, "", False, True
        .AddItem 0, "Auto", True, False, XPM_Window_Auto
        .AddItem 0, "", False, True
        .AddItem 0, "Remember", True, False, XPM_Window_Remember
        .AddItem 0, "Forget", True, False, XPM_Window_Forget
        .AddItem 0, "Reset", True, False, XPM_Window_Reset
        .AddItem 0, "", False, True
        Dim i As Integer, j As Integer
        For i = 1 To MAX_CONNS
            'Set XPM_ServerMenu(i) = New clsXPMenu
            .SetVisible .AddItem(14, "", True, False, XPM_ServerMenu(i)), False
            'Set XPM_ServerMenu(i) = New clsXPMenu
            XPM_ServerMenu(i).Init "Server " & i, CLIENT.ilTreeView
            XPM_ServerMenu(i).AddItem 1, "Status", False, False
            XPM_ServerMenu(i).AddItem 0, "", False, True
            For j = 1 To MAX_CHANS
                XPM_ServerMenu(i).SetVisible XPM_ServerMenu(i).AddItem(2, "", False, False), False
            Next j
            XPM_ServerMenu(i).AddItem 0, "", False, True
            For j = 1 To MAX_QUERIES
                XPM_ServerMenu(i).SetVisible XPM_ServerMenu(i).AddItem(3, "", False, False), False
            Next j
        Next i
    End With
        
End Sub


Sub CreateMenu_Edit()
    With XPM_Edit
        .Init "Edit", CLIENT.ilMenu
        .AddItem 6, "Undo", False, False
        .AddItem 0, "", False, True
        .AddItem 7, "Cut", False, False
        .AddItem 8, "Copy", False, False
        .AddItem 9, "Paste", False, False
        .AddItem 10, "Delete", False, False
        .AddItem 0, "", False, True
        .AddItem 0, "Select All", False, False
    End With
        
End Sub
Function MenuVisible() As Boolean
    If XPM_Connect.IsVisible Then
        MenuVisible = True
    ElseIf XPM_Tools.IsVisible Then
        MenuVisible = True
    ElseIf XPM_Edit.IsVisible Then
        MenuVisible = True
    ElseIf XPM_View.IsVisible Then
        MenuVisible = True
    ElseIf XPM_Format.IsVisible Then
        MenuVisible = True
    ElseIf XPM_Commands.IsVisible Then
        MenuVisible = True
    ElseIf XPM_Window.IsVisible Then
        MenuVisible = True
    ElseIf XPM_Help.IsVisible Then
        MenuVisible = True
    Else
        MenuVisible = False
    End If
End Function


