VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* Menu properties
Private mnuName As String

Private dim_TPCwidth    As Long

'* Menu constants
Const XBuffer As Long = 3
Const YBuffer As Long = 3

Const dim_MarginWidth As Long = 23
Const fnt_MenuItem  As String = "Tahoma"

'* Width
Private mnuWidth As Long
Private theTextHeight As Long

Private frmMenu As frmXPMenu
Private ActivePopup As New clsXPMenu

Private bVisible As Boolean
Private bPopupShown As Boolean
Private Yhilight As Long

'* image list
Private imageLst As ImageList

'* Menu array
Private MenuItems()     As typMenuItem
Private MenuItemCount   As Long
Private TextItemCnt     As Long
Private SepItemCnt      As Long
Private hilightedItem   As Long

'* Types
Private Type typMenuItem
    IconNum     As Long
    Text        As String
    bPopupmenu  As Boolean
    mnuSubMenu  As clsXPMenu
    bSeperator  As Boolean
    bDisabled   As Boolean
    bChecked    As Boolean
    bVisible    As Boolean
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function AddItem(IconNum As Long, Text As String, bPopupmenu As Boolean, bSeperator As Boolean, Optional mnuSubitem As clsXPMenu = Nothing) As Integer
    
    MenuItemCount = MenuItemCount + 1
    ReDim Preserve MenuItems(1 To MenuItemCount) As typMenuItem
    
    With MenuItems(MenuItemCount)
        .IconNum = IconNum
        .Text = Text
        .bPopupmenu = bPopupmenu
        .bSeperator = bSeperator
        .bVisible = True
        If (mnuSubitem Is Nothing) Then Else Set .mnuSubMenu = mnuSubitem
    End With
        
    If bSeperator Then
        SepItemCnt = SepItemCnt + 1
    Else
        TextItemCnt = TextItemCnt + 1
    End If
    
    Dim theWidth As Integer
    With frmMenu
        theWidth = .textWidth(Text) + (XBuffer * 4) + 2 + dim_MarginWidth + 2 '2=border
        
        If bPopupmenu Then
            theWidth = theWidth + (XBuffer * 2) + frmMenu.textWidth("4")
        End If
        
        If theWidth > mnuWidth Then mnuWidth = theWidth
    End With
    
    AddItem = MenuItemCount
    
End Function



Function GetHilightNum() As Integer
    GetHilightNum = hilightedItem
End Function

Public Function GetItemText(itemNum As Integer) As String
    If itemNum > MenuItemCount Then
        GetItemText = ""
        Exit Function
    End If
    
    GetItemText = MenuItems(itemNum).Text
End Function

Public Function GetMaxWidth() As Long
    Dim i As Integer, maxWidth As Integer, theWidth As Integer
    
    frmMenu.FontName = strMenuFontName
    frmMenu.FontSize = intMenuFontSize
    For i = 1 To MenuItemCount
        With frmMenu
            theWidth = .textWidth(MenuItems(i).Text) + (XBuffer * 4) + 2 + dim_MarginWidth + 2  '2=border
            If MenuItems(i).bPopupmenu Then
                theWidth = theWidth + (XBuffer * 2) + frmMenu.textWidth("4")
            End If
            If theWidth > maxWidth Then maxWidth = theWidth
        End With
    Next i
    
    GetMaxWidth = maxWidth
    'when i am typing this i noticed that it is a bit lagged.
End Function

Public Function GetMenuName()
    GetMenuName = mnuName
End Function

Public Sub HideMenu()
    frmMenu.Visible = False
    bVisible = False
    
    hilightedItem = 0
    
    frmMenu.tmrActive.Enabled = False
    frmMenu.tmrHover.Enabled = False
End Sub

Public Function IsTextItem(itemNum As Integer) As Boolean
    If itemNum > MenuItemCount Then
        IsTextItem = False
        Exit Function
    End If
    
    If MenuItems(itemNum).bPopupmenu Or MenuItems(itemNum).bSeperator Then
        IsTextItem = False
    Else
        IsTextItem = True
    End If
End Function

Function IsVisible() As Boolean
    IsVisible = bVisible
End Function

Public Sub KillAllMenus()
    On Error Resume Next

    Dim frm As Form
    For Each frm In Forms
        If frm.Tag = "XPMenu" Then
            If frm.XPMenuClass.IsVisible Then
                frm.XPMenuClass.KillPopupMenus
                frm.XPMenuClass.UnloadMenu
            End If
        End If
    Next frm
    
    CLIENT.bMenuShown = False
    CLIENT.mnuOverWhich = -1
    CLIENT.DrawMenu
End Sub


Public Sub KillAllMenusExcept(strName As String)
    Dim frm As Form
    For Each frm In Forms
        If frm.Tag = "XPMenu" Then
            If frm.XPMenuClass.IsVisible And frm.XPMenuClass.GetMenuName <> strName Then
                frm.XPMenuClass.KillPopupMenus
                frm.XPMenuClass.UnloadMenu
            End If
        End If
    Next frm
    
    CLIENT.bMenuShown = False
    CLIENT.mnuOverWhich = -1
    
    CLIENT.DrawMenu
End Sub


Public Sub KillPopupMenus()
    Dim i As Long
    
    For i = 1 To MenuItemCount
        If MenuItems(i).bPopupmenu Then
            If MenuItems(i).mnuSubMenu.IsVisible Then
                MenuItems(i).mnuSubMenu.KillPopupMenus
                MenuItems(i).mnuSubMenu.UnloadMenu
            End If
        End If
    Next i
End Sub

Public Sub KillSpecPopup(intItem As Integer)
    On Error Resume Next
    MenuItems(intItem).mnuSubMenu.KillPopupMenus
    MenuItems(intItem).mnuSubMenu.UnloadMenu

End Sub

Public Function PopupShown() As Boolean
     PopupShown = bPopupShown
End Function

Public Function SetCheck(itemNum As Integer, bValue As Boolean)
    If itemNum > MenuItemCount Then Exit Function
    
    MenuItems(itemNum).bChecked = bValue
End Function

Public Sub SetNextHilightItem()
    If hilightedItem < 1 Then
        hilightedItem = 0
    End If
    
    Do
        If hilightedItem = MenuItemCount Then
            hilightedItem = 1
        Else
            hilightedItem = hilightedItem + 1
        End If
    Loop Until Not MenuItems(hilightedItem).bSeperator
    DrawMenu
End Sub

Public Sub SetPreviousHilightItem()
    
    If hilightedItem < 1 Then
        hilightedItem = MenuItemCount + 1
    End If
    
    Do
        If hilightedItem = 1 Then
            hilightedItem = MenuItemCount
        Else
            hilightedItem = hilightedItem - 1
        End If
    Loop Until Not MenuItems(hilightedItem).bSeperator
    DrawMenu
End Sub
Public Function SetText(itemNum As Integer, strText As String)
    If itemNum > MenuItemCount Then Exit Function
    MenuItems(itemNum).Text = strText
    mnuWidth = GetMaxWidth()
End Function
Public Function SetDisable(itemNum As Integer, bValue As Boolean)
    If itemNum > MenuItemCount Then Exit Function
    
    MenuItems(itemNum).bDisabled = bValue
End Function
Public Function SetVisible(itemNum As Integer, bValue As Boolean)
    If itemNum > MenuItemCount Then Exit Function
    
    MenuItems(itemNum).bVisible = bValue
End Function
Public Function GetCheck(itemNum As Integer) As Boolean
    If itemNum > MenuItemCount Then Exit Function
    
    GetCheck = MenuItems(itemNum).bChecked
End Function
Public Function GetVisible(itemNum As Integer) As Boolean
    If itemNum > MenuItemCount Then Exit Function
    
    GetVisible = MenuItems(itemNum).bVisible
End Function
Public Function GetDisable(itemNum As Integer)
    If itemNum > MenuItemCount Then Exit Function
    
    GetDisable = MenuItems(itemNum).bDisabled
End Function

Sub ShowMenu(x As Long, y As Long, Optional lngTPCwidth As Long = 0)
    dim_TPCwidth = lngTPCwidth
    'On Error GoTo errorHandler

    KillPopupMenus
    If bVisible = True Then Exit Sub

    On Error GoTo errorHandler
    frmMenu.Left = x '* 15
    frmMenu.Top = y '* 15
    frmMenu.bVisible = True
    
    If (GetHeight() * 15 + frmMenu.Top) > Screen.Height Then
        frmMenu.Top = frmMenu.Top - (GetHeight * 15)
        dim_TPCwidth = 0
    End If
    
    frmMenu.Tag = "XPMenu"

    mnuWidth = GetMaxWidth
    DrawMenu
    SendMessage frmMenu.hwnd, WM_SETFOCUS, 0&, vbNullString
    
    bVisible = True
    frmMenu.tmrActive.Enabled = True
    frmMenu.tmrHover.Enabled = True
    
    Exit Sub
errorHandler:
    MsgBox "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & "Source: " & Err.Source
End Sub

Public Sub DrawMenu()
    
    Dim tBrush As Long, Area As Rect, tPen As Long, hOldFont As Long
    Dim BMP As BitmapStruc, hFont As Long, hFontSym As Long
    
    '* Create the fonts to be used
    hFont = CreateFont(intMenuFontHeight, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, strMenuFontName)
    If hFont = 0 Then Exit Sub
    hFontSym = CreateFont(14, 12, 0, 0, FW_LIGHT, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, DEFAULT_PITCH, "Marlett")
    hFontSym2 = CreateFont(17, 14, 0, 0, FW_LIGHT, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, DEFAULT_PITCH, "Marlett")
    
    If hFontSym = 0 Then DeleteObject hFont: Exit Sub
    
    With frmMenu
        frmMenu.Height = GetHeight() * 15
        frmMenu.Width = mnuWidth * 15
        
        BMP.Area.Left = 0
        BMP.Area.Top = 0
        BMP.Area.Right = mnuWidth
        BMP.Area.Bottom = frmMenu.Height
        
        '* Create bitmap
        BMP.hDcMemory = CreateCompatibleDC(frmMenu.hdc)
        BMP.hDcBitmap = CreateCompatibleBitmap(frmMenu.hdc, frmMenu.ScaleWidth, frmMenu.ScaleHeight)
        BMP.hDcPointer = SelectObject(BMP.hDcMemory, BMP.hDcBitmap)
        
        If BMP.hDcMemory = 0 Or BMP.hDcBitmap = 0 Then
            DeleteObject BMP.hDcBitmap
            DeleteDC BMP.hDcMemory
            DeleteObject hFont
            DeleteObject hFontSym
            Exit Sub
        End If
        
        '* SaveDC State
        SaveDC BMP.hDcMemory
        
        '* Set The Font
        hOldFont = SelectObject(BMP.hDcMemory, hFont)
        
        '* background of text transparent
        SetBkMode BMP.hDcMemory, 0
     
        '* Margin
        tBrush = CreateSolidBrush(clrLeftMargin)
        oldObj = SelectObject(BMP.hDcMemory, tBrush)
        tPen = CreatePen(5, 0, clrMenuBorder)
        oldObj2 = SelectObject(BMP.hDcMemory, tPen)
        If dim_TPCwidth > 0 Then
            Rectangle BMP.hDcMemory, 1, 2, dim_MarginWidth + 2, frmMenu.ScaleHeight
        Else
            Rectangle BMP.hDcMemory, 1, 1, dim_MarginWidth + 2, frmMenu.ScaleHeight
        End If
        SelectObject BMP.hDcMemory, oldObj
        SelectObject BMP.hDcMemory, oldObj2
        DeleteObject tBrush
        DeleteObject tPen
        
        '* if TPC width, draw connector
        'If dim_TPCwidth > 0 Then
        '    tBrush = CreateSolidBrush(clrLeftMargin)
        '    oldobj = SelectObject (BMP.hDcMemory, tBrush)
        '    tPen = CreatePen(0, 1, clrLeftMargin)
        '    oldobj2 = SelectObject (BMP.hDcMemory, tPen)
        '    Rectangle BMP.hDcMemory, 1, 0, dim_TPCwidth + 1, 1
        '    SelectObject BMP.hDcMemory, oldObj
        '    SelectObject BMP.hDcMemory, oldObj2
        '    DeleteObject tBrush
        '    DeleteObject tPen
        'End If
        
        '* X, Y info
        Dim Xcur As Long, Ycur As Long, Index As Integer
        Ycur = 0
        
        DrawButton BMP.hDcMemory, mc_HilightOff, mc_HilightOff, mc_HilightOff, 0, 0, .Width, .Height
        
        For Index = 1 To MenuItemCount
            If MenuItems(Index).bVisible = False Then GoTo drawNext
        
            Xcur = dim_MarginWidth + (XBuffer * 2)  '* 1 for the border
        
            SetTextColor BMP.hDcMemory, mc_TextOff
        
            '* hilighted?
            If hilightedItem = Index And MenuItems(Index).bSeperator = False Then 'And MenuItems(Index).bDisabled = False Then
                Yhilight = Ycur
                
                DrawButton BMP.hDcMemory, mc_HilightOver, mc_HilightOver, mc_HilightOver, lngMenuHilightStart, Ycur, mnuWidth - 3, YBuffer * 2 + theTextHeight + 1
                'DrawButton BMP.hDcMemory, mc_HilightOff, mc_BHilightOver, mc_BShadowOver, 0, 0, 0, 0
                
                frmMenu.picIcon.BackColor = mc_HilightOff
                SetTextColor BMP.hDcMemory, mc_TextOver
            Else
                'DrawButton BMP.hDcMemory, mc_HilightOff, mc_BHilightOff, mc_BShadowOff, lngMenuHilightStart, Ycur, mnuWidth - 3, YBuffer * 2 + theTextHeight + 2
                
                frmMenu.picIcon.BackColor = clrLeftMargin
            End If
            
            '* bit icon OR draw checked value
            If imageLst Is Nothing Or MenuItems(Index).bChecked = True Then
                If MenuItems(Index).bChecked = True Then
                    '* fix code, need to draw check
                    oldObj = SelectObject(BMP.hDcMemory, hFontSym2)
                    TextOut BMP.hDcMemory, (dim_MarginWidth - 16) \ 2 + 3, Ycur + (((theTextHeight + (YBuffer * 2)) - 16) \ 2) + 2, "b", 1
                    SelectObject BMP.hDcMemory, oldObj
                End If
            Else
                If MenuItems(Index).IconNum <> 0 Then
                    .picIcon.Picture = .picIcon.Image
                    
                    If hilightedItem = Index Then
                        DrawButton BMP.hDcMemory, mc_HilightOff, mc_HilightOff, mc_HilightOff, 0, Ycur, dim_MarginWidth + 2, 20
                        DrawButton BMP.hDcMemory, mc_HilightOff, mc_BHilightOver, mc_BShadowOver, 1, Ycur, 23, 20
                    End If
                    frmMenu.picIcon.Picture = imageLst.ListImages.item(MenuItems(Index).IconNum).Picture
                    BitBlt BMP.hDcMemory, (dim_MarginWidth - 16) \ 2 + 2, Ycur + (((theTextHeight + (YBuffer * 2)) - 16) \ 2) + 1, 16, 16, frmMenu.picIcon.hdc, 0, 0, SRCCOPY
                    
                End If
            End If
                
            '* popup menu
            If MenuItems(Index).bPopupmenu Then
                '* fix, need to draw popup arrow
                oldObj = SelectObject(BMP.hDcMemory, hFontSym)
                TextOut BMP.hDcMemory, .ScaleWidth - .textHeight("4") - XBuffer, Ycur + (((theTextHeight + (YBuffer * 2)) - 16) \ 2) + 1, "4", 1
                SelectObject BMP.hDcMemory, oldObj
            End If
            
            '* draw item
            If MenuItems(Index).bSeperator Then
                DrawButton BMP.hDcMemory, mc_shadow, mc_shadow, mc_shadow, 1, Ycur + YBuffer, mnuWidth - 1, 1
                DrawButton BMP.hDcMemory, mc_hilight, mc_hilight, mc_hilight, 1, Ycur + YBuffer + 1, mnuWidth - 1, 1
                
                Ycur = Ycur + 1 + (YBuffer * 2)
            Else
                If MenuItems(Index).bDisabled Then
                    If hilightedItem <> Index Then
                        SetTextColor BMP.hDcMemory, mc_hilight
                        TextOut BMP.hDcMemory, Xcur + 1, Ycur + YBuffer + 1, MenuItems(Index).Text, Len(MenuItems(Index).Text)
                    End If
                    
                    SetTextColor BMP.hDcMemory, mc_TextDisabled
                Else
                    If hilightedItem = Index Then
                        SetTextColor BMP.hDcMemory, mc_TextOver
                    Else
                        SetTextColor BMP.hDcMemory, mc_TextOff
                    End If
                End If
                
                TextOut BMP.hDcMemory, Xcur, Ycur + YBuffer, MenuItems(Index).Text, Len(MenuItems(Index).Text)
                Ycur = Ycur + theTextHeight + (YBuffer * 2)
            End If
drawNext:
        Next Index
        
    End With
    
    BitBlt frmMenu.hdc, BMP.Area.Left, BMP.Area.Top, BMP.Area.Right, BMP.Area.Bottom, BMP.hDcMemory, 0, 0, SRCCOPY
    'InvalidateRect frmMenu.hDc, BMP.Area, False
    
    SelectObject BMP.hDcMemory, hOldFont
    
    '* RestoreDC State
    RestoreDC BMP.hDcMemory, -1
    
    DeleteObject hOldFont
    DeleteObject tPen
    DeleteObject tBrush
    DeleteObject hFont
    DeleteObject hFontSym
    DeleteObject hFontSym2
    DeleteObject BMP.hDcBitmap
    DeleteDC BMP.hDcMemory
    
    If frmMenu.Visible = False Then
        Load frmMenu
        frmMenu.Show
    End If
   
End Sub

Function GetHeight() As Long
    Dim lngHeight As Long
    
    With frmMenu
        Dim Ycur As Long, Index As Integer
        Ycur = 3
        
        For Index = 1 To MenuItemCount
            '* draw item
            If MenuItems(Index).bVisible Then
                If MenuItems(Index).bSeperator Then
                    Ycur = Ycur + 1 + (YBuffer * 2)
                Else
                    Ycur = Ycur + theTextHeight + (YBuffer * 2)
                End If
            End If
        Next Index
    End With
    
    lngHeight = Ycur + 4
    GetHeight = lngHeight
End Function

Public Function GetHilightedItem(y As Single) As Integer
    On Error GoTo endd
    
    With frmMenu
        '* X, Y info
        Dim Ycur As Long, Index As Integer
        Ycur = 3
        
        For Index = 1 To MenuItemCount
            If MenuItems(Index).bVisible = False Then
                '* nothing :D!
            ElseIf MenuItems(Index).bSeperator Then
                If y >= Ycur And (y <= Ycur + (YBuffer * 2) + 1) Then
                    GetHilightedItem = Index
                    Exit Function
                End If
                Ycur = Ycur + 1 + (YBuffer * 2)
            Else
                'TextOut .hdc, Xcur, Ycur + YBuffer, MenuItems(index).Text, Len(MenuItems(index).Text)
                If y >= Ycur And (y <= Ycur + theTextHeight + (YBuffer * 2)) Then
                    GetHilightedItem = Index
                    Exit Function
                End If
                Ycur = Ycur + theTextHeight + (YBuffer * 2)
            End If
            
        Next Index
        
    End With
    Exit Function
endd:
End Function

Sub Init(strMenuName As String, Optional imageListBind As ImageList)
    
    mnuName = strMenuName
    
    Set frmMenu = New frmXPMenu
    Set frmMenu.XPMenuClass = Me
    
    'SetClassLong frmMenu.hwnd, GCL_STYLE, GetClassLong(frmMenu.hwnd, GCL_STYLE) Or CS_DROPSHADOW
    'SetClassLong frmMenu.hwnd, GCL_STYLE, GetClassLong(frmMenu.hwnd, GCL_STYLE) And Not CS_OWNDC
    'SetWindowLong frmMenu.hwnd, GWL_STYLE, WS_POPUP Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_OVERLAPPED
    
    If imageListBind Is Nothing Then Else Set imageLst = imageListBind
    
    theTextHeight = frmMenu.textHeight("gW")

    MenuItemCount = 0
    SepItemCnt = 0
    TextItemCnt = 0
    hilightedItem = 0
    'ReDim MenuItems(MenuItemCount) As typMenuItem
End Sub


Public Sub MoveMenu(Lft As Long, Tp As Long)
    frmMenu.Left = Lft
    frmMenu.Top = Tp
End Sub

Public Sub setHilightedItem(item As Integer)
    If item = 0 Or hilightedItem = item Then Exit Sub
    
    Dim startTime As Long
    
    If frmMenu.tmrKillPopup.Enabled = False Then
        frmMenu.tmrKillPopup.Tag = hilightedItem
        frmMenu.tmrKillPopup.Enabled = True
    End If
    
    If item = -1 Then
        hilightedItem = -1
        DrawMenu
        Exit Sub
    End If
    
    hilightedItem = item
    bPopupShown = False
    
    If MenuItems(item).bPopupmenu Then
        bPopupShown = True
        Set ActivePopup = MenuItems(item).mnuSubMenu
        If ActivePopup.IsVisible Then Exit Sub
        
        'frmXPMenu.setFocus
        SendMessage frmXPMenu.hwnd, &H7, 0, vbNullString
        DrawMenu
        
        frmMenu.tmrKillPopup.Enabled = True
        
        startTime = Timer * 1000

        Do While (Timer * 1000) - startTime < 300
            DoEvents
        Loop
        
        If hilightedItem = item Then
            ActivePopup.ShowMenu frmMenu.Left + frmMenu.Width - 75, frmMenu.Top + (Yhilight * 15)
        End If
        Exit Sub
    Else
        DrawMenu
    End If
    'KillPopupMenus
    frmMenu.tmrHover.Enabled = True
   
End Sub


Public Sub UnloadMenu()
    
    'AnimateWindow frmMenu.hwnd, 100, AW_HIDE
    On Error Resume Next
    Unload frmMenu
    frmMenu.bVisible = False
    bVisible = False
    hilightedItem = 0
    
    frmMenu.tmrActive.Enabled = False
    frmMenu.tmrHover.Enabled = False
        
End Sub


