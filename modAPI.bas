Attribute VB_Name = "modAPI"
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const AW_HOR_POSITIVE = &O1
Public Const AW_HOR_NEGATIVE = &H2
Public Const AW_VER_POSITIVE = &H4
Public Const AW_VER_NEGATIVE = &H8
Public Const AW_CENTER = &H10
Public Const AW_HIDE = &H10000
Public Const AW_ACTIVATE = &H20000
Public Const AW_SLIDE = &H40000
Public Const AW_BLEND = &H80000
Public Const CS_DROPSHADOW = &H20000
Public Const CS_OWNDC = 32
Public Const CS_NOCLOSE = &H200
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_3DFACE = COLOR_BTNFACE
Public Const COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DHILIGHT = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_3DSHADOW = COLOR_BTNSHADOW
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_DESKTOP = COLOR_BACKGROUND
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_HOTLIGHT = 26
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_INFOBK = 24
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE
Public Const EM_UNDO As Long = &HC7
Public Const GCL_STYLE = (-26)
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = (-4)
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_TOPMOST As Long = -1
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const LVS_OWNERDRAWFIXED = &H400
Public Const ODT_LISTVIEW = 102
Public Const SIZE_MAXIMIZED = 2
Public Const SW_HIDE As Long = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW As Long = &H40
Public Const WA_ACTIVE = 1
Public Const WM_ACTIVATE = &H6
Public Const WM_DRAWITEM = &H2B
Public Const WM_INITMENU As Long = &H116
Public Const WM_LBUTTONUP As Long = &H202&
Public Const WM_MDISETMENU = &H230
Public Const WM_NOTIFY As Long = &H4E&
Public Const WM_PAINT = &HF
Public Const WM_SETFOCUS As Long = &H7
Public Const WM_SIZE = &H5
Public Const WM_SYSCOMMAND As Long = &H112
Public Const WM_UNDO As Long = &H304
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As Rect
    itemData As Long
End Type


Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Sub DrawButton(hdc As Long, colorFace As Long, colorHilight As Long, colorShadow As Long, x As Long, y As Long, w As Long, h As Long)
    Dim brushFace As Long, brushHilight As Long, brushShadow As Long, brushShadow2 As Long
    Dim oldObjBrush As Long, oldObjBrush2  As Long
    
    '* Draw face and hilight
    brushFace = CreateSolidBrush(colorFace)
    oldObjBrush = SelectObject(hdc, brushFace)
    brushHilight = CreatePen(0, 1, colorHilight)
    oldObjBrush2 = SelectObject(hdc, brushHilight)
    Rectangle hdc, x, y, (x + w), y + h
    SelectObject hdc, oldObjBrush2
    SelectObject hdc, oldObjBrush
    DeleteObject brushFace
    DeleteObject brushHilight
    
    '* Draw shadow
    brushShadow = CreateSolidBrush(colorShadow)
    oldObjBrush = SelectObject(hdc, brushShadow)
    brushShadow2 = CreatePen(0, 0, colorShadow)
    oldObjBrush2 = SelectObject(hdc, brushShadow2)
    Rectangle hdc, x, y + h - 1, (x + w), y + h
    Rectangle hdc, x + w - 1, y, (x + w), y + h
    SelectObject hdc, oldObjBrush2
    SelectObject hdc, oldObjBrush
    DeleteObject brushShadow
    DeleteObject brushShadow2
End Sub
Sub DropShadow(hwnd As Long)
    'SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub


Public Function GetCharHeight(intPoint As Integer)
    GetCharHeight = -MulDiv(intPoint, GetDeviceCaps(CLIENT.picMenu.hdc, LOGPIXELSY), 72)
    'GetCharHeight = -11
End Function


