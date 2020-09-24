Attribute VB_Name = "modSettings"
Option Explicit
Public Const MAX_TEXT_HISTORY = 30

Public lngBuild                 As Long

Public strGlobalINI             As String
Public strINI                   As String
Public strProfile               As String
Public winINI                   As String

Public strFontName              As String
Public intFontSize              As Integer
Public intIndent                As Integer

Public bKeepChildrenInBounds    As Boolean
Public bTimestamp               As Boolean
Public strTimeFormat            As String

Public bStretchButtons          As Boolean
Global Const ICON_SIZE As Integer = 16
Public intButtonWidth           As Integer

'* Scripting Options
Public strDefScriptFolder       As String
Public strScripts()             As String

'* Server Options
Public ServerAddr           As String
Public ServerPort           As Long

'* User info
Public strNicks()           As String
Public strFName             As String
Public strEmail             As String

'* Channel Options
Public bCloseOnKick        As Boolean
Public bRejoinOnKick       As Boolean

'* Taskbar
Public bNoDown      As Boolean

'* Display settings for taskbar
Public tb_ButtonOff     As Long
Public tb_ShadowOff     As Long
Public tb_HilightOff    As Long
Public tb_TextOff       As Long
Public tb_ButtonOver    As Long
Public tb_ShadowOver    As Long
Public tb_HilightOver   As Long
Public tb_TextOver      As Long
Public tb_ButtonDown    As Long
Public tb_ShadowDown    As Long
Public tb_HilightDown   As Long
Public tb_TextDown      As Long
Public tb_ButtonActive  As Long
Public tb_ShadowActive  As Long
Public tb_HilightActive As Long
Public tb_TextActive    As Long
Public b_DualSwitch     As Boolean

Public clrBackground    As Long
Public clrMenuBackground    As Long
Public clrMenuBorder    As Long
Public clrLeftMargin    As Long
Public clrLines         As Long
Public clrSeperator     As Long
Public clrTextDisabled  As Long
Public lngMenuHilightStart  As Long

'Display settings for menu
Public mc_TextOff       As Long
Public mc_TextOver      As Long
Public mc_TextDown      As Long
Public mc_BShadowOff    As Long
Public mc_BHilightOff   As Long
Public mc_HilightOff    As Long
Public mc_BShadowOver   As Long
Public mc_BHilightOver  As Long
Public mc_HilightOver   As Long
Public mc_BShadowDown   As Long
Public mc_BHilightDown  As Long
Public mc_HilightDown   As Long
Public mc_TextDisabled  As Long
Public mc_hilight       As Long
Public mc_shadow        As Long
Public mv_bSink         As Boolean
Public strMenuFontName  As String
Public intMenuFontSize  As Integer
Public intMenuFontHeight    As Integer

Function GetINI(strINIx As String, strSection As String, strSetting As String, strDefault As String)
    Dim lngReturn As Long, strReturn As String, lngSize As Long
    lngSize = 255
    strReturn = String(lngSize, 0)
    lngReturn = GetPrivateProfileString(strSection, strSetting, strDefault, strReturn, lngSize, strINIx)
    
    GetINI = LeftOf(strReturn, Chr(0))
    
End Function
Sub LoadUserOptions()
    
    '* Scripts
    strDefScriptFolder = ReadINI("scripting", "def_sf", PATH)
    strScripts = Split(ReadINI("scripting", "scripts", "scripts\default.sex"), ",")
    If strScripts(UBound(strScripts)) = "" Then ReDim Preserve strScripts(UBound(strScripts) - 1) As String
    
    '* Servers
    ServerAddr = ReadINI("server", "address", "irc.liek.net")
    ServerPort = CLng(ReadINI("server", "port", "6667"))
    
    '* User Info
    strNicks = Split(ReadINI("user", "nicks", "sIRCuser,sIRC_user,sIRC_user-"), ",")
    strFName = ReadINI("user", "name", "sIRC User")
    strEmail = ReadINI("user", "email", "I am using sIRC for Win32 (http://vcv.ath.cx/sirc)")
    
    '* Display Settings
    tb_ButtonOff = CLng(ReadINI("display", "tb_off_face", CStr(GetSysColor(COLOR_3DFACE))))
    tb_ShadowOff = CLng(ReadINI("display", "tb_off_shadow", CStr(GetSysColor(COLOR_3DFACE))))
    tb_HilightOff = CLng(ReadINI("display", "tb_off_hilight", CStr(GetSysColor(COLOR_3DFACE))))
    tb_TextOff = CLng(ReadINI("display", "tb_off_text", CStr(vbBlack)))
    tb_ButtonOver = CLng(ReadINI("display", "tb_over_face", CStr(GetSysColor(COLOR_3DFACE))))
    tb_ShadowOver = CLng(ReadINI("display", "tb_over_shadow", CStr(GetSysColor(COLOR_BTNSHADOW))))
    tb_HilightOver = CLng(ReadINI("display", "tb_over_hilight", CStr(GetSysColor(COLOR_3DHILIGHT))))
    tb_TextOver = CLng(ReadINI("display", "tb_over_text", CStr(vbBlack)))
    tb_ButtonDown = CLng(ReadINI("display", "tb_down_face", CStr(GetSysColor(COLOR_3DFACE))))
    tb_ShadowDown = CLng(ReadINI("display", "tb_down_shadow", CStr(GetSysColor(COLOR_3DHILIGHT))))
    tb_HilightDown = CLng(ReadINI("display", "tb_down_hilight", CStr(GetSysColor(COLOR_BTNSHADOW))))
    tb_TextDown = CLng(ReadINI("display", "tb_down_text", CStr(vbBlack)))
    tb_ButtonActive = CLng(ReadINI("display", "tb_active_face", CStr(RGB(235, 235, 235))))
    tb_ShadowActive = CLng(ReadINI("display", "tb_active_shadow", CStr(GetSysColor(COLOR_3DHILIGHT))))
    tb_HilightActive = CLng(ReadINI("display", "tb_active_hilight", CStr(GetSysColor(COLOR_BTNSHADOW))))
    tb_TextActive = CLng(ReadINI("display", "tb_active_text", CStr(vbBlack)))
    b_DualSwitch = CBool(ReadINI("display", "dualswitch", CStr(True)))
    
    clrLines = CLng(ReadINI("display", "lines", CStr(&HA0A0A0)))
    clrBackground = CLng(ReadINI("display", "background", CStr(GetSysColor(COLOR_3DFACE))))
    clrMenuBackground = CLng(ReadINI("display", "menubackground", CStr(GetSysColor(COLOR_3DFACE))))
    clrMenuBorder = CLng(ReadINI("display", "menuborder", CStr(GetSysColor(COLOR_3DFACE)))) 'CStr(&H868686)))
    clrLeftMargin = CLng(ReadINI("display", "menuleftmargin", CStr(GetSysColor(COLOR_3DFACE))))
    clrSeperator = CLng(ReadINI("display", "seperator", CStr(GetSysColor(12))))
    clrTextDisabled = CLng(ReadINI("display", "seperator", CStr(GetSysColor(COLOR_GRAYTEXT))))
    lngMenuHilightStart = CLng(ReadINI("display", "menuhilistart", CStr(0)))
    
    mc_TextOff = CLng(ReadINI("display", "m_textoff", CStr(vbBlack)))
    mc_TextOver = CLng(ReadINI("display", "m_textover", CStr(GetSysColor(COLOR_HIGHLIGHTTEXT))))
    mc_TextDown = CLng(ReadINI("display", "m_textdown", CStr(vbBlack)))
    mc_TextDisabled = CLng(ReadINI("display", "m_textdisabled", CStr(GetSysColor(COLOR_GRAYTEXT))))
    mc_BShadowOff = CLng(ReadINI("display", "m_bshadowoff", CStr(GetSysColor(COLOR_3DFACE))))
    mc_BHilightOff = CLng(ReadINI("display", "m_bhilightoff", CStr(GetSysColor(COLOR_3DFACE))))
    mc_HilightOff = CLng(ReadINI("display", "m_hilightoff", CStr(GetSysColor(COLOR_3DFACE))))
    mc_BShadowOver = CLng(ReadINI("display", "m_bshadowover", CStr(GetSysColor(COLOR_BTNSHADOW))))
    mc_BHilightOver = CLng(ReadINI("display", "m_bhilightover", CStr(GetSysColor(COLOR_3DHILIGHT))))
    mc_HilightOver = CLng(ReadINI("display", "m_hilightover", CStr(GetSysColor(COLOR_HIGHLIGHT))))
    mc_BShadowDown = CLng(ReadINI("display", "m_bshadowdown", CStr(GetSysColor(COLOR_3DHILIGHT)))) 'CStr(GetSysColor(COLOR_3DHILIGHT))))
    mc_BHilightDown = CLng(ReadINI("display", "m_bhilightdown", CStr(GetSysColor(COLOR_BTNSHADOW))))
    mc_HilightDown = CLng(ReadINI("display", "m_hilightdown", CStr(GetSysColor(COLOR_3DFACE))))
    mc_shadow = CLng(ReadINI("display", "m_shadow", CStr(GetSysColor(COLOR_BTNSHADOW))))
    mc_hilight = CLng(ReadINI("display", "m_hilight", CStr(GetSysColor(COLOR_3DHILIGHT))))
    mv_bSink = CBool(ReadINI("display", "m_bsink", CStr(True)))
    strMenuFontName = ReadINI("display", "m_fontname", "Tahoma")
    intMenuFontSize = CInt(ReadINI("display", "m_fontsize", 8))
    intMenuFontHeight = GetCharHeight(intMenuFontSize)
End Sub

Sub PutINI(strINIx As String, strSection As String, strLValue As String, strRValue As String)
    Dim lngReturn As Long
    lngReturn = WritePrivateProfileString(strSection, strLValue, strRValue, strINIx)
End Sub

Sub FlushINI(strFile As String)
    Dim lngReturn As Long
    lngReturn = WritePrivateProfileString("", "", "", strFile)
End Sub

Function ReadINI(strSection As String, strKey As String, strDefault As String)
    Dim lngReturn As Long, strReturn As String, lngSize As Long
    lngSize = 255
    strReturn = String(lngSize, 0)
    lngReturn = GetPrivateProfileString(strSection, strKey, strDefault, strReturn, lngSize, strINI)

    If strReturn = "" Then
        WriteINI strSection, strKey, strDefault
        ReadINI = strReturn
    Else
        ReadINI = LeftOf(strReturn, Chr(0))
    End If
End Function


Sub WriteINI(strSection As String, strKey As String, strValue As String)
    PutINI strINI, strSection, strKey, strValue
End Sub


