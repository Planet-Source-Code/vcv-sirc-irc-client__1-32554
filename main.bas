Attribute VB_Name = "modmain"
'* DEBUGGING
Public Const bDebug As Boolean = False

'* ANSI Formatting character values
Global Const Cancel = 15
Global Const BOLD = 2
Global Const UNDERLINE = 31
Global Const Color = 3
Global Const REVERSE = 22
Global Const ACTION = 1

Global Const COMMANDCHAR = "/"

Global strVersion As String

'* ANSI Formatting characters
Global strBold As String
Global strUnderline As String
Global strColor As String
Global strReverse As String
Global strAction As String
Global PATH As String

Public Type RectAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Public CurrentServerID  As Integer
Public scriptEngine     As clsSSE_Main
Public strEmptyParams(0)    As String

'* Variables for incoming commands
Type ParsedData
    bHasPrefix   As Boolean
    strParams()  As String
    intParams    As Integer
    strFullHost  As String
    strCommand   As String
    strNick      As String
    strIdent     As String
    strHost      As String
    AllParams    As String
    bFromServer  As Boolean
End Type

Sub ParseMode(strSetBy As String, strChannel As String, strData As String, serverID As Integer)
    Dim strModes() As String, strChar As String, theChannel As Form
    Dim i As Integer, intParam As Integer
    Dim bAdd As Boolean
    Set theChannel = GetFormByName(strChannel, serverID)
    
    On Error Resume Next
    
    bAdd = True
    strModes = Split(strData, " ")
    For i = 1 To Len(strModes(0))
        strChar = Mid$(strModes(0), i, 1)
        If Left$(strChannel, 1) = "#" Or Left$(strChannel, 1) = "&" Then
            Select Case strChar
                Case "+"
                    bAdd = True
                Case "-"
                    bAdd = False
                Case "v", "b", "o", "h", "k", "l", "q", "a"
                    intParam = intParam + 1
                    theChannel.DoMode bAdd, strChar, strModes(intParam)
                Case Else
                    'DoMode strChannel, bAdd, strChar, ""
            End Select
        Else    'server
            Select Case strChar
                Case "+"
                    bAdd = True
                Case "-"
                    bAdd = False
                Case "" 'ignore this
                    intParam = intParam + 1
                    'DoMode strChannel, bAdd, strChar, strModes(intParam)
                Case Else
                    'DoMode strChannel, bAdd, strChar, ""
            End Select

        End If
    Next i
        
    If Left$(strChannel, 1) = "#" Then
        theChannel.UpdateNickList
        Dim argsX(3) As String, vars(5) As String
        argsX(1) = strChannel
        argsX(2) = strSetBy
        argsX(3) = strData
        vars(0) = "chan:" & strChannel
        vars(1) = "channel:" & strChannel
        vars(5) = "target:" & strChannel
        vars(2) = "data:" & strData
        vars(3) = "modes:" & strData
        vars(4) = "nick:" & strSetBy
        
        scriptEngine.ExecuteEvent "mode", argsX, serverID, vars
    End If
End Sub


Function DoCommandLine(paramlist() As String, strWindow As String, serverIDx As Integer)
    On Error Resume Next
    
    Dim args_pass() As String, strCom As String
    ReDim args_pass(0) As String
    strCom = paramlist(0)
    
    Select Case LCase(paramlist(0))
        Case "msg"
            ReDim Preserve args_pass(2) As String
            args_pass(1) = paramlist(1)
            args_pass(2) = JoinArray(paramlist, " ", 3)
        Case "me"
            ReDim Preserve args_pass(2) As String
            args_pass(1) = strWindow
            args_pass(2) = JoinArray(paramlist, " ", 2)
        Case Else
            ReDim Preserve args_pass(UBound(paramlist)) As String
            Dim i As Integer
            For i = LBound(paramlist) To UBound(paramlist)
                args_pass(i) = paramlist(i)
            Next i
            strCom = paramlist(0)
    End Select
    
    Dim strRet As String, vars(1) As String
    vars(1) = "source:" & strWindow
    strRet = scriptEngine.ExecuteAlias(strCom, args_pass, serverIDx, vars)
    
    If scriptEngine.bLastExec Then
        DoCommandLine = True
    Else
        DoCommandLine = False
    End If
End Function


Sub ParseData(ByVal strData As String, ByRef parsed As ParsedData)

    '* Declare variables
    Dim strTmp As String, i As Integer
    
    '* Reset variables
    bHasPrefix = False
    parsed.strNick = ""
    parsed.strIdent = ""
    parsed.strHost = ""
    parsed.strCommand = ""
    parsed.intParams = 1
    ReDim parsed.strParams(1 To 1) As String
    
    parsed.bFromServer = True
    
    '* Check for prefix, if so, parse nick, ident and host (or just host)
    If Left$(strData, 1) = ":" Then
        bHasPrefix = True
        strData = Right$(strData, Len(strData) - 1)
        '* Put data left of " " in strHost, data right of " "
        '* into strData
        Seperate strData, " ", parsed.strHost, strData
        parsed.strFullHost = parsed.strHost
        
        '* Check to see if client host name
        If InStr(parsed.strHost, "!") Then
            Seperate parsed.strHost, "!", parsed.strNick, parsed.strHost
            Seperate parsed.strHost, "@", parsed.strIdent, parsed.strHost
            parsed.bFromServer = False
        End If
    End If
        
    
    '* If any params, parse
    If InStr(strData, " ") Then
        Seperate strData, " ", parsed.strCommand, strData
        
        parsed.AllParams = strData
       '* Let's parse all the parameters.. yummy
begin: '* OH NO I USED A LABEL!

        '* If begginning of param is :, indicates that its the last param
        If Left$(strData, 1) = ":" Then
            parsed.strParams(parsed.intParams) = Right$(strData, Len(strData) - 1)
            GoTo finish
        End If
        '* If there is a space still, there is more params
        If InStr(strData, " ") Then
            Seperate strData, " ", parsed.strParams(parsed.intParams), strData
            parsed.intParams = parsed.intParams + 1
            ReDim Preserve parsed.strParams(1 To parsed.intParams) As String
            GoTo begin
        Else
            parsed.strParams(parsed.intParams) = strData
        End If
    Else
        '* No params, strictly command
        parsed.intParams = 0
        parsed.strCommand = strData
    End If
finish:
End Sub
Public Sub LoadServers(tv As TreeView)
    Dim strBuffer As String, lngSize As Long, strServ As String, strPort As String
    Dim nod() As Node, intGroup As Integer, i As Integer
    
    tv.Nodes.Clear
    
    On Error GoTo errhandler
    Open PATH & "servers.inf" For Input As #1
        Do
            Input #1, strBuffer
            
            ReDim Preserve nod(1 To tv.Nodes.Count + 1)
            If strBuffer <> "" Then
                If Left$(strBuffer, 1) = "g" Then
                    Set nod(tv.Nodes.Count) = tv.Nodes.Add(, , , Right$(strBuffer, Len(strBuffer) - 1))
                    'nod(tv.Nodes.Count).Expanded = False
                    intGroup% = tv.Nodes.Count
                Else
                    If tv.Nodes.Count > 0 Then
                        strBuffer = Right$(strBuffer, Len(strBuffer) - 1)
                        strServ = LeftOf(strBuffer, ":")
                        strPort = RightOf(strBuffer, ":")
                        If InStr(strBuffer, ":") = 0 Then
                            strPort = ""
                        Else
                            strPort = " (" & strPort & ")"
                        End If
                        Set nod(tv.Nodes.Count) = tv.Nodes.Add(nod(intGroup%), tvwChild, , strServ & strPort)
                        'nod(tv.Nodes.Count).Expanded = False
                        'nod(tv.Nodes.Count).EnsureVisible
                    End If
                End If
            End If
        Loop Until EOF(1)
    Close #1
  
    Exit Sub
  
errhandler:
    MsgBox "Loading servers: " & Err & ":" & Error
End Sub

Sub HandleKeypress(rtf As RichTextBox, chrKP As String)
    If rtf.Name = "rt_Output" Then Exit Sub

    Dim selStart As Long, selLength As Long, seltext As String
    selStart = rtf.selStart
    selLength = rtf.selLength
    seltext = rtf.seltext
    Select Case chrKP
    Case strUnderline, strReverse, strBold, Chr(15)
        If selLength = 0 Then
            rtf.seltext = chrKP
        Else
            If Left$(seltext, 1) = chrKP And Right$(seltext, 1) = chrKP Then
                rtf.seltext = Mid$(seltext, 2, Len(seltext) - 2)
                rtf.selStart = selStart
                rtf.selLength = selLength - 2
            ElseIf Left$(seltext, 1) <> chrKP And Right$(seltext, 1) <> chrKP Then
                rtf.seltext = chrKP & seltext & chrKP
                rtf.selStart = selStart
                rtf.selLength = selLength + 2
            ElseIf Left$(seltext, 1) = chrKP And Right$(seltext, 1) <> chrKP Then
                rtf.seltext = seltext & chrKP
                rtf.selStart = selStart
                rtf.selLength = selLength + 1
            ElseIf Left$(seltext, 1) <> chrKP And Right$(seltext, 1) = chrKP Then
                rtf.seltext = chrKP & seltext
                rtf.selStart = selStart
                rtf.selLength = selLength + 1
            End If
        End If
    Case strColor
        If selLength = 0 Then
            rtf.seltext = chrKP
        Else
            If Left$(seltext, 1) = chrKP And Right$(seltext, 1) = chrKP Then
                rtf.seltext = Mid$(seltext, 2, Len(seltext) - 2)
                rtf.selStart = selStart
                rtf.selLength = selLength - 2
            ElseIf Left$(seltext, 1) <> chrKP And Right$(seltext, 1) <> chrKP Then
                rtf.seltext = chrKP & seltext & chrKP
                rtf.selStart = selStart + 1
                rtf.selLength = 0
            ElseIf Left$(seltext, 1) = chrKP And Right$(seltext, 1) <> chrKP Then
                rtf.seltext = seltext & chrKP
                rtf.selStart = selStart + 1
                rtf.selLength = 0
            ElseIf Left$(seltext, 1) <> chrKP And Right$(seltext, 1) = chrKP Then
                rtf.seltext = chrKP & seltext
                rtf.selStart = selStart + 1
                rtf.selLength = 0
            End If
        End If
    Case Chr(8)
    
    Case Else
        rtf.seltext = chrKP
    End Select
End Sub


Sub Main()

    dblStart = Timer

    '* XP theme compatbility
    InitCommonControls
    
    '* Initialization for paths
    If Right$(App.PATH, 1) <> "\" Then slash$ = "\"
    PATH = App.PATH & slash$
    strGlobalINI = PATH & "sIRC.ini"
    
    '* Create the menus
    CreateMenus
    
    If FileExists(PATH & "version.data") Then
        lngBuild = CLng(GetINI(PATH & "version.data", "data", "build", "0"))
        lngBuild = lngBuild + 1
        PutINI PATH & "version.data", "data", "build", CStr(lngBuild)
    End If
    
    '* uncomment this and change before compiling
    'lngBuild = 3416
        
    
    '* REMOVE OR DIE
    'frmChanCentral.Show
    'Exit Sub
    
    
    
    '* Load Color Table
    LoadColors
    
    '* Initialize scripting engine
    Set scriptEngine = New clsSSE_Main
        
    Dim strAL As String, strName As String
    strAL = GetSetting("sIRC", "options", "loadauto")
    strName = GetSetting("sIRC", "options", "loadautoname")
    
    If strAL = "true" Then
        strINI = PATH & strName & "-settings.ini"
        strProfile = strName
        winINI = PATH & strName & "-windows.ini"
        CLIENT.Show
    Else
        Load frmLoadProfile
        frmLoadProfile.Show
    End If
    
End Sub

Sub TimeOut(dblDuration)

    Dim startTime As Double
    startTime = Timer

    Do While Timer - startTime < dblDuration
        DoEvents
    Loop

End Sub
Sub PutText_Reset(rtf As RichTextBox, strData As String)
    
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String
    Dim clr As Integer, bclr As Integer, dftclr As Integer, strRTFBuff As String
    Dim bbbold As Boolean, bbunderline As Boolean, bbreverse As Boolean, strTmp As String
    Dim lngFC As String, lngBC As String, lngStart As Long, lngLength As Long, strPlaceHolder As String
    
    lngStart = rtf.selStart
    lngLength = rtf.selLength
    
    rtf.Text = ""
    
    '* if not inialized, set font, intialiaze
    Dim btCharSet As Long
    Dim strRTF As String
    If rtf.Tag <> "init'd" Then
        rtf.Tag = "init'd"
        strFontName = rtf.Font.Name
        rtf.parent.FontName = strFontName
        btCharSet = GetTextCharset(rtf.parent.hdc)
        strRTF = ""
        strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
        strRTF = strRTF & ColorTable & vbCrLf
        strRTF = strRTF & "\viewkind4\uc1\pard\cf0\fi-" & intIndent & "\li" & intIndent & "\f0\fs" & CInt(intFontSize * 2) & vbCrLf
        strPlaceHolder = "\n"
        For i = 0 To DefinedColors - 1
            strRTF = strRTF & "\cf" & i & " " & strPlaceHolder
        Next
        strRTF = strRTF & "}"
        rtf.TextRTF = strRTF
        
        '* New session for window... call
        '# LogData rtf.Parent.Caption, "blah", strData, True
    Else
        '# LogData rtf.Parent.Caption, "blah", strData, False
        
    End If
    
    rtf.parent.FontName = strFontName
    btCharSet = GetTextCharset(rtf.parent.hdc)
    strRTF = ""
    strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
    strRTF = strRTF & ColorTable & vbCrLf
    strRTF = strRTF & "\viewkind4\uc1\pard\cf0\fi-" & intIndent & "\li" & intIndent & "\f0\fs" & CInt(intFontSize * 2) & vbCrLf
        
    strRTFBuff = "\b0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1 & "\i0\ulnone"
    dftclr = RAnsiColor(lngForeColor)
    
    Length = Len(strData)
    i = 1
    
    Do
        strChar = Mid$(strData, i, 1)
        Select Case strChar
            Case Chr(Cancel)    'cancel code
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                lngFC = CStr(RAnsiColor(lngForeColor))
                lngBC = CStr(RAnsiColor(lngBackColor))
                strRTFBuff = strRTFBuff & strBuffer & "\b0\ul0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                strBuffer = ""
                i = i + 1
            Case strBold
                bbbold = Not bbbold
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\b"
                If bbbold = False Then strRTFBuff = strRTFBuff & "0"
                strBuffer = ""
                i = i + 1
            Case strUnderline
                bbunderline = Not bbunderline
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\ul"
                If bbunderline = False Then strRTFBuff = strRTFBuff & "none"
                strBuffer = ""
                i = i + 1
            Case strReverse
                bbreverse = Not bbreverse
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " " ' & strBuffer & "\"
                If bbreverse = False Then
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                Else
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngBackColor) + 1 & "\highlight" & RAnsiColor(lngForeColor) + 1
                End If
                
                strBuffer = ""
                i = i + 1
            Case strColor
                
                strTmp = ""
                i = i + 1

                Do Until Not ValidColorCode(strTmp) Or i > Length
                    strTmp = strTmp & Mid$(strData, i, 1)
                    i = i + 1
                Loop
                
                strTmp = LeftR(strTmp, 1)
                If strTmp = "" Then
                    lngFC = CStr(RAnsiColor(lngForeColor))
                    lngBC = CStr(RAnsiColor(lngBackColor))
                Else
                    lngFC = LeftOf(strTmp, ",")
                    lngFC = CStr(CInt(lngFC))
                    If InStr(strTmp, ",") Then
                        lngBC = RightOf(strTmp, ",")
                        If lngBC <> "" Then lngBC = CStr(CInt(lngBC)) Else lngBC = CStr(RAnsiColor(lngBackColor))
                    Else
                        lngBC = ""
                    End If
                End If
                
                If lngFC = "" Then lngFC = CStr(lngForeColor)
                lngFC = Int(lngFC) + 1
                If lngBC <> "" Then lngBC = Int(lngBC) + 1
                
                
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer
                strRTFBuff = strRTFBuff & "\cf" & lngFC
                If lngBC <> "" Then strRTFBuff = strRTFBuff & "\highlight" & lngBC
                
                i = i - 1
                strBuffer = ""
                If i >= Length Then GoTo TheEnd
                
            Case Else
                Select Case strChar
                Case "}", "{", "\"
                    strBuffer = strBuffer & "\" & strChar
                Case Else
                    strBuffer = strBuffer & strChar
                End Select
                i = i + 1
        End Select
        
    Loop Until i > Length
    
   
TheEnd:
    If strBuffer <> "" Then
        strRTFBuff = strRTFBuff & " " & strBuffer
    End If

    Dim strR As String
    
    strRTFBuff = strRTFBuff '& vbCrLf
    rtf.selStart = Len(rtf.Text)
    rtf.selLength = 0
    rtf.SelRTF = strRTF & strRTFBuff & " }"
    rtf.selStart = 0
End Sub

Sub ReLoadScripts()
    scriptEngine.KillAllScripts
    
    '* Load scripts
    Dim i As Integer
    For i = LBound(strScripts) To UBound(strScripts)
        If FileExists(strScripts(i)) Then
            scriptEngine.LoadScript scriptEngine.NewScript(), strScripts(i)
        ElseIf FileExists(strDefScriptFolder & strScripts(i)) Then
            scriptEngine.LoadScript scriptEngine.NewScript(), strDefScriptFolder & strScripts(i)
        End If
    Next i
End Sub

Function ValidColorCode(strCode As String) As Boolean

    If strCode = "" Then ValidColorCode = True: Exit Function
    Dim c1 As Integer, c2 As Integer
    If strCode Like "" Or _
       strCode Like "#" Or _
       strCode Like "##" Or _
       strCode Like "#,#" Or _
       strCode Like "##,#" Or _
       strCode Like "#,##" Or _
       strCode Like "#," Or _
       strCode Like "##," Or _
       strCode Like "##,##" Or _
       strCode Like ",#" Or _
       strCode Like ",##" Then
        Dim strCol() As String
        strCol = Split(strCode, ",")
        '
        If UBound(strCol) = -1 Then
            ValidColorCode = True
        ElseIf UBound(strCol) = 0 Then
            If strCol(0) = "" Then strCol(0) = 0
            If Int(strCol(0)) >= 0 And Int(strCol(0)) <= 99 Then
                ValidColorCode = True
                Exit Function
            Else
                ValidColorCode = False
                Exit Function
            End If
        Else
            If strCol(0) = "" Then strCol(0) = lngForeColor
            If strCol(1) = "" Then strCol(1) = 0
            c1 = Int(strCol(0))
            c2 = Int(strCol(1))
            If Int(c2) < 0 Or Int(c2) > 99 Then
                ValidColorCode = False
                Exit Function
            Else
                ValidColorCode = True
                Exit Function
            End If
        End If
        ValidColorCode = True
        Exit Function
    Else
        ValidColorCode = False
        Exit Function
    End If
End Function



Sub PutText(rtf As RichTextBox, strData As String)
    
    If strData = "" Then Exit Sub
    
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String
    Dim clr As Integer, bclr As Integer, dftclr As Integer, strRTFBuff As String
    Dim bbbold As Boolean, bbunderline As Boolean, bbreverse As Boolean, strTmp As String
    Dim lngFC As String, lngBC As String, lngStart As Long, lngLength As Long, strPlaceHolder As String
    
    lngStart = rtf.selStart
    lngLength = rtf.selLength
    
    
    '* if not inialized, set font, intialiaze
    Dim btCharSet As Long
    Dim strRTF As String
    
    
    If rtf.Tag <> "init'd" Then
        rtf.Tag = "init'd"
        strFontName = rtf.Font.Name
        rtf.parent.FontName = strFontName
        btCharSet = GetTextCharset(rtf.parent.hdc)
        strRTF = ""
        strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
        strRTF = strRTF & ColorTable & vbCrLf
        strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
        strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & CInt(intFontSize * 2) & vbCrLf
        strPlaceHolder = " \n"
        For i = 0 To DefinedColors - 1
            strRTF = strRTF & "\cf" & i & strPlaceHolder
        Next
        strRTF = strRTF & "}"
        rtf.TextRTF = strRTF
        
        '* New session for window... call
        '# LogData rtf.Parent.Caption, "blah", strData, True
    Else
        '# LogData rtf.Parent.Caption, "blah", strData, False
        'rtf.selStart = Len(rtf.Text)
        'rtf.selLength = 0
        'rtf.seltext = vbCrLf
    End If
    
    rtf.parent.FontName = strFontName
    btCharSet = GetTextCharset(rtf.parent.hdc)
    strRTF = ""
    strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
    strRTF = strRTF & ColorTable & vbCrLf
    strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
    strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & CInt(intFontSize * 2) & "\fi-" & intIndent & "\li" & intIndent & "\ffprot1 "
    
    strRTFBuff = "\b0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1 & "\i0\ulnone "
    
    dftclr = RAnsiColor(lngForeColor)
    
    If bTimestamp Then
        strData = "15[" & Format(Time, strTimeFormat) & "] " & strData
    End If
        
    Length = Len(strData)
    i = 1
    
    Do
        strChar = Mid$(strData, i, 1)
        Select Case strChar
            Case Chr(Cancel)    'cancel code
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                lngFC = CStr(RAnsiColor(lngForeColor))
                lngBC = CStr(RAnsiColor(lngBackColor))
                strRTFBuff = strRTFBuff & strBuffer & "\b0\ul0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                strBuffer = ""
                i = i + 1
            Case strBold
                bbbold = Not bbbold
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\b"
                If bbbold = False Then strRTFBuff = strRTFBuff & "0"
                strBuffer = ""
                i = i + 1
            Case strUnderline
                bbunderline = Not bbunderline
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\ul"
                If bbunderline = False Then strRTFBuff = strRTFBuff & "none"
                strBuffer = ""
                i = i + 1
            Case strReverse
                bbreverse = Not bbreverse
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " " ' & strBuffer & "\"
                If bbreverse = False Then
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                Else
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngBackColor) + 1 & "\highlight" & RAnsiColor(lngForeColor) + 1
                End If
                
                strBuffer = ""
                i = i + 1
            Case strColor
                
                strTmp = ""
                i = i + 1

                Do Until Not ValidColorCode(strTmp) Or i > Length
                    strTmp = strTmp & Mid$(strData, i, 1)
                    i = i + 1
                Loop
                
                strTmp = LeftR(strTmp, 1)
                If strTmp = "" Then
                    lngFC = CStr(RAnsiColor(lngForeColor))
                    lngBC = CStr(RAnsiColor(lngBackColor))
                Else
                    lngFC = LeftOf(strTmp, ",")
                    lngFC = CStr(CInt(lngFC))
                    If InStr(strTmp, ",") Then
                        lngBC = RightOf(strTmp, ",")
                        If lngBC <> "" Then lngBC = CStr(CInt(lngBC)) Else lngBC = CStr(RAnsiColor(lngBackColor))
                    Else
                        lngBC = ""
                    End If
                End If
                
                If lngFC = "" Then lngFC = CStr(lngForeColor)
                lngFC = Int(lngFC) + 1
                If lngBC <> "" Then lngBC = Int(lngBC) + 1
                
                
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer
                strRTFBuff = strRTFBuff & "\cf" & lngFC
                If lngBC <> "" Then strRTFBuff = strRTFBuff & "\highlight" & lngBC
                
                i = i - 1
                strBuffer = ""
                If i >= Length Then GoTo TheEnd
                
            Case Else
                Select Case strChar
                Case "}", "{", "\"
                    strBuffer = strBuffer & "\" & strChar
                Case Else
                    strBuffer = strBuffer & strChar
                End Select
                i = i + 1
        End Select
        
    Loop Until i > Length
    
   
TheEnd:
    If strBuffer <> "" Then
        strRTFBuff = strRTFBuff & " " & strBuffer
    End If
    
    'Clipboard.SetText rtf.TextRTF & vbCrLf & vbCrLf & vbCrLf & strRTF & strRTFBuff & vbCrLf & "}", 1
    
    strRTFBuff = strRTFBuff & vbCrLf
    rtf.selStart = Len(rtf.Text)
    rtf.selLength = 0
    If rtf.Text = "" Then
        rtf.SelRTF = strRTF & strRTFBuff & vbCrLf & " }" & vbCrLf
    Else
        rtf.SelRTF = strRTF & "\par " & strRTFBuff & vbCrLf & " }" & vbCrLf
    End If
    
    rtf.selStart = Len(rtf.Text) - 2
    rtf.selLength = 0
    
End Sub
Sub SetText(rtf As RichTextBox, strData As String)
    
    DoEvents
    '* Not Finished
    If strData = "" Then Exit Sub
    
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String
    Dim clr As Integer, bclr As Integer, dftclr As Integer, strRTFBuff As String
    Dim bbbold As Boolean, bbunderline As Boolean, bbreverse As Boolean, strTmp As String
    Dim lngFC As String, lngBC As String, lngStart As Long, lngLength As Long, strPlaceHolder As String
    
    lngStart = rtf.selStart
    lngLength = rtf.selLength
    
    
    '* if not inialized, set font, intialiaze
    Dim btCharSet As Long
    Dim strRTF As String
    If rtf.Tag <> "init'd" Then
        rtf.Tag = "init'd"
        strFontName = rtf.Font.Name
        'CLIENT.picToolBar.FontName = strFontName
        btCharSet = GetTextCharset(CLIENT.picMenu.hdc)
        strRTF = ""
        strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
        strRTF = strRTF & ColorTable & vbCrLf
        strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
        strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & (intFontSize * 2) & vbCrLf
        strPlaceHolder = "\n"
        For i = 0 To DefinedColors - 1
            strRTF = strRTF & "\cf" & i & " " & strPlaceHolder
        Next
        'strRTF = strRTF & vbCrLf
        strRTF = strRTF & "}"
        rtf.TextRTF = strRTF
        
        '* New session for window... call
        '# LogData rtf.Parent.Caption, "blah", strData, True
    Else
        '# LogData rtf.Parent.Caption, "blah", strData, False
    End If
    
    rtf.parent.FontName = strFontName
    btCharSet = GetTextCharset(rtf.parent.hdc)
    strRTF = ""
    strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
    strRTF = strRTF & ColorTable & vbCrLf
    strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
    strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & (intFontSize * 2) & "\fi-" & "270" & "\li" & "270" & "\ffprot1 "
        
    strRTFBuff = "\b0\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1 & "\i0\ulnone "
    dftclr = RAnsiColor(lngForeColor)
    
    If bTimestamp Then
        strData = "15[" & Format(Time, strTimeFormat) & "] " & strData
    End If
    
    Length = Len(strData)
    i = 1
    
    Do
        strChar = Mid$(strData, i, 1)
        Select Case strChar
            Case strBold
                bbbold = Not bbbold
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\b"
                If bbbold = False Then strRTFBuff = strRTFBuff & "0"
                strBuffer = ""
                i = i + 1
            Case strUnderline
                bbunderline = Not bbunderline
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer & "\ul"
                If bbunderline = False Then strRTFBuff = strRTFBuff & "none"
                strBuffer = ""
                i = i + 1
            Case strReverse
                bbreverse = Not bbreverse
                strRTFBuff = strRTFBuff & " " & strBuffer & "\"
                If bbreverse = False Then
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngForeColor) + 1 & "\highlight" & RAnsiColor(lngBackColor) + 1
                Else
                    If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                    strRTFBuff = strRTFBuff & strBuffer & "\cf" & RAnsiColor(lngBackColor) + 1 & "\highlight" & RAnsiColor(lngForeColor) + 1
                    'strRTFBuff = strRTFBuff & " " & strBuffer & "\cf1\highlight3"
                End If
                
                strBuffer = ""
                i = i + 1
            Case strColor, Chr(15)
                
                strTmp = ""
                i = i + 1

                Do Until Not ValidColorCode(strTmp) Or i > Length
                    strTmp = strTmp & Mid$(strData, i, 1)
                    i = i + 1
                Loop
                
                strTmp = LeftR(strTmp, 1)
                If strTmp = "" Then
                    lngFC = CStr(RAnsiColor(lngForeColor))
                    lngBC = CStr(RAnsiColor(lngBackColor))
                Else
                    lngFC = LeftOf(strTmp, ",")
                    lngFC = CStr(CInt(lngFC))
                    If InStr(strTmp, ",") Then
                        lngBC = RightOf(strTmp, ",")
                        If lngBC <> "" Then lngBC = CStr(CInt(lngBC)) Else lngBC = CStr(RAnsiColor(lngBackColor))
                    Else
                        lngBC = ""
                    End If
                End If
                
                If lngFC = "" Then lngFC = CStr(lngForeColor)
                lngFC = Int(lngFC) + 1
                If lngBC <> "" Then lngBC = Int(lngBC) + 1
                
                
                If Right$(strRTFBuff, 1) <> " " Then strRTFBuff = strRTFBuff & " "
                strRTFBuff = strRTFBuff & strBuffer
                strRTFBuff = strRTFBuff & "\cf" & lngFC
                If lngBC <> "" Then strRTFBuff = strRTFBuff & "\highlight" & lngBC
                
                i = i - 1
                strBuffer = ""
                If i >= Length Then GoTo TheEnd
                
            Case Else
                Select Case strChar
                Case "}", "{", "\"
                    strBuffer = strBuffer & "\" & strChar
                Case Else
                    strBuffer = strBuffer & strChar
                End Select
                i = i + 1
        End Select
        
    Loop Until i > Length
    
   
TheEnd:
    If strBuffer <> "" Then
        strRTFBuff = strRTFBuff & " " & strBuffer
    End If

    Dim strR As String
    
    strRTFBuff = strRTFBuff & vbCrLf
    rtf.selStart = Len(rtf.Text)
    rtf.selLength = 0
    HideCaret rtf.hwnd
    rtf.SelRTF = strRTF & strRTFBuff & vbCrLf & " }" & vbCrLf
    HideCaret rtf.hwnd
    'rtf.seltext = vbCrLf
    
    'If GetKeyState(vbKeyScrollLock) Then
    '    rtf.SelStart = lngStart
    '    rtf.SelLength = lngLength
    'Else
        rtf.selStart = Len(rtf.Text) - 3
        rtf.selLength = 0
    'End If
    
    
End Sub


