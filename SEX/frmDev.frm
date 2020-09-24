VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSexIDE 
   Caption         =   "sIRC Script Editor"
   ClientHeight    =   5670
   ClientLeft      =   2610
   ClientTop       =   2475
   ClientWidth     =   8385
   Icon            =   "frmDev.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   378
   ScaleMode       =   0  'User
   ScaleWidth      =   559
   Begin MSComctlLib.StatusBar sbScripts 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5370
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12595
            MinWidth        =   1191
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   714
            MinWidth        =   714
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   926
            MinWidth        =   926
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picResize 
      BorderStyle     =   0  'None
      Height          =   5730
      Left            =   2205
      MousePointer    =   9  'Size W E
      ScaleHeight     =   382
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   2
      Top             =   -15
      Width           =   60
   End
   Begin MSComctlLib.ImageList imgScripts 
      Left            =   1065
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDev.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDev.frx":15DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDev.frx":262E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDev.frx":3680
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvScripts 
      Height          =   5325
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   9393
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   317
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgScripts"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   1140
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Script"
      Filter          =   "SEX Script (*.sex)|*.sex|"
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   5340
      Left            =   2250
      TabIndex        =   0
      Top             =   15
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   9419
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmDev.frx":46D2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBMPC"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_File_Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_File_SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_File_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_ApplyChanges 
         Caption         =   "A&pply Script Changes"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu frm_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_Edit_Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_Edit_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_Cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_Edit_Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_Edit_Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_Edit_Delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu_Edit_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_Find 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_Edit_FindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_Edit_LB03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnu_Run 
      Caption         =   "&Run"
      Visible         =   0   'False
      Begin VB.Menu mnu_Run_Start 
         Caption         =   "&Start"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu_Run_Break 
         Caption         =   "&Break"
         Enabled         =   0   'False
         Shortcut        =   +^{F5}
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Options_ 
         Caption         =   "&Font..."
      End
   End
End
Attribute VB_Name = "frmSexIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bHasChanged     As Boolean
Private strFileName     As String
Public bShowInTaskbar As Boolean

Private nIndent         As Integer
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private bInVar      As Boolean
Private bInComm     As Boolean

Private Declare Function GetTextCharset Lib "gdi32" (ByVal hdc As Long) As Long

Private LastX   As Single, LastY As Single
Public serverID As Integer
Public strTitle As String
Public bNewData As Boolean
Sub DoComment()
    txtCode.SelColor = RGB(150, 0, 0)
End Sub

Sub LoadScripts()
        
    'this procedure WAS used to load buddies from the ini to treeview.
    'taken from Chad Cox's AIM example, thanks Chad. (ass)
    
    Dim strBuffer As String * 600, lngSize As Long, arrBuddies() As String, lngDo As Long
    Dim nod() As Node, intGroup As Integer, strOptions As String
    Dim i As Integer, j As Integer
    
    For i = LBound(strScripts) To UBound(strScripts)
        If strScripts(i) <> "" Then
            strOptions = strOptions & "g " & strScripts(i) & Chr(1)
            For j = 1 To scriptEngine.AliasCount(i + 1)
                strOptions = strOptions & scriptEngine.GetAlias(i + 1, j) & Chr(1)
            Next j
        End If
    Next i
                 
    With tvScripts
        arrBuddies$ = Split(strOptions, Chr(1))
        .Nodes.Clear
        For lngDo& = LBound(arrBuddies$) To UBound(arrBuddies$)
            'MsgBox arrBuddies(lngDo)
            ReDim Preserve nod(1 To .Nodes.Count + 1)
            If arrBuddies$(lngDo&) <> "" Then
                If Left$(arrBuddies$(lngDo&), 1) = "g" Then
                    Set nod(.Nodes.Count) = .Nodes.Add(, , , Right$(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 1, 1)
                    intGroup% = .Nodes.Count
                ElseIf Left$(arrBuddies$(lngDo&), 1) = "a" Then
                    If .Nodes.Count > 0 Then
                        Set nod(.Nodes.Count) = .Nodes.Add(nod(intGroup%), tvwChild, , Right$(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 2)
                        nod(.Nodes.Count).EnsureVisible
                    End If
                ElseIf Left$(arrBuddies$(lngDo&), 1) = "e" Then
                    If .Nodes.Count > 0 Then
                        Set nod(.Nodes.Count) = .Nodes.Add(nod(intGroup%), tvwChild, , Right$(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 4)
                        nod(.Nodes.Count).EnsureVisible
                    End If
                ElseIf Left$(arrBuddies$(lngDo&), 1) = "c" Then
                    If .Nodes.Count > 0 Then
                        Set nod(.Nodes.Count) = .Nodes.Add(nod(intGroup%), tvwChild, , Right$(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 3)
                        nod(.Nodes.Count).EnsureVisible
                    End If
                End If
            End If
        Next
    End With
End Sub

Sub NoComment()
    txtCode.SelColor = vbBlack
End Sub


Sub SyntaxHighlight(rtf As RichTextBox, strData As String)
    
    DoEvents
    '* Not Finished
    If strData = "" Then Exit Sub
    
    Dim i As Long, Length As Integer, strChar As String, strBuffer As String, prevChar As String
    Dim clr As Integer, bclr As Integer, dftclr As Integer, strRTFBuff As String
    Dim bInVar As Boolean, strTmp As String, strbufferx As String
    Dim lngFC As String, lngBC As String, lngStart As Long, lngLength As Long, strPlaceHolder As String
    Dim strFontName As String, strSData() As String, j As Integer
    Dim bMLComment As Boolean, bSLComment As Boolean, strPar As String
    strFontName = txtCode.Font.Name
    
    strSData = Split(strData, vbCrLf)
    
    
    lngStart = rtf.selStart
    lngLength = rtf.selLength
        
    '* if not initialized, set font, initialiaze
    Dim btCharSet As Long
    Dim strRTF As String

    strFontName = rtf.Font.Name
    Me.FontName = txtCode.Font.Name
    btCharSet = GetTextCharset(Me.hdc)
    strRTF = ""
    strRTF = strRTF & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fcharset" & btCharSet & " " & strFontName & ";}}" & vbCrLf
    strRTF = strRTF & "{\colortbl ;\red0\green0\blue0;\red0\green0\blue255;\red150\green0\blue0;\red0\green150\blue0;\red0\green150\blue0;}" & vbCrLf
    strRTF = strRTF & "{\stylesheet{ Normal;}{\s1 heading 1;}{\s2 heading 2;}}" & vbCrLf
    strRTF = strRTF & "\viewkind4\uc1\pard\cf0\f0\fs" & CInt(txtCode.Font.Size * 2) & " "
    strPlaceHolder = "\n"
    For i = 0 To 3
        strRTF = strRTF & "\cf" & i & " " & strPlaceHolder & " "
    Next
    
    Length = Len(strData)
    i = 1
    
    prevChar = ""
        
    For j = LBound(strSData) To UBound(strSData)
        strData = strSData(j)
        prevChar = ""
        Length = Len(strData)
        ' parse dat line
        For i = 1 To Length
            strChar = Mid$(strData, i, 1)
            
            If bMLComment = True Then
                If strChar = "`" And prevChar <> "\" Then
                    bMLComment = False
                    strRTFBuff = strRTFBuff & "\cf4 " & strBuffer & "\`\cf0"
                    strBuffer = ""
                    GoTo nextchar
                Else
                    strBuffer = strBuffer & strChar
                    GoTo nextchar
                End If
            End If
            
            If bSLComment = True Then
                Select Case strChar
                    Case "\", "}", "{"
                        strBuffer = strBuffer & "\" & strChar
                    Case Else
                        strBuffer = strBuffer & strChar
                End Select
                GoTo nextchar
            End If
            
            If prevChar <> "\" Then
                If strChar = "[" Then
                    strRTFBuff = strRTFBuff & strBuffer & "\cf0\b \[\b0 "
                    strBuffer = ""
                    GoTo nextchar
                ElseIf strChar = "]" Then
                    strRTFBuff = strRTFBuff & strBuffer & "\cf0\b \]\b0 "
                    strBuffer = ""
                    GoTo nextchar
                ElseIf strChar = "`" Then
                    bMLComment = True
                    strRTFBuff = strRTFBuff & strBuffer & "\cf4 \`"
                    strBuffer = ""
                    GoTo nextchar
                ElseIf strChar = ";" Then
                    bSLComment = True
                    strRTFBuff = strRTFBuff & strBuffer & "\cf4 ;"
                    strBuffer = ""
                    GoTo nextchar
                ElseIf strChar = "$" Then
                    strRTFBuff = strRTFBuff & strBuffer & "\cf3 \$"
                    strBuffer = ""
                    bInVar = True
                    GoTo nextchar
                ElseIf strChar = " " Then
                    strbufferx = Trim(strBuffer)
                    
                    If strbufferx = "endif" Or _
                       strbufferx = "if" Or _
                       strbufferx = "elseif" Or _
                       strbufferx = "else" Or _
                       strbufferx = "endalias" Or _
                       strbufferx = "alias" Or _
                       strbufferx = "endevent" Or _
                       strbufferx = "endctcp" Or _
                       strbufferx = "ctcp" Or _
                       strbufferx = "event" Or _
                       strbufferx = "endhotkey" Or _
                       strbufferx = "hotkey" Or _
                       strbufferx = "endwhile" Or _
                       strbufferx = "while" Or _
                       strbufferx = "endloop" Or _
                       strbufferx = "loop" Then
                       
                        
                        strRTFBuff = strRTFBuff & "\cf2 " & strBuffer & "\cf0 "
                        If InStr(strBuffer, " ") = False Then
                            strRTFBuff = strRTFBuff & " "
                        End If
                        
                        strBuffer = ""
                        GoTo nextchar
                    ElseIf strChar = "\" Then
                    
                        strRTFBuff = strRTFBuff & strBuffer & "\\"
                        strBuffer = ""
                    
                        GoTo nextchar
                        
                    Else
                    
                        'MsgBox "~" & strBuffer & "~"
                        If strbufferx = "end" Then
                            strBuffer = strBuffer & " "
                            GoTo nextchar
                        Else
                            strRTFBuff = strRTFBuff & strBuffer & "\cf0 " & strChar
                            strBuffer = ""
                            bInVar = False
                            GoTo nextchar
                        End If
                    End If
                    
                End If
            

            End If
            
            Select Case strChar
                Case "}", "{"
                    strBuffer = strBuffer & "\" & strChar
                Case "`", ",", ".", "/", "!", "@", "(", ")", "=", "+", "&", "^", "%", "*", "/", """"
                    strRTFBuff = strRTFBuff & strBuffer & "\cf0 " & strChar
                    strBuffer = ""
                    bInVar = False
                Case "\"
                    strRTFBuff = strRTFBuff & strBuffer & "\cf0\\"
                    strBuffer = ""
                Case Else
                    strBuffer = strBuffer & strChar
            End Select
            
nextchar:
            prevChar = strChar
        Next i
                
        strPar = "\par"
        If UBound(strSData) = j Then strPar = ""
        
        strbufferx = LTrim(strBuffer)
        If strbufferx = "if" Or _
            strbufferx = "elseif" Or _
            strbufferx = "else" Or _
            strbufferx = "end if" Or _
            strbufferx = "alias" Or _
            strbufferx = "end alias" Or _
            strbufferx = "event" Or _
            strbufferx = "end event" Or _
            strbufferx = "ctcp" Or _
            strbufferx = "end ctcp" Or _
            strbufferx = "hotkey" Or _
            strbufferx = "end hotkey" Or _
            strbufferx = "while" Or _
            strbufferx = "end while" Or _
            strbufferx = "loop" Or _
            strbufferx = "end loop" Then
               
            strRTFBuff = strRTFBuff & "\cf2 " & strBuffer & strPar & "\cf0"
            strBuffer = ""
        Else
            If bMLComment = True Then
                strRTFBuff = strRTFBuff & strBuffer & strPar & "\cf4"
            Else
                strRTFBuff = strRTFBuff & strBuffer & strPar & "\cf0"
            End If
            strBuffer = ""
        End If
        
        bSLComment = False
        bInVar = False
    Next j
    
    If strBuffer <> "" Then
        strRTFBuff = strRTFBuff & " " & strBuffer
    End If
    
    rtf.Text = ""
    rtf.SelRTF = strRTF & strRTFBuff & " }" & vbCrLf
    rtf.selStart = 0
    
End Sub

Function ColorTable() As String
    Dim i As Integer, strTable As String
    Dim r As Integer, B As Integer, g As Integer
    strTable = "{\colortbl ;"
    strTable = strTable & "\red0\green0\blue0;"
    strTable = strTable & "\red150\green0\blue0;"
    strTable = strTable & "\red0\green0\blue255;"
    strTable = strTable & "}"
    ColorTable = strTable
End Function

Sub BoldPrevChars(txtCode As RichTextBox, inum As Integer, Optional offSet As Integer = 0)
    If inum > txtCode.selStart + offSet Then Exit Sub
    
    Dim selStrt As Integer
    selStrt = txtCode.selStart
    txtCode.selStart = txtCode.selStart - inum + offSet
    txtCode.selLength = inum
    Call DoKW
    'txtCode.SelStart = selStrt - 1
    txtCode.selLength = 0
    Call NoKW
    txtCode.selStart = selStrt
    txtCode.selLength = 0
    Call NoKW

End Sub

Sub DoFunc()
    txtCode.SelBold = True
End Sub

Sub DoKW()
    txtCode.SelColor = vbBlue
End Sub

Sub DoVar()
    txtCode.SelColor = RGB(150, 0, 0)
End Sub

Function GetFont(strKey As String) As String
    GetFont = GetSetting("sexeditor", "fonts", strKey)
End Function


Function GetLine(strText As String, curPos As Integer) As String
    Dim strRet As String
    If curPos = 0 Then Exit Function
    
    Do
        If Mid$(strText, curPos, 1) = Chr(10) Then Exit Do
        strRet = Mid$(strText, curPos, 1) & strRet
        curPos = curPos - 1
    Loop Until curPos = 0
    GetLine = strRet
End Function

Sub InsertStruct(txtBox As RichTextBox, nMiddle As Integer, strEnding As String)
    
    If strEnding = "" Then
        strEnding = ""
    Else
        strEnding = vbCrLf & strEnding
    End If
    
    Call DoKW
    txtBox.seltext = vbCrLf & strrepeat(" ", nIndent + nMiddle) & strEnding
    txtBox.selStart = txtBox.selStart - Len(strEnding)
    Call NoKW
    
    
End Sub

Sub NoFunc()
    txtCode.SelBold = False
End Sub

Public Sub NoKW()
    If bInVar Then
    Else
        txtCode.SelColor = vbBlack
    End If
End Sub

Sub NoVar()
    txtCode.SelColor = vbBlack
End Sub

Sub OpenFile(strFileName As String)
    On Error GoTo noLoad
    Dim strFileData As String
    
    Open strFileName For Binary As #1
        strFileData = String(LOF(1), 0)
        Get #1, 1, strFileData
    Close #1
    
    mnu_Run_Start.Enabled = True
    
    SyntaxHighlight txtCode, strFileData
    'txtCode.Text = strFileData
    
    strFileData = ""
'    Delete strFileData
    
    Exit Sub
noLoad:
    MsgBox "An error has occured while trying to load the file: [" & Err & "]" & vbCrLf & vbCrLf, vbCritical
End Sub

Function PrevChars(txtCode As RichTextBox, inum As Integer) As String
    If inum > txtCode.selStart Then
        PrevChars = ""
    Else
        PrevChars = Mid$(txtCode.Text, txtCode.selStart - inum + 1, inum)
    End If
End Function

Sub SaveFile(strFileName As String)
    
    
    On Error GoTo noSaveAs
    Me.MousePointer = 11
    Open strFileName For Output As #1
        Print #1, txtCode.Text
    Close #1
    
    LoadScripts
    
    Me.MousePointer = 0
    mnu_Run_Start.Enabled = True
    Exit Sub
    
noSaveAs:
    Me.MousePointer = 0
    MsgBox "An error has occured while trying to save the file, [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation

End Sub

Sub SaveFont(strKey As String, strVal As String)
    SaveSetting "sexeditor", "fonts", strKey, strVal
End Sub

Private Sub Form_Load()
    
    nIndent = 4
    Center Me
    
    If GetFont("bold") = "" Then GoTo skipFont
    If GetFont("italic") = "" Then GoTo skipFont
    If GetFont("size") = "" Then GoTo skipFont
    If GetFont("name") = "" Then GoTo skipFont
    If GetFont("strikethru") = "" Then GoTo skipFont
    If GetFont("underline") = "" Then GoTo skipFont
    
    txtCode.Font.BOLD = CBool(GetFont("bold"))
    txtCode.Font.Name = GetFont("name")
    txtCode.Font.Italic = CBool(GetFont("italic"))
    txtCode.Font.Size = CInt(GetFont("size"))
    txtCode.Font.Strikethrough = CBool(GetFont("strikethru"))
    txtCode.Font.UNDERLINE = CBool(GetFont("underline"))
    
skipFont:
    Me.Visible = True
    If FileExists(Replace$(Command$, """", "")) Then
        OpenFile Replace$(Command$, """", "")
        strFileName = Replace$(Command$, """", "")
    End If
    
    LoadScripts
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtCode.Move tvScripts.Width + 3, 0, Me.ScaleWidth - tvScripts.Width - 2, Me.ScaleHeight - sbScripts.Height
    tvScripts.Move 0, 1, tvScripts.Width, Me.ScaleHeight - 2 - sbScripts.Height
    picResize.Move tvScripts.Width, 0
End Sub


Private Sub mnu_Edit_Copy_Click()
    Clipboard.SetText txtCode.seltext
End Sub

Private Sub mnu_Edit_Cut_Click()
    Clipboard.SetText txtCode.seltext
    txtCode.seltext = ""
End Sub

Private Sub mnu_Edit_Delete_Click()
    If txtCode.selLength > 0 Then
        txtCode.seltext = ""
    Else
        txtCode.selLength = 1
        txtCode.seltext = ""
    End If
End Sub

Private Sub mnu_Edit_Paste_Click()
'    txtCode.SelText = Clipboard.GetText()
End Sub

Private Sub mnu_Edit_SelectAll_Click()
    txtCode.selStart = 0
    txtCode.selLength = Len(txtCode.Text)
End Sub

Private Sub mnu_File_ApplyChanges_Click()
    If bHasChanged Then
        Dim intReturn As Integer
        intReturn = MsgBox("The current script you are editing has been changed and not saved.  Would you like to save it?", vbYesNoCancel Or vbQuestion)
        If intReturn = vbYes Then
            SaveFile strFileName
        ElseIf intReturn = vbCancel Then
            Exit Sub
        End If
    End If
    
    sbScripts.Panels(1).Text = "Reloading scripts, this may take a moment..."
    ReLoadScripts
    sbScripts.Panels(1).Text = "Scripts reloaded."
    
End Sub

Private Sub mnu_File_Exit_Click()
    If bHasChanged Then
        Dim intReturn As Integer
        intReturn = MsgBox("The current script you are editing has been changed and not saved.  Would you like to save it?", vbYesNoCancel Or vbQuestion)
        If intReturn = vbYes Then
            SaveFile strFileName
        ElseIf intReturn = vbCancel Then
            Exit Sub
        End If
    End If
    
    sbScripts.Panels(1).Text = "Reloading scripts, this may take a moment..."
    ReLoadScripts
    sbScripts.Panels(1).Text = "Scripts reloaded."
    Unload Me
End Sub

Private Sub mnu_File_New_Click()
    If bHasChanged = False Then
        txtCode.Text = ""
        strFileName = ""
        mnu_Run_Start.Enabled = False
    End If
End Sub

Private Sub mnu_File_Open_Click()
    On Error GoTo noopen
    cmDialog.DialogTitle = "Open Script"
    cmDialog.FileName = strFileName
    cmDialog.ShowOpen
    strFileName = cmDialog.FileName
    Me.MousePointer = 11
    
    OpenFile strFileName
    
    sbScripts.Panels(1).Text = "Opened - " & strFileName & " (" & Format(FileLen(strFileName), "###,###,###,###") & " bytes)"
    
    Me.MousePointer = 0
    Exit Sub
noopen:
    If Err = 32755 Then Exit Sub
    MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation

End Sub

Private Sub mnu_File_Save_Click()
    Me.MousePointer = 11
    If strFileName = "" Then
        Call mnu_File_SaveAs_Click
    Else
        SaveFile strFileName
        sbScripts.Panels(1).Text = "Saved Script - " & strFileName & " (" & Format(FileLen(strFileName), "###,###,###,###") & " bytes)"
    End If
    mnu_File_Save.Enabled = False
    Me.MousePointer = 0
End Sub

Private Sub mnu_File_SaveAs_Click()
    On Error GoTo noSaveAs
    cmDialog.DialogTitle = "Save Script"
    cmDialog.Filter = "SEX Script (*.sex)|*.sex|"
    cmDialog.FilterIndex = 0
    cmDialog.FileName = strFileName
    cmDialog.ShowSave
    strFileName = cmDialog.FileName
    SaveFile strFileName
    mnu_File_Save.Enabled = False
    
    Exit Sub
noSaveAs:
    If Err = 32755 Then Exit Sub
    MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation

End Sub


Private Sub mnu_Options__Click()
    On Error GoTo final
    cmDialog.Flags = &H100 Or &H3
    cmDialog.FontBold = txtCode.Font.BOLD
    cmDialog.FontName = txtCode.Font.Name
    cmDialog.FontItalic = txtCode.Font.Italic
    cmDialog.FontSize = txtCode.Font.Size
    cmDialog.FontStrikethru = txtCode.Font.Strikethrough
    cmDialog.FontUnderline = txtCode.Font.UNDERLINE
    cmDialog.ShowFont
    
    txtCode.Font.BOLD = cmDialog.FontBold
    txtCode.Font.Name = cmDialog.FontName
    txtCode.Font.Italic = cmDialog.FontItalic
    txtCode.Font.Size = cmDialog.FontSize
    txtCode.Font.Strikethrough = cmDialog.FontStrikethru
    txtCode.Font.UNDERLINE = cmDialog.FontUnderline
    
    SaveFont "bold", CStr(cmDialog.FontBold)
    SaveFont "name", CStr(cmDialog.FontName)
    SaveFont "italic", CStr(cmDialog.FontItalic)
    SaveFont "size", CStr(cmDialog.FontSize)
    SaveFont "strikethru", CStr(cmDialog.FontStrikethru)
    SaveFont "underline", CStr(cmDialog.FontUnderline)
    
    SyntaxHighlight txtCode, txtCode.Text
    
    Exit Sub
final:
    
End Sub

Private Sub mnu_Run_Break_Click()
'    Stop
End Sub


Private Sub mnu_Run_Start_Click()
    On Error GoTo errorHandler
    
    Me.MousePointer = 11
    mnu_Run_Break.Enabled = True
    Dim engine  As New clsSSE_Main
    engine.NewScript
    engine.LoadScript engine.ScriptCount, strFileName
    Dim paramlist(0)
    'engine.ExecuteAlias "main", paramlist()
    Me.MousePointer = 0
    Exit Sub
errorHandler:
    MsgBox "An error has occured while trying to run the script, [" & Err & "]:" & vbCrLf & vbCrLf & Error
    
End Sub



Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastX = x: LastY = y
    picResize.Visible = False
End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        tvScripts.Width = tvScripts.Width - (LastX - x)
        txtCode.Move tvScripts.Width + 3, 0, Me.ScaleWidth - tvScripts.Width - 2
        LastX = x
    End If
End Sub


Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picResize.Visible = True
    picResize.Left = tvScripts.Width
End Sub

Private Sub tvScripts_Click()
    Dim strFile As String
    On Error GoTo nulllabel
    strFile = tvScripts.selectedItem.Text
    
    On Error GoTo isscript
    If tvScripts.selectedItem.parent <> "" Then
        
        If tvScripts.selectedItem.Image = 2 Then
            If InStr(txtCode.Text, "alias " & strFile) Then
                txtCode.setFocus
                txtCode.selStart = InStr(txtCode.Text, "alias " & strFile) - 1
            End If
        ElseIf tvScripts.selectedItem.Image = 1 Then
            If InStr(txtCode.Text, "event " & strFile) Then
                txtCode.setFocus
                txtCode.selStart = InStr(txtCode.Text, "event " & strFile) - 1
            End If
        ElseIf tvScripts.selectedItem.Image = 3 Then
            If InStr(txtCode.Text, "ctcp " & strFile) Then
                txtCode.setFocus
                txtCode.selStart = InStr(txtCode.Text, "ctcp " & strFile) - 1
            End If
        Else
            If InStr(txtCode.Text, "hotkey " & strFile) Then
                txtCode.setFocus
                txtCode.selStart = InStr(txtCode.Text, "hotkey " & strFile) - 1
            End If
        End If
        
    End If
    
    Exit Sub
isscript:
    Me.MousePointer = 11
    If FileExists(strFile) Then
        OpenFile strFile
        strFileName = strFile
        sbScripts.Panels(1).Text = "Now viewing Script - " & strFile & " (" & Format(FileLen(strFile), "###,###,###,###") & " bytes)"
    ElseIf FileExists(PATH & strFile) Then
        OpenFile PATH & strFile
        strFileName = PATH & strFile
        sbScripts.Panels(1).Text = "Now viewing Script - " & strFile & " (" & Format(FileLen(PATH & strFile), "###,###,###,###") & " bytes)"
    Else
        MsgBox "File does not exist, please go to options and fix this problem.", vbCritical, "ERROR"
    End If
    Me.MousePointer = 0
nulllabel:
End Sub


Private Sub txtCode_Change()
    bHasChanged = True
    mnu_File_Save.Enabled = True
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
    
    If txtCode.SelColor = RGB(0, 150, 0) And KeyAscii <> 13 Then
        Exit Sub
    End If
    
    NoKW
    NoFunc
    
    If Not bInVar Then NoVar
    
    If Chr(KeyAscii) = "$" Then
        If PrevChars(txtCode, 1) = "\" Then Exit Sub
        DoVar
        bInVar = True
    End If
    
    If Chr(KeyAscii) = " " Then
        If bInVar Then
            bInVar = False
            NoVar
        End If
        
        If PrevChars(txtCode, 5) = "alias" Then
            BoldPrevChars txtCode, 5
        ElseIf PrevChars(txtCode, 5) = "event" Then
            BoldPrevChars txtCode, 5
        ElseIf PrevChars(txtCode, 4) = "ctcp" Then
            BoldPrevChars txtCode, 4
        ElseIf PrevChars(txtCode, 6) = "hotkey" Then
            BoldPrevChars txtCode, 6
        ElseIf PrevChars(txtCode, 6) = "elseif" Then
            BoldPrevChars txtCode, 6
        ElseIf PrevChars(txtCode, 4) = "else" Then
            BoldPrevChars txtCode, 4
        ElseIf PrevChars(txtCode, 2) = "if" Then
            BoldPrevChars txtCode, 2
        ElseIf PrevChars(txtCode, 5) = "while" Then
            BoldPrevChars txtCode, 5
        ElseIf PrevChars(txtCode, 4) = "loop" Then
            BoldPrevChars txtCode, 4
        End If
    End If
    
    Dim pnt As POINTAPI, retVal As Long
    If Chr(KeyAscii) = "[" Then
        If bInVar Then
            bInVar = False
            NoVar
        End If
        
        If PrevChars(txtCode, 1) = "\" Then Exit Sub
        KeyAscii = 0
        DoFunc
        txtCode.seltext = "[]"
        txtCode.selStart = txtCode.selStart - 1
        NoFunc
        retVal = GetCaretPos(pnt)
        frmFuncList.Move Me.Left + (pnt.x * 15), Me.Top + (pnt.y * 15) + 900
        frmFuncList.Show
        Exit Sub
    End If

    If KeyAscii = 9 Then
        If bInVar Then
            bInVar = False
            NoVar
        End If
        
        KeyAscii = 0
        Dim selSt As Integer, selLn As Integer
        selSt = txtCode.selStart
        selLn = txtCode.selLength
        txtCode.seltext = strrepeat(" ", 4) & txtCode.seltext
        If selLn <> 0 Then
            txtCode.selStart = selSt
            txtCode.selLength = selLn + 4
        End If
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        If bInVar Then
            bInVar = False
            NoVar
        End If
        
        If PrevChars(txtCode, 4) = "else" Then
            BoldPrevChars txtCode, 4
            txtCode.seltext = strrepeat(" ", nIndent)
        ElseIf PrevChars(txtCode, 9) = "end alias" Then
            BoldPrevChars txtCode, 9
        ElseIf PrevChars(txtCode, 9) = "end event" Then
            BoldPrevChars txtCode, 9
        ElseIf PrevChars(txtCode, 8) = "end ctcp" Then
            BoldPrevChars txtCode, 8
        ElseIf PrevChars(txtCode, 10) = "end hotkey" Then
            BoldPrevChars txtCode, 10
        ElseIf PrevChars(txtCode, 9) = "end while" Then
            BoldPrevChars txtCode, 9
        ElseIf PrevChars(txtCode, 8) = "end loop" Then
            BoldPrevChars txtCode, 8
        ElseIf PrevChars(txtCode, 6) = "end if" Then
            BoldPrevChars txtCode, 6
        End If
        
        KeyAscii = 0
        
        Dim strLine As String, strFull As String
        strFull = GetLine(txtCode.Text, txtCode.selStart)
        strLine = Trim(strFull)
        
        If txtCode.selLength > 0 Then
            txtCode.seltext = vbCrLf & txtCode.seltext
            Call NoKW
            Exit Sub
        End If
        
        '* auto-complete
        If strLine Like "alias *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "alias")), LeftOf(strFull, "alias") & "end alias"
        ElseIf strLine Like "event *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "event")), LeftOf(strFull, "event") & "end event"
        ElseIf strLine Like "ctcp *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "event")), LeftOf(strFull, "ctcp") & "end ctcp"
        ElseIf strLine Like "hotkey *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "hotkey")), LeftOf(strFull, "hotkey") & "end hotkey"
        ElseIf strLine Like "if *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "if")), LeftOf(strFull, "if") & "end if"
        ElseIf strLine Like "elseif *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "elseif")), ""
        ElseIf strLine Like "else *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "else")), ""
        ElseIf strLine Like "while *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "while")), LeftOf(strFull, "while") & "end while"
        ElseIf strLine Like "loop *" Then
            InsertStruct txtCode, Len(LeftOf(strFull, "loop")), LeftOf(strFull, "loop") & "end loop"
        Else
            If strLine = "" Then
                If strFull = "" Then
                    txtCode.seltext = vbCrLf
                    Call NoKW
                    Exit Sub
                End If
                txtCode.selStart = txtCode.selStart - 1
                If Mid$(txtCode.Text, txtCode.selStart + 1, 1) = Chr(10) Then Exit Sub
                txtCode.selStart = txtCode.selStart - 1
                If Mid$(txtCode.Text, txtCode.selStart + 1, 1) = Chr(10) Then Exit Sub
                txtCode.selStart = txtCode.selStart - 1
                If Mid$(txtCode.Text, txtCode.selStart + 1, 1) = Chr(10) Then Exit Sub
                txtCode.selStart = txtCode.selStart - 1
                If Mid$(txtCode.Text, txtCode.selStart + 1, 1) = Chr(10) Then Exit Sub
                txtCode.selLength = 4
                txtCode.seltext = ""
            Else
                txtCode.seltext = vbCrLf & strrepeat(" ", Len(strFull) - Len(strLine))
                Call NoKW
            End If
        
        End If
        Exit Sub
    End If
    
    If InStr("?,./!\@()=+&^%*", Chr(KeyAscii)) Then
        If bInVar Then
            bInVar = False
            NoVar
        End If
    End If
    
    If frmFuncList.Visible = True Then
        retVal = GetCaretPos(pnt)
        frmFuncList.Move Me.Left + (pnt.x * 15), Me.Top + (pnt.y * 15) + 900
    End If
End Sub

