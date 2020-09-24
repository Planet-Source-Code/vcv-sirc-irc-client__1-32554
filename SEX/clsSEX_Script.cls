VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSE_Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private alias()         As clsSSE_Alias
Private alias_count     As Integer
Public parent   As clsSSE_Main

Public returnVal As String
Public bExecuted As Boolean
Private Sub ClearScript()
    'ReDim Functions(1 To 1) As clsSSE_Alias
    Erase functions
End Sub


Public Function ExecuteAlias(strName As String, ByRef params() As String, AliasOrEvent As Integer, xServerID As Integer, localVars() As String, Optional strKey As String = "") As String
    
    Dim i As Integer
    bExecuted = False
    
    For i = 1 To alias_count
        If strName Like alias(i).GetName Then
            If alias(i).GetType = AliasOrEvent Then
                If AliasOrEvent = 3 Then    'hotkey
                    If strKey Like alias(i).GetModifier() Then
                    Else
                        GoTo finishoff
                    End If
                End If
                
                Dim tempAlias As New clsSSE_Alias
                alias(i).CopyAlias alias(i), tempAlias
            
                tempAlias.serverID = xServerID
                ExecuteAlias = tempAlias.Execute(params, localVars)
                bExecuted = True
                returnVal = ""
            End If
        End If
finishoff:
    Next i
End Function


Public Function GetAliasCount() As Integer
    GetAliasCount = alias_count
End Function

Public Function GetAlias(AliasNum As Integer) As String
    If alias(AliasNum).GetType = 0 Then 'alias
        GetAlias = "a " & alias(AliasNum).GetName
    ElseIf alias(AliasNum).GetType = 1 Then    'event
        GetAlias = "e " & alias(AliasNum).GetName
    Else
        GetAlias = "c " & alias(AliasNum).GetName
    End If
End Function


Public Sub LoadScript(strFileName As String, ByRef parent As clsSSE_Main)
    '* If file doesnt exist, exit
    If FileExists(strFileName) = False Then
        '* ...? echo...
        Exit Sub
    End If
    
    '* Variable declarations
    Dim FF As Integer, strLine As String, lineCount As Integer, i As Integer
    Dim inAlias As Boolean, strAlias() As String, strLines() As String, strBuffer As String
    
    'On Error GoTo error_handler
    FF = FreeFile
    Open strFileName For Input As #FF
        Do
            Line Input #FF, strLine
            strLine = TrimLeft(strLine)
        
            If Left(strLine, 1) = ";" Then GoTo ignorefullline
        
            If InStr(strLine, "|") = 0 Then
                strLine = strLine & "|;"
            End If
            
            strLines = Split(strLine, "|")
            i = 0
            strBuffer = ""
            Do
                If Right(strLines(i), 1) = "\" Then
                    strBuffer = strBuffer & strLines(i)
                    If i <> UBound(strLines) Then strBuffer = strBuffer & "|"
                    If i = UBound(strLines) Then
                    Else
                        GoTo ignoreline
                    End If
                Else
                    strBuffer = strBuffer & strLines(i)
                End If
                
                strBuffer = TrimLeft(strBuffer)
                If strBuffer = "" Then GoTo ignoreline
                '* Now lets check the line for an alias or event or ctcp tag
                'strLine = LCase(strLine)
                If strBuffer Like "alias *" Or strBuffer Like "event *" Or strBuffer Like "ctcp *" Or strBuffer Like "hotkey *" Then
                    If inAlias Then
                    Else
                        inAlias = True
                        strAlias = Split(strBuffer, " ")
                        If strAlias(0) = "alias" Then
                            NewAlias LeftOf(strAlias(1), "    "), 0, JoinArray(strAlias, " ", 3)
                        ElseIf strAlias(0) = "event" Then
                            NewAlias LeftOf(strAlias(1), "    "), 1, JoinArray(strAlias, " ", 3)
                        ElseIf strAlias(0) = "ctcp" Then
                            NewAlias LeftOf(strAlias(1), "    "), 2, JoinArray(strAlias, " ", 3)
                        ElseIf strAlias(0) = "hotkey" Then
                            NewAlias LeftOf(strAlias(1), "    "), 3, JoinArray(strAlias, " ", 3)
                        End If
                        Set alias(alias_count).rootEngine = parent
                        lineCount = 0
                    End If
                ElseIf strBuffer = "end alias" Or strBuffer = "end event" Or strBuffer = "end ctcp" Or strBuffer = "end hotkey" Then
                    inAlias = False
                ElseIf Left$(strBuffer, 1) = ";" Then
                Else
                    If inAlias Then
                        alias(alias_count).AddCodeLine strBuffer
                    End If
                End If
                strBuffer = ""
ignoreline:
            i = i + 1
            Loop Until i > UBound(strLines)
ignorefullline:
        Loop Until EOF(1)
    Close #FF
    Exit Sub
error_handler:
    '* echo an error message
    
End Sub



Private Sub NewAlias(strName As String, the_type As Integer, Optional strExtraParams As String)
    alias_count = alias_count + 1
    ReDim Preserve alias(1 To alias_count) As clsSSE_Alias
    Set alias(alias_count) = New clsSSE_Alias
    
    If the_type = 3 Then 'hotkey
        alias(alias_count).SetInfo strName, the_type, RightOf(strExtraParams, " "), LeftOf(strExtraParams, " ")
    Else
        alias(alias_count).SetInfo strName, the_type, strExtraParams
    End If
    
End Sub


