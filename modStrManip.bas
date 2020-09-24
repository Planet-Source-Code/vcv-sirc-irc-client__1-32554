Attribute VB_Name = "modStrManip"
Option Explicit
Public Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Sub Seperate(strData As String, strDelim As String, ByRef strLeft As String, ByRef strRight As String)
    '* Seperates strData into 2 variables based on strDelim
    '* Ex: strData is "Bill Clinton"
    '*     Dim strFirstName As String, strLastName As String
    '*     Seperate strData, " ", strFirstName, strLastName
    
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        strLeft = Left$(strData, intPos - 1)
        strRight = Mid$(strData, intPos + 1, Len(strData) - intPos)
    Else
        strLeft = strData
        strRight = strData
    End If
End Sub


Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left$(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function
Function LeftR(strData As String, intMin As Integer)
    
    On Error Resume Next
    LeftR = Left$(strData, Len(strData) - intMin)
End Function


Function RightOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        RightOf = Mid$(strData, intPos + 1, Len(strData) - intPos)
    Else
        RightOf = strData
    End If
End Function

Function RightR(strData As String, intMin As Integer)
    On Error Resume Next
    RightR = Right$(strData, Len(strData) - intMin)
End Function
