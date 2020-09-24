Attribute VB_Name = "modIAL"
Option Explicit

Private Type typeIAL
    Nick As String
    Host As String
    ident As String
    fullname As String
End Type

Global IAL() As typeIAL
Private IAL_Length As Long

Private IAL_LowFree As Long
Private IAL_HighFree As Long
Sub IAL_AddAddress(strNick As String, Optional strHost As String = "", Optional strIdent As String = "", Optional strFullName As String = "")
    
    Dim i As Long
    i = 1
    Do
        If IAL(i).Nick = "" Then
            IAL(i).Nick = strNick
            IAL(i).Host = strHost
            IAL(i).fullname = strFullName
            IAL(i).ident = strIdent
            
            If i < IAL_LowFree Then IAL_LowFree = i
            If i > IAL_HighFree Then IAL_HighFree = i
            
            Exit Sub
        End If
        i = i + 1
    Loop Until i >= IAL_Length
    
    IAL_Length = IAL_Length + 1
    ReDim Preserve IAL(1 To IAL_Length)
    IAL(IAL_Length).Nick = strNick
    IAL(IAL_Length).Host = strHost
    IAL(IAL_Length).fullname = strFullName
    IAL(IAL_Length).ident = strIdent
    
    If IAL_LowFree = 0 Or IAL_LowFree = IAL_Length - 1 Then
        IAL_LowFree = IAL_LowFree + 1
        IAL_HighFree = IAL_HighFree + 1
    End If
    
End Sub

Sub IAL_Update(strNick As String, strNewNick As String, Optional strHost As String = "", Optional strIdent As String = "", Optional strFullName As String = "")
    
    Dim i As Long
    i = IAL_LowFree
    Do
        If IAL(i).Nick = strNick Then
            If strNewNick <> "" Then IAL(i).Nick = strNewNick
            If strHost <> "" Then IAL(i).Host = strHost
            If strIdent <> "" Then IAL(i).ident = strIdent
            If strFullName <> "" Then IAL(i).fullname = strFullName
            Exit Do
        End If
        i = i + 1
    Loop Until i >= IAL_HighFree
    
    
End Sub


