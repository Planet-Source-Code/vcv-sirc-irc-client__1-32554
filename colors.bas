Attribute VB_Name = "modColors"
Global Const COLOR_LGRAY = &HE0E0E0
Global Const COLOR_TBOUTLINE = 14212574
Global Const COLOR_DGRAY = 10526880
Global Const COLOR_HLBG = 14203830
Global Const COLOR_HLBORDER = 4268824

Public lngForeColor As Long
Public lngBackColor As Long

Public Type RGB
    r As Integer
    g As Integer
    B As Integer
End Type

Public Colors(99) As RGB
Public DefinedColors As Integer

Function ColorTable() As String
    Dim i As Integer, strTable As String
    Dim r As Integer, B As Integer, g As Integer
    strTable = "{\colortbl ;"
    'MsgBox DefinedColors
    For i = 0 To DefinedColors - 1
        r = Colors(i).r
        g = Colors(i).g
        B = Colors(i).B
        strTable = strTable & "\red" & r & "\green" & g & "\blue" & B & ";"
    Next i
    strTable = strTable & "}"
    ColorTable = strTable
End Function

Sub LoadColors()
    Dim i As Integer, strFile As String
    strFile = PATH & "colors.inf"
    
    On Error GoTo errorHandler
    DefinedColors = 0
    Open strFile For Input As #1
        Do
            Input #1, Colors(DefinedColors).r
            Input #1, Colors(DefinedColors).g
            Input #1, Colors(DefinedColors).B
            DefinedColors = DefinedColors + 1
        Loop Until EOF(1) Or DefinedColors >= 99
    Close #1
    Exit Sub
    
errorHandler:
    Select Case Err
        Case 53:
            MsgBox "(ERROR 53) The color information file does not exist, and therefore could not be loaded.  Please check the manual on how to fix this problem.  The program will now exit.", vbCritical
            End
        Case 76:
            MsgBox "(ERROR 76) The color information file does not exist, and therefore could not be loaded.  Please check the manual on how to fix this problem.  The program will now exit.", vbCritical
            End
        Case 62:
            MsgBox "(ERROR 62) The color information file is not complete.  Please check the manual for more information.", vbCritical
            End
        Case Else:
            MsgBox "An unknown error has occured.  The following information has been obtained, but is not documented in the manual." & vbCrLf & vbCrLf & "Error #" & Err & " : " & Error, vbCritical
    End Select

End Sub

Function RAnsiColor(lngColor As Long) As Integer

    Dim i As Integer
    For i = 0 To 15
        If RGB(Colors(i).r, Colors(i).g, Colors(i).B) = lngColor Then
            RAnsiColor = i
            Exit Function
        End If
    Next i
    
    If lngColor = lngForeColor Then
        RAnsiColor = 1
    ElseIf lngColor = lngBackColor Then
        RAnsiColor = 0
    Else
        RAnsiColor = 99
    End If
    
End Function


Sub ConvertToGrayscale(picColor As PictureBox, picBW As PictureBox)
    Dim x As Integer, y As Integer, AvgCol As Integer, lngColor As Long
    For x = 0 To picColor.ScaleWidth
        For y = 0 To picColor.ScaleHeight
            lngColor = picColor.point(x, y)
            If lngColor = 13160660 Then
            Else
                AvgCol = (Red(lngColor) + Green(lngColor) + Blue(lngColor)) / 3
                picBW.ForeColor = RGB(AvgCol, AvgCol, AvgCol)
                picBW.PSet (x, y)
            End If
        Next y
    Next x
End Sub

Function Red(ByVal Color As Long)
    Red = Color Mod 256
End Function


Function Green(ByVal Color As Long)
    Green = (Color / 256) Mod 256
End Function


Function Blue(ByVal Color As Long)
    Blue = Color / 65536
End Function

Sub SetRGB(lngColor As Long, ByRef r As Integer, ByRef g As Integer, ByRef B As Integer)
    
    r = Red(lngColor)
    g = Green(lngColor)
    B = Blue(lngColor)
    Exit Sub
End Sub


