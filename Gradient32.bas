Attribute VB_Name = "Gradient32"
Option Explicit
'**********************************
'* CODE BY: PATRICK MOORE (ZELDA) *
'* Feel free to re-distribute or  *
'* Use in your own projects.      *
'* Giving credit to me would be   *
'* nice :)                        *
'*                                *
'* Please vote for me if you find *
'* this code useful :]   -Patrick *
'**********************************
'
'PS: Please look for more submissions to PSC by me
'    shortly.  I've recently been working on a lot
'    :))  All my submissions are under author name
'    "Patrick Moore (Zelda)"

Global rRed As Long, rBlue As Long, rGreen As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Sub FormDrag(Frm As Form)
    'Release capture
    Call ReleaseCapture
    
    'Send API messages to the form
    Call SendMessage(Frm.hWnd, &H112, &HF012, 0)
End Sub

Function RGBfromLONG(LongCol As Long)
' Get The Red, Blue And Green Values Of A Colour From The Long Value
Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double
Blue = Fix((LongCol / 256) / 256)
Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
rRed = Red: rBlue = Blue: rGreen = Green
End Function

Public Sub Gradient(picBox, StartColor As Long, EndColor As Long)
Dim x%, x2!, y%, i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
x% = 0
' find the length of the picturebox and
'     cut it into 100 pieces
x2 = picBox.ScaleWidth / 100
y% = picBox.ScaleHeight
' setting how much red, green, and blue
'     goes into each of the two colors
RGBfromLONG StartColor
red1% = Abs(rRed) ' the amount of red In color one
green1% = Abs(rGreen)
blue1% = Abs(rBlue)

RGBfromLONG EndColor
red2% = Abs(rRed) ' the amount of red In color two
green2% = Abs(rGreen)
blue2% = Abs(rBlue)
' cut the difference between the two col
'     ors into 100 pieces
pat1 = (red2% - red1%) / 100
pat2 = (green2% - green1%) / 100
pat3 = (blue2% - blue1%) / 100
' set the c variables at the starting co
'     lors
c1 = red1%
c2 = green1%
c3 = blue1%
' draw 100 different lines on the pictur
'     ebox


For i% = 1 To 100
    picBox.Line (x%, 0)-(x% + x2, y%), RGB(c1, c2, c3), BF
    x% = x% + x2 ' draw the Next line one step up from the old step
    c1 = c1 + pat1 ' make the c variable equal 2 it's Next step
    c2 = c2 + pat2
    c3 = c3 + pat3
Next

End Sub

