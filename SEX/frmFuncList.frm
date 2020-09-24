VERSION 5.00
Begin VB.Form frmFuncList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   5565
   ClientTop       =   8925
   ClientWidth     =   2445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFunctions 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      IntegralHeight  =   0   'False
      ItemData        =   "frmFuncList.frx":0000
      Left            =   0
      List            =   "frmFuncList.frx":002E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2445
   End
End
Attribute VB_Name = "frmFuncList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nCount  As Integer

Public bShowInTaskbar As Boolean
Public serverID As Integer
Private Sub lstFunctions_KeyPress(KeyAscii As Integer)
    Dim strSel As String, tash As String
        
    If Chr(KeyAscii) = "]" Then
        Me.Visible = False
        frmSexIDE.txtCode.setFocus
        strSel = lstFunctions.List(lstFunctions.ListIndex)
        nCount = 0
    End If
    
    If KeyAscii = 13 Or Chr(KeyAscii) = " " Then
        tash = Chr(KeyAscii)
        If tash <> " " Then tash = ""
        strSel = lstFunctions.List(lstFunctions.ListIndex)
        On Error Resume Next
        frmSexIDE.txtCode.seltext = Right$(strSel, Len(strSel) - nCount) & tash
        Me.Visible = False
        frmSexIDE.txtCode.setFocus
        KeyAscii = 0
        nCount = 0
        Exit Sub
    End If
    
    If KeyAscii <> 8 Then
        frmSexIDE.txtCode.seltext = Chr(KeyAscii)
        nCount = nCount + 1
    End If
End Sub


