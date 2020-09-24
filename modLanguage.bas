Attribute VB_Name = "modLanguage"
Option Explicit
Private LanguagePath As String


'**********************************
'*
'* Client Menu
'*
'**********************************
Public MI_FILE      As String
Public MI_EDIT      As String
Public MI_VIEW      As String
Public MI_FORMAT    As String
Public MI_COMMANDS  As String
Public MI_WINDOW    As String
Public MI_HELP      As String

' File menu
Public MI_FILE_NEWSERVER    As String
Public MI_FILE_CONNECT      As String
Public MI_FILE_DISCONNECT   As String

' Edit menu
Public MI_EDIT_UNDO         As String
Public MI_EDIT_CUT          As String
Public MI_EDIT_COPY         As String
Public MI_EDIT_PASTE        As String
Public MI_EDIT_DELETE       As String
Public MI_EDIT_SELECTALL    As String

' Format menu
Public MI_FORMAT_BOLD       As String
Public MI_FORMAT_COLOR      As String
Public MI_FORMAT_REVERSE    As String
Public MI_FORMAT_UNDERLINE  As String

' Window menu
Public MI_WINDOW_CASCADE    As String
Public MI_WINDOW_TILEH      As String
Public MI_WINDOW_TILEV      As String

'************************************

Sub Language_Init()
    Dim slash As String
    If Right$(App.PATH, 1) <> "/" Then slash = "/"
    
    LanguagePath = App.PATH & slash & "languagepacks/"
End Sub

Sub LoadLanguage(strLang As String)
    Dim LangFile As String, Trash As String
    LangFile = LanguagePath & strLang & ".lng"
    
    On Error GoTo errorHandler
    Open LangFile For Input As #1
        Input #1, MI_FILE, Trash
        Input #1, MI_EDIT, Trash
        Input #1, MI_VIEW, Trash
        Input #1, MI_FORMAT, Trash
        Input #1, MI_COMMANDS, Trash
        Input #1, MI_WINDOW, Trash
        Input #1, MI_HELP, Trash
        Input #1, Trash, Trash
        Input #1, MI_FILE_NEWSERVER, Trash
        Input #1, MI_FILE_CONNECT, Trash
        Input #1, MI_FILE_DISCONNECT, Trash
        Input #1, Trash, Trash
        Input #1, MI_EDIT_UNDO, Trash
        Input #1, MI_EDIT_CUT, Trash
        Input #1, MI_EDIT_COPY, Trash
        Input #1, MI_EDIT_PASTE, Trash
        Input #1, MI_EDIT_DELETE, Trash
        Input #1, MI_EDIT_SELECTALL, Trash
        Input #1, Trash, Trash
        Input #1, MI_FORMAT_BOLD, Trash
        Input #1, MI_FORMAT_COLOR, Trash
        Input #1, MI_FORMAT_REVERSE, Trash
        Input #1, MI_FORMAT_UNDERLINE, Trash
        Input #1, Trash, Trash
        Input #1, MI_WINDOW_CASCADE, Trash
        Input #1, MI_WINDOW_TILEH, Trash
        Input #1, MI_WINDOW_TILEV, Trash
    Close #1
    
    Exit Sub
errorHandler:
    Select Case Err
        Case 76:
            MsgBox "(ERROR 76) The language pack specified does not exist, and therefore could not be loaded.  Please check the manual on how to fix this problem.  The program will now exit.", vbCritical
            End
        Case 62:
            MsgBox "(ERROR 62) The language pack is not complete.  Please check the manual for more information. The program will now exit", vbCritical
            End
        Case Else:
            MsgBox "An unknown error has occured.  The following information has been obtained, but is not documented in the manual." & vbCrLf & vbCrLf & "Error #" & Err & " : " & Error, vbCritical
    End Select
    
End Sub

Public Sub SetLang_Menu()
    CLIENT.mnu_Connect.Caption = MI_FILE
    CLIENT.mnu_Edit.Caption = MI_EDIT
    CLIENT.mnu_View.Caption = MI_VIEW
    CLIENT.mnu_Format.Caption = MI_FORMAT
    CLIENT.mnu_Commands.Caption = MI_COMMANDS
    CLIENT.mnu_Window.Caption = MI_WINDOW
    CLIENT.mnu_Help.Caption = MI_HELP
    
    CLIENT.mnu_Connect_NewServer.Caption = MI_FILE_NEWSERVER
    CLIENT.mnu_Connect_Connect.Caption = MI_FILE_CONNECT
    CLIENT.mnu_Connect_Disconnect.Caption = MI_FILE_DISCONNECT
    
    CLIENT.mnu_Edit_Undo.Caption = MI_EDIT_UNDO
    CLIENT.mnu_Edit_Cut.Caption = MI_EDIT_CUT
    CLIENT.mnu_Edit_Copy.Caption = MI_EDIT_COPY
    CLIENT.mnu_Edit_Paste.Caption = MI_EDIT_PASTE
    CLIENT.mnu_Edit_Delete.Caption = MI_EDIT_DELETE
    CLIENT.mnu_Edit_SelectAll.Caption = MI_EDIT_SELECTALL
    
    CLIENT.mnu_Format_Bold.Caption = MI_FORMAT_BOLD
    CLIENT.mnu_Format_Color.Caption = MI_FORMAT_COLOR
    CLIENT.mnu_Format_Reverse.Caption = MI_FORMAT_REVERSE
    CLIENT.mnu_Format_Underline.Caption = MI_FORMAT_UNDERLINE
    
    CLIENT.mnu_Window_Cascade.Caption = MI_WINDOW_CASCADE
    CLIENT.mnu_Window_TileH.Caption = MI_WINDOW_TILEH
    CLIENT.mnu_Window_TileV.Caption = MI_WINDOW_TILEV
    
    
End Sub



