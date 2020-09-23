Attribute VB_Name = "mHTMLHelp"
Option Explicit

' Microsoft Knowledge Base Article - 183434
' HOWTO: Use HTML Help API in a Visual Basic Application

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_TP_HELP_CONTEXTMENU = &H10
Public Const HH_TP_HELP_WM_HELP = &H11

Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
         (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long


Public Sub ShowHTMLHelp(ByVal hWnd As Long)

    Dim hwndHelp As Long
    
    'The return value is the window handle of the created help window.
    hwndHelp = HtmlHelp(hWnd, App.Path & "\CompuFont8.chm", HH_DISPLAY_TOPIC, 0)
'    hwndHelp = HtmlHelp(hWnd, App.Path & "\HelpdeskSoftware.chm", HH_HELP_CONTEXT, 9000)

End Sub

