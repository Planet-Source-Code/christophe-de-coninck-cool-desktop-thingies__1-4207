Attribute VB_Name = "DeskTopRoutines"
Option Explicit
'
' Declarations
'

Public Const SW_HIDE = 0    ' Hide Window
Public Const SW_SHOW = 5    ' Show Window

Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
   
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Sub DisplayTaskBar(ByVal bShow As Boolean)
    Dim lTaskBarHWND As Long
    Dim lRet As Long
    Dim lFlags As Long
'
' Show / hide the taskbar
'
    On Error GoTo vbErrorHandler
    
    lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
    
    
    lTaskBarHWND = FindWindow("Shell_TrayWnd", "")

    lRet = ShowWindow(lTaskBarHWND, lFlags)
    
    If lRet < 0 Then
    '
    ' Handle error from api
    '
    End If

    Exit Sub
    
vbErrorHandler:
'
' Handle Errors here
'
End Sub

Public Sub DisplayDeskTopIcons(ByVal bShow As Boolean)
    Dim lDesktopHwnd As Long
    Dim lFlags As Long
'
' Show / Hide the Desktop Icons
'
    On Error Resume Next

    lDesktopHwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)

    If lDesktopHwnd = 0 Then
        ' raise an error ! You have no desktop !!!
        Exit Sub
    End If
    
    lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
    
    ShowWindow lDesktopHwnd, lFlags
    
End Sub
