Attribute VB_Name = "Module1"
'PHREAKY!!!'

'Ok people... here's some KINDA useful, but mostly lamely cool
'Visual Basic options i decided to throw together...
'if you want me to make new verzion, send me some feedback!
'worldfamouskr0q@phreaker.net'

'Also,if you find any bugs, email me at the eMail address
'above, and I'll fix them!

'thanks and enjoy my first module released to the public!!!

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long


Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long



Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)


Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Const SW_HIDE = 0    ' Hide Window
Public Const SW_SHOW = 5    ' Show Window


Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_USER = &H400
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_SYSCOMMAND = &H112
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3

Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE




Public Const WM_ENABLE = &HA



Public Const HIDE_WINDOW = 0
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_SELECTSTRING = &H18C

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT



Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_UP = &H26

Public Const VK_SPACE = &H20


Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000


Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40



Sub CDRom_Open()
'Open's your CD Rom drive...AKA cup holder...

    retvalue = MciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Sub CDRom_Close()
'Closes your CD Rom drive

    retvalue = MciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub
Sub CRRom_Ghost()
'Pretty stupid... might be funny if you leave this running on someone's computer
    'while they're out to coffe... SET THE PAUSE INTERVALS TO WHATEVER YOU WANT,
    'DEPENDING ON WHAT YOU WANT THIS TO ACT LIKE...
    
Do
    Call CDRom_Open
    Pause 1
    Call CDRom_Close
    Pause 3
Loop
End Sub

Sub Computer_Restart()
'Will restart the computer

    ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub
Sub Computer_Shutdown()
'Will shut-down the computer

    StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub
Sub Computer_ForceShutdown()
'Forces a shut-down of the computer

    ForcedShutdown = ExitWindowsEx(EWX_FORCE, 0&)
End Sub
Sub Mouse_Show()
'Shows your mouse cursor

    Hid$ = ShowCursor(True)
End Sub
Sub Mouse_Hide()
'Hides your mouse cursor

    Hid$ = ShowCursor(False)
End Sub

Sub Mouse_Insane()
'Crzy mouse movement.. could be useful to make a prog like Fake Surf...

Do
    boob = (Rnd * 400)
    boob2 = (Rnd * 400)
    DoEvents
Loop
End Sub
Sub StartMenu_Hide()
'Hides the Start Menu

End Sub
Sub StartMenu_Show()
'Shows the Start Menu
' doesn't work for me, but maybe that's cause i'm on Windows NT...

End Sub

Function CtrlAltDel_Enable()
'This re-enables CTRL+ALT+Delete

    Dim ret As Integer
    Dim pOld As Boolean
     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Function
Function CtrlAltDel_Disable()
'This will disable the CTRL+ALT+DELETE function of Windows.
    'Make sure you re-enable this before your prog ends,
    'or the person using this is screwed!
    
    Dim ret As Integer
    Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Function

Sub ScreenSaver_On()

End Sub
Sub beep()
'Not much... tossed this in for those who didn't know about this..
'you don't need this module to use this, just type beep, and it'll
'make that beep sound...

End Sub
Sub PlayWav(DiR)
'Plays the specified WAV file...


End Sub
Sub Pause(interval)
'Don't modify this!
    'This just puts a delay between actions...
    'like...

'Call CtrlAltDel_Disable
'Pause 15
'Call CtrlAltDel_Enable

'That will put a 15 second pause between the
    'CTRL ALT DELETE Disable and Enable
    
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub
