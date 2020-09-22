VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "CoOl tHiNgS"
   ClientHeight    =   7545
   ClientLeft      =   2115
   ClientTop       =   525
   ClientWidth     =   4050
   Height          =   7950
   Icon            =   "Form1.frx":0000
   Left            =   2055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   4050
   Top             =   180
   Width           =   4170
   Begin VB.CommandButton Command21 
      Caption         =   "&Show Desktop Icons"
      Height          =   375
      Left            =   -120
      TabIndex        =   22
      Top             =   7200
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Hide Desktop Icons"
      Height          =   375
      Left            =   -120
      TabIndex        =   21
      Top             =   6840
      Width           =   3615
   End
   Begin VB.CommandButton Command20 
      Caption         =   "&Enable Ctrl+Alt+Del"
      Height          =   375
      Left            =   -120
      TabIndex        =   20
      Top             =   6480
      Width           =   3615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "&Disable Ctrl+Alt+Del"
      Height          =   375
      Left            =   -240
      TabIndex        =   19
      Top             =   6120
      Width           =   3735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "&Shutdown"
      Height          =   375
      Left            =   -120
      TabIndex        =   18
      Top             =   5760
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Restart"
      Height          =   375
      Left            =   -120
      TabIndex        =   17
      Top             =   5400
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&About"
      Height          =   3735
      Left            =   3480
      TabIndex        =   16
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command23 
      Caption         =   "&Hide Taskbar Icons"
      Height          =   375
      Left            =   -120
      TabIndex        =   15
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton Command22 
      Caption         =   "&Show Taskbar Icons"
      Height          =   375
      Left            =   -120
      TabIndex        =   14
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "&Show Clock"
      Height          =   375
      Left            =   -120
      TabIndex        =   13
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&Hide Clock"
      Height          =   375
      Left            =   -240
      TabIndex        =   12
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Show Startbutton"
      Height          =   375
      Left            =   -120
      TabIndex        =   11
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&Hide Startbutton"
      Height          =   375
      Left            =   -120
      TabIndex        =   10
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Mouse Pointer &R"
      Height          =   375
      Left            =   -120
      TabIndex        =   9
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Mouse Pointer &W"
      Height          =   375
      Left            =   -120
      TabIndex        =   8
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Show Systray"
      Height          =   375
      Left            =   -120
      TabIndex        =   7
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Hide Systray"
      Height          =   375
      Left            =   -120
      TabIndex        =   6
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Show Mouse"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&Hide Mouse"
      Height          =   375
      Left            =   -120
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&CD CLOSE"
      Height          =   375
      Left            =   -120
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&CD OPEN"
      Height          =   375
      Left            =   -120
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Exit"
      Height          =   3975
      Left            =   3480
      TabIndex        =   1
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Loop CD "
      Height          =   375
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Command1_Click()
Do
    Call CDRom_Open
    Call CDRom_Close
 Loop

End Sub


Private Sub Command10_Click()
    Call CDRom_Close

End Sub

Private Sub Command11_Click()
Unload Me
End Sub

Private Sub Command12_Click()

    Hid$ = ShowCursor(False)
End Sub

Private Sub Command13_Click()

    Hid$ = ShowCursor(True)
End Sub


Private Sub Command14_Click()
Dim Handle As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle&, 0
End Sub

Private Sub Command15_Click()
Dim Handle As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle&, 1
End Sub


Private Sub Command16_Click()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 0
End Sub

Private Sub Command17_Click()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 1
End Sub


Private Sub Command18_Click()
Call Computer_Shutdown
End Sub

Private Sub Command19_Click()
Call DisableCtrlAltDel
End Sub


Private Sub Command2_Click()
Screen.MousePointer = vbHourglass
PreventFromClosing
DisableCtrlAltDel
End Sub


Private Sub Command20_Click()
Call EnableCtrlAltDel
End Sub

Private Sub Command21_Click()
    DisplayDeskTopIcons True
   End Sub

Private Sub Command22_Click()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle&, 1
End Sub


Private Sub Command23_Click()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle&, 0
End Sub


Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command4_Click()
    DisplayDeskTopIcons False
End Sub


Private Sub Command5_Click()
 Call Computer_Restart
End Sub


Private Sub Command6_Click()
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 0
 
End Sub


Private Sub Command7_Click()
   Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1
End Sub


Private Sub Command8_Click()
Screen.MousePointer = vbArrow
UnPreventFromClosing
EnableCtrlAltDel
End Sub

Private Sub Command9_Click()
    Call CDRom_Open

End Sub


Private Sub MaskEdBox1_ValidationError(InvalidText As String, StartPosition As Integer)
End Sub


