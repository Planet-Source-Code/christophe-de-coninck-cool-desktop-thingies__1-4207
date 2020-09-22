VERSION 4.00
Begin VB.Form Form3 
   Caption         =   "Enter"
   ClientHeight    =   1050
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   2595
   Height          =   1455
   Icon            =   "Form3.frx":0000
   Left            =   1080
   LinkTopic       =   "Form3"
   ScaleHeight     =   1050
   ScaleWidth      =   2595
   Top             =   1170
   Width           =   2715
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim Counter As Integer
    Dim Workarea(2500) As String
    ProgressBar1.Min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.Min
       Form1.Show
    Unload Me
End Sub

Private Sub Form_Load()
Dim Counter As Integer
    Dim Workarea(250) As String
    ProgressBar1.Min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter
    Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.Min
End Sub

Private Sub Form()
    ProgressBar1.Align = vbAlignBottom
    ProgressBar1.Visible = False
    Command1.Caption = "Initialize array"
    Form1.Show
    Unload Me

End Sub


Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Form1.Show
    Unload Me
End Sub


