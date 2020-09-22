VERSION 4.00
Begin VB.Form Form2 
   Caption         =   "AbOuT"
   ClientHeight    =   750
   ClientLeft      =   1215
   ClientTop       =   1785
   ClientWidth     =   4710
   Height          =   1155
   Icon            =   "Form2.frx":0000
   Left            =   1155
   LinkTopic       =   "Form2"
   ScaleHeight     =   750
   ScaleWidth      =   4710
   Top             =   1440
   Width           =   4830
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "About"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Label Label1 
         Caption         =   "All copyrights reserved by Christophe De Coninck      This program is made in VB 4.0/32 bit"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub


