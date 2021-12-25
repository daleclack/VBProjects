VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4185
   ScaleWidth      =   3960
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   3
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   480
      Picture         =   "Form2.frx":038A
      Top             =   360
      Width           =   240
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub CmdExit_Click()
Me.Hide
Form1.Show
End Sub

Private Sub CmdStart_Click()
Label1.Caption = ""
CmdStart.Caption = "Go on"
Command1(a).DisabledPicture = LoadPicture()
For i = 0 To 3
Command1(i).Enabled = True
Next
End Sub

Private Sub Command1_Click(Index As Integer)
Randomize
a = Int(4 * Rnd)
If a = Index Then
Label1.Caption = "Good luck!"
End If
For i = 0 To 3
Command1(i).Enabled = False
Command1(i).DisabledPicture = LoadPicture()
Next i
Command1(a).DisabledPicture = LoadPicture(App.Path + "\1.jpg")
End Sub
