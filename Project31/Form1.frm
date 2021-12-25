VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   2115
   ClientTop       =   1425
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4200
   Begin VB.CommandButton Command5 
      Caption         =   "STOP"
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Btn4"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Btn3"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Btn2"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Btn1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "hello"
End Sub

Private Sub Command2_Click()
Command2.Visible = False
End Sub

Private Sub Command3_Click()
Command3.Move 200, 100
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Rnd() * (Form1.Width - Command4.Width)
Y = Rnd() * (Form1.Height - Command4.Height)
Command4.Left = X
Command4.Top = Y
End Sub

Private Sub Command5_Click()
End
End Sub

