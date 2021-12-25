VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4020
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   4020
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   1335
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Cmd4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1320
      ItemData        =   "Form2.frx":0000
      Left            =   0
      List            =   "Form2.frx":001F
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "描述:"
      Height          =   180
      Left            =   1800
      TabIndex        =   9
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "总数:"
      Height          =   180
      Left            =   1800
      TabIndex        =   8
      Top             =   360
      Width           =   450
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd3_Click()
List1.Clear
End Sub

Private Sub Cmd4_Click()
Me.Hide
Form1.Show
End Sub

Private Sub List1_Click()
Text3.Text = List1.List(i)
End Sub
