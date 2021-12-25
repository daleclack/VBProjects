VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6300
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Cmd2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "显示"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.PictureBox Pic1 
      Height          =   3855
      Left            =   240
      ScaleHeight     =   3795
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd1_Click()
Dim a, b, i
a = 1
b = 2
i = 1
Do
Pic1.Print a; "  "; b;
Pic1.Print
a = a + b
b = a + b
i = i + 1
Loop While i < 15
End Sub

Private Sub Cmd2_Click()
End
End Sub
