VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   2820
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdClear 
      Caption         =   "清除"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "显示"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClear_Click()
Text1.Text = ""
CmdShow.Enabled = True
CmdClear.Enabled = False
End Sub

Private Sub CmdShow_Click()
Text1.Text = "Adobe Flash Professional CC"
CmdShow.Enabled = False
CmdClear.Enabled = True
End Sub
