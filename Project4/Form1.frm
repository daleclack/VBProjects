VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   2595
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdJubu 
      Caption         =   "静态变量"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtJubu 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdJubu_Click()
Static i As Integer
i = i + 1
TxtJubu.Text = Str(i)
End Sub
