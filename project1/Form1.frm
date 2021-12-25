VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   2505
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdJuBu 
      Caption         =   "局部变量"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TxtJubu 
      Height          =   495
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
Private Sub CmdJuBu_Click()
Dim i As Integer
i = i + 1
TxtJubu.Text = Str(i)
End Sub
