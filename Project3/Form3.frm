VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdExit 
      Caption         =   "Return"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Password wrong!"
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
Form1.Show
Form1.TxtPass.Text = ""
End Sub
