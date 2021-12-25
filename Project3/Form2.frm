VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4200
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdReturn 
      Caption         =   "BACK"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome!"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdReturn_Click()
Unload Me
Form1.Show
Form1.TxtPass.Text = ""
End Sub
