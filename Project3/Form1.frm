VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3420
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "GO"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdLogin_Click()
If TxtPass.Text = "daleclack" Then
Unload Me
Form2.Show
Else
Unload Me
Form3.Show
End If
End Sub
