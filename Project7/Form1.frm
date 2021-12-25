VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   3750
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox TxtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton CmdGO 
      Caption         =   "GO"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
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

Private Sub CmdGO_Click()
If TxtPass.Text = "daleclack" Then
Unload Me
Form2.Show
Else
Form1.Label2.Caption = "Password wrong!"
End If
End Sub
