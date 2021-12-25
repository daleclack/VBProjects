VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3480
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Cmddenglu 
      Caption         =   "GO"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "password"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmddenglu_Click()
If Text1.Text = "daleclack" Then
Unlond Me
Form2.Show
Else
Unlond Me
Form3.Show
End If
End Sub

Private Sub Command2_Click()
Unlond Me
End
End Sub
