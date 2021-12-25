VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   7155
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   1560
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "BACK"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1110
      Index           =   1
      Left            =   360
      Picture         =   "Form2.frx":0000
      Top             =   1920
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   1110
      Index           =   0
      Left            =   360
      Picture         =   "Form2.frx":1045
      Top             =   1920
      Width           =   1110
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBack_Click()
Unload Me
Form1.Show
Form1.TxtPass.Text = ""
End Sub

Private Sub CmdStart_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Static i, x, y As Integer
 Image1(i).Visible = True
 x = Image1(i).Left + 370 * Rnd
 y = Image1(i).Top - 80 * Rnd
 i = (i + 1) Mod 2
 Image1(i).Visible = False
 Image1(i).Move x, y
End Sub
