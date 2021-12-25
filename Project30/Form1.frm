VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2565
   ClientLeft      =   2745
   ClientTop       =   1860
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   3705
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "print"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Timer1.Enabled = False Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
Dim a, b, c
Randomize
a = Int(256 * Rnd)
b = Int(256 * Rnd)
c = Int(256 * Rnd)
Form1.ForeColor = RGB(a, b, c)
CurrentX = Form1.Width * Rnd
CurrentY = Form1.Height * Rnd
Print "*"
End Sub
