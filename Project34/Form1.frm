VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1035
   ClientLeft      =   4170
   ClientTop       =   3480
   ClientWidth     =   2265
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   2265
   Begin VB.TextBox Txt2 
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Text            =   "120"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   1560
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e

Private Sub Command1_Click()
End
End Sub

Private Sub Timer1_Timer()
a = Hour(Time)
b = Minute(Time)
e = Val(Txt1.Text)
If Combo1.Text = "W" Then
c = a - Int((e + 120) / 15)
d = b - Rnd((e + 120) / 15) * 4 + 1
ElseIf Combo1.Text = "E" Then
c = a + Int((e - 120) / 15)
d = b + Rnd((e - 120) / 15) * 4
End If
If c < 0 Then
c = -c
End If
Txt2.Text = Str(Int(c)) + ":" + Str(Int(d))
End Sub

Private Sub Txt1_Change()
e = Txt1.Text
End Sub
