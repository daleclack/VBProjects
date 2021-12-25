VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   975
   ClientLeft      =   2910
   ClientTop       =   3225
   ClientWidth     =   2895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   2895
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":31C95
      Left            =   2040
      List            =   "Form1.frx":31C9F
      TabIndex        =   3
      Text            =   "E"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "120"
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Timer1_Timer()
Dim h, m, h1, m1, e, a, h2
h = Hour(Time) - 8
h2 = Hour(Time) + 4
m = Minute(Time)
e = Val(Txt1.Text)
If Combo1.Text = "W" Then
a = m + (180 - e) * 4
h1 = h2 + Int(a / 60)
m1 = a Mod 60
ElseIf Combo1.Text = "E" Then
a = m + e * 4
h1 = h + Int(a / 60)
m1 = a Mod 60
End If
If h1 < 0 Then
h1 = h1 + 24
ElseIf h1 >= 24 Then
h1 = h1 - 24
End If
Label1.Caption = Format(h1, "00") + ":" + Format(m1, "00")
End Sub

