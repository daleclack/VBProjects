VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   1950
   ClientTop       =   1980
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5580
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox TxtV 
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "暂停"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "开始"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2640
      Top             =   1440
   End
   Begin VB.ListBox LstX 
      Height          =   2940
      ItemData        =   "Form1.frx":0000
      Left            =   3360
      List            =   "Form1.frx":0002
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.ListBox LstV 
      Height          =   2400
      ItemData        =   "Form1.frx":0004
      Left            =   720
      List            =   "Form1.frx":0006
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox TxtS 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox TxtA 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "初始速度"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "x(m)"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "V(m/s)"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "记录时间(s)"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "a(m/s )"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t
Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdStart_Click()
Timer1.Enabled = True
Timer1.Interval = Val(TxtS.Text) * 1000
End Sub

Private Sub CmdStop_Click()
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
t = 0
LstV.Clear
LstX.Clear
v = 0
x = 0
End Sub

Private Sub Timer1_Timer()
Dim a, s, v0, v, x, t1
t = t + 1
s = Val(TxtS.Text)
a = Val(TxtA.Text)
v0 = Val(TxtV.Text)
t1 = t * s
v = v0 + a * t1
x = v0 * t1 + (a * (t1 ^ 2)) / 2
LstV.AddItem Str(v)
LstX.AddItem Str(x)
End Sub
