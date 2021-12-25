VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   8865
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Cmdjisuan 
      Caption         =   "计算"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Pic1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim a(10) As Single, i As Single, sum As Single, c As Integer, p As Single
Private Sub Cmdjisuan_Click()
    p = sum / 10
     For i = 1 To 10
          If a(i) > p Then c = c + 1
     Next i
    Label1.Caption = "总数" + Str(sum) + "  " + "平均" + Str(p) + "  " + "超过平均的人数" + Str(c)
End Sub

Private Sub Form_Load()
i = 0
c = 0
sum = 0
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then
          i = i + 1
          If i = 10 Then Text1.Enabled = False: Cmdjisuan.Enabled = True
          a(i) = Text1.Text
          Pic1.Print a(i); Spc(1);
          sum = sum + a(i)
          Text1.Text = ""
      End If
End Sub
