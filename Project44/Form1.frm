VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "学生成绩管理"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6585
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   240
      ScaleHeight     =   3315
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
   End
   Begin VB.CommandButton CmdLu 
      Caption         =   "录入"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "请点击‘成绩录入’按钮，进行学生成绩录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "学生成绩管理系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e
Private Sub CmdLu_Click()
Dim a, a1, b, b1, c, c1, d, i
a1 = 0: b1 = 0: c1 = 0: e = 1
Do While e <= 5
  a = InputBox("课程1", "输入成绩")
  b = InputBox("课程2", "输入成绩")
  c = InputBox("课程3", "输入成绩")
    If IsNumeric(a) = True And IsNumeric(b) = True And IsNumeric(c) = True Then
      a = Val(a): b = Val(b): c = Val(c)
      d = Int((a + b + c) / 3)
      Pic1.Print Space(3) + Str(e) + Space(7) + Str(a) + Space(7) + Str(b) + Space(7) + Str(c) + Space(7) + Str(d)
      a1 = a1 + a: b1 = b1 + b: c1 = c1 + c
      e = e + 1
    Else
      i = MsgBox("Error", , "提示")
    End If
Loop
Pic1.Print Space(1) + "各科平均分" + Space(1) + Str(Int(a1 / 5)) + Space(7) + Str(Int(b1 / 5)) + Space(7) + Str(Int(c1 / 5))
End Sub

Private Sub Form_Load()
Pic1.Print Space(3) + "学号" + Space(5) + "课程1" + Space(5) + "课程2" + Space(5) + "课程3" + Space(5) + "平均成绩"
End Sub

