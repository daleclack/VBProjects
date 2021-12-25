VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7920
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdClear 
      Caption         =   "清除"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "显示"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim mf(), n, x As Integer, y As Integer, k As Integer

Private Sub CmdClear_Click()
Pic1.Cls
End Sub

Private Sub CmdShow_Click()
n = InputBox("输入阶数", "提示")
If IsNumeric(n) = True Then
  ReDim mf(n, n)
  For x = 1 To n
    For y = 1 To n
      mf(x, y) = 0
    Next y
  Next x
  y = Int(n / 2) + 1
  mf(1, y) = 1
  For k = 2 To n * n
    x = x - 1
    y = y + 1
    If (x < 1 And y > n) Then
      x = x + 2
      y = y - 1
    Else
      If x < 1 Then x = n
      If y > n Then y = 1
    End If
    If mf(x, y) = 0 Then
      mf(x, y) = k
    Else
      x = x + 2
      y = y - 1
      mf(x, y) = k
    End If
  Next k
  For x = 1 To n
    Pic1.Print
    For y = 1 To n
      Pic1.Print Tab(5 * y); mf(x, y);
    Next y
  Next x
Else
  MsgBox "Error"
End If
End Sub

Private Sub Form_Load()
Pic1.FontSize = 10
End Sub

