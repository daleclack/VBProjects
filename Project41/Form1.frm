VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   6090
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Cmd2 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "计算"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, y, t
Private Sub Cmd1_Click()
x = InputBox("输入X", "X")
y = InputBox("输入Y", "Y")
If IsNumeric(x) = True And IsNumeric(y) = True Then
Label1.Caption = Str(x) & " 和 " & Str(y)
 If x < y Then
 t = y: y = x: x = t
 End If
t = x Mod y
Do While t <> 0
 x = y
 y = t
 t = x Mod y
Loop
Label1.Caption = Label1.Caption & " :" & Str(y)
Else
Label1.Caption = ""
MsgBox "Error"
End If
End Sub

Private Sub Cmd2_Click()
End
End Sub

Private Sub Form_Load()
Label1.FontSize = 10
Cmd1.FontSize = 10
End Sub
