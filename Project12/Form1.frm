VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   8370
   ClientTop       =   4410
   ClientWidth     =   3855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   3855
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox TxtY 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox TxtX 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim x!
Dim y!
Dim i As String
x = TxtX.Text
Select Case x
Case Is < 0
y = x ^ 2
Case Is = 0
y = 1.25
Case Is > 0
y = Sqr(x)
End Select
TxtY.Text = Str(y)
If TxtX.Text = "" Then
i = MsgBox("x值不能为空！", vbOKOnly, "提示")
End If
End Sub

Private Sub TxtX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
End If
End Sub
