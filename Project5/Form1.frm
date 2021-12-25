VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Cmdjisuan 
      Caption         =   "计算"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Txthe 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "计算结果："
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdjisuan_Click()
Dim sum As Integer
Dim i As Integer
sum = 0
For i = 1 To 100
sum = sum + i
Next i
Txthe.Text = sum
End Sub
