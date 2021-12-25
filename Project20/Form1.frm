VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4290
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "无边框"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "有边框"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd1_Click()
Label1.BorderStyle = 1
End Sub

Private Sub Command1_Click()
Label1.BorderStyle = 0
End Sub
