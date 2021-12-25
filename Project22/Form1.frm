VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":5C12
   ScaleHeight     =   4860
   ScaleWidth      =   7980
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   4080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "关机"
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   7560
      TabIndex        =   2
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label Label2 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " 开始 "
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path + "\1.ico")
Form1.Picture = LoadPicture(App.Path + "\.1.jpg")
End Sub

Private Sub Label1_Click()
If Label2.Visible = False Then
Label2.Visible = True
Label1.BorderStyle = 1
Label5.Visible = True
Else
Label2.Visible = False
Label1.BorderStyle = 0
Label5.Visible = False
End If
End Sub

Private Sub Label5_Click()
End
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
End Sub
