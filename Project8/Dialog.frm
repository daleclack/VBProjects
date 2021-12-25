VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPass 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   4560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   180
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub OKButton_Click()
If TxtPass.Text = "daleclack" Then
Me.Hide
Form1.Hide
Form2.Show
Else
Label1.Caption = "Password wrong!"
End If
End Sub

Private Sub TxtPass_Change()
Label1.Caption = ""
End Sub

Private Sub TxtPass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
OKButton_Click
End If
End Sub
