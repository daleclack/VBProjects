VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPass 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "                           "
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   2430
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
Form1.Show
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
