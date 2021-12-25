VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   1770
   ClientLeft      =   7695
   ClientTop       =   4935
   ClientWidth     =   5160
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "+"
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   3015
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
