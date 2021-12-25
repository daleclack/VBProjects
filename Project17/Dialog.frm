VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   0  'None
   Caption         =   "对话框标题"
   ClientHeight    =   1560
   ClientLeft      =   7650
   ClientTop       =   5160
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPass 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2415
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
Dialog.Hide
Form1.Hide
Form2.Show
Else
Label1.Caption = "Password wrong!"
End If
End Sub

Private Sub TxtPass_Change()
Label1.Caption = ""
End Sub
