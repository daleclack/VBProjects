VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   555
   ClientLeft      =   3705
   ClientTop       =   2685
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   2325
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text Like "#" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "##" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "###" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "####" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "#####" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "######" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "#######" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "########" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "#########" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "##########" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "###########" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "############" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "#############" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "##############" = True Then
MsgBox "ok"
ElseIf Text1.Text Like "###############" = True Then
MsgBox "ok"
Else
MsgBox "Error"
End If
End Sub
