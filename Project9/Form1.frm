VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4215
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Btn"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Index > 0 Then
Unload Command1(Index)
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mouse As POINTAPI
GetCursorPos mouse
If Button = 1 Then
i = Command1.ubound + 1
Load Command1(i)
Command1(i).Left = Val(mouse.X) * 15 - Me.Left - 100
Command1(i).Top = Val(mouse.Y) * 15 - Me.Top - 500
Command1(i).Caption = "Btn" & Str(i)
Command1(i).Visible = True
End If
End Sub


