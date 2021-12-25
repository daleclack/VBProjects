VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   4785
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label2 
      Caption         =   "±¾³ÌĞòÓÃÓÚ½ØÈ¡¼üÅÌ¼üÂë"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Esc=ÍË³ö"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static i As Integer
i = i + 1
If i >= 4 Then
Me.Cls
i = 0
End If
CurrentX = 100
If CurrentY < 800 Then
CurrentY = 800
End If
Select Case Shift
Case 0:
Print Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 1:
Print "Shift+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 2:
Print "Ctrl+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 3:
Print "Shift + Ctrl+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 4:
Print "Alt+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 5:
Print "Shift + Alt+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 6:
Print "Ctrl + Alt+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
Case 7:
Print "Ctrl + Alt + Shift+" + "" + Chr(KeyCode) + "" + "¼üÂë:" + Format(KeyCode) + ";"
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
Form1.Show
End If
End Sub
