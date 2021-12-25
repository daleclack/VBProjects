VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   900
   ClientLeft      =   4020
   ClientTop       =   3645
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Days"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Timer1_Timer()
Dim a, b
If Year(Date) < 2020 Then
a = (2020 - Year(Date)) * 365 + 159
Select Case Month(Date)
Case 1
b = a - Day(Date)
Case 2
b = a - 31 - Day(Date)
Case 3
b = a - 59 - Day(Date)
Case 4
b = a - 90 - Day(Date)
Case 5
b = a - 120 - Day(Date)
Case 6
b = a - 151 - Day(Date)
Case 7
b = a - 181 - Day(Date)
Case 8
b = a - 212 - Day(Date)
Case 9
b = a - 243 - Day(Date)
Case 10
b = a - 273 - Day(Date)
Case 11
b = a - 304 - Day(Date)
Case 12
b = a - 334 - Day(Date)
End Select
ElseIf Year(Date) = 2020 Then
Select Case Month(Date)
Case 1
b = 159 - Day(Date)
Case 2
b = 128 - Day(Date)
Case 3
b = 99 - Day(Date)
Case 4
b = 68 - Day(Date)
Case 5
b = 38 - Day(Date)
Case 6
b = 7 - Day(Date)
End Select
End If
Label2.Caption = Str(b)
End Sub

