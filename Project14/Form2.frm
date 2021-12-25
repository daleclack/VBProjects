VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   8070
   ClientTop       =   3795
   ClientWidth     =   4560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CheckBox Check2 
      Caption         =   "斜体"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "下划线"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "黑体"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Times New Romen"
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Modern"
      Height          =   300
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Arial"
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Learning Visual Basic "
      Height          =   180
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   1980
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If Check1.Value = 1 Then
Lbl1.FontBold = True
Else
Lbl1.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Lbl1.FontItalic = True
Else
Lbl1.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Lbl1.FontUnderline = True
Else
Lbl1.FontUnderline = False
End If
End Sub

Private Sub CmdExit_Click()
Me.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path + "\1.ico")
Form1.Picture = LoadPicture(App.Path + "\1.ico")
Dialog.Icon = LoadPicture(App.Path + "\1.ico")
Form2.Icon = LoadPicture(App.Path + "\1.ico")
End Sub

Private Sub Option1_Click()
Lbl1.FontName = "Arial"
End Sub

Private Sub Option2_Click()
Lbl1.FontName = "Modern"
End Sub

Private Sub Option3_Click()
Lbl1.FontName = "Times New Roman"
End Sub
