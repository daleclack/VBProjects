VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3015
   ClientLeft      =   8070
   ClientTop       =   4110
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.TextBox TxtC 
      Height          =   1575
      Left            =   720
      TabIndex        =   2
      Text            =   "There is a pen"
      Top             =   240
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Modern"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Arial"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Option1_Click()
TxtC.FontName = "Arial"
End Sub

Private Sub Option2_Click()
TxtC.FontName = "Modern"
End Sub
