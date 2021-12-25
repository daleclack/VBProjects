VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4125
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6465
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Menu menu 
      Caption         =   "menu"
      Begin VB.Menu f1 
         Caption         =   "1"
      End
      Begin VB.Menu f2 
         Caption         =   "2"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub f1_Click()
Form1.Show
End Sub

Private Sub f2_Click()
Form2.Show
End Sub

