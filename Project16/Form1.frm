VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Pic"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   1560
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Image1.Picture = LoadPicture(App.Path + "\2.gif")
End Sub

Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path + "\1.ico")
End Sub
