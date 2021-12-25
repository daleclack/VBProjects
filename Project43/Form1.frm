VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   7050
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton CmdShow 
      Caption         =   "œ‘ æ"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.PictureBox Pic1 
      Height          =   4215
      Left            =   240
      ScaleHeight     =   4155
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdShow_Click()
Dim a, b, c
c = 1
For a = 1 To 5
Pic1.Print Tab(5);
  For b = 1 To 5
  Pic1.Print b * c;
  Next b
  c = c + 1
  Pic1.Print
Next a
End Sub

Private Sub Form_Load()
Pic1.FontSize = 10
CmdShow.FontSize = 10
End Sub
