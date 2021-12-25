VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   5205
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   975
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   975
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox Pic1 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b
For a = 1 To 16
If a > 4 And a < 12 Then
 Pic1.Print Tab(17 - a);
   For b = 1 To a
   Pic1.Print "*" + " ";
   Next b
   Pic1.Print
 Else
   Pic1.Print
 End If
Next a
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Pic1.FontSize = 10
Command1.FontSize = 10
Command2.FontSize = 10
End Sub
