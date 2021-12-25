VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   13995
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   855
      Left            =   8760
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   855
      Left            =   8640
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.PictureBox Pic1 
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, b
Private Sub Command1_Click()
Do
b = b + 1
If b Mod 3 <> 0 Then
Pic1.Print b;
i = i + 1
If i Mod 5 = 0 Then Pic1.Print
End If
Loop While i < 20
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
i = 0
b = 1
End Sub
