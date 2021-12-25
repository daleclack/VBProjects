VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   2850
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1860
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0043
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3000
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()
Dim i As Integer
i = List1.ListIndex
Text1.Text = List1.List(i)
End Sub
