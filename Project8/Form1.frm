VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   2190
   ScaleWidth      =   4020
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton CmdGO 
      Caption         =   "GO"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEXIT_Click()
Unload Me
End
End Sub

Private Sub CmdGO_Click()
Dialog.Show
Dialog.TxtPass.Text = ""
End Sub
