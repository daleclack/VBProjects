VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3300
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdGO 
      Caption         =   "GO"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End
End Sub

Private Sub CmdGO_Click()
Dialog.Show
Dialog.TxtPass.Text = ""
End Sub
