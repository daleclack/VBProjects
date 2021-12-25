VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000014&
   ClientHeight    =   930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3000
   HasDC           =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   3000
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton CmdGO 
      Caption         =   "GO"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
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
