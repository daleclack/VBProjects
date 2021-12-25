VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   690
   ClientLeft      =   8370
   ClientTop       =   5625
   ClientWidth     =   3480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":25CA
   ScaleHeight     =   690
   ScaleWidth      =   3480
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdGO 
      Caption         =   "GO"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdGO_Click()
Dialog.Show
Dialog.TxtPass.Text = ""
End Sub
