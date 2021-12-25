VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   8070
   ClientTop       =   4110
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5040
   Begin VB.FileListBox File1 
      Height          =   2070
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2190
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub


