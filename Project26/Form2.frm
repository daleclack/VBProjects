VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   1710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Shut Off"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "3"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "2"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Label1_Click()
Form3.Show
End Sub

Private Sub Label2_Click()
Form4.Show
End Sub

Private Sub Label3_Click()
Form5.Show
End Sub
