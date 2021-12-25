VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   3615
   ClientTop       =   2865
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   1950
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim e As String
Dim a As Integer
e = Txt1.Text
a = PDTF(e)
If a = -1 Then
MsgBox "Error"
End If
End Sub

Function PDTF(b As String) As Integer
PDTF = 0
If b Like "#" Then
ElseIf b Like "##" Then
ElseIf b Like "###" Then
Else
PDTF = -1
End If
End Function
