VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3870
   ClientLeft      =   7950
   ClientTop       =   3330
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3870
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.VScrollBar VscZise 
      Height          =   2895
      Left            =   4080
      Max             =   100
      TabIndex        =   4
      Top             =   840
      Value           =   100
      Width           =   255
   End
   Begin VB.HScrollBar HsbRed 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin VB.HScrollBar HsbBlue 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin VB.HScrollBar HsbGreen 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   1
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Main"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      Height          =   180
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      Height          =   180
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Green"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Red"
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   270
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Iwidth As Integer
Dim Iheight As Integer


Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
Iwidth = Text1.Width
Iheight = Text1.Height
Label5.Caption = "100%"
End Sub

Private Sub HsbBlue_Change()
Text1.BackColor = RGB(HsbRed.Value, HsbGreen.Value, HsbBlue.Value)
End Sub

Private Sub HsbBlue_Scroll()
Text1.BackColor = RGB(HsbRed.Value, HsbGreen.Value, HsbBlue.Value)
End Sub

Private Sub HsbGreen_Change()
Text1.BackColor = RGB(HsbRed.Value, HsbGreen.Value, HsbBlue.Value)
End Sub

Private Sub HsbGreen_Scroll()
Text1.BackColor = RGB(HsbRed.Value, HsbGreen.Value, HsbBlue.Value)
End Sub

Private Sub HsbRed_Change()
Text1.BackColor = RGB(HsbRed.Value, HsbGreen.Value, HsbBlue.Value)
End Sub

Private Sub HsbRed_Scroll()
Text1.BackColor = RGB(HsbRed.Value, HsbGreen.Value, HsbBlue.Value)
End Sub

Private Sub VscZise_Change()
Text1.Width = Iwidth * (VscZise.Value / 100)
Text1.Height = Iheight * (VscZise.Value / 100)
Label5.Caption = VscZise.Value & "%"
End Sub

Private Sub VscZise_Scroll()
Text1.Width = Iwidth * (VscZise.Value / 100)
Text1.Height = Iheight * (VscZise.Value / 100)
Label5.Caption = VscZise.Value & "%"
End Sub
