VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   8565
   ClientTop       =   5760
   ClientWidth     =   3645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":25CA
   ScaleHeight     =   780
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "GO"
      Default         =   -1  'True
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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

Private Sub CmdGo_Click()
Dialog.Show
Dialog.TxtPass.Text = ""
End Sub

Private Sub Form_Load()
'System Tray Begin
With nfIconData
.hWnd = Me.hWnd
.uID = Me.Icon
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle
.szTip = App.Title + "(版本 " & App.Major & "." & App.Minor & "." & App.Revision & ")" & vbNullChar
.cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
Call Shell_NotifyIcon(NIM_ADD, nfIconData)
'System Tray End
Form1.Icon = LoadPicture(App.Path + "\1.ico")
Form1.Picture = LoadPicture(App.Path + "\1.ico")
Form1.Caption = ""
Dialog.Icon = LoadPicture(App.Path + "\1.ico")
Dialog.Caption = "Password"
Form2.Caption = ""
Form2.Icon = LoadPicture(App.Path + "\1.ico")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
Case WM_LBUTTONUP
'MsgBox "请用鼠标右键点击图标!", vbInformation, "实时播音专家"
'单击左键，显示窗体
ShowWindow Me.hWnd, SW_RESTORE
'下面两句的目的是把窗口显示在窗口最顶层
'Me.Show
'Me.SetFocus
'' Case WM_RBUTTONUP
'' PopupMenu MenuTray '如果是在系统Tray图标上点右键，则弹出菜单MenuTray
'' Case WM_MOUSEMOVE
'' Case WM_LBUTTONDOWN
'' Case WM_LBUTTONDBLCLK
'' Case WM_RBUTTONDOWN
'' Case WM_RBUTTONDBLCLK
'' Case Else
End Select
End Sub
