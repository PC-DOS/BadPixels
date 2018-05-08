VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form12"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   LinkTopic       =   "Form12"
   MouseIcon       =   "Form12.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   5700
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_Activate()
On Error Resume Next
Dim a
Cls
For a = 0 To Me.ScaleWidth Step 30
Me.Line (a, 0)-(a, Me.Height)
Next
End Sub
Private Sub Form_Click()
On Error Resume Next
Dim a
Cls
For a = 0 To Me.ScaleWidth Step 30
Me.Line (a, 0)-(a, Me.Height)
Next
End Sub
Private Sub Form_Load()
On Error Resume Next
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
With Me
.Caption = ""
.Height = Screen.Height
.Width = Screen.Width
.Left = 0
.Top = 0
.WindowState = 2
.BackColor = vbWhite
End With
End Sub
Private Sub Form_DblClick()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub
