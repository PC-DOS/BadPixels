VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form11"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   MouseIcon       =   "Form11.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_DblClick()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
HWND_TOPMOST = -1
SWP_NOSIZE = &H1
SWP_NOREDRAW = &H8
SWP_NOMOVE = &H2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.WindowState = 2
Me.Left = 0
Me.Top = 0
End Sub
