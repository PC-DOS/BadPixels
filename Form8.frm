VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   MouseIcon       =   "Form8.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_DblClick()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
HWND_TOPMOST = -1
SWP_NOSIZE = &H1
SWP_NOREDRAW = &H8
SWP_NOMOVE = &H2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
Shape1.Left = 0
Shape1.Width = 0
Shape1.Height = Screen.Height
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Shape1.Width = Shape1.Width + 50
If Shape1.Width >= Screen.Width Then
If Shape1.FillStyle <= 6 Then
Shape1.Width = 0
Shape1.FillStyle = Shape1.FillStyle + 1
Shape1.Width = Shape1.Width + 50
Else
Shape1.FillStyle = 1
Shape1.Width = Shape1.Width + 50
End If
End If
End Sub
