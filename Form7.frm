VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   LinkTopic       =   "Form7"
   MouseIcon       =   "Form7.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6990
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   3840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "10twip/ms"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "50twip/ms"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "100twip/ms"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   735
      Left            =   0
      Top             =   2880
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   735
      Left            =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tms As Integer
Dim tmss As Integer
Dim tmsss As Integer
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
Me.Height = Screen.Height
Me.Width = Screen.Width
tms = 0
tmss = 0
tmsss = 0
Shape1.Left = 0
Shape2.Left = 0
Shape2.Left = 0
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
tmss = tmss + 100
If tmss <= Screen.Width Then
Shape1.Left = tmss
Else
tmss = 0
Shape1.Left = 0
End If
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
tms = tms + 50
If tms <= Screen.Width Then
Shape2.Left = tms
Else
tms = 0
Shape2.Left = 0
End If
End Sub
Private Sub Timer3_Timer()
On Error Resume Next
tmsss = tmsss + 10
If tmsss <= Screen.Width Then
Shape3.Left = tmsss
Else
tmsss = 0
Shape3.Left = 0
End If
End Sub
