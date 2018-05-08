VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   LinkTopic       =   "Form10"
   MouseIcon       =   "Form10.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   4965
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "显示器测试示例文本ABCDEFGHIJKLMNOPQR STUVWXYZabcdefghij klmnopqrstuvwxyz12 34567890-=_+|\!@#$ %^&*(){}[]:"";'<>,. ?/`~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Index           =   4
      Left            =   4950
      TabIndex        =   4
      Top             =   3345
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "显示器测试示例文本ABCDEFGHIJKLMNOPQR STUVWXYZabcdefghij klmnopqrstuvwxyz12 34567890-=_+|\!@#$ %^&*(){}[]:"";'<>,. ?/`~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   3495
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "显示器测试示例文本ABCDEFGHIJKLMNOPQR STUVWXYZabcdefghij klmnopqrstuvwxyz12 34567890-=_+|\!@#$ %^&*(){}[]:"";'<>,. ?/`~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Index           =   2
      Left            =   5085
      TabIndex        =   2
      Top             =   0
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "显示器测试示例文本ABCDEFGHIJKLMNOPQR STUVWXYZabcdefghij klmnopqrstuvwxyz12 34567890-=_+|\!@#$ %^&*(){}[]:"";'<>,. ?/`~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Index           =   1
      Left            =   -75
      TabIndex        =   1
      Top             =   -15
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "显示器测试示例文本ABCDEFGHIJKLMNOPQR STUVWXYZabcdefghij klmnopqrstuvwxyz12 34567890-=_+|\!@#$ %^&*(){}[]:"";'<>,. ?/`~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   3120
      Width           =   1890
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_DblClick()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
With Me
.WindowState = 2
.Left = 0
.Top = 0
.Width = Screen.Width
.Height = Screen.Height
.BackColor = RGB(0, 0, 0)
End With
With Label1(0)
.Top = Screen.Height / 2 - Label1(0).Height / 2
.Left = Screen.Width / 2 - Label1(0).Width / 2
End With
With Label1(1)
.Top = 0
.Left = 0
End With
With Label1(2)
.Top = 0
.Left = Me.Width - Label1(2).Width
End With
With Label1(3)
.Top = Me.Height - Label1(0).Height
.Left = 0
End With
With Label1(4)
.Top = Me.Height - Label1(4).Height
.Left = Me.Width - Label1(4).Width
End With
HWND_TOPMOST = -1
SWP_NOSIZE = &H1
SWP_NOREDRAW = &H8
SWP_NOMOVE = &H2
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Label1_DblClick(Index As Integer)
Unload Me
End Sub
