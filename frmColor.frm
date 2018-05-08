VERSION 5.00
Begin VB.Form FrmColorPanel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Panel"
   ClientHeight    =   3480
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   9795
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar HScroll4 
      Height          =   285
      LargeChange     =   5
      Left            =   780
      Max             =   255
      Min             =   155
      TabIndex        =   54
      Top             =   2715
      Value           =   155
      Width           =   7800
   End
   Begin VB.Frame Frame1 
      Caption         =   "常用颜色"
      Height          =   885
      Left            =   120
      TabIndex        =   13
      Top             =   1785
      Width           =   6630
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   135
         Top             =   180
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   38
         Left            =   5925
         TabIndex        =   53
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   37
         Left            =   6240
         TabIndex        =   52
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   36
         Left            =   5925
         TabIndex        =   51
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   31
         Left            =   6240
         TabIndex        =   50
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   40
         Left            =   4995
         TabIndex        =   49
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   35
         Left            =   4680
         TabIndex        =   48
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   34
         Left            =   5625
         TabIndex        =   47
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   33
         Left            =   5610
         TabIndex        =   46
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   32
         Left            =   5310
         TabIndex        =   45
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   30
         Left            =   4695
         TabIndex        =   44
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   29
         Left            =   5295
         TabIndex        =   43
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   28
         Left            =   4980
         TabIndex        =   42
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   27
         Left            =   4380
         TabIndex        =   41
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   26
         Left            =   2430
         TabIndex        =   40
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   4380
         TabIndex        =   39
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   4050
         TabIndex        =   38
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   3720
         TabIndex        =   37
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   3720
         TabIndex        =   36
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   3405
         TabIndex        =   35
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   3075
         TabIndex        =   34
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   3075
         TabIndex        =   33
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   2775
         TabIndex        =   32
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   4035
         TabIndex        =   31
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   3420
         TabIndex        =   30
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   2775
         TabIndex        =   29
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   2430
         TabIndex        =   28
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   2100
         TabIndex        =   27
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   165
         TabIndex        =   26
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   2100
         TabIndex        =   25
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   1770
         TabIndex        =   24
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1455
         TabIndex        =   23
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   1455
         TabIndex        =   22
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   1140
         TabIndex        =   21
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   825
         TabIndex        =   20
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   825
         TabIndex        =   19
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   510
         TabIndex        =   18
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1770
         TabIndex        =   17
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1155
         TabIndex        =   16
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   15
         Top             =   210
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6705
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "0"
      Top             =   1320
      Width           =   1290
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "0"
      Top             =   705
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "0"
      Top             =   135
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   870
      Left            =   6810
      Picture         =   "frmColor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1170
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   360
      LargeChange     =   15
      Left            =   1425
      Max             =   255
      TabIndex        =   7
      Top             =   1320
      Width           =   5085
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   360
      LargeChange     =   15
      Left            =   1425
      Max             =   255
      TabIndex        =   5
      Top             =   720
      Width           =   5085
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   360
      LargeChange     =   15
      Left            =   1440
      Max             =   255
      TabIndex        =   3
      Top             =   150
      Width           =   5085
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "透明度"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8655
      TabIndex        =   55
      Top             =   2745
      Width           =   1080
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3060
      UseMnemonic     =   0   'False
      Width           =   9750
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1110
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Top             =   150
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Height          =   255
      Left            =   8190
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2370
      Left            =   8160
      TabIndex        =   0
      Top             =   285
      Width           =   1575
   End
   Begin VB.Menu copy 
      Caption         =   "Copyit"
      Visible         =   0   'False
      Begin VB.Menu CRGB 
         Caption         =   "复制RGB颜色值(&R)"
         Shortcut        =   ^C
      End
      Begin VB.Menu chex 
         Caption         =   "复制16进制颜色值(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu CWEB 
         Caption         =   "复制网页颜色值(&W)"
         Shortcut        =   ^W
      End
      Begin VB.Menu CPS 
         Caption         =   "复制PS颜色值(&P)"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "FrmColorPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pi1&
Dim red As Integer
Dim green As Integer
Dim blue As Integer
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Sub chex_Click()
On Error Resume Next
Clipboard.Clear
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Clipboard.SetText "0x" & Hex(pi1&)
MsgBox "Copy Successfully", vbExclamation, "Info"
End Sub
Private Sub Command1_Click()
Me.Tag = Label1.BackColor
Me.Hide
End Sub
Private Sub CPS_Click()
On Error Resume Next
Clipboard.Clear
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Clipboard.SetText "#000000" & Hex(pi1&)
Case 1
Clipboard.SetText "#00000" & Hex(pi1&)
Case 2
Clipboard.SetText "#0000" & Hex(pi1&)
Case 3
Clipboard.SetText "#000" & Hex(pi1&)
Case 4
Clipboard.SetText "#00" & Hex(pi1&)
Case 5
Clipboard.SetText "#0" & Hex(pi1&)
Case 6
Clipboard.SetText "#" & Hex(pi1&)
End Select
MsgBox "Copy Successfully", vbExclamation, "Info"
End Sub
Private Sub CRGB_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Text1.Text & "," & Text2.Text & "," & Text3.Text
MsgBox "Copy Successfully", vbExclamation, "Info"
End Sub
Private Sub CWEB_Click()
On Error Resume Next
Clipboard.Clear
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Clipboard.SetText "&H00000000" & Hex(pi1&) & "&"
Case 1
Clipboard.SetText "&H0000000" & Hex(pi1&) & "&"
Case 2
Clipboard.SetText "&H000000" & Hex(pi1&) & "&"
Case 3
Clipboard.SetText "&H00000" & Hex(pi1&) & "&"
Case 4
Clipboard.SetText "&H0000" & Hex(pi1&) & "&"
Case 5
Clipboard.SetText "&H000" & Hex(pi1&) & "&"
Case 6
Clipboard.SetText "&H00" & Hex(pi1&) & "&"
Case 7
Clipboard.SetText "&H0" & Hex(pi1&) & "&"
Case 8
Clipboard.SetText "&H" & Hex(pi1&) & "&"
End Select
MsgBox "Copy Successfully", vbExclamation, "Info"
End Sub
Private Sub Form_Load()
With Text1
.Text = "0"
.MaxLength = 3
End With
With Text2
.Text = "0"
.MaxLength = 3
End With
With Text3
.Text = "0"
.MaxLength = 3
End With
With Me.HScroll1
.Min = 0
.Max = 255
End With
With Me.HScroll2
.Min = 0
.Max = 255
End With
With Me.HScroll3
.Min = 0
.Max = 255
End With
With Me.HScroll4
.Max = 255
.Min = 155
.LargeChange = 5
.SmallChange = 1
.Value = 230
End With
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, Me.HScroll4.Value, LWA_ALPHA
Label12.Caption = Me.HScroll4.Value
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Text1.Text = Me.HScroll1.Value
Text2.Text = Me.HScroll2.Value
Text3.Text = Me.HScroll3.Value
End Sub
Private Sub HScroll2_Change()
On Error Resume Next
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Text1.Text = Me.HScroll1.Value
Text2.Text = Me.HScroll2.Value
Text3.Text = Me.HScroll3.Value
End Sub
Private Sub HScroll3_Change()
On Error Resume Next
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Text1.Text = Me.HScroll1.Value
Text2.Text = Me.HScroll2.Value
Text3.Text = Me.HScroll3.Value
End Sub
Private Sub HScroll4_Change()
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, Me.HScroll4.Value, LWA_ALPHA
Label12.Caption = Me.HScroll4.Value
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Me.copy
End If
End Sub
Private Sub Label12_Click()
On Error Resume Next
Dim alp As Integer
Dim oldalp As Integer
Dim rtn As Long
oldalp = Label12.Caption
alp = Val(InputBox$("请输入透明度" & vbCrLf & "范围:155-255", "Alpha", 230))
If Val(alp) = 0 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, oldalp, LWA_ALPHA
Me.HScroll1.Value = oldalp
Label12.Caption = Me.HScroll1.Value
Exit Sub
End If
If 155 <= alp And alp <= 255 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, alp, LWA_ALPHA
Me.HScroll1.Value = alp
Label12.Caption = Me.HScroll1.Value
Else
MsgBox "无效透明度数值", vbCritical, "Error"
End If
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Me.copy
End If
End Sub
Private Sub Label7_Click(Index As Integer)
Dim a
Dim r As Integer
Dim g As Integer
Dim b As Integer
Shape1.Visible = True
Shape1.Left = Label7(Index).Left
Shape1.Top = Label7(Index).Top
a = Label7(Index).BackColor
r = a Mod 256
g = ((a And &HFF00) / 256&) Mod 256&
b = (a And &HFF0000) / 65536
Me.HScroll1.Value = r
Me.HScroll2.Value = g
Me.HScroll3.Value = b
Text1.Text = Me.HScroll1.Value
Text2.Text = Me.HScroll2.Value
Text3.Text = Me.HScroll3.Value
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Label1.BackColor = a
End Sub
Private Sub Text2_Change()
On Error Resume Next
If Val(Text2.Text) >= 0 And Val(Text2.Text) <= 255 Then
Me.HScroll2.Value = Val(Text2.Text)
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Else
Exit Sub
End If
End Sub
Private Sub Text1_Change()
On Error Resume Next
If Val(Text1.Text) >= 0 And Val(Text1.Text) <= 255 Then
Me.HScroll1.Value = Val(Text1.Text)
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Else
Exit Sub
End If
End Sub
Private Sub Text3_Change()
On Error Resume Next
If Val(Text3.Text) >= 0 And Val(Text3.Text) <= 255 Then
Me.HScroll3.Value = Val(Text3.Text)
Label1.BackColor = RGB(Me.HScroll1.Value, Me.HScroll2.Value, Me.HScroll3.Value)
pi1& = Label1.BackColor
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000000" & Hex(pi1&) & "&"
Case 1
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "RGB: " & Me.HScroll1.Value & "," & Me.HScroll2.Value & "," & Me.HScroll3.Value & "  Hex: 0x" & Hex(pi1&) & "  Web: &H" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 0
Label6.Caption = Label6.Caption & "  PS:#000000" & Hex(pi1&)
Case 1
Label6.Caption = Label6.Caption & "  PS:#00000" & Hex(pi1&)
Case 2
Label6.Caption = Label6.Caption & "  PS:#0000" & Hex(pi1&)
Case 3
Label6.Caption = Label6.Caption & "  PS:#000" & Hex(pi1&)
Case 4
Label6.Caption = Label6.Caption & "  PS:#00" & Hex(pi1&)
Case 5
Label6.Caption = Label6.Caption & "  PS:#0" & Hex(pi1&)
Case 6
Label6.Caption = Label6.Caption & "  PS:#" & Hex(pi1&)
End Select
Else
Exit Sub
End If
End Sub
