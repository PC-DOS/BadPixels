VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmOpenFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请选择要打开的JPG/BMP/GIF格式图片文件"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   3045
      Left            =   12945
      TabIndex        =   15
      Top             =   6000
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   705
      TabIndex        =   14
      Top             =   9105
      Width           =   12225
   End
   Begin VB.PictureBox Picture3 
      Height          =   3075
      Left            =   720
      ScaleHeight     =   3015
      ScaleWidth      =   12150
      TabIndex        =   13
      Top             =   6015
      Width           =   12210
      Begin VB.Image Image1 
         Height          =   3000
         Left            =   -15
         Top             =   0
         Width           =   12150
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   6990
      ScaleHeight     =   5010
      ScaleWidth      =   6375
      TabIndex        =   10
      Top             =   15
      Width           =   6375
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   4965
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   6330
         ExtentX         =   11165
         ExtentY         =   8758
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   12225
      TabIndex        =   9
      Top             =   5400
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开(&O)"
      Default         =   -1  'True
      Height          =   420
      Left            =   11040
      TabIndex        =   8
      Top             =   5400
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   690
      TabIndex        =   7
      Top             =   5040
      Width           =   12660
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   2895
      TabIndex        =   4
      Top             =   270
      Width           =   4080
   End
   Begin VB.DirListBox Dir1 
      Height          =   4290
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   2865
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Top             =   15
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "預覽"
      Height          =   180
      Left            =   45
      TabIndex        =   12
      Top             =   6000
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件名"
      Height          =   180
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   5100
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件"
      Height          =   180
      Index           =   2
      Left            =   2895
      TabIndex        =   5
      Top             =   45
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件夹"
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   2
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "驱动器"
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "frmOpenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Dim tmp As Long
Public pth As String
Private Sub Form_Activate()
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.Path
End With
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub file1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Me.WebBrowser1.Navigate "About:Operations Are Not Allowed "
End Sub
Private Sub text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub WebBrowser1_GotFocus()
On Error Resume Next
Me.Dir1.SetFocus
On Error Resume Next
Me.WebBrowser1.Navigate "About:Operations Are Not Allowed "
End Sub
Private Sub Command1_Click()
On Error GoTo ep
Me.Dir1.Enabled = False
Me.File1.Enabled = False
Me.Text1.Enabled = False
Me.Command1.Enabled = False
Me.Drive1.Enabled = False
Me.Command2.Enabled = False
If Trim(Text1.Text) = "" Then
MsgBox "您没有选择一个有效的文件!", vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
End If
If Right(Dir1.Path, 1) = "\" Then
pth = Dir1.Path & Trim(Text1.Text)
Else
pth = Dir1.Path & "\" & Trim(Text1.Text)
End If
If Dir(pth) = "" Then
MsgBox "Windows找不到文件 " & pth & " 请确定文件是否存在,如果您想亲自查找,请使用Windows搜索", vbExclamation, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
End If
If 1 = 245 Then
tmp = 245
If tmp = 0 Then
MsgBox pth & "不包含图标,请选择一个包含图标的文件", vbExclamation, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Me.Hide
With Form1
.Tag = Me.pth
.SetFocus
End With
End If
End If
With Form1
.Tag = Me.pth
End With
Me.Hide
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
Form1.Tag = ""
Form1.SetFocus
End Sub
Private Sub Dir1_Change()
On Error GoTo ep
Drive1.Drive = Left$(Dir1.Path, 2)
With File1
.Pattern = "*.jpg;*.bmp;*.gif;*.ico;*.jpeg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.WebBrowser1
.Navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.Path
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
With Me.WebBrowser1
.Navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.Path
End With
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Dir1.Path = Drive1.Drive
With File1
.Pattern = "*.jpg;*.bmp;*.gif;*.ico;*.jpeg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.WebBrowser1
.Navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.Path
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Drive1.Drive = "C:"
Dir1.Path = Drive1.Drive
With File1
.Pattern = "*.jpg;*.bmp;*.gif;*.ico;*.jpeg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.WebBrowser1
.Navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.Path
End With
End Sub
Private Sub Drive1_GotFocus()
On Error Resume Next
Me.Command1.Default = False
Me.Command2.Cancel = False
End Sub
Private Sub Drive1_LostFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub Dir1_GotFocus()
On Error Resume Next
Me.Command1.Default = False
Me.Command2.Cancel = True
End Sub
Private Sub Dir1_LostFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub File1_GotFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub File1_LostFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub File1_Click()
On Error Resume Next
If File1.ListIndex >= 0 Then
Me.Text1.Text = File1.List(File1.ListIndex)
End If
Dim lpPath As String
If Right(Dir1.Path, 1) = "\" Then
lpPath = Dir1.Path & Trim(Text1.Text)
Else
lpPath = Dir1.Path & "\" & Trim(Text1.Text)
End If
Image1.Picture = LoadPicture(lpPath)
Picture3.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.Top = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.Picture1.Height And Me.Image1.Width <= Me.Picture1.Width Then
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Exit Sub
End If
If Me.Image1.Width <= Me.Picture3.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.Picture3.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.Picture3.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.Top = Me.Picture3.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.Picture3.Height
Me.HScroll1.Max = Me.Image1.Width - Me.Picture3.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub File1_DblClick()
On Error GoTo ep
Me.Dir1.Enabled = False
Me.File1.Enabled = False
Me.Text1.Enabled = False
Me.Command1.Enabled = False
Me.Drive1.Enabled = False
Me.Command2.Enabled = False
If Trim(Text1.Text) = "" Then
MsgBox "您没有选择一个有效的文件!", vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
End If
If Right(Dir1.Path, 1) = "\" Then
pth = Dir1.Path & Trim(Text1.Text)
Else
pth = Dir1.Path & "\" & Trim(Text1.Text)
End If
If Dir(pth) = "" Then
MsgBox "Windows找不到文件 " & pth & " 请确定文件是否存在,如果您想亲自查找,请使用Windows搜索", vbExclamation, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
End If
If 1 = 245 Then
tmp = 245
If tmp = 0 Then
MsgBox pth & "不包含图标,请选择一个包含图标的文件", vbExclamation, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Me.Hide
With Form1
.Tag = Me.pth
.SetFocus
End With
End If
End If
With Form1
.Tag = Me.pth
End With
Me.Hide
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Text1.Enabled = True
Me.Command1.Enabled = True
Me.Drive1.Enabled = True
Me.Command2.Enabled = True
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
With File1
.Refresh
End With
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
With Me
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
.KeyPreview = True
End With
With File1
.Path = Me.Dir1.Path
.Pattern = "*.jpg;*.bmp;*.gif;*.ico;*.jpeg"
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
Me.Text1.Text = ""
With Me.WebBrowser1
.Navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.Path
End With
End Sub
Private Sub VScroll1_Change()
On Error Resume Next
Image1.Top = -VScroll1.Value
If -VScroll1.Value > 0 Then
Image1.Top = VScroll1.Value
End If
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Image1.Left = -HScroll1.Value
If -HScroll1.Value > 0 Then
Image1.Left = HScroll1.Value
End If
End Sub
