VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LCD ����/����/���� ���Գ���(LCD Test) Version 1.0 - PC-DOS Workshop"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11280
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   5085
      Left            =   9195
      TabIndex        =   41
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Command15 
         Caption         =   "�����߻��߲���(&L)"
         Height          =   495
         Left            =   60
         TabIndex        =   49
         Top             =   4455
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
         Caption         =   "����������ק����(&F)"
         Height          =   495
         Left            =   60
         TabIndex        =   48
         Top             =   255
         Width           =   1935
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ͼ�λ�ͼ����(&D)"
         Height          =   495
         Left            =   60
         TabIndex        =   47
         Top             =   855
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         Caption         =   "ɫ��ͨ͸�Բ���(&C)"
         Height          =   495
         Left            =   60
         TabIndex        =   46
         Top             =   1455
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         Caption         =   "�ҶȲ���(&G)"
         Height          =   495
         Left            =   60
         TabIndex        =   45
         Top             =   2055
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         Caption         =   "������ʾ����(&O)"
         Height          =   495
         Left            =   60
         TabIndex        =   44
         Top             =   2655
         Width           =   1935
      End
      Begin VB.CommandButton Command13 
         Caption         =   "ͼƬ����(&P)"
         Height          =   495
         Left            =   60
         TabIndex        =   43
         Top             =   3255
         Width           =   1935
      End
      Begin VB.CommandButton Command14 
         Caption         =   "��ֱ�߻��߲���(&L)"
         Height          =   495
         Left            =   60
         TabIndex        =   42
         Top             =   3855
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command11 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&E)"
      Height          =   495
      Left            =   9240
      TabIndex        =   40
      Top             =   5160
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0442
      Left            =   6705
      List            =   "Form1.frx":0473
      TabIndex        =   39
      Text            =   "1000"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ12"
      Height          =   495
      Index           =   11
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ11"
      Height          =   495
      Index           =   10
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ6"
      Height          =   495
      Index           =   5
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ5"
      Height          =   495
      Index           =   4
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ4"
      Height          =   495
      Index           =   3
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ3"
      Height          =   495
      Index           =   2
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�Զ�����ɫ2"
      Height          =   495
      Index           =   1
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "�û��Զ�����ɫ(�����ťָ����ɫ)"
      Height          =   5655
      Left            =   4785
      TabIndex        =   23
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command6 
         Caption         =   "�Զ�ѭ������(&U)"
         Height          =   735
         Left            =   240
         TabIndex        =   37
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "�ֶ�ѭ������(&N)"
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   3960
         Width           =   3855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�Զ�����ɫ10"
         Height          =   495
         Index           =   9
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�Զ�����ɫ9"
         Height          =   495
         Index           =   8
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�Զ�����ɫ8"
         Height          =   495
         Index           =   7
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�Զ�����ɫ7"
         Height          =   495
         Index           =   6
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "�Զ�����ɫ1"
         Height          =   495
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "ѭ��ʱ��(ms,1s=1000ms)"
         Height          =   255
         Left            =   1920
         TabIndex        =   38
         Top             =   4800
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "15.����ɫ"
      Height          =   375
      Index           =   15
      Left            =   1980
      TabIndex        =   16
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "14.����ɫ"
      Height          =   375
      Index           =   14
      Left            =   165
      TabIndex        =   15
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "13.�����ɫ"
      Height          =   375
      Index           =   13
      Left            =   1980
      TabIndex        =   14
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "12.����ɫ"
      Height          =   375
      Index           =   12
      Left            =   165
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "11.����ɫ"
      Height          =   375
      Index           =   11
      Left            =   1980
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "10.����ɫ"
      Height          =   375
      Index           =   10
      Left            =   165
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9.����ɫ"
      Height          =   375
      Index           =   9
      Left            =   1980
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8.��ɫ"
      Height          =   375
      Index           =   8
      Left            =   165
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7.��ɫ"
      Height          =   375
      Index           =   7
      Left            =   1980
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6.��ɫ"
      Height          =   375
      Index           =   6
      Left            =   165
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5.���ɫ"
      Height          =   375
      Index           =   5
      Left            =   1980
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4.��ɫ"
      Height          =   375
      Index           =   4
      Left            =   165
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3.��ɫ"
      Height          =   375
      Index           =   3
      Left            =   1980
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2.��ɫ"
      Height          =   375
      Index           =   2
      Left            =   165
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1.��ɫ"
      Height          =   375
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "��׼16ɫ(������ť���е������)"
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      Begin VB.CommandButton Command3 
         Caption         =   "16ɫ�ֶ�ѭ��(&M)"
         Height          =   615
         Left            =   165
         TabIndex        =   22
         Top             =   4920
         Width           =   4335
      End
      Begin VB.PictureBox Picture1 
         Height          =   3450
         Left            =   3645
         ScaleHeight     =   3390
         ScaleWidth      =   810
         TabIndex        =   21
         ToolTipText     =   "�ڴ�Ԥ����ɫ"
         Top             =   660
         Width           =   870
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form1.frx":04D7
         Left            =   2325
         List            =   "Form1.frx":0508
         TabIndex        =   19
         Text            =   "1000"
         Top             =   4470
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "16ɫ�Զ�ѭ��(&A)"
         Height          =   615
         Left            =   165
         TabIndex        =   18
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0.��ɫ"
         Height          =   375
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "ѭ��ʱ��(ms,1s=1000ms)"
         Height          =   255
         Left            =   2325
         TabIndex        =   20
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "��ɫԤ��"
         Height          =   270
         Left            =   3630
         TabIndex        =   17
         Top             =   345
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Combo1_KeyPress(KeyAscii As Integer)
On Error GoTo ep
If KeyAscii = vbKeyReturn Then
Form3.Timer1.Interval = Val(Combo1.Text)
con:
If Form3.Timer1.Interval = 0 Then Form3.Timer1.Interval = 1000
Form3.Show (1)
End If
Exit Sub
ep:
MsgBox "��������:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
On Error GoTo ep
If KeyAscii = vbKeyReturn Then
Form5.BackColor = Command4(0).BackColor
Form5.Timer1.Interval = Val(Combo2.Text)
If Form5.Timer1.Interval = 0 Then Form5.Timer1.Interval = 1000
Form5.Show (1)
End If
Exit Sub
ep:
MsgBox "��������:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command1_Click(Index As Integer)
Load Form2
Form2.BackColor = QBColor(Index)
Form2.WindowState = 2
Form2.singlecolour = True
Form2.Show (1)
End Sub
Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BackColor = QBColor(Index)
End Sub
Private Sub Command10_Click()
Form9.Show
End Sub
Private Sub Command11_Click()
End
End Sub
Private Sub Command12_Click()
Form10.Show 1
End Sub
Private Sub Command13_Click()
On Error GoTo ep
Me.Tag = ""
frmOpenFile.Show 1
If Trim(Tag) <> "" Then
Form11.Picture = LoadPicture(Me.Tag)
Form11.Show (1)
End If
Exit Sub
ep:
If Err.Number <> 32755 Then
MsgBox "Error:" & vbCrLf & Err.Description, vbCritical, "Error"
Else
Exit Sub
End If
End Sub
Private Sub Command14_Click()
Form12.Show (1)
End Sub
Private Sub Command15_Click()
Form13.Show 1
End Sub
Private Sub Command2_Click()
On Error GoTo ep
Form3.Timer1.Interval = Val(Combo1.Text)
con:
If Form3.Timer1.Interval = 0 Then Form3.Timer1.Interval = 1000
Form3.Show (1)
Exit Sub
ep:
MsgBox "��������:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command3_Click()
Form2.BackColor = QBColor(0)
Form2.Show (1)
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo ep
FrmColorPanel.Show 1
Command4(Index).BackColor = CLng(FrmColorPanel.Tag)
Command4(Index).Caption = CLng(FrmColorPanel.Tag)
CommonDialog1.ShowColor
Command4(Index).BackColor = CommonDialog1.Color
Command4(Index).Caption = CommonDialog1.Color
Index = Index + 1
If Index <= 11 Then
Command4(Index).SetFocus
Else
Command5.SetFocus
End If
Exit Sub
ep:
If Err.Number = 32755 Then Exit Sub
End Sub
Private Sub Command5_Click()
Form4.BackColor = Command4(0).BackColor
Form4.Show (1)
End Sub
Private Sub Command6_Click()
On Error GoTo ep
Form5.BackColor = Command4(0).BackColor
Form5.Timer1.Interval = Val(Combo2.Text)
If Form5.Timer1.Interval = 0 Then Form5.Timer1.Interval = 1000
Form5.Show (1)
Exit Sub
ep:
MsgBox "��������:" & Chr(13) & Err.Description, vbCritical, "Error"
End Sub
Private Sub Command7_Click()
Form6.Show (1)
End Sub
Private Sub Command8_Click()
Form7.Show (1)
End Sub
Private Sub Command9_Click()
Form8.Show (1)
End Sub
Private Sub Form_Activate()
Picture1.BackColor = vbWhite
End Sub
Private Sub Form_Deactivate()
Picture1.BackColor = vbWhite
End Sub
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = False Then
Me.Show
Picture1.BackColor = &H8000000F
Else
MsgBox "������������ͬʱ����2����2�����ϵ�ʵ��,���򼴽��˳�...", vbExclamation, "Error"
End
End If
End Sub
