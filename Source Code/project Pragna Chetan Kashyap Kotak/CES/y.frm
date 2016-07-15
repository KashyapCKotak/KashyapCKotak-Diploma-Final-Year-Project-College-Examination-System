VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Offline Examiner"
   ClientHeight    =   9990
   ClientLeft      =   360
   ClientTop       =   900
   ClientWidth     =   20250
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox StatusBar3 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   20190
      TabIndex        =   36
      Top             =   9240
      Width           =   20250
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7440
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      BackColor       =   &H000080FF&
      Height          =   7605
      Left            =   18435
      ScaleHeight     =   7545
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command12 
         BackColor       =   &H008080FF&
         Height          =   615
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   8160
         Width           =   302
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF80FF&
         Height          =   615
         Left            =   1207
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8160
         Width           =   302
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00F486C0&
         Height          =   615
         Left            =   907
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   8160
         Width           =   302
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080C0FF&
         Height          =   615
         Left            =   606
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   8160
         Width           =   302
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FF80&
         Height          =   615
         Left            =   303
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   8160
         Width           =   302
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FF8080&
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   8160
         Width           =   302
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00534FFD&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Picture         =   "y.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
      Begin VB.PictureBox Picture6 
         Height          =   735
         Left            =   0
         Picture         =   "y.frx":08A5
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox Picture5 
         Height          =   735
         Left            =   0
         Picture         =   "y.frx":13ED
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox Picture4 
         Height          =   735
         Left            =   0
         Picture         =   "y.frx":1DAF
         ScaleHeight     =   1.191
         ScaleMode       =   0  'User
         ScaleWidth      =   1.191
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00DB71A0&
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Picture         =   "y.frx":27E8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00DB71A0&
         Caption         =   "TIME TABLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Picture         =   "y.frx":303A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00DB71A0&
         Caption         =   "FEEDBACK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Picture         =   "y.frx":34FE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6960
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "LOGOUT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Picture         =   "y.frx":3A02
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Picture         =   "y.frx":3EAA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   33
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0082F5F9&
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   32
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   31
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "THEMES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   8760
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H000080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   29
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H000080FF&
         Caption         =   "EXAMINER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   0
         TabIndex        =   28
         Top             =   9120
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   27
         Top             =   8040
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   14
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   13
         Top             =   4440
         Width           =   1815
      End
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   20190
      TabIndex        =   0
      Top             =   9615
      Width           =   20250
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   7605
      Left            =   0
      Picture         =   "y.frx":4352
      ScaleHeight     =   7545
      ScaleWidth      =   18555
      TabIndex        =   12
      Tag             =   "1240     523"
      Top             =   0
      Width           =   18615
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BackColor       =   &H008080FF&
      Height          =   1640
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   20190
      TabIndex        =   3
      Top             =   7605
      Width           =   20250
      Begin VB.PictureBox StatusBar2 
         Height          =   700
         Left            =   500
         ScaleHeight     =   645
         ScaleWidth      =   20190
         TabIndex        =   1
         Top             =   9500
         Width           =   20250
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   80
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1140
         Width           =   20300
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   80
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   750
         Width           =   20300
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   80
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   20300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "News"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "&General"
      Begin VB.Menu mnuView 
         Caption         =   "&View"
         Begin VB.Menu mnuViewTheme 
            Caption         =   "&Change Theme"
            Begin VB.Menu mnuViewOr 
               Caption         =   "Orange"
            End
            Begin VB.Menu mnuViewGr 
               Caption         =   "Green"
            End
            Begin VB.Menu mnuViewPnk 
               Caption         =   "Pink"
            End
            Begin VB.Menu mnuViewBl 
               Caption         =   "Blue"
            End
            Begin VB.Menu mnuViewPl 
               Caption         =   "Purple"
            End
            Begin VB.Menu mnuViewRd 
               Caption         =   "Red"
            End
         End
      End
      Begin VB.Menu mnudh7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminLogin 
         Caption         =   "Log&in"
      End
      Begin VB.Menu mnuAdminLogout 
         Caption         =   "Log&out"
      End
      Begin VB.Menu mnudh3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminExit 
         Caption         =   "Log out and &Exit"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Administrator"
      WindowList      =   -1  'True
      Begin VB.Menu mnuadminCrtUsr 
         Caption         =   "&Create Teacher Account"
      End
      Begin VB.Menu mnuExStdReg 
         Caption         =   "&Student Registration"
      End
      Begin VB.Menu mnuAdminTt 
         Caption         =   "Create &Time Table"
      End
      Begin VB.Menu mnuAdminNews 
         Caption         =   "Edit News-lines"
      End
      Begin VB.Menu mnudh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminCrtBrnh 
         Caption         =   "Create &Branch"
      End
      Begin VB.Menu mnuAdminAdSub 
         Caption         =   "Add &Subject"
      End
      Begin VB.Menu mnudh6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminDltUsr 
         Caption         =   "Delete &Teacher Account"
      End
      Begin VB.Menu mnuAdminDltBranch 
         Caption         =   "Delete B&ranch"
      End
      Begin VB.Menu mnuAdminDltSub 
         Caption         =   "&Delete Subject"
      End
      Begin VB.Menu mnudh10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminFeedback 
         Caption         =   "View &Feedback"
      End
      Begin VB.Menu mnuAdminData 
         Caption         =   "Configure DataBase &Source"
      End
   End
   Begin VB.Menu mnuExam 
      Caption         =   "&Teacher"
      Begin VB.Menu mnuAdminchgPass 
         Caption         =   "Change &Password"
      End
      Begin VB.Menu mnuExamAdQst 
         Caption         =   "Create &Exam"
      End
      Begin VB.Menu mnudh4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExamQst 
         Caption         =   "&View all Questions"
      End
      Begin VB.Menu mnuExVQBEXID 
         Caption         =   "View &Questons based on ExamID"
      End
   End
   Begin VB.Menu mnuStudent 
      Caption         =   "&Student"
      Begin VB.Menu mnuExAttEx 
         Caption         =   "&Attend Exam"
      End
      Begin VB.Menu mnudh5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTt 
         Caption         =   "View Time &Table"
      End
      Begin VB.Menu mnuRslt 
         Caption         =   "&View Result"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout"
      Begin VB.Menu mnuAboutAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer

Private Sub login_Click()

End Sub

Private Sub Command1_Click()
cnt = 1
MDIForm1.Picture1.Visible = False
MDIForm1.Picture3.Visible = False
Load frmLoginSelect
frmLoginSelect.Show
'StatusBar1.Panels(1).Text = "Login or choose to exit"
End Sub

Private Sub Command10_Click()
MDIForm1.BackColor = &HF486C0
color = &HEA1585
End Sub

Private Sub Command11_Click()
MDIForm1.BackColor = &HFF80FF
color = &HFF00FF
End Sub

Private Sub Command12_Click()
MDIForm1.BackColor = &H8080FF
color = &HFF&
End Sub

Private Sub Command2_Click()
MDIForm1.Picture1.Visible = True
MDIForm1.Picture3.Visible = True
mnuExam.Enabled = False
mnuAdminAdSub.Enabled = False
mnuAdminchgPass.Enabled = False
mnuAdminCrtBrnh.Enabled = False
mnuadminCrtUsr.Enabled = False
mnuAdminDltUsr.Enabled = False
mnuAdminLogout.Enabled = False
mnuStudent.Enabled = True
mnuAdminLogin.Enabled = True
MDIForm1.Command1.Enabled = True
MDIForm1.Command2.Enabled = False
MDIForm1.Picture4.Visible = False
MDIForm1.Picture5.Visible = False
MDIForm1.Picture6.Visible = False
Load frmLoginSelect
frmLoginSelect.Show
StatusBar1.Panels(1).Text = "Status: Logged out successfully. Please login as another user or choose to exit"
MDIForm1.StatusBar1.Panels(2).Text = ""
MDIForm1.Label4.Caption = ""
End Sub

Private Sub Command3_Click()
MDIForm1.Picture1.Visible = False
MDIForm1.Picture3.Visible = False
xyz = 15
Load frmFeedback
frmFeedback.Show
End Sub

Private Sub Command4_Click()
MDIForm1.Picture1.Visible = False
MDIForm1.Picture3.Visible = False
sh = 5
Load frmTT
frmTT.Show
End Sub

Private Sub Command5_Click()
If cnt = 0 Then
MDIForm1.Picture1.Visible = True
MDIForm1.Picture3.Visible = True
cnt = 1
ElseIf cnt = 1 Then
    MDIForm1.Picture1.Visible = False
    MDIForm1.Picture3.Visible = False
    cnt = 0
End If
End Sub

Private Sub Command6_Click()
Unload MDIForm1
'Load frmSplash2
'frmSplash2.Show
End Sub

Private Sub Command7_Click()
MDIForm1.BackColor = &HFF8080
color = &HFF0000
End Sub

Private Sub Command8_Click()
MDIForm1.BackColor = &H80FF80
color = &HFF00&
End Sub

Private Sub Command9_Click()
MDIForm1.BackColor = &H80C0FF
color = &H80FF&
End Sub

Private Sub MDIForm_Load()
counter = 1
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String
strFilename = App.Path & "\loc.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
 strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
 Close #iFile
 dbloc = strTheData

'MDIForm1.StatusBar3.Panels(1).Text = dbloc & " Data Source"

sh = 1

 strFilename = App.Path & "\dbsource.txt"
 iFile = FreeFile
 Open strFilename For Input As #iFile
 strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
 Close #iFile
 str = strTheData


MDIForm1.BackColor = &H80C0FF
color = &H80FF&



cnt = 0
time = 0
data = ""
mnuAdmin.Enabled = False
mnuExam.Enabled = False
mnuAdminAdSub.Enabled = False
mnuAdminchgPass.Enabled = False
mnuAdminCrtBrnh.Enabled = False
mnuadminCrtUsr.Enabled = False
mnuAdminDltUsr.Enabled = False
mnuAdminLogout.Enabled = False
mnuStudent.Enabled = False
mnuAdminLogin.Enabled = True
'Load frmLoginSelect
'frmLoginSelect.Show
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\OfflineExaminer.mdb;Persist Security Info=False"
'StatusBar1.Panels(1).Text = "Please Login"


strFilename = App.Path & "\news1.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text1.Text = strTheData
  
strFilename = App.Path & "\news2.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text2.Text = strTheData
  
strFilename = App.Path & "\news3.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text3.Text = strTheData


End Sub

Private Sub MDIForm_Terminate()
Load frmSplash2
frmSplash2.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Load frmSplash2
frmSplash2.Show
End Sub

Private Sub mnuAboutAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuAdminAdSub_Click()
Load frmSubjects
frmSubjects.Show
End Sub

Private Sub mnuAdminchgPass_Click()
Load frmChangePass
frmChangePass.Show
End Sub

Private Sub mnuAdminCrtBrnh_Click()
Load frmBranch
frmBranch.Show
End Sub

Private Sub mnuadminCrtUsr_Click()
Load frmAddUser
frmAddUser.Show
End Sub

Private Sub mnuAdminData_Click()
Load frmAskdb
frmAskdb.Show
End Sub

Private Sub mnuAdminDltBranch_Click()
Load frmDelBranch
frmDelBranch.Show
End Sub

Private Sub mnuAdminDltSub_Click()
Load frmDelSub
frmDelSub.Show
End Sub

Private Sub mnuAdminDltUsr_Click()
Load frmDeleteUser
frmDeleteUser.Show
End Sub

Private Sub mnuAdminExit_Click()
Load frmSplash2
frmSplash2.Show
End Sub

Private Sub mnuAdminFeedback_Click()
xyz = 0
Load frmFeedback
frmFeedback.Show

End Sub

Private Sub mnuAdminLogin_Click()
cnt = 1
MDIForm1.Picture1.Visible = False
MDIForm1.Picture3.Visible = False
Load frmLoginSelect
frmLoginSelect.Show
StatusBar1.Panels(1).Text = "Login or choose to exit"
End Sub

Private Sub mnuAdminLogout_Click()
MDIForm1.Picture1.Visible = True
MDIForm1.Picture3.Visible = True
mnuExam.Enabled = False
mnuAdminAdSub.Enabled = False
mnuAdminchgPass.Enabled = False
mnuAdminCrtBrnh.Enabled = False
mnuadminCrtUsr.Enabled = False
mnuAdminDltUsr.Enabled = False
mnuAdminLogout.Enabled = False
mnuStudent.Enabled = True
mnuAdminLogin.Enabled = True
Load frmLoginSelect
frmLoginSelect.Show
StatusBar1.Panels(1).Text = "Status: Logged out successfully. Please login as another user or choose to exit"
MDIForm1.StatusBar1.Panels(2).Text = ""
MDIForm1.Label4.Caption = ""
End Sub

Private Sub mnuAdminNews_Click()
Load frmNews
frmNews.Show
End Sub

Private Sub mnuAdminTt_Click()
sh = 1
Load frmTT
frmTT.Show
End Sub

Private Sub mnuExamAdQst_Click()
Load frmAddQst
frmAddQst.Show vbModal
End Sub

Private Sub mnuExamQst_Click()
Load DataReportQst
DataReportQst.Show
End Sub

Private Sub mnuExAttEx_Click()
Load frmStartExam
frmStartExam.Show
End Sub

Private Sub mnuExStdReg_Click()
Load frmStudReg
frmStudReg.Show
End Sub


Private Sub mnuExVQBEXID_Click()
Load frmViewRptQst
frmViewRptQst.Show
End Sub

Private Sub mnuRslt_Click()
Load frmStudRslt
frmStudRslt.Show
End Sub

Private Sub mnuViewBl_Click()
MDIForm1.BackColor = &HFF8080
color = &HFF0000
End Sub

Private Sub mnuViewGr_Click()
MDIForm1.BackColor = &H80FF80
color = &HFF00&
End Sub

Private Sub mnuViewOr_Click()

MDIForm1.BackColor = &H80C0FF
color = &H80FF&
End Sub

Private Sub mnuViewPl_Click()
MDIForm1.BackColor = &HF486C0
color = &HEA1585
End Sub

Private Sub mnuViewPnk_Click()
MDIForm1.BackColor = &HFF80FF
color = &HFF00FF
End Sub

Private Sub mnuViewRd_Click()
MDIForm1.BackColor = &H8080FF
color = &HFF&
End Sub

Private Sub mnuViewTt_Click()
sh = 5
Load frmTT
frmTT.Show
End Sub


Private Sub Timer1_Timer()
time = 0
If MDIForm1.StatusBar1.Panels(4).Text = "10:00 AM" Then
    time = 1
ElseIf MDIForm1.StatusBar1.Panels(4).Text = "11:00 AM" Then
    time = 0
Else
    time = 0
End If
If time = 0 Then
    MDIForm1.mnuExAttEx.Enabled = False
Else
    MDIForm1.mnuExAttEx.Enabled = True
End If
End Sub
