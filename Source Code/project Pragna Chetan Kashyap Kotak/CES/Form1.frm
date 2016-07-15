VERSION 5.00
Begin VB.Form frmLoginSelect 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Login"
   ClientHeight    =   6030
   ClientLeft      =   5775
   ClientTop       =   1125
   ClientWidth     =   12300
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   12300
   Begin VB.Frame frmLogin 
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H0080C0FF&
      Height          =   4095
      Left            =   5520
      TabIndex        =   9
      Top             =   1320
      Width           =   5535
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton cmdAdminLog 
         BackColor       =   &H008080FF&
         Caption         =   "Admin Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Student Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTchLog 
         BackColor       =   &H008080FF&
         Caption         =   "Teacher Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   600
         TabIndex        =   16
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " LOGIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   5535
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   420
         Left            =   4080
         Picture         =   "Form1.frx":0442
         Top             =   2160
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   4080
         Picture         =   "Form1.frx":0A38
         Top             =   2160
         Width           =   630
      End
   End
   Begin VB.Frame mainFrame 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":0EF6
         ScaleHeight     =   735
         ScaleWidth      =   735
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2880
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":18E9
         ScaleHeight     =   1.191
         ScaleMode       =   0  'User
         ScaleWidth      =   1.191
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":2322
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":2E6A
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H008080FF&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdStudent 
         BackColor       =   &H00FF8080&
         Caption         =   "STUDENT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdTeacher 
         BackColor       =   &H00FF8080&
         Caption         =   "TEACHER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdmin 
         BackColor       =   &H00FF8080&
         Caption         =   "ADMINISTRATOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Height          =   135
         Left            =   0
         TabIndex        =   24
         Top             =   2640
         Width           =   3135
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label non3 
      BackColor       =   &H00C00000&
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label non 
      BackColor       =   &H00C00000&
      Height          =   3615
      Left            =   4920
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label non2 
      BackColor       =   &H00C00000&
      Height          =   3615
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label name 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CollegeExaminationSystem"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   9855
   End
End
Attribute VB_Name = "frmLoginSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdmin_Click()

Command2.Enabled = True
cmdAdmin.BackColor = &HC0C0C0
cmdAdmin.Enabled = False
cmdTeacher.BackColor = &HC0C0C0
cmdTeacher.Enabled = False
cmdStudent.BackColor = &HC0C0C0
cmdStudent.Enabled = False
Image1.Enabled = True
frmLogin.BackColor = &HC00000
Label1.ForeColor = &HFF00&
Label2.ForeColor = &HFF00&
Label3.ForeColor = &HFF00&
Command2.BackColor = &H8080FF
frmLogin.Enabled = True
frmLogin.Visible = True
cmdAdminLog.Visible = True
End Sub

Private Sub cmdAdminLog_Click()
Dim un As String
Dim pw As String
Dim decide As Integer
decide = 0
un = "admin"
pw = "admin"
una = txtUsername.Text
pwa = txtPassword.Text
sh = 5
If un = una Then
    decide = 1
End If
If pw = pwa Then
    decide = 1
End If
'Set rs = con.Execute("select * from Adminlogin where Username='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
If decide = 1 Then
    decide = 0
    MsgBox "Login Success", vbInformation, " CES"
    MDIForm1.mnuAdmin.Enabled = True
    MDIForm1.mnuExam.Enabled = True
    MDIForm1.mnuAdminAdSub.Enabled = True
    MDIForm1.mnuAdminchgPass.Enabled = True
    MDIForm1.mnuAdminCrtBrnh.Enabled = True
    MDIForm1.mnuadminCrtUsr.Enabled = True
    MDIForm1.mnuAdminDltUsr.Enabled = True
    MDIForm1.mnuAdminLogout.Enabled = True
    MDIForm1.mnuStudent.Enabled = True
    MDIForm1.mnuAdminLogout.Enabled = True
    MDIForm1.mnuAdminLogin.Enabled = False
    MDIForm1.Command1.Enabled = False
    MDIForm1.Command2.Enabled = True
    MDIForm1.Picture4.Visible = True
    MDIForm1.mnuAdmin.Enabled = True
MDIForm1.mnuExam.Enabled = True
MDIForm1.mnuStudent.Enabled = True
    MDIForm1.StatusBar1.Panels(1).Text = "Status: Logged in as Administrator"
    MDIForm1.StatusBar1.Panels(2).Text = txtUsername.Text
    MDIForm1.Label4.Caption = txtUsername.Text
    Unload Me
Else
    MsgBox "Invalid Username or Password", vbCritical, " CES"
End If
'rs.Close
MDIForm1.mnuAdminLogin.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload MDIForm1
'Load frmSplash2
'frmSplash2.Show
End Sub

Private Sub cmdStudent_Click()

Command2.Enabled = True
cmdAdmin.BackColor = &HC0C0C0
cmdAdmin.Enabled = False
cmdTeacher.BackColor = &HC0C0C0
cmdTeacher.Enabled = False
cmdStudent.BackColor = &HC0C0C0
cmdStudent.Enabled = False
frmLogin.BackColor = &HC00000
Image1.Enabled = True
Label1.ForeColor = &HFF00&
Label2.ForeColor = &HFF00&
Label3.ForeColor = &HFF00&
Command2.BackColor = &H8080FF
frmLogin.Enabled = True
frmLogin.Visible = True
Command1.Visible = True
End Sub

Private Sub cmdTchLog_Click()
connectdb
sh = 1
Set rs = con.Execute("select * from Userlogin where Username='" + txtUsername.Text + "' and UPassword='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Login Success", vbInformation, " CES"
    MDIForm1.mnuStudent.Enabled = False
    MDIForm1.mnuAdminLogout.Enabled = True
    MDIForm1.mnuAdminLogin.Enabled = False
    MDIForm1.mnuExam.Enabled = True
    MDIForm1.mnuAdminchgPass.Enabled = True
    MDIForm1.Command1.Enabled = False
    MDIForm1.Command2.Enabled = True
    MDIForm1.Picture5.Visible = True
    MDIForm1.mnuAdmin.Enabled = False
MDIForm1.mnuExam.Enabled = True
MDIForm1.mnuStudent.Enabled = False
    MDIForm1.StatusBar1.Panels(1).Text = "Status: Logged in as Teacher"
    MDIForm1.StatusBar1.Panels(2).Text = txtUsername.Text
    MDIForm1.Label4.Caption = txtUsername.Text
    Unload Me
Else
    MsgBox "Invalid Username or Password", vbCritical, " CES"
End If
MDIForm1.mnuAdminLogin.Enabled = False

End Sub

Private Sub cmdTeacher_Click()

Command2.Enabled = True
cmdAdmin.BackColor = &HC0C0C0
cmdAdmin.Enabled = False
cmdTeacher.BackColor = &HC0C0C0
cmdTeacher.Enabled = False
cmdStudent.BackColor = &HC0C0C0
cmdStudent.Enabled = False
frmLogin.BackColor = &HC00000
Image1.Enabled = True
Label1.ForeColor = &HFF00&
Label2.ForeColor = &HFF00&
Label3.ForeColor = &HFF00&
Command2.BackColor = &H8080FF
frmLogin.Enabled = True
frmLogin.Visible = True
cmdTchLog.Visible = True
End Sub

Private Sub Command1_Click()
connectdb
sh = 1
Set rs = con.Execute("select * from Student where Username='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Login Success", vbInformation, " CES"
MDIForm1.mnuAdminLogin.Enabled = False
MDIForm1.mnuAdminLogout.Enabled = True
MDIForm1.StatusBar1.Panels(1).Text = "Status: Logged in as Student"
MDIForm1.mnuExam.Enabled = False
MDIForm1.StatusBar1.Panels(2).Text = txtUsername.Text
MDIForm1.Label4.Caption = txtUsername.Text
MDIForm1.mnuStudent = True
MDIForm1.Command1.Enabled = False
MDIForm1.Command2.Enabled = True
MDIForm1.Picture6.Visible = True
MDIForm1.mnuAdmin.Enabled = False
MDIForm1.mnuExam.Enabled = False
MDIForm1.mnuStudent.Enabled = True
Unload Me
Else
    MsgBox "Invalid Username or Password", vbCritical, " CES"
End If
rs.Close
MDIForm1.mnuAdminLogin.Enabled = False

End Sub

Private Sub Command2_Click()
cmdAdmin.BackColor = &HFF8080
cmdAdmin.Enabled = True
cmdTeacher.BackColor = &HFF8080
cmdTeacher.Enabled = True
cmdStudent.BackColor = &HFF8080
cmdStudent.Enabled = True
frmLogin.BackColor = &HC0C0C0
Label1.ForeColor = &H808080
Label2.ForeColor = &H808080
Label3.ForeColor = &H808080
Command2.BackColor = &HC0C0C0
cmdAdminLog.Visible = False
cmdTchLog.Visible = False
Command1.Visible = False
frmLogin.Enabled = False
End Sub



Private Sub Command4_Click()
End
End Sub



Private Sub Command3_Click()

'ShellExecute hWnd, "open", "D:\wallpapers", vbNullString, vbNullString, SW_SHOWNORMAL
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = (Asc("'")) Then
    KeyAscii = 0
    MsgBox "Character not allowed", vbInformation
End If
If KeyAscii = (Asc("-")) Then
    KeyAscii = 0
    MsgBox "Character not allowed", vbInformation
End If
If KeyAscii = (Asc("*")) Then
    KeyAscii = 0
    MsgBox "Character not allowed", vbInformation
End If
'If KeyAscii = (Asc("'")) Then
'    KeyAscii = 0
'    MsgBox "Character not allowed", vbInformation
'End If
End Sub

Private Sub Form_Load()
BackColor = color
'connectdb
MDIForm1.mnuAdminLogin.Enabled = False
frmLoginSelect.Top = 1550
frmLoginSelect.Left = 3900
End Sub

Private Sub Form_Terminate()
MDIForm1.mnuAdminLogin.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.mnuAdminLogin.Enabled = True
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtPassword.PasswordChar = ""
txtPassword.Refresh
Image1.Visible = False
Image2.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtPassword.PasswordChar = "*"
txtPassword.Refresh
Image1.Visible = True
Image2.Visible = False
End Sub

