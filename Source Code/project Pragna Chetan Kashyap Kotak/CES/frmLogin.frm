VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   6120
   ClientLeft      =   6915
   ClientTop       =   2505
   ClientWidth     =   7605
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7605
   Begin VB.Frame frmLogin 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H0080C0FF&
      Height          =   4095
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
      Begin VB.CommandButton Command1 
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
         Left            =   240
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdminLog 
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
         Left            =   3960
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdTchLog 
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
         Left            =   2040
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   2160
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
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
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   4440
         Picture         =   "frmLogin.frx":0442
         Top             =   2160
         Width           =   630
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
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   5535
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
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   1485
         Width           =   1425
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
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   2250
         Width           =   1485
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   4440
         Picture         =   "frmLogin.frx":0A38
         Top             =   2160
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_KeyPress(KeyAscii As Integer)
txtPassword.PasswordChar = ""
txtPassword.Refresh
End Sub

Private Sub A_KeyUp(KeyCode As Integer, Shift As Integer)
txtPassword.PasswordChar = "*"
txtPassword.Refresh
End Sub

Private Sub A_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtPassword.PasswordChar = ""
txtPassword.Refresh
End Sub

Private Sub A_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtPassword.PasswordChar = "*"
txtPassword.Refresh
End Sub

Private Sub cmdAdminLog_Click()
Set rs = con.Execute("select * from Adminlogin where Username='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Login Success", vbInformation, " CES"
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
    Unload Me
Else
    MsgBox "Invalid Username or Password", vbCritical, " CES"
End If
rs.Close
End Sub

Private Sub cmdCancel_Click()
End
End Sub
Private Sub cmdTchLog_Click()
Set rs = con.Execute("select * from Userlogin where Username='" + txtUsername.Text + "' and UPassword='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Login Success", vbInformation, " CES"
    MDIForm1.mnuStudent.Enabled = False
    MDIForm1.mnuAdminLogout.Enabled = True
    MDIForm1.mnuAdminLogin.Enabled = False
    MDIForm1.mnuExam.Enabled = True
    MDIForm1.mnuAdminchgPass.Enabled = True
    Unload Me
Else
    MsgBox "Invalid Username or Password", vbCritical, " CES"
End If
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
End Sub

Private Sub Form_Load()
BackColor = color
connectdb
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

