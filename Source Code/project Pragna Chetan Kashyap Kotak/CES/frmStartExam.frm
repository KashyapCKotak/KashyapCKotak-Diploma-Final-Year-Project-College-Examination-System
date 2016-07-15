VERSION 5.00
Begin VB.Form frmStartExam 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start Exam"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "frmStartExam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7875
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   1410
      TabIndex        =   0
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtUname 
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
         Left            =   2340
         TabIndex        =   1
         Top             =   840
         Width           =   2445
      End
      Begin VB.TextBox txtPass 
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
         Left            =   2340
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   1845
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H008080FF&
         Caption         =   "Enter"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H008080FF&
         Caption         =   "Cancel"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
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
         Height          =   555
         Left            =   600
         TabIndex        =   5
         Top             =   1800
         Width           =   1770
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   4150
         Picture         =   "frmStartExam.frx":0442
         Top             =   1680
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   4150
         Picture         =   "frmStartExam.frx":0A38
         Top             =   1680
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmStartExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub cmdEnter_Click()
Set rs = con.Execute("select ucounter from Student where Username='" + txtUname.Text + "'")
If (Not rs.EOF) Then
    cnt = rs(0)
End If
rs.Close
If (cnt < 3) Then
Set rs = con.Execute("select * from Student where Username='" + txtUname.Text + "' and Password='" + txtPass.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Login Success", vbInformation, " CES"
    uname = txtUname.Text
    MDIForm1.StatusBar1.Panels(2).Text = txtUname.Text
    Load frmSelectExam
    frmSelectExam.Show
    Unload Me
    'rs.Close
Else
    MsgBox "Invalid Username or Password", vbCritical, " CES"
End If

Else
    MsgBox "Your number of attempts over, Please reregister", vbCritical, " CES"
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

Private Sub Label1_Click()

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtPass.PasswordChar = ""
txtPass.Refresh
Image1.Visible = False
Image2.Visible = True

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtPass.PasswordChar = "*"
txtPass.Refresh
Image1.Visible = True
Image2.Visible = False
End Sub

