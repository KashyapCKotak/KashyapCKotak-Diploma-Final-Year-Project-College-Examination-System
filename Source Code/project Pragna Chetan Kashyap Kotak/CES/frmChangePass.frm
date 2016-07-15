VERSION 5.00
Begin VB.Form frmChangePass 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmChangePass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8475
   Begin VB.Frame frmLogin 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   5280
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   5280
      Begin VB.ComboBox cmbUsername 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2340
         TabIndex        =   10
         Top             =   1485
         Width           =   2490
      End
      Begin VB.TextBox txtConfrmpass 
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
         TabIndex        =   8
         Top             =   3510
         Width           =   1875
      End
      Begin VB.TextBox txtNewpass 
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
         TabIndex        =   6
         Top             =   2835
         Width           =   1875
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdChange 
         BackColor       =   &H008080FF&
         Caption         =   "Change"
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtcurpassword 
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
         TabIndex        =   1
         Top             =   2160
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE TEACHER PASSWORD"
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
         TabIndex        =   11
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   450
         TabIndex        =   9
         Top             =   3600
         Width           =   1860
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "NewPassword"
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
         Left            =   450
         TabIndex        =   7
         Top             =   2925
         Width           =   1860
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
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
         Left            =   450
         TabIndex        =   5
         Top             =   2250
         Width           =   1860
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
         Height          =   420
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   1905
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   4200
         Picture         =   "frmChangePass.frx":0442
         Top             =   2160
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   4200
         Picture         =   "frmChangePass.frx":0A38
         Top             =   2160
         Width           =   630
      End
      Begin VB.Image Image4 
         Height          =   420
         Left            =   4200
         Picture         =   "frmChangePass.frx":0EF6
         Top             =   2835
         Width           =   630
      End
      Begin VB.Image Image3 
         Height          =   420
         Left            =   4200
         Picture         =   "frmChangePass.frx":14EC
         Top             =   2835
         Width           =   630
      End
      Begin VB.Image Image5 
         Height          =   420
         Left            =   4200
         Picture         =   "frmChangePass.frx":19AA
         Top             =   3510
         Width           =   630
      End
      Begin VB.Image Image6 
         Height          =   420
         Left            =   4200
         Picture         =   "frmChangePass.frx":1FA0
         Top             =   3510
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim s As Boolean
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdChange_Click()
If (txtConfrmpass.Text = txtNewpass.Text) Then
x = cmbUsername.Text
    Set rs = con.Execute("select * from Userlogin where Username='" + cmbUsername.Text + "' and UPassword='" + txtcurpassword.Text + "'")
    If (Not rs.EOF) Then
    s = True
        'con.Execute ("UPDATE Userlogin set Password='anup' where Username='" & cmbUsername.Text & "'")
        
        'MsgBox "Password successfully updated!!", vbInformation, "Offline Examiner"
   
    End If
Else
    MsgBox "Password Mismatch!!", vbInformation, " CES"
    txtConfrmpass.Text = ""
    txtNewpass.Text = ""
    txtNewpass.SetFocus
End If

If (s = True) Then
On Error Resume Next
con.Execute ("UPDATE Userlogin set UPassword='" + txtNewpass.Text + "' where Username='" + cmbUsername.Text + "'")

MsgBox "Password successfully updated!!", vbInformation, " CES"
cmdChange.Enabled = False
 Else
        MsgBox "Invalid Password", vbCritical, " CES"
        txtcurpassword.Text = ""
        txtcurpassword.SetFocus
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
Set rs = con.Execute("select * from Userlogin")
While (Not rs.EOF)
    cmbUsername.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
s = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtcurpassword.PasswordChar = ""
txtcurpassword.Refresh
Image1.Visible = False
Image2.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtcurpassword.PasswordChar = "*"
txtcurpassword.Refresh
Image1.Visible = True
Image2.Visible = False
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtNewpass.PasswordChar = ""
txtNewpass.Refresh
Image4.Visible = False
Image3.Visible = True
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtNewpass.PasswordChar = "*"
txtNewpass.Refresh
Image4.Visible = True
Image3.Visible = False
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
txtConfrmpass.PasswordChar = ""
txtConfrmpass.Refresh
Image5.Visible = False
Image6.Visible = True
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
txtConfrmpass.PasswordChar = "*"
txtConfrmpass.Refresh
Image5.Visible = True
Image6.Visible = False
End Sub
