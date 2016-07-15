VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmAddUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8475
   Begin VB.Frame frmLogin 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   5535
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H008080FF&
         Caption         =   "Add"
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
         TabIndex        =   2
         Top             =   3360
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
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADD TEACHER ACCOUNT"
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
         BackColor       =   &H80000018&
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
         TabIndex        =   6
         Top             =   1485
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   2280
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
Set rs = con.Execute("select * from Userlogin where Username='" + txtUsername.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Sorry!! User already exists. Try another username", vbCritical, " CES"
    txtPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
Else
    con.Execute ("insert into Userlogin values('" + txtUsername.Text + "','" + txtPassword.Text + "')")
    MsgBox "User added sucessfully", vbInformation, " CES"
    txtPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
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
frmAddUser.Top = 1550
frmAddUser.Left = 5000
End Sub
