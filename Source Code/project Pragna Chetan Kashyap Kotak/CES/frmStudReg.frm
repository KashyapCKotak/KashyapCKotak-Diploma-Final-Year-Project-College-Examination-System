VERSION 5.00
Begin VB.Form frmStudReg 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Registration"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmStudReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8475
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   5535
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdReg 
         BackColor       =   &H008080FF&
         Caption         =   "Register"
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
         TabIndex        =   5
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
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
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtUname 
         Appearance      =   0  'Flat
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
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STUDENT REGISTRATION"
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
         Top             =   240
         Width           =   5535
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
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   2070
         Width           =   1410
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
         Left            =   840
         TabIndex        =   1
         Top             =   1350
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmStudReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dec, n As Integer
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdReg_Click()
dec = 1
n = 0
If (txtPass.Text = "" Or txtUname.Text = "") Then
    MsgBox "Missing Fields", vbInformation, " CES"
Else
    Set rs = con.Execute("select * from Student where Username='" + txtUname.Text + "'")
    If (Not rs.EOF) Then
        MsgBox "Username already exist! please try another name", vbInformation, " CES"
        dec = 0
                Set rs = con.Execute("select ucounter from Student where Username='" + txtUname.Text + "'")
                If (Not rs.EOF) Then
                    cnt = rs(0)
                End If
                rs.Close
                If (cnt >= 3) Then
                    con.Execute ("update Student set ucounter=" & n & " where Username='" & txtUname.Text & "'")
                MsgBox "Re-Registered successfully! Login and attend the exam", vbInformation, " CES"
                dec = 0
                End If
    End If
    If dec = 1 Then
        con.Execute ("insert into Student values('" + txtUname.Text + "','" + txtPass.Text + "','0',NULL,NULL,NULL)")
        MsgBox "Registered successfully! Login and attend the exam", vbInformation, " CES"
    End If
    If dec = 0 And n = 1 Then
            Load frmAsk
            frmAsk.Show
    End If
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
