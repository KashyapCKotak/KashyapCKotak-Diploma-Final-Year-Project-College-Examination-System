VERSION 5.00
Begin VB.Form frmBranch 
   Appearance      =   0  'Flat
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Branch"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmBranch.frx":0000
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
      Height          =   4335
      Left            =   1560
      TabIndex        =   0
      Top             =   840
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
         TabIndex        =   9
         Top             =   3720
         Width           =   1455
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtsem 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2835
         Width           =   2670
      End
      Begin VB.TextBox txtbrcode 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   2025
         Width           =   2670
      End
      Begin VB.TextBox txtbrname 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1215
         Width           =   2670
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "(Please enter capital letters)"
         ForeColor       =   &H00796BF8&
         Height          =   240
         Left            =   2685
         TabIndex        =   10
         Top             =   2400
         Width           =   2325
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         TabIndex        =   6
         Top             =   2865
         Width           =   1230
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   720
         TabIndex        =   3
         Top             =   2055
         Width           =   1320
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1245
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Branch"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         TabIndex        =   1
         Top             =   315
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dec As Integer
Private Sub cmdadd_Click()
dec = 1
Dim num As Integer
If (txtbrname.Text = "" Or txtbrcode.Text = "" Or txtsem.Text = "") Then
    MsgBox "Missing Fields", vbInformation, " CES"
Else
    Set rs = con.Execute("select * from Branch where Branchname='" + txtbrname.Text + "'")
    If (Not rs.EOF) Then
        MsgBox "Sorry!! Branch already exists. Try another branch name", vbCritical, " CES"
        dec = 0
        End If
        Set rs = con.Execute("select * from Branch where Branchcode='" + txtbrcode.Text + "'")
    If (Not rs.EOF) Then
        MsgBox "Sorry!! Branch code already exists. Try another branch code", vbCritical, " CES"
        dec = 0
    End If
End If
If dec = 1 Then
    rs.Close
    con.Execute ("insert into Branch values('" + txtbrname.Text + "','" + txtbrcode.Text + "'," + txtsem.Text + ")")
    MsgBox "Record added successfully", vbInformation, " CES"
    txtbrname.Text = ""
    txtbrcode.Text = ""
    txtsem.Text = ""
    txtbrname.SetFocus
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
End Sub

