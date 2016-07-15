VERSION 5.00
Begin VB.Form frmSubjects 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Subjects"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   Icon            =   "frmSubjects.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9615
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   6465
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4320
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox cmbSem 
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
         Left            =   2745
         TabIndex        =   9
         Top             =   2115
         Width           =   2955
      End
      Begin VB.ComboBox cmbBranch 
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
         Left            =   2745
         TabIndex        =   8
         Top             =   1485
         Width           =   2955
      End
      Begin VB.TextBox txtSubcode 
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
         Left            =   2745
         TabIndex        =   7
         Top             =   3465
         Width           =   2955
      End
      Begin VB.TextBox txtSubname 
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
         Left            =   2745
         TabIndex        =   6
         Top             =   2750
         Width           =   2955
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code"
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
         Left            =   765
         TabIndex        =   5
         Top             =   3555
         Width           =   1860
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
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
         Left            =   765
         TabIndex        =   4
         Top             =   2880
         Width           =   1860
      End
      Begin VB.Label Label3 
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
         Height          =   420
         Left            =   765
         TabIndex        =   3
         Top             =   2205
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
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
         Left            =   765
         TabIndex        =   2
         Top             =   1530
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Subjects"
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
         TabIndex        =   1
         Top             =   360
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmSubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim i As Integer
Dim dec, decname, decsub, deccode As Integer



Private Sub cmbBranch_Click()
    cmbSem.Clear
    Set rs = con.Execute("select Sem from Branch where Branchcode='" + cmbBranch.Text + "'")
    If (Not rs.EOF) Then
        x = rs(0)
        For i = 1 To x
         cmbSem.AddItem i
        Next
    End If
    
End Sub

Private Sub cmdadd_Click()
dec = 1
If (cmbBranch.Text = "" Or cmbSem.Text = "" Or txtSubname.Text = "" Or txtSubCode.Text = "") Then
    MsgBox "Missing Fields", vbInformation, " CES"
    dec = 0
Else
    Set rs = con.Execute("select * from Subjects where Branchcode='" + cmbBranch.Text + "' AND Subjectname='" + txtSubname.Text + "'")
    If (Not rs.EOF) Then
        MsgBox "Sorry!! The Specified subject already exists in the given branch. Try another branch name or subject name", vbCritical, " CES"
        dec = 0
    Else
        Set rs = con.Execute("select * from Subjects where Subjectcode='" + txtSubCode.Text + "'")
            If (Not rs.EOF) Then
                MsgBox "Sorry!! The Specified branch code already exists. Branch codes must be unique irrespective of their branch/subject", vbCritical, " CES"
                dec = 0
            Else
                dec = 1
            End If
    End If
End If
rs.Close
If dec = 1 Then
    con.Execute ("insert into Subjects values('" + cmbBranch.Text + "'," + cmbSem.Text + ",'" + txtSubname.Text + "','" + txtSubCode.Text + "')")
    MsgBox "Record added sucessfully", vbInformation, " CES"
    txtSubname.Text = ""
    txtSubCode.Text = ""
    txtSubname.SetFocus
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
Set rs = con.Execute("select * from Branch")
While (Not rs.EOF)
    cmbBranch.AddItem rs(1)
    rs.MoveNext
Wend
rs.Close
End Sub

