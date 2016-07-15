VERSION 5.00
Begin VB.Form frmSelectExam 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Exam"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmSelectExam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7560
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   5655
      Begin VB.ComboBox cmbExID 
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
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3600
         Width           =   2400
      End
      Begin VB.ComboBox cmbSub 
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
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2400
         Width           =   2400
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
         Left            =   2250
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   2400
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1200
         Width           =   2400
      End
      Begin VB.TextBox txtTime 
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
         Height          =   330
         Left            =   2250
         TabIndex        =   0
         Top             =   4200
         Width           =   555
      End
      Begin VB.CommandButton Command1 
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H008080FF&
         Caption         =   "Start Exam"
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
         TabIndex        =   3
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtSubCode 
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
         Height          =   330
         Left            =   2250
         TabIndex        =   2
         Top             =   3000
         Width           =   2400
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Your Exam"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Select ExamID"
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
         Left            =   720
         TabIndex        =   16
         Top             =   3680
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Time"
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
         Left            =   720
         TabIndex        =   11
         Top             =   4200
         Width           =   1140
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "(Time in minutes)"
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
         TabIndex        =   10
         Top             =   4250
         Width           =   1770
      End
      Begin VB.Label Label19 
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
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Branch"
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
         Left            =   720
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Semester"
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
         Left            =   720
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Subject"
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
         Left            =   720
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
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
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSelectExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt As Date
Dim x As Integer

Private Sub cmbBranch_Click()
cmbSem.Clear
Set rs = con.Execute("select Sem from Branch where BranchCode='" + cmbBranch.Text + "'")
    If (Not rs.EOF) Then
        x = rs(0)
        For i = 1 To x
            cmbSem.AddItem i
        Next
    End If
    rs.Close
End Sub

Private Sub cmbExID_Click()
Set rs = con.Execute("select Time from ExamDetails where ExamId='" + cmbExID.Text + "'")
If (Not rs.EOF) Then
    txtTime.Text = rs(0)
End If
rs.Close
End Sub

Private Sub cmbSem_Click()
    cmbSub.Clear
    Set rs = con.Execute("select Subjectname from Subjects where Branchcode='" + cmbBranch.Text + "' and Sem=" + cmbSem.Text + "")
    While (Not rs.EOF)
        cmbSub.AddItem rs(0)
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cmbSub_Click()

Set rs = con.Execute("select Subjectcode from Subjects where Subjectname='" + cmbSub.Text + "'")
    If (Not rs.EOF) Then
        txtSubCode.Text = rs(0)
    End If
    rs.Close
Set rs = con.Execute("select distinct(ExamID) from ExamDetails where BranchCode='" + cmbBranch.Text + _
                    "' and Sem=" + cmbSem.Text + " and SubjectCode='" + txtSubCode.Text + "' and ExDate=#" & Date & "#")
                
    While (Not rs.EOF)
        cmbExID.AddItem rs(0)
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = con.Execute("select Time from ExamDetails where BranchCode='" + cmbBranch.Text + _
                    "' and Sem=" + cmbSem.Text + " and SubjectCode='" + txtSubCode.Text + "' and ExDate=#" & Date & "#")
        If (Not rs.EOF) Then
                    t = rs(0)
        End If
            rs.Close
End Sub

Private Sub cmdStart_Click()
'abc = 1
'If (cmbBranch.Text = "" Or cmbExID.Text = "" Or cmbSem.Text = "" Or cmbSub.Text = "") Then
'    MsgBox "Missing fields, Please fill up all", vbInformation, " Examner"
'Else
'    bcode = cmbBranch.Text
'    sem = cmbSem.Text
'    subcode = txtSubCode.Text
'    exid = cmbExID.Text
'    Unload Me
'    Set rs = con.Execute("select count(*) from Questions where BranchCode='" + bcode + "' and Sem=" & sem & " and SubjectCode='" + subcode + "' and ExamID='" + exid + "'  ")
'If (Not rs.EOF) Then
'    x = rs(0)
'End If
'rs.Close
'If (x < 5) Then
'     MsgBox "Exam not ready!!!!", vbCritical, " CES"
'     abc = 0
'     ter = 1
'Else
'        Load frmExam
'    End If
'End If
If (cmbBranch.Text = "" Or cmbExID.Text = "" Or cmbSem.Text = "" Or cmbSub.Text = "") Then
    MsgBox "Missing fields, Please fill up all", vbInformation, " Examner"
Else
    bcode = cmbBranch.Text
    sem = cmbSem.Text
    subcode = txtSubCode.Text
    exid = cmbExID.Text
    On Error GoTo localerror
    Load frmExam
    
    frmExam.Show
    Unload Me
End If
localerror:
Unload frmExam
Unload frmSelectExam
End Sub

Private Sub Command1_Click()
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

