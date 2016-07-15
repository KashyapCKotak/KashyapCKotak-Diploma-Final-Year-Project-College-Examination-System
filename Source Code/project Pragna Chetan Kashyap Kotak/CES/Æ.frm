VERSION 5.00
Begin VB.Form frmAddQst 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Questions"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   9420
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   10425
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   20055
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   8895
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   19815
         Begin VB.PictureBox txtAns 
            Height          =   1095
            Left            =   2160
            ScaleHeight     =   1035
            ScaleWidth      =   17475
            TabIndex        =   41
            Top             =   7560
            Width           =   17535
         End
         Begin VB.PictureBox txtOpt4 
            Height          =   1095
            Left            =   2160
            ScaleHeight     =   1035
            ScaleWidth      =   17475
            TabIndex        =   40
            Top             =   6240
            Width           =   17535
         End
         Begin VB.PictureBox txtOpt3 
            Height          =   1095
            Left            =   2160
            ScaleHeight     =   1035
            ScaleWidth      =   17475
            TabIndex        =   39
            Top             =   5040
            Width           =   17535
         End
         Begin VB.PictureBox txtOpt2 
            Height          =   1095
            Left            =   2160
            ScaleHeight     =   1035
            ScaleWidth      =   17475
            TabIndex        =   38
            Top             =   3840
            Width           =   17535
         End
         Begin VB.PictureBox txtOpt1 
            Height          =   1095
            Left            =   2160
            ScaleHeight     =   1035
            ScaleWidth      =   17475
            TabIndex        =   37
            Top             =   2640
            Width           =   17535
         End
         Begin VB.PictureBox txtQst 
            Height          =   2415
            Left            =   2160
            ScaleHeight     =   2355
            ScaleWidth      =   17475
            TabIndex        =   36
            Top             =   120
            Width           =   17535
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Answer Key"
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
            Left            =   360
            TabIndex        =   24
            Top             =   7920
            Width           =   1095
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0E0FF&
            Caption         =   "2"
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
            TabIndex        =   23
            Top             =   4200
            Width           =   375
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0E0FF&
            Caption         =   "3"
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
            TabIndex        =   22
            Top             =   5520
            Width           =   375
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0E0FF&
            Caption         =   "4"
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
            TabIndex        =   21
            Top             =   6720
            Width           =   375
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Options"
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
            Left            =   120
            TabIndex        =   20
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0E0FF&
            Caption         =   "1"
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
            TabIndex        =   19
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label lblQno 
            BackColor       =   &H00C0E0FF&
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
            Left            =   1320
            TabIndex        =   18
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Question"
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
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdcancel 
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
         Left            =   12480
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   9960
         Width           =   1455
      End
      Begin VB.CommandButton cmdadd 
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   9960
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   6360
         TabIndex        =   11
         Top             =   480
         Width           =   7440
         Begin VB.Label lblSubCode 
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
            Height          =   255
            Left            =   6435
            TabIndex        =   15
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Code:"
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
            Left            =   5160
            TabIndex        =   14
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label lblSubName 
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
            Left            =   1380
            TabIndex        =   13
            Top             =   0
            Width           =   3705
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            Caption         =   "Subject Name:"
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
            Left            =   150
            TabIndex        =   12
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Add Question"
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
         TabIndex        =   34
         Top             =   0
         Width           =   20055
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2"
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
         Left            =   12480
         TabIndex        =   10
         Top             =   120
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   6090
      Left            =   7560
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
      Begin VB.TextBox txtTime 
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
         TabIndex        =   32
         Top             =   4545
         Width           =   555
      End
      Begin VB.TextBox txtDate 
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
         TabIndex        =   30
         Top             =   3825
         Width           =   2400
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
         TabIndex        =   28
         Top             =   3060
         Width           =   2400
      End
      Begin VB.ComboBox cmbSubjects 
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
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
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
         Left            =   2280
         TabIndex        =   7
         Top             =   1500
         Width           =   2415
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
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H008080FF&
         Caption         =   "Next"
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
         TabIndex        =   5
         Top             =   5220
         Width           =   1335
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
         Left            =   3015
         TabIndex        =   33
         Top             =   4590
         Width           =   1770
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Set Time"
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
         TabIndex        =   31
         Top             =   4545
         Width           =   1140
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Set Date"
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
         TabIndex        =   29
         Top             =   3840
         Width           =   1140
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
         TabIndex        =   27
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1"
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
         TabIndex        =   4
         Top             =   360
         Width           =   615
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
         TabIndex        =   3
         Top             =   2280
         Width           =   1320
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
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
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
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum 26 questons must be entered to take an exam."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5400
      TabIndex        =   35
      Top             =   10560
      Width           =   9135
   End
End
Attribute VB_Name = "frmAddQst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, i As Integer
Dim exid As String
Dim eID As String
Dim Qno As Integer

Private Sub cmbBranch_Click()
cmbSem.Clear
    Set rs = con.Execute("select Sem from Branch where Branchcode='" + cmbBranch.Text + "'")
    If (Not rs.EOF) Then
        X = rs(0)
        For i = 1 To X
         cmbSem.AddItem i
        Next
    End If
    rs.Close
End Sub


Private Sub cmbSem_Click()
    cmbSubjects.Clear
    Set rs = con.Execute("select Subjectname from Subjects where Branchcode='" + cmbBranch.Text + "' and Sem=" + cmbSem.Text + "")
    While (Not rs.EOF)
        cmbSubjects.AddItem rs(0)
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub cmbSubjects_Click()
    Set rs = con.Execute("select Subjectcode from Subjects where Subjectname='" + cmbSubjects.Text + "' and BranchCode='" + cmbBranch.Text + "'")
    If (Not rs.EOF) Then
        txtSubCode.Text = rs(0)
    End If
    rs.Close
End Sub

Private Sub cmdadd_Click()
If (txtAns.Text = "" Or txtOpt1.Text = "" Or txtOpt2.Text = "" Or txtOpt3.Text = "" Or txtOpt4.Text = "" Or txtQst.Text = "") Then
    MsgBox "Missing Fields", vbInformation, "Offline Examiner"
Else
    If (txtOpt1.Text <> txtAns.Text And txtOpt2.Text <> txtAns.Text And txtOpt3.Text <> txtAns.Text And txtOpt4.Text <> txtAns.Text) Then
        MsgBox "Answer key does not match any one of 4 options", vbCritical, "Offline Examiner"
        txtAns.Text = ""
        txtAns.SetFocus
    Else
    exid = cmbBranch.Text + txtDate.Text + txtSubCode.Text
    con.Execute ("insert into Questions values('" + cmbBranch.Text + "'," + cmbSem.Text + _
                ",'" + txtSubCode.Text + "','" + exid + "','" + lblQno.Caption + _
                "','" + txtQst.Text + "','" + txtOpt1.Text + "', '" + txtOpt2.Text + _
                "','" + txtOpt3.Text + "','" + txtOpt4.Text + "','" + txtAns.Text + "')")
    con.Execute ("insert into ExamDetails values('" + cmbBranch.Text + "'," + cmbSem.Text + _
                ",'" + txtSubCode.Text + "','" + exid + "','" + txtDate.Text + "','" + txtTime.Text + "')")
            MsgBox "Question Added Successfully", vbInformation, "Offline Examiner"
            txtAns.Text = ""
            txtOpt1.Text = ""
            txtOpt2.Text = ""
            txtOpt3.Text = ""
            txtOpt4.Text = ""
            txtQst.Text = ""
            txtQst.SetFocus
            lblQno.Caption = lblQno.Caption + 1
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdNext_Click()

If (txtSubCode.Text = "" Or txtDate.Text = "" Or txtTime.Text = "") Then
    MsgBox "Please Select and fill all the fields"
    Frame1.Visible = True
    Frame2.Visible = False
Else
    
    lblSubName.Caption = cmbSubjects.Text
    lblSubCode.Caption = txtSubCode.Text
    eID = cmbBranch.Text + txtDate.Text + txtSubCode.Text
    Set rs = con.Execute("select count(*) from Questions where BranchCode='" + cmbBranch.Text + _
                        "' and Sem=" + cmbSem.Text + " and SubjectCode='" + txtSubCode.Text + _
                        "' and ExamID='" + eID + "'")
    If (Not rs.EOF) Then
        Qno = rs(0)
        Qno = Qno + 1
    Else
        Qno = 1
    End If
    lblQno.Caption = Qno
    Frame1.Visible = False
    Frame2.Visible = True
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = (Asc("'")) Then
    KeyAscii = 0
    MsgBox "Character not allowed", vbInformation
End If
'If KeyAscii = (Asc("-")) Then
'    KeyAscii = 0
'    MsgBox "Character not allowed", vbInformation
'End If
'If KeyAscii = (Asc("*")) Then
'    KeyAscii = 0
'    MsgBox "Character not allowed", vbInformation
'End If
End Sub

Private Sub Form_Load()

'If KeyAscii = (Asc("-")) Then
'    KeyAscii = 0
'    MsgBox "Character not allowed", vbInformation
'End If
'If KeyAscii = (Asc("*")) Then
'    KeyAscii = 0
'    MsgBox "Character not allowed", vbInformation
'End If
BackColor = color
Frame1.Visible = True
Frame2.Visible = False
connectdb
Set rs = con.Execute("select * from Branch")
While (Not rs.EOF)
    cmbBranch.AddItem rs(1)
    rs.MoveNext
Wend
rs.Close
End Sub

