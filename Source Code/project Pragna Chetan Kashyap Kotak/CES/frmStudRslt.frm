VERSION 5.00
Begin VB.Form frmStudRslt 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Result"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   Icon            =   "frmStudRslt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7245
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   4830
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H008080FF&
         Caption         =   "Load ExamID"
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
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdVwRpt 
         BackColor       =   &H008080FF&
         Caption         =   "View Report"
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
         TabIndex        =   5
         Top             =   3240
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
         Height          =   375
         Left            =   1890
         TabIndex        =   3
         Top             =   1200
         Width           =   2715
      End
      Begin VB.ComboBox CmbExamID 
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
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2640
         Width           =   2715
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW RESULT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   4770
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   4
         Top             =   2680
         Width           =   1500
      End
      Begin VB.Label Label1 
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
         Height          =   420
         Left            =   240
         TabIndex        =   1
         Top             =   1215
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmStudRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As String

Private Sub cmdLoad_Click()
If (txtUsername.Text = "") Then
    MsgBox "Enter Username", vbInformation, " CES"
Else
   Set rs = con.Execute("select distinct(ExamID) from Result where Username='" + txtUsername.Text + "' ")
While (Not rs.EOF)
    CmbExamID.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End If
End Sub

Private Sub cmdVwRpt_Click()
a = txtUsername.Text
b = CmbExamID.Text
If (DataEnvironment1.rsCommand3.State = 1) Then
    DataEnvironment1.rsCommand3.Close
Else
   DataEnvironment1.Command3 (b)
    Load DataReportStudRslt
    DataReportStudRslt.Show
    
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
connectdb
BackColor = color
End Sub
