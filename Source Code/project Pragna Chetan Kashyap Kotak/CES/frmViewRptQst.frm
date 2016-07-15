VERSION 5.00
Begin VB.Form frmViewRptQst 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Questions"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   Icon            =   "frmViewRptQst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7395
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   4575
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
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
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
         Left            =   1800
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW REPORT"
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
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Select ExamId"
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
         TabIndex        =   3
         Top             =   1360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmViewRptQst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Private Sub cmdNext_Click()

End Sub

Private Sub cmdVwRpt_Click()
x = cmbBranch.Text
If (DataEnvironment1.rsCommand2.State = 1) Then
    DataEnvironment1.rsCommand2.Close
Else
DataEnvironment1.Command2 (x)
Load DataReportQstExID
DataReportQstExID.Show
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
Set rs = con.Execute("select distinct(ExamID) from ExamDetails")
While (Not rs.EOF)
    cmbBranch.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub
