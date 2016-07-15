VERSION 5.00
Begin VB.Form frmAsk 
   BackColor       =   &H000080FF&
   Caption         =   "Re-register"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   3270
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "No"
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Yes"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Do you want to re-register student?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
connectdb
 Set rs = con.Execute("select ucounter from Student where Username='" + txtUname.Text + "'")
    If (Not rs.EOF) Then
        cnt = rs(0)
    End If
    rs.Close
    If (cnt >= 3) Then
        con.Execute ("update Student set ucounter=" & n & " where Username='" & txtUname.Text & "'")
        MsgBox "Registered successfully! Login and attend the exam", vbInformation, "Offline Examiner"
    ElseIf (cnt < 3) Then
        MsgBox "Username already exist! please try another", vbInformation, "Offline Examiner"
        dec = 0
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
hieght = 3315
Width = 3510
connectdb
End Sub
