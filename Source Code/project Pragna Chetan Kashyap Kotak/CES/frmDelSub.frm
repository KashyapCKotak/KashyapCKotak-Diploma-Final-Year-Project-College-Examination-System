VERSION 5.00
Begin VB.Form frmDelSub 
   BackColor       =   &H000080FF&
   Caption         =   "Delete Subject"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   Icon            =   "frmDelSub.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   8505
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H008080FF&
         Caption         =   "Delete"
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
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox cmbSubjectcode 
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
         ItemData        =   "frmDelSub.frx":0442
         Left            =   1800
         List            =   "frmDelSub.frx":0444
         TabIndex        =   4
         Top             =   1880
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4200
         TabIndex        =   3
         Top             =   1875
         Width           =   735
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
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DELETE SUBJECT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmDelSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
con.Execute ("delete from Subjects where Subjectcode='" + cmbSubjectcode.Text + "'")
MsgBox "User deleted sucessfully!!", vbInformation, " CES"
cmbSubjectcode.Text = ""
End Sub

Private Sub Form_Load()
BackColor = color
connectdb
Width = 8625
Height = 6885
Set rs = con.Execute("select distinct Subjectcode from Subjects")
While (Not rs.EOF)
    cmbSubjectcode.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

Private Sub Label3_Click()
Unload Me
Load frmDelSub
frmDelSub.Width = 8745
hieght = 6945
frmDelSub.Top = 0
frmDelSub.Left = -60
frmDelSub.Show

End Sub
