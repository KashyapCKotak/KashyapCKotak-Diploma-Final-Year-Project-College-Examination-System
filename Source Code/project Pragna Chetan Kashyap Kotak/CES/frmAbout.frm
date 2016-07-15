VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FF0000&
   Caption         =   "About"
   ClientHeight    =   8445
   ClientLeft      =   3825
   ClientTop       =   645
   ClientWidth     =   12765
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12765
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   12255
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H008080FF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6480
         Width           =   2655
      End
      Begin VB.PictureBox Picture1 
         Height          =   3375
         Left            =   240
         Picture         =   "frmAbout.frx":0442
         ScaleHeight     =   3315
         ScaleWidth      =   4155
         TabIndex        =   2
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF80&
         Caption         =   "COLLEGE EXAMINATION SYSTEM"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   2270
         Width           =   6495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " kckotak99@gmail.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5640
         TabIndex        =   5
         Top             =   5400
         Width           =   6255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                  CONTACT AT:"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   5400
         Width           =   11655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmAbout.frx":2DD20
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3375
         Left            =   4560
         TabIndex        =   3
         Top             =   1800
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "College Examination System"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12015
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Unload Me
End Sub

Private Sub Command1_Click()

Dim iFile As Long
  Dim strFilename As String
  Dim strTheData As String

  strFilename = "C:\Documents and Settings\PragnaChetanKashyap.KOTAK-B43F5C7CD\Desktop\OfflineExaminer12\Offline CES\kck.txt"

  iFile = FreeFile

  Open strFilename For Input As #iFile
   strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text1.Text = strTheData


End Sub

Private Sub Form_Load()
Width = 12885
Height = 8955
Top = 195
Left = 3765
End Sub
