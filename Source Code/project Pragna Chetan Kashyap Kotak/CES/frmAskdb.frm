VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAskdb 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Source"
   ClientHeight    =   5475
   ClientLeft      =   4635
   ClientTop       =   3315
   ClientWidth     =   5805
   Icon            =   "frmAskdb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5805
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   33023
      ForeColor       =   8421631
      TabCaption(0)   =   "Local"
      TabPicture(0)   =   "frmAskdb.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Remote"
      TabPicture(1)   =   "frmAskdb.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(2)=   "Text2"
      Tab(1).Control(3)=   "Picture2"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "ACCEPT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73560
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   -72600
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         Height          =   975
         Left            =   -74880
         Picture         =   "frmAskdb.frx":047A
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "ACCEPT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   4095
      End
      Begin VB.PictureBox Picture1 
         Height          =   975
         Index           =   0
         Left            =   3480
         Picture         =   "frmAskdb.frx":0F49
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Path:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Enter IP add.:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73800
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Path:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select the location of the Database"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmAskdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub Check1_Click()

End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Command1_Click()
str = Text1.Text
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String
'data = text1.Text
 'text1.Tag = data
iFile = FreeFile
Open App.Path & "\dbsource.txt" For Output As #iFile
Print #iFile, str
Close #iFile
dbloc = "Local"
iFile = FreeFile
Open App.Path & "\loc.txt" For Output As #iFile
Print #iFile, dbloc
Close #iFile
MDIForm1.StatusBar3.Panels(1).Text = dbloc & " Data Source"


strFilename = str & "\news1.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  MDIForm1.Text1.Text = strTheData
  
strFilename = str & "\news2.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  MDIForm1.Text2.Text = strTheData
  
strFilename = str & "\news3.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  MDIForm1.Text3.Text = strTheData


Unload Me
End Sub

Private Sub Command2_Click()
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String
str = "\\" & Text2.Text & Text3.Text
'data = text1.Text
 'text1.Tag = data
iFile = FreeFile
Open str & "\dbsource.txt" For Output As #iFile
Print #iFile, str
Close #iFile
dbloc = "Remote"
iFile = FreeFile
Open str & "\loc.txt" For Output As #iFile
Print #iFile, dbloc
Close #iFile
MDIForm1.StatusBar3.Panels(1).Text = dbloc & " Data Source"

strFilename = str & "\news1.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  MDIForm1.Text1.Text = strTheData
  
strFilename = str & "\news2.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  MDIForm1.Text2.Text = strTheData
  
strFilename = str & "\news3.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  MDIForm1.Text3.Text = strTheData


Unload Me
End Sub

