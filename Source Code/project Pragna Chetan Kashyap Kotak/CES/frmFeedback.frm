VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFeedback 
   BackColor       =   &H000080FF&
   Caption         =   "FEEDBACK"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8010
   Icon            =   "frmFeedback.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8010
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "OK"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin RichTextLib.RichTextBox text1 
         Height          =   2655
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
         _Version        =   393217
         Enabled         =   0   'False
         ScrollBars      =   3
         TextRTF         =   $"frmFeedback.frx":0442
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter your name before feedback text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   3840
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FEEDBACK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String
'data = text1.Text
 'text1.Tag = data
iFile = FreeFile
Open str & "\feedback.txt" For Output As #iFile
Print #iFile, Text1.Text
Close #iFile
'  strFilename = str & "\timetable.txt"
'  iFile = FreeFile
'  Open strFilename For Output As #iFile
'  'StrConv(InputB(LOF(iFile), iFile), vbUnicode) = strTheData
'  Close #iFile
'  text1.Text = strTheData

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
frmFeedback.Width = 8250
BackColor = color
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String
If xyz = 15 Then
    Label2.Visible = True
   Command1.Enabled = True
  frmFeedback.Text1.Enabled = True
  ' data = text1.Tag
   strFilename = str & "\feedback.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text1.Text = strTheData
   'text1.Text = data
ElseIf xyz = 0 Then
    Command1.Enabled = True
    frmFeedback.Text1.Enabled = False
     'data = text1.Tag
    'text1.Text = data
  strFilename = str & "\feedback.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
 strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
 Close #iFile
  Text1.Text = strTheData
  'frmTT.text1.Enabled = False
'  ElseIf xyz = 0 Then
'  Label2.Visible = False
'  text1.Enabled = False
  
End If
End Sub
