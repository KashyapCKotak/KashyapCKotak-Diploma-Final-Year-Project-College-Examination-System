VERSION 5.00
Begin VB.Form frmTT 
   BackColor       =   &H000080FF&
   Caption         =   "Time Table"
   ClientHeight    =   9150
   ClientLeft      =   1425
   ClientTop       =   825
   ClientWidth     =   11925
   Icon            =   "frmTT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   11925
   Begin VB.PictureBox Picture1 
      Height          =   1165
      Left            =   240
      Picture         =   "frmTT.frx":0442
      ScaleHeight     =   1110
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   7920
      Width           =   3730
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   11415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL THE BEST"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   8160
      Width           =   9615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIME TABLE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmTT"
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
Open str & "\timetable.txt" For Output As #iFile
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
frmTT.Width = 12165
frmTT.Height = 9690
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String
If sh = 1 Then
   Command1.Enabled = True
   frmTT.Text1.Enabled = True
  ' data = text1.Tag
   strFilename = str & "\timetable.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text1.Text = strTheData
   'text1.Text = data
Else
    Command1.Enabled = True
     data = Text1.Tag
    Text1.Text = data
  strFilename = str & "\timetable.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
 strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
 Close #iFile
  Text1.Text = strTheData
  frmTT.Text1.Enabled = False
  
End If
Width = 12960
Height = 9960
Top = 0
Left = 3510
'Dim iFile As Long
  'Dim strFilename As String
  'Dim strTheData As String

  'strFilename = "C:\Documents and Settings\PragnaChetanKashyap.KOTAK-B43F5C7CD\Desktop\OfflineExaminer\Offline Examiner\timetable.txt"

  'iFile = FreeFile

  'Open strFilename For Input As #iFile
  ' strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  'Close #iFile
  'text1.Text = strTheData


End Sub

