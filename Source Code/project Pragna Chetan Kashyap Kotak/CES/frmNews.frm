VERSION 5.00
Begin VB.Form frmNews 
   BackColor       =   &H0080C0FF&
   Caption         =   "Edit News"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "frmNews.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   20250
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   20300
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   20300
   End
   Begin VB.TextBox text1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   20300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT NEWS-LINES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   20295
   End
End
Attribute VB_Name = "frmNews"
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
Open str & "\news1.txt" For Output As #iFile
Print #iFile, Text1.Text
Close #iFile
'  strFilename = str & "\timetable.txt"
'  iFile = FreeFile
'  Open strFilename For Output As #iFile
'  'StrConv(InputB(LOF(iFile), iFile), vbUnicode) = strTheData
'  Close #iFile
'  text1.Text = strTheData
iFile = FreeFile
Open str & "\news2.txt" For Output As #iFile
Print #iFile, Text2.Text
Close #iFile

iFile = FreeFile
Open str & "\news3.txt" For Output As #iFile
Print #iFile, Text3.Text
Close #iFile
MDIForm1.Text1.Refresh
MDIForm1.Text2.Refresh
MDIForm1.Text3.Refresh
Unload Me
End Sub

Private Sub frmNews_Change()

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
Dim iFile As Long
Dim strFilename As String
Dim strTheData As String

strFilename = str & "\news1.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text1.Text = strTheData
  
strFilename = str & "\news2.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text2.Text = strTheData
  
  strFilename = str & "\news3.txt"
  iFile = FreeFile
  Open strFilename For Input As #iFile
  strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
  Text3.Text = strTheData
  
End Sub

