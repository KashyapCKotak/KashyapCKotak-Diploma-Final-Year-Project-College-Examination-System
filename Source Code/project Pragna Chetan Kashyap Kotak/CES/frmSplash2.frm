VERSION 5.00
Begin VB.Form frmSplash2 
   BackColor       =   &H000040C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5880
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "GO BACK"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "EXIT"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   7080
      Begin VB.Label lblWarning 
         BackColor       =   &H80000014&
         Caption         =   " Warning: Copyright Protected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   3840
         Width           =   5295
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Version: 3.0.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5445
         TabIndex        =   2
         Top             =   3720
         Width           =   1560
      End
      Begin VB.Image imgLogo 
         Height          =   1665
         Left            =   120
         Picture         =   "frmSplash2.frx":0442
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "College Examination System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   4380
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSplash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()

End Sub

