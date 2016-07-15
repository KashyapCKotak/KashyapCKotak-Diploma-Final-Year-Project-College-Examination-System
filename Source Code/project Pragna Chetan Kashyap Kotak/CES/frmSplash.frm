VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H000040C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6750
      Top             =   4185
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   225
      ScaleHeight     =   330
      ScaleWidth      =   6855
      TabIndex        =   5
      Top             =   4635
      Width           =   6855
      Begin VB.Image Image1 
         Height          =   285
         Left            =   0
         Picture         =   "frmSplash.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   3990
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
         TabIndex        =   1
         Top             =   3660
         Width           =   4455
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
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Platform: Windows XP and Later"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   3
         Top             =   3240
         Width           =   4890
      End
      Begin VB.Label lblProductName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "College Examination System "
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
         Left            =   2640
         TabIndex        =   4
         Top             =   0
         Width           =   4440
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   9030
         Left            =   0
         Picture         =   "frmSplash.frx":076B
         Top             =   120
         Width           =   13845
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   1200
      TabIndex        =   8
      Top             =   4200
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Files:"
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1050
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim i As Integer
Option Explicit

Private Sub Form_Load()
    
    File1.FileName = str
    x = File1.ListCount
End Sub

Private Sub Frame1_Click()
   
    Load MDIForm1
    MDIForm1.Show
    Unload Me
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Timer1_Timer()
time2 = time2 + 1
If (Image1.Left <= 6435) Then
        Image1.Left = Image1.Left + 100
    Else
        Image1.Left = 0
End If
If (i <= x) Then
Label2.Caption = File1.List(i)
    i = i + 1
Else
    Load MDIForm1
    MDIForm1.Show
    Unload Me
End If
If time2 = 50 Then
    Load MDIForm1
    MDIForm1.Show
    Unload Me
    End If

End Sub
