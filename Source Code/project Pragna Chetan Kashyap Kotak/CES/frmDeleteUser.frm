VERSION 5.00
Begin VB.Form frmDeleteUser 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete User"
   ClientHeight    =   6120
   ClientLeft      =   10365
   ClientTop       =   2970
   ClientWidth     =   8475
   Icon            =   "frmDeleteUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8475
   Begin VB.Frame frmLogin 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
      Begin VB.ComboBox cmbUsername 
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
         ItemData        =   "frmDeleteUser.frx":0442
         Left            =   1800
         List            =   "frmDeleteUser.frx":0444
         TabIndex        =   4
         Top             =   1880
         Width           =   2295
      End
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   3240
         Width           =   1455
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
         TabIndex        =   6
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DELETE TEACHER ACCOUNT"
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
         TabIndex        =   5
         Top             =   360
         Width           =   5535
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
con.Execute ("delete from Userlogin where Username='" + cmbUsername.Text + "'")
MsgBox "User deleted sucessfully!!", vbInformation, " CES"
cmbUsername.Text = ""
End Sub

Private Sub Command1_Click()
Unload Me
Load frmDeleteUser
frmDeleteUser.Show
Command1.Visible = True
End Sub

Private Sub Form_Load()
BackColor = color
connectdb
frmDeleteUser.Top = 1550
frmDeleteUser.Left = 5000
Set rs = con.Execute("select * from Userlogin")
While (Not rs.EOF)
    cmbUsername.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

Private Sub Label3_Click()
Unload Me
Load frmDelSub
frmDelSub.Show
End Sub
