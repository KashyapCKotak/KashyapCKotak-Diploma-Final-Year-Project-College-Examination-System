VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmExam 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Exam"
   ClientHeight    =   11520
   ClientLeft      =   4095
   ClientTop       =   1905
   ClientWidth     =   20400
   Icon            =   "frmExam.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      PlayEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      DeviceType      =   ""
      FileName        =   "C:\Documents and Settings\PragnaChetanKashyap\Desktop\Alarm activationkashyap.wav"
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H000C9D04&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4200
      TabIndex        =   10
      Top             =   10560
      Width           =   11055
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0000FF00&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   10815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15480
      Top             =   360
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000106DC&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   855
      Left            =   18000
      TabIndex        =   13
      Top             =   10560
      Width           =   2295
      Begin VB.CommandButton cmdFinish 
         BackColor       =   &H00181EFE&
         Caption         =   "Finish"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3570
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   19005
      Begin VB.Label lblQst 
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   1560
         TabIndex        =   9
         Top             =   450
         Width           =   17280
      End
      Begin VB.Label lblQno 
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   810
         TabIndex        =   8
         Top             =   450
         Width           =   555
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   180
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "FOUL"
      Height          =   5520
      Left            =   2400
      TabIndex        =   1
      Top             =   4680
      Width           =   17820
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   4
         Top             =   2880
         Width           =   17160
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   5
         Top             =   4200
         Width           =   17160
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   17160
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   17160
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Stop Alarm And Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: Do not switch to any other app as it will rise a siren in exam hall"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   25
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "          FOUL        RESULT:FAIL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   20055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   3975
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000106DC&
      Caption         =   "Label7"
      ForeColor       =   &H000000FF&
      Height          =   1410
      Left            =   1200
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   8925
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   18240
      TabIndex        =   20
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "You CANNOT Go To Previous Question"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   11040
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click For Next Question"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   10680
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select To View Result And Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   16080
      TabIndex        =   15
      Top             =   10800
      Width           =   2175
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   18960
      TabIndex        =   12
      Top             =   360
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Exam"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   540
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   20160
   End
End
Attribute VB_Name = "frmExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer
Dim ar(25) As Integer
Dim y, i As Integer
Dim Qcnt As Integer
Dim z As Integer
Dim ans As String
Dim selected As String
Dim mark As Integer
Dim rslt As String
Dim min As Integer
Dim sec As Long
Dim check As Long
Dim ter As Integer

Private Declare Function GetActiveWindow Lib "user32" () As Long

'Option Explicit
'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type
'Dim m_CursorPos As POINTAPI
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Sub cmdFinish_Click()
ter = 1
If (mark >= 3) Then
    rslt = "Passed"
Else
    rslt = "Failed"
End If
con.Execute ("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')")
MsgBox mark
Unload Me
End Sub

Private Sub cmdNext_Click()
    If (Option1.Value = True) Then
        selected = Option1.Caption
    ElseIf (Option2.Value = True) Then
        selected = Option2.Caption
    ElseIf (Option3.Value = True) Then
        selected = Option3.Caption
    ElseIf (Option4.Value = True) Then
        selected = Option4.Caption
    Else
        selected = "(UnAnswered)"
    End If
    If (ans = selected) Then
        mark = mark + 1
    End If
   Label8.Caption = "Mark: " & mark & "--Correct: " & ans & "--Your Ans.: " & selected
    Option5.Value = True
    Qcnt = Qcnt + 1
    If (Qcnt <= 25) Then
        Set rs = con.Execute("select * from Questions where BranchCode='" & bcode & "' and Sem=" & sem & " and SubjectCode='" & subcode & "' and ExamID='" & exid & "' and Qno=" & ar(Qcnt - 1) & "  ")
        If (Not rs.EOF) Then
            lblQno.Caption = Qcnt
            lblQst.Caption = rs(5)
            Option1.Caption = rs(6)
            Option2.Caption = rs(7)
            Option3.Caption = rs(8)
            Option4.Caption = rs(9)
            ans = rs(10)
        End If
    Else
        MsgBox "Exam Completed", vbInformation, " CES"
        cmdNext.Enabled = False
        cmdFinish.Enabled = True
        Option1.Enabled = False
        Option2.Enabled = False
        Option3.Enabled = False
        Option4.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
MMControl1.Command = "stop"
Unload Me
End Sub
'Private Sub GetCursor()
'Dim LonCStat As Long
   ' LonCStat = GetCursorPos&(m_CursorPos)
    'to use this result, the data must be converted into Pixel

Private Sub Form_Deactivate()
mark = 0
rslt = "Failed"
MMControl1.Command = "Play"
con.Execute ("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')")
Frame1.Visible = False
Frame2.Visible = False
frmExam.BackColor = &H106DC
Frame3.Visible = False
Frame4.Visible = False
Label4.Visible = False
Label5.Visible = False
Label3.Visible = False
Label9.Visible = True
Command1.Visible = True
End Sub

    'm_CursorPos.x = m_CursorPos.x * Screen.TwipsPixelX
    'm_CursorPos.y = m_CursorPos.y * Screen.TwipsPixelY
'End Sub

Private Sub Form_Load()
Set rs = con.Execute("select count(*) from Questions where BranchCode='" + bcode + "' and Sem=" & sem & " and SubjectCode='" + subcode + "' and ExamID='" + exid + "'  ")
If (Not rs.EOF) Then
    x = rs(0)
End If
rs.Close
If (x < 26) Then
     MsgBox "Exam not ready!!!!", vbCritical, " CES"
     abc = 0
     ter = 1
Else
examdecide = 1
cnt = cnt + 1
    con.Execute ("update Student set ucounter=" & cnt & " where Username='" & MDIForm1.StatusBar1.Panels(2).Text & "'")
abc = 1
check = 0
ter = 0
MMControl1.Command = "Open"
Dim y, i As Integer
connectdb
mark = 0
Qcnt = 1
min = 0
sec = 0
Set rs = con.Execute("select count(*) from Questions where BranchCode='" + bcode + "' and Sem=" & sem & " and SubjectCode='" + subcode + "' and ExamID='" + exid + "'  ")
If (Not rs.EOF) Then
    x = rs(0)
End If
rs.Close
'If (x < 26) Then
'     MsgBox "Exam not ready!!!!", vbCritical, " CES"
'     abc = 0
'     ter = 1
'End If
If (abc = 1) Then
y = RandomNumbers(x, 2, 25)
For i = LBound(y) To UBound(y)
ar(i) = y(i)
Next
 
Set rs = con.Execute("select * from Questions where BranchCode='" & bcode & "' and Sem=" & sem & " and SubjectCode='" & subcode & "' and ExamID='" & exid & "' and Qno=" & ar(0) & "  ")
If (Not rs.EOF) Then
    lblQno.Caption = Qcnt
    lblQst.Caption = rs(5)
    Option1.Caption = rs(6)
    Option2.Caption = rs(7)
    Option3.Caption = rs(8)
    Option4.Caption = rs(9)
    ans = rs(10)
End If
rs.Close
cmdFinish.Enabled = False
frmExam.Show vbModal
End If
End If
End Sub
Public Function RandomNumbers(Upper As Integer, _
   Optional Lower As Integer = 1, _
   Optional HowMany As Integer = 1, _
   Optional Unique As Boolean = True) As Variant
'*******************************************************
    'This Function generates random array of
    'Numbers between Lower & Upper
    'In Addition parameters can include whether
    'UNIQUE values are required
 
   'Note the Result is INCLUSIVE of the Range

    'Debug Example:
    'x = RandomNumbers(49, 1, 7)
    'For n = LBound(x) To UBound(x): Debug.Print x(n);: Next n
    'WARNING HowMany MUST be greater than (Higher - Lower)
    '******************************************************

    On Error GoTo localerror
    If HowMany > ((Upper + 1) - (Lower - 1)) Then Exit Function
    Dim x           As Integer
    Dim n           As Integer
    Dim arrNums()   As Variant
    Dim colNumbers  As New Collection
    
    ReDim arrNums(HowMany - 1)
    With colNumbers
        'First populate the collection
        For x = Lower To Upper
            .Add x
        Next x
        For x = 0 To HowMany - 1
            n = RandomNumber(0, colNumbers.Count + 1)
            arrNums(x) = colNumbers(n)
            If Unique Then
                colNumbers.Remove n
            End If
        Next x
    End With
    Set colNumbers = Nothing
    RandomNumbers = arrNums
Exit Function
localerror:
    'Justin (just in case)
    RandomNumbers = ""
End Function

Public Function RandomNumber(Upper As Integer, _
     Lower As Integer) As Integer
    'Generates a Random Number BETWEEN the LOWER and UPPER values
    Randomize
    RandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
    
End Function



Private Sub Form_LostFocus()
mark = 0
rslt = "Failed"
MMControl1.Command = "Play"
con.Execute ("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')")
Frame1.Visible = False
Frame2.Visible = False
frmExam.BackColor = &H106DC
Frame3.Visible = False
Frame4.Visible = False
Label4.Visible = False
Label5.Visible = False
Label3.Visible = False
Label9.Visible = True
Command1.Visible = True
End Sub

Private Sub Form_Terminate()
If ter = 0 Then
If (mark >= 10) Then
    rslt = "Passed"
Else
    rslt = "Failed"
End If

con.Execute ("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')")
MsgBox ("YOU CANCELED YOUR EXAM BEFORE IT WAS COMPLETED. THE MARKS OF THE QUESTIONS YOU ANSWERED WERE ENTERED IN THE DATABASE")
MsgBox mark
Unload frmExam
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If examdecide = 1 Then
If ter = 0 Then
If (mark >= 10) Then
    rslt = "Passed"
Else
    rslt = "Failed"
End If

con.Execute ("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')")
MsgBox ("YOU CANCELED YOUR EXAM BEFORE IT WAS COMPLETED. THE MARKS OF THE QUESTIONS YOU ANSWERED WERE ENTERED IN THE DATABASE")
MsgBox mark
Unload frmExam
End If
'End If
End Sub

Private Sub Timer1_Timer()
If abc = 0 Then
Unload Me
End If
If check = 0 Then
    sec = sec + 1
    End If

    If sec = 60 Then
        min = min + 1
        sec = 0
    End If
    If min = t Then
        MsgBox "Exam Time out.", vbInformation, " CES"
        cmdNext.Enabled = False
        cmdFinish.Enabled = True
        Option1.Enabled = False
        Option2.Enabled = False
        Option3.Enabled = False
        Option4.Enabled = False
        Timer1.Enabled = False
    End If
    lblTime.Caption = min & ":" & sec

        If frmExam.hWnd <> GetActiveWindow() Then
        check = 1
            mark = 0
rslt = "Failed"
MMControl1.FileName = str & "\Alarmskashyap.wav"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
con.Execute ("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')")
Frame1.Visible = False
Frame2.Visible = False
frmExam.BackColor = &H106DC
Frame3.Visible = False
Frame4.Visible = False
Label4.Visible = False
Label5.Visible = False
Label3.Visible = False
Label9.Visible = True
Command1.Visible = True
        End If
        If abc = 0 Then
        Unload Me
        End If
End Sub


