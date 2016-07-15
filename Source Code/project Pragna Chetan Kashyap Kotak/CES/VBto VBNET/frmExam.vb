Imports VB = Microsoft.VisualBasic

Public Class frmExam
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Frame3 As System.Windows.Forms.Panel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Frame4 As System.Windows.Forms.Panel
    Friend WithEvents cmdFinish As System.Windows.Forms.Button
    Friend WithEvents Frame2 As System.Windows.Forms.Panel
    Friend WithEvents lblQno As System.Windows.Forms.Label
    Friend WithEvents Frame1 As System.Windows.Forms.Panel
    Friend WithEvents Option3 As System.Windows.Forms.RadioButton
    Friend WithEvents Option5 As System.Windows.Forms.RadioButton
    Friend WithEvents Option4 As System.Windows.Forms.RadioButton
    Friend WithEvents Option1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmExam))
        Me.Frame3 = New System.Windows.Forms.Panel()
        Me.Timer1 = New System.Windows.Forms.Timer()
        Me.Frame4 = New System.Windows.Forms.Panel()
        Me.cmdFinish = New System.Windows.Forms.Button()
        Me.Frame2 = New System.Windows.Forms.Panel()
        Me.lblQno = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.Option3 = New System.Windows.Forms.RadioButton()
        Me.Option5 = New System.Windows.Forms.RadioButton()
        Me.Option4 = New System.Windows.Forms.RadioButton()
        Me.Option1 = New System.Windows.Forms.RadioButton()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame3
        '
        Me.Frame3.Name = "Frame3"
        Me.Frame3.TabIndex = 10
        Me.Frame3.Location = New System.Drawing.Point(283, 712)
        Me.Frame3.Size = New System.Drawing.Size(745, 58)
        Me.Frame3.BackColor = System.Drawing.Color.FromArgb(CType(4, Byte), CType(157, Byte), CType(12, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Frame4
        '
        Me.Frame4.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdFinish})
        Me.Frame4.Name = "Frame4"
        Me.Frame4.TabIndex = 13
        Me.Frame4.Location = New System.Drawing.Point(1213, 712)
        Me.Frame4.Size = New System.Drawing.Size(155, 58)
        Me.Frame4.BackColor = System.Drawing.Color.FromArgb(CType(220, Byte), CType(6, Byte), CType(1, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdFinish
        '
        Me.cmdFinish.Name = "cmdFinish"
        Me.cmdFinish.TabIndex = 14
        Me.cmdFinish.Location = New System.Drawing.Point(24, 8)
        Me.cmdFinish.Size = New System.Drawing.Size(114, 41)
        Me.cmdFinish.Text = "Finish"
        Me.cmdFinish.BackColor = System.Drawing.Color.FromArgb(CType(254, Byte), CType(30, Byte), CType(24, Byte))
        Me.cmdFinish.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFinish.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Frame2
        '
        Me.Frame2.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblQno})
        Me.Frame2.Name = "Frame2"
        Me.Frame2.TabIndex = 6
        Me.Frame2.Location = New System.Drawing.Point(81, 65)
        Me.Frame2.Size = New System.Drawing.Size(1281, 241)
        Me.Frame2.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'lblQno
        '
        Me.lblQno.Name = "lblQno"
        Me.lblQno.TabIndex = 8
        Me.lblQno.Location = New System.Drawing.Point(55, 30)
        Me.lblQno.Size = New System.Drawing.Size(37, 28)
        Me.lblQno.Text = ""
        Me.lblQno.BackColor = System.Drawing.Color.Transparent
        Me.lblQno.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.lblQno.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Option3, Me.Option5, Me.Option4, Me.Option1})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 1
        Me.Frame1.Location = New System.Drawing.Point(162, 316)
        Me.Frame1.Size = New System.Drawing.Size(1201, 372)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Option3
        '
        Me.Option3.Name = "Option3"
        Me.Option3.TabIndex = 4
        Me.Option3.Location = New System.Drawing.Point(32, 194)
        Me.Option3.Size = New System.Drawing.Size(1157, 74)
        Me.Option3.Text = ""
        Me.Option3.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Option3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Option5
        '
        Me.Option5.Name = "Option5"
        Me.Option5.Visible = False
        Me.Option5.TabIndex = 18
        Me.Option5.Location = New System.Drawing.Point(396, 162)
        Me.Option5.Size = New System.Drawing.Size(50, 25)
        Me.Option5.Text = "Option5"
        Me.Option5.BackColor = System.Drawing.SystemColors.Control
        Me.Option5.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Option4
        '
        Me.Option4.Name = "Option4"
        Me.Option4.TabIndex = 5
        Me.Option4.Location = New System.Drawing.Point(32, 283)
        Me.Option4.Size = New System.Drawing.Size(1157, 74)
        Me.Option4.Text = ""
        Me.Option4.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Option4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option4.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Option1
        '
        Me.Option1.Name = "Option1"
        Me.Option1.TabIndex = 2
        Me.Option1.Location = New System.Drawing.Point(32, 16)
        Me.Option1.Size = New System.Drawing.Size(1157, 74)
        Me.Option1.Text = ""
        Me.Option1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Option1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Option1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label9
        '
        Me.Label9.Name = "Label9"
        Me.Label9.Visible = False
        Me.Label9.TabIndex = 24
        Me.Label9.Location = New System.Drawing.Point(8, 57)
        Me.Label9.Size = New System.Drawing.Size(1352, 114)
        Me.Label9.Text = "          FOUL        RESULT:FAIL"
        Me.Label9.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 39.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label7
        '
        Me.Label7.Name = "Label7"
        Me.Label7.Visible = False
        Me.Label7.TabIndex = 21
        Me.Label7.Location = New System.Drawing.Point(81, 65)
        Me.Label7.Size = New System.Drawing.Size(602, 95)
        Me.Label7.Text = "Label7"
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(220, Byte), CType(6, Byte), CType(1, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label8
        '
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 23
        Me.Label8.Location = New System.Drawing.Point(8, 324)
        Me.Label8.Size = New System.Drawing.Size(139, 268)
        Me.Label8.Text = ""
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Label5
        '
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 17
        Me.Label5.Location = New System.Drawing.Point(8, 744)
        Me.Label5.Size = New System.Drawing.Size(284, 25)
        Me.Label5.Text = "You CANNOT Go To Previous Question"
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(255, Byte), CType(0, Byte))
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 15
        Me.Label3.Location = New System.Drawing.Point(1084, 728)
        Me.Label3.Size = New System.Drawing.Size(147, 41)
        Me.Label3.Text = "Select To View Result And Exit"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmExam
        '
        Me.ClientSize = New System.Drawing.Size(1375, 777)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame3, Me.Frame4, Me.Frame2, Me.Frame1, Me.Label9, Me.Label7, Me.Label8, Me.Label5, Me.Label3})
        Me.Name = "frmExam"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ShowInTaskbar = False
        Me.MinimizeBox = False
        Me.MaximizeBox = True
        Me.Icon = CType(Resources.GetObject("frmExam.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exam"
        Me.Frame3.ResumeLayout(False)
        Me.cmdFinish.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.lblQno.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Option3.ResumeLayout(False)
        Me.Option4.ResumeLayout(False)
        Me.Option1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.Label9.ResumeLayout(False)
        Me.Label5.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    ' VBto upgrade warning: x As Short	OnWrite(VBtoRecordSet)
    Dim x As Short
    Dim ar(25) As Short
    Dim y As Object, i As Short
    Dim Qcnt As Short
    Dim z As Short
    ' VBto upgrade warning: ans As String	OnWrite(VBtoRecordSet)
    Dim ans As String
    Dim selected As String
    Dim mark As Short
    Dim rslt As String
    Dim min As Short
    Dim sec As Integer
    Dim check As Integer
    Dim ter As Short

    Private Declare Function GetActiveWindow Lib "user32"  As Integer

    'Option Explicit
    'Private Type POINTAPI
    '        x As Long
    '        y As Long
    'End Type
    'Dim m_CursorPos As POINTAPI
    'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


    Private Sub cmdFinish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdFinish.Click
'#Const def_cmdFinish_Click = True
#If def_cmdFinish_Click
        ter = 1
        If (mark>=3) Then
            rslt = "Passed"
        Else
            rslt = "Failed"
        End If
        con.Execute(("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')"))
        MsgBox(mark)
        Close()
#End If	' def_cmdFinish_Click
    End Sub

    Private Sub cmdNext_Click()
'#Const def_cmdNext_Click = True
#If def_cmdNext_Click
        If (Option1.Checked=True) Then
            selected = Option1.Text
        ElseIf (Option2.Checked=True) Then
            selected = Option2.Text
        ElseIf (Option3.Checked=True) Then
            selected = Option3.Text
        ElseIf (Option4.Checked=True) Then
            selected = Option4.Text
        Else
            selected = "(UnAnswered)"
        End If
        If (ans=selected) Then
            mark += 1
        End If
        Label8.Text = "Mark: " & mark & "--Correct: " & ans & "--Your Ans.: " & selected
        Option5.Checked = True
        Qcnt += 1
        If (Qcnt<=25) Then
            rs = con.Execute("select * from Questions where BranchCode='" & bcode & "' and Sem=" & sem & " and SubjectCode='" & subcode & "' and ExamID='" & exid & "' and Qno=" & ar(Qcnt-1) & "  ")
            If ( Not rs.EOF) Then
                lblQno.Text = Qcnt
                lblQst.Text = rs(5)
                Option1.Text = rs(6)
                Option2.Text = rs(7)
                Option3.Text = rs(8)
                Option4.Text = rs(9)
                ans = rs(10)
            End If
        Else
            MsgBox("Exam Completed", MsgBoxStyle.Information, " CES")
            cmdNext.Enabled = False
            cmdFinish.Enabled = True
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
        End If
#End If	' def_cmdNext_Click
    End Sub

    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click
        MMControl1.Command = "stop"
        Close()
#End If	' def_Command1_Click
    End Sub
    'Private Sub GetCursor()
    'Dim LonCStat As Long
    ' LonCStat = GetCursorPos&(m_CursorPos)
    'to use this result, the data must be converted into Pixel

    Private Sub frmExam_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Deactivate
'#Const def_Form_Deactivate = True
#If def_Form_Deactivate
        mark = 0
        rslt = "Failed"
        MMControl1.Command = "Play"
        con.Execute(("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')"))
        Frame1.Visible = False
        Frame2.Visible = False
        Me.BackColor = ColorTranslator.FromOle(&H106DC)
        Frame3.Visible = False
        Frame4.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label3.Visible = False
        Label9.Visible = True
        Command1.Visible = True
#End If	' def_Form_Deactivate
    End Sub

    'm_CursorPos.x = m_CursorPos.x * Screen.TwipsPixelX
    'm_CursorPos.y = m_CursorPos.y * Screen.TwipsPixelY
    'End Sub

    Private Sub frmExam_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        rs = con.Execute("select count(*) from Questions where BranchCode='"+bcode+"' and Sem=" & sem & " and SubjectCode='"+subcode+"' and ExamID='"+exid+"'  ")
        If ( Not rs.EOF) Then
            x = rs(0)
        End If
        rs.Close()
        If (x<26) Then
            MsgBox("Exam not ready!!!!", MsgBoxStyle.Critical, " CES")
            abc = 0
            ter = 1
        Else
            examdecide = 1
            cnt += 1
            con.Execute(("update Student set ucounter=" & cnt & " where Username='" & MDIForm1.StatusBar1.Panels.Item(2 - 1).Text & "'"))
            abc = 1
            check = 0
            ter = 0
            MMControl1.Command = "Open"
            ' VBto upgrade warning: i As Short	OnWrite(Integer)
            Dim y As Object, i As Short
            connectdb()
            mark = 0
            Qcnt = 1
            min = 0
            sec = 0
            rs = con.Execute("select count(*) from Questions where BranchCode='"+bcode+"' and Sem=" & sem & " and SubjectCode='"+subcode+"' and ExamID='"+exid+"'  ")
            If ( Not rs.EOF) Then
                x = rs(0)
            End If
            rs.Close()
            'If (x < 26) Then
            '     MsgBox "Exam not ready!!!!", vbCritical, " CES"
            '     abc = 0
            '     ter = 1
            'End If
            If (abc=1) Then
                y = RandomNumbers(x, 2, 25)
                For i = LBound(y) To UBound(y)
                    ar(i) = y(i)
                Next

                rs = con.Execute("select * from Questions where BranchCode='" & bcode & "' and Sem=" & sem & " and SubjectCode='" & subcode & "' and ExamID='" & exid & "' and Qno=" & ar(0) & "  ")
                If ( Not rs.EOF) Then
                    lblQno.Text = Qcnt
                    lblQst.Text = rs(5)
                    Option1.Text = rs(6)
                    Option2.Text = rs(7)
                    Option3.Text = rs(8)
                    Option4.Text = rs(9)
                    ans = rs(10)
                End If
                rs.Close()
                cmdFinish.Enabled = False
                Me.ShowDialog()
            End If
        End If
#End If	' def_Form_Load
    End Sub
    Public Function RandomNumbers(ByVal Upper As Short, Optional ByVal Lower As Short = 1, Optional ByVal HowMany As Short = 1, Optional ByVal Unique As Boolean = True) As Object
        RandomNumbers = 0
'#Const def_RandomNumbers = True
#If def_RandomNumbers
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
            If HowMany>((Upper+1)-(Lower-1)) Then Exit Function
            Dim x As Short
            Dim n As Short
            ' VBto upgrade warning: arrNums As Object --> As Collection
            Dim arrNums() As Collection
            ' VBto upgrade warning: colNumbers As Collection	OnWrite(Integer)
            Dim colNumbers As New Collection

            ReDim arrNums(HowMany-1)
            'First populate the collection
            For x = Lower To Upper
                colNumbers.Add(x)
            Next x
            For x = 0 To HowMany-1
                n = RandomNumber(0, colNumbers.Count+1)
                arrNums(x) = colNumbers(n)
                If Unique Then
                    colNumbers.Remove(n)
                End If
            Next x

            colNumbers = Nothing
            RandomNumbers = arrNums
            Exit Function
         localerror:
            'Justin (just in case)
            RandomNumbers = ""
#End If	' def_RandomNumbers
    End Function

    Public Function RandomNumber(ByVal Upper As Short, ByVal Lower As Short) As Short
        RandomNumber = 0
'#Const def_RandomNumber = True
#If def_RandomNumber
        'Generates a Random Number BETWEEN the LOWER and UPPER values
        Randomize()
        RandomNumber = Int((Upper-Lower+1)*Rnd()+Lower)

#End If	' def_RandomNumber
    End Function



    Private Sub frmExam_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Leave
'#Const def_Form_LostFocus = True
#If def_Form_LostFocus
        mark = 0
        rslt = "Failed"
        MMControl1.Command = "Play"
        con.Execute(("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')"))
        Frame1.Visible = False
        Frame2.Visible = False
        Me.BackColor = ColorTranslator.FromOle(&H106DC)
        Frame3.Visible = False
        Frame4.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label3.Visible = False
        Label9.Visible = True
        Command1.Visible = True
#End If	' def_Form_LostFocus
    End Sub

    Private Sub frmExam_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
'#Const def_Form_Terminate = True
#If def_Form_Terminate
        If ter=0 Then
            If (mark>=10) Then
                rslt = "Passed"
            Else
                rslt = "Failed"
            End If

            con.Execute(("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')"))
            MsgBox(("YOU CANCELED YOUR EXAM BEFORE IT WAS COMPLETED. THE MARKS OF THE QUESTIONS YOU ANSWERED WERE ENTERED IN THE DATABASE"))
            MsgBox(mark)
            Me.Close()
        End If
#End If	' def_Form_Terminate
    End Sub

    Private Sub Form_Unload(ByRef Cancel As Short)
'#Const def_Form_Unload = True
#If def_Form_Unload
        'If examdecide = 1 Then
        If ter=0 Then
            If (mark>=10) Then
                rslt = "Passed"
            Else
                rslt = "Failed"
            End If

            con.Execute(("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')"))
            MsgBox(("YOU CANCELED YOUR EXAM BEFORE IT WAS COMPLETED. THE MARKS OF THE QUESTIONS YOU ANSWERED WERE ENTERED IN THE DATABASE"))
            MsgBox(mark)
            Me.Close()
        End If
        'End If
#End If	' def_Form_Unload
    End Sub

    Private Sub frmExam_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim Cancel As Short = 0

        Form_Unload(Cancel)
        If Cancel <> 0 Then e.Cancel = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
'#Const def_Timer1_Timer = True
#If def_Timer1_Timer
        If abc=0 Then
            Close()
        End If
        If check=0 Then
            sec += 1
        End If

        If sec=60 Then
            min += 1
            sec = 0
        End If
        If min=t Then
            MsgBox("Exam Time out.", MsgBoxStyle.Information, " CES")
            cmdNext.Enabled = False
            cmdFinish.Enabled = True
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            Timer1.Enabled = False
        End If
        lblTime.Text = min & ":" & sec

        If Me.Handle.ToInt32<>GetActiveWindow() Then
            check = 1
            mark = 0
            rslt = "Failed"
            MMControl1.FileName = str & "\Alarmskashyap.wav"
            MMControl1.Command = "Open"
            MMControl1.Command = "Play"
            con.Execute(("insert into Result values('" & uname & "','" & bcode & "'," & sem & ",'" & subcode & "','" & exid & "'," & mark & ",'" & rslt & "')"))
            Frame1.Visible = False
            Frame2.Visible = False
            Me.BackColor = ColorTranslator.FromOle(&H106DC)
            Frame3.Visible = False
            Frame4.Visible = False
            Label4.Visible = False
            Label5.Visible = False
            Label3.Visible = False
            Label9.Visible = True
            Command1.Visible = True
        End If
        If abc=0 Then
            Close()
        End If
#End If	' def_Timer1_Timer
    End Sub



End Class