Imports VB = Microsoft.VisualBasic

Public Class frmSelectExam
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
    Friend WithEvents Frame1 As System.Windows.Forms.Panel
    Friend WithEvents cmbSub As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSem As System.Windows.Forms.ComboBox
    Friend WithEvents txtTime As System.Windows.Forms.TextBox
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents txtSubCode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSelectExam))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmbSub = New System.Windows.Forms.ComboBox()
        Me.cmbSem = New System.Windows.Forms.ComboBox()
        Me.txtTime = New System.Windows.Forms.TextBox()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.txtSubCode = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmbSub, Me.cmbSem, Me.txtTime, Me.cmdStart, Me.txtSubCode, Me.Label10, Me.Label19, Me.Label2, Me.Label3, Me.Label5})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 1
        Me.Frame1.Location = New System.Drawing.Point(65, 57)
        Me.Frame1.Size = New System.Drawing.Size(381, 397)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmbSub
        '
        Me.cmbSub.Name = "cmbSub"
        Me.cmbSub.TabIndex = 14
        Me.cmbSub.Location = New System.Drawing.Point(152, 162)
        Me.cmbSub.Size = New System.Drawing.Size(162, 23)
        Me.cmbSub.Text = ""
        Me.cmbSub.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSub.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSub.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSub.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmbSem
        '
        Me.cmbSem.Name = "cmbSem"
        Me.cmbSem.TabIndex = 13
        Me.cmbSem.Location = New System.Drawing.Point(152, 121)
        Me.cmbSem.Size = New System.Drawing.Size(162, 23)
        Me.cmbSem.Text = ""
        Me.cmbSem.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSem.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSem.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtTime
        '
        Me.txtTime.Name = "txtTime"
        Me.txtTime.Enabled = False
        Me.txtTime.TabIndex = 0
        Me.txtTime.Location = New System.Drawing.Point(152, 283)
        Me.txtTime.Size = New System.Drawing.Size(37, 22)
        Me.txtTime.Text = ""
        Me.txtTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTime.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdStart
        '
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.TabIndex = 3
        Me.cmdStart.Location = New System.Drawing.Point(65, 340)
        Me.cmdStart.Size = New System.Drawing.Size(98, 25)
        Me.cmdStart.Text = "Start Exam"
        Me.cmdStart.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdStart.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStart.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtSubCode
        '
        Me.txtSubCode.Name = "txtSubCode"
        Me.txtSubCode.Enabled = False
        Me.txtSubCode.TabIndex = 2
        Me.txtSubCode.Location = New System.Drawing.Point(152, 202)
        Me.txtSubCode.Size = New System.Drawing.Size(162, 22)
        Me.txtSubCode.Text = ""
        Me.txtSubCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubCode.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label10
        '
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 11
        Me.Label10.Location = New System.Drawing.Point(49, 283)
        Me.Label10.Size = New System.Drawing.Size(77, 25)
        Me.Label10.Text = "Total Time"
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label19
        '
        Me.Label19.Name = "Label19"
        Me.Label19.TabIndex = 9
        Me.Label19.Location = New System.Drawing.Point(49, 202)
        Me.Label19.Size = New System.Drawing.Size(86, 25)
        Me.Label19.Text = "Subject Code"
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 7
        Me.Label2.Location = New System.Drawing.Point(49, 121)
        Me.Label2.Size = New System.Drawing.Size(122, 25)
        Me.Label2.Text = "Select Semester"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 6
        Me.Label3.Location = New System.Drawing.Point(49, 162)
        Me.Label3.Size = New System.Drawing.Size(122, 25)
        Me.Label3.Text = "Select Subject"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label5
        '
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 5
        Me.Label5.Location = New System.Drawing.Point(324, 24)
        Me.Label5.Size = New System.Drawing.Size(41, 25)
        Me.Label5.Text = ""
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmSelectExam
        '
        Me.ClientSize = New System.Drawing.Size(510, 508)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmSelectExam"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmSelectExam.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Select Exam"
        Me.cmbSub.ResumeLayout(False)
        Me.cmbSem.ResumeLayout(False)
        Me.txtTime.ResumeLayout(False)
        Me.cmdStart.ResumeLayout(False)
        Me.txtSubCode.ResumeLayout(False)
        Me.Label10.ResumeLayout(False)
        Me.Label19.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label5.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Dim dt As Date
    ' VBto upgrade warning: x As Short	OnWrite(VBtoRecordSet)
    Dim x As Short

    Private Sub cmbBranch_Click()
'#Const def_cmbBranch_Click = True
#If def_cmbBranch_Click
        cmbSem.Items.Clear()
        rs = con.Execute("select Sem from Branch where BranchCode='"+cmbBranch.Text+"'")
        If ( Not rs.EOF) Then
            x = rs(0)
            For i = 1 To x
                cmbSem.Items.Add(i)
            Next
        End If
        rs.Close()
#End If	' def_cmbBranch_Click
    End Sub

    Private Sub cmbExID_Click()
'#Const def_cmbExID_Click = True
#If def_cmbExID_Click
        rs = con.Execute("select Time from ExamDetails where ExamId='"+cmbExID.Text+"'")
        If ( Not rs.EOF) Then
            txtTime.Text = rs(0)
        End If
        rs.Close()
#End If	' def_cmbExID_Click
    End Sub

    Private Sub cmbSem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSem.SelectedIndexChanged
'#Const def_cmbSem_Click = True
#If def_cmbSem_Click
        cmbSub.Items.Clear()
        rs = con.Execute("select Subjectname from Subjects where Branchcode='"+cmbBranch.Text+"' and Sem="+cmbSem.Text+"")
        While ( Not rs.EOF)
            cmbSub.Items.Add(rs(0))
            rs.MoveNext()
        End While
        rs.Close()
#End If	' def_cmbSem_Click
    End Sub

    Private Sub cmbSub_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSub.SelectedIndexChanged
'#Const def_cmbSub_Click = True
#If def_cmbSub_Click

        rs = con.Execute("select Subjectcode from Subjects where Subjectname='"+cmbSub.Text+"'")
        If ( Not rs.EOF) Then
            txtSubCode.Text = rs(0)
        End If
        rs.Close()
        rs = con.Execute("select distinct(ExamID) from ExamDetails where BranchCode='"+cmbBranch.Text+"' and Sem="+cmbSem.Text+" and SubjectCode='"+txtSubCode.Text+"' and ExDate=#" & Today & "#")

        While ( Not rs.EOF)
            cmbExID.Items.Add(rs(0))
            rs.MoveNext()
        End While

        rs.Close()
        rs = con.Execute("select Time from ExamDetails where BranchCode='"+cmbBranch.Text+"' and Sem="+cmbSem.Text+" and SubjectCode='"+txtSubCode.Text+"' and ExDate=#" & Today & "#")
        If ( Not rs.EOF) Then
            t = rs(0)
        End If
        rs.Close()
#End If	' def_cmbSub_Click
    End Sub

    Private Sub cmdStart_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdStart.Click
'#Const def_cmdStart_Click = True
#If def_cmdStart_Click
            'abc = 1
            'If (cmbBranch.Text = "" Or cmbExID.Text = "" Or cmbSem.Text = "" Or cmbSub.Text = "") Then
            '    MsgBox "Missing fields, Please fill up all", vbInformation, " Examner"
            'Else
            '    bcode = cmbBranch.Text
            '    sem = cmbSem.Text
            '    subcode = txtSubCode.Text
            '    exid = cmbExID.Text
            '    Unload Me
            '    Set rs = con.Execute("select count(*) from Questions where BranchCode='" + bcode + "' and Sem=" & sem & " and SubjectCode='" + subcode + "' and ExamID='" + exid + "'  ")
            'If (Not rs.EOF) Then
            '    x = rs(0)
            'End If
            'rs.Close
            'If (x < 5) Then
            '     MsgBox "Exam not ready!!!!", vbCritical, " CES"
            '     abc = 0
            '     ter = 1
            'Else
            '        Load frmExam
            '    End If
            'End If
            If (cmbBranch.Text="" Or cmbExID.Text="" Or cmbSem.Text="" Or cmbSub.Text="") Then
                MsgBox("Missing fields, Please fill up all", MsgBoxStyle.Information, " Examner")
            Else
                bcode = cmbBranch.Text
                sem = cmbSem.Text
                subcode = txtSubCode.Text
                exid = cmbExID.Text
                On Error GoTo localerror
                LoadUnUsed(frmExam)

                ShowModeless(frmExam)
                Close()
            End If
         localerror:
            frmExam.Close()
            Me.Close()
#End If	' def_cmdStart_Click
    End Sub

    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click
        Close()
#End If	' def_Command1_Click
    End Sub

    Private Sub frmSelectExam_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
'#Const def_Form_KeyPress = True
#If def_Form_KeyPress
        If e.KeyAscii=(Asc("'")) Then
            e.KeyAscii = 0
            MsgBox("Character not allowed", MsgBoxStyle.Information)
        End If
        If e.KeyAscii=(Asc("-")) Then
            e.KeyAscii = 0
            MsgBox("Character not allowed", MsgBoxStyle.Information)
        End If
        If e.KeyAscii=(Asc("*")) Then
            e.KeyAscii = 0
            MsgBox("Character not allowed", MsgBoxStyle.Information)
        End If
#End If	' def_Form_KeyPress
    End Sub

    Private Sub frmSelectExam_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
        rs = con.Execute("select * from Branch")
        While ( Not rs.EOF)
            cmbBranch.Items.Add(rs(1))
            rs.MoveNext()
        End While
        rs.Close()
#End If	' def_Form_Load
    End Sub


End Class