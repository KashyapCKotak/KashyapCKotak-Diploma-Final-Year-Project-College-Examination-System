Imports VB = Microsoft.VisualBasic

Imports System.Data.OleDb
Imports VBto
Public Class frmAddQst
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
    Friend WithEvents Frame2 As System.Windows.Forms.Panel
    Friend WithEvents Frame4 As System.Windows.Forms.Panel
    Friend WithEvents txtOpt4 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents lblQno As System.Windows.Forms.Label
    Friend WithEvents cmdcancel As System.Windows.Forms.Button
    Friend WithEvents Frame3 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblSubName As System.Windows.Forms.Label
    Friend WithEvents Frame1 As System.Windows.Forms.Panel
    Friend WithEvents txtSubCode As System.Windows.Forms.TextBox
    Friend WithEvents cmbSubjects As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSem As System.Windows.Forms.ComboBox
    Friend WithEvents cmbBranch As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddQst))
        Me.Frame2 = New System.Windows.Forms.Panel()
        Me.Frame4 = New System.Windows.Forms.Panel()
        Me.txtOpt4 = New System.Windows.Forms.RichTextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.lblQno = New System.Windows.Forms.Label()
        Me.cmdcancel = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.Panel()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lblSubName = New System.Windows.Forms.Label()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.txtSubCode = New System.Windows.Forms.TextBox()
        Me.cmbSubjects = New System.Windows.Forms.ComboBox()
        Me.cmbSem = New System.Windows.Forms.ComboBox()
        Me.cmbBranch = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame2
        '
        Me.Frame2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame4, Me.cmdcancel, Me.Frame3})
        Me.Frame2.Name = "Frame2"
        Me.Frame2.TabIndex = 9
        Me.Frame2.Location = New System.Drawing.Point(8, 8)
        Me.Frame2.Size = New System.Drawing.Size(1352, 703)
        Me.Frame2.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Frame4
        '
        Me.Frame4.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtOpt4, Me.Label16, Me.lblQno})
        Me.Frame4.Name = "Frame4"
        Me.Frame4.TabIndex = 16
        Me.Frame4.Location = New System.Drawing.Point(0, 65)
        Me.Frame4.Size = New System.Drawing.Size(1336, 600)
        Me.Frame4.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'txtOpt4
        '
        Me.txtOpt4.Name = "txtOpt4"
        Me.txtOpt4.Enabled = True
        Me.txtOpt4.TabIndex = 40
        Me.txtOpt4.Location = New System.Drawing.Point(146, 421)
        Me.txtOpt4.Size = New System.Drawing.Size(1182, 74)
        '
        'Label16
        '
        Me.Label16.Name = "Label16"
        Me.Label16.TabIndex = 22
        Me.Label16.Location = New System.Drawing.Point(81, 372)
        Me.Label16.Size = New System.Drawing.Size(25, 25)
        Me.Label16.Text = "3"
        Me.Label16.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label16.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblQno
        '
        Me.lblQno.Name = "lblQno"
        Me.lblQno.TabIndex = 18
        Me.lblQno.Location = New System.Drawing.Point(89, 32)
        Me.lblQno.Size = New System.Drawing.Size(41, 33)
        Me.lblQno.Text = ""
        Me.lblQno.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.lblQno.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblQno.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdcancel
        '
        Me.cmdcancel.Name = "cmdcancel"
        Me.cmdcancel.TabIndex = 26
        Me.cmdcancel.Location = New System.Drawing.Point(841, 671)
        Me.cmdcancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdcancel.Text = "Cancel"
        Me.cmdcancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdcancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdcancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Frame3
        '
        Me.Frame3.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label9, Me.lblSubName})
        Me.Frame3.Name = "Frame3"
        Me.Frame3.TabIndex = 11
        Me.Frame3.Location = New System.Drawing.Point(429, 32)
        Me.Frame3.Size = New System.Drawing.Size(502, 33)
        Me.Frame3.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Label9
        '
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 14
        Me.Label9.Location = New System.Drawing.Point(348, 8)
        Me.Label9.Size = New System.Drawing.Size(77, 17)
        Me.Label9.Text = "Subject Code:"
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblSubName
        '
        Me.lblSubName.Name = "lblSubName"
        Me.lblSubName.TabIndex = 13
        Me.lblSubName.Location = New System.Drawing.Point(93, 0)
        Me.lblSubName.Size = New System.Drawing.Size(250, 25)
        Me.lblSubName.Text = ""
        Me.lblSubName.BackColor = System.Drawing.Color.Transparent
        Me.lblSubName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSubName.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtSubCode, Me.cmbSubjects, Me.cmbSem, Me.cmbBranch, Me.Label10, Me.Label8, Me.Label19, Me.Label5, Me.Label2})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(510, 162)
        Me.Frame1.Size = New System.Drawing.Size(381, 411)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'txtSubCode
        '
        Me.txtSubCode.Name = "txtSubCode"
        Me.txtSubCode.Enabled = False
        Me.txtSubCode.TabIndex = 28
        Me.txtSubCode.Location = New System.Drawing.Point(152, 206)
        Me.txtSubCode.Size = New System.Drawing.Size(162, 22)
        Me.txtSubCode.Text = ""
        Me.txtSubCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubCode.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmbSubjects
        '
        Me.cmbSubjects.Name = "cmbSubjects"
        Me.cmbSubjects.TabIndex = 8
        Me.cmbSubjects.Location = New System.Drawing.Point(154, 154)
        Me.cmbSubjects.Size = New System.Drawing.Size(163, 23)
        Me.cmbSubjects.Text = ""
        Me.cmbSubjects.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSubjects.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSubjects.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmbSem
        '
        Me.cmbSem.Name = "cmbSem"
        Me.cmbSem.TabIndex = 7
        Me.cmbSem.Location = New System.Drawing.Point(154, 101)
        Me.cmbSem.Size = New System.Drawing.Size(163, 23)
        Me.cmbSem.Text = ""
        Me.cmbSem.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSem.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSem.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmbBranch
        '
        Me.cmbBranch.Name = "cmbBranch"
        Me.cmbBranch.TabIndex = 6
        Me.cmbBranch.Location = New System.Drawing.Point(154, 49)
        Me.cmbBranch.Size = New System.Drawing.Size(163, 23)
        Me.cmbBranch.Text = ""
        Me.cmbBranch.BackColor = System.Drawing.SystemColors.Window
        Me.cmbBranch.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbBranch.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label10
        '
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 31
        Me.Label10.Location = New System.Drawing.Point(49, 306)
        Me.Label10.Size = New System.Drawing.Size(77, 25)
        Me.Label10.Text = "Set Time"
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label8
        '
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 29
        Me.Label8.Location = New System.Drawing.Point(49, 259)
        Me.Label8.Size = New System.Drawing.Size(77, 25)
        Me.Label8.Text = "Set Date"
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label19
        '
        Me.Label19.Name = "Label19"
        Me.Label19.TabIndex = 27
        Me.Label19.Location = New System.Drawing.Point(49, 206)
        Me.Label19.Size = New System.Drawing.Size(122, 25)
        Me.Label19.Text = "Subject Code"
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label5
        '
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 4
        Me.Label5.Location = New System.Drawing.Point(324, 24)
        Me.Label5.Size = New System.Drawing.Size(41, 25)
        Me.Label5.Text = "Step 1"
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 2
        Me.Label2.Location = New System.Drawing.Point(49, 105)
        Me.Label2.Size = New System.Drawing.Size(122, 25)
        Me.Label2.Text = "Select Semester"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmAddQst
        '
        Me.ClientSize = New System.Drawing.Size(635, 603)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame2, Me.Frame1})
        Me.Name = "frmAddQst"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = True
        Me.Icon = CType(Resources.GetObject("frmAddQst.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add Questions"
        Me.Label16.ResumeLayout(False)
        Me.lblQno.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.cmdcancel.ResumeLayout(False)
        Me.Label9.ResumeLayout(False)
        Me.lblSubName.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.txtSubCode.ResumeLayout(False)
        Me.cmbSubjects.ResumeLayout(False)
        Me.cmbSem.ResumeLayout(False)
        Me.cmbBranch.ResumeLayout(False)
        Me.Label10.ResumeLayout(False)
        Me.Label8.ResumeLayout(False)
        Me.Label19.ResumeLayout(False)
        Me.Label5.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    ' VBto upgrade warning: x As Object --> As VBtoRecordSet
    ' VBto upgrade warning: i As Short	OnWrite(Short, VBtoRecordSet)
    Dim x As VBtoRecordSet, i As Short
    Dim exid As String
    Dim eID As String
    ' VBto upgrade warning: Qno As Short	OnWrite(VBtoRecordSet, Short)
    Dim Qno As Short

    Private Sub cmbBranch_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBranch.SelectedIndexChanged
'#Const def_cmbBranch_Click = True
#If def_cmbBranch_Click
        cmbSem.Items.Clear()
        rs = con.Execute("select Sem from Branch where Branchcode='"+cmbBranch.Text+"'")
        If ( Not rs.EOF) Then
            x = rs(0)
            For i = 1 To x
                cmbSem.Items.Add(i)
            Next
        End If
        rs.Close()
#End If	' def_cmbBranch_Click
    End Sub


    Private Sub cmbSem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSem.SelectedIndexChanged
'#Const def_cmbSem_Click = True
#If def_cmbSem_Click
        cmbSubjects.Items.Clear()
        rs = con.Execute("select Subjectname from Subjects where Branchcode='"+cmbBranch.Text+"' and Sem="+cmbSem.Text+"")
        While ( Not rs.EOF)
            cmbSubjects.Items.Add(rs(0))
            rs.MoveNext()
        End While
        rs.Close()
#End If	' def_cmbSem_Click
    End Sub
    Private Sub cmbSubjects_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSubjects.SelectedIndexChanged
'#Const def_cmbSubjects_Click = True
#If def_cmbSubjects_Click
        rs = con.Execute("select Subjectcode from Subjects where Subjectname='"+cmbSubjects.Text+"' and BranchCode='"+cmbBranch.Text+"'")
        If ( Not rs.EOF) Then
            txtSubCode.Text = rs(0)
        End If
        rs.Close()
#End If	' def_cmbSubjects_Click
    End Sub

    Private Sub cmdadd_Click()
'#Const def_cmdadd_Click = True
#If def_cmdadd_Click
        If (txtAns.Text="" Or txtOpt1.Text="" Or txtOpt2.Text="" Or txtOpt3.Text="" Or txtOpt4.Text="" Or txtQst.Text="") Then
            MsgBox("Missing Fields", MsgBoxStyle.Information, " CES")
        Else
            If (txtOpt1.Text<>txtAns.Text And txtOpt2.Text<>txtAns.Text And txtOpt3.Text<>txtAns.Text And txtOpt4.Text<>txtAns.Text) Then
                MsgBox("Answer key does not match any one of 4 options", MsgBoxStyle.Critical, " CES")
                txtAns.Text = ""
                txtAns.Focus()
            Else
                exid = cmbBranch.Text+txtDate.Text+txtSubCode.Text
                con.Execute(("insert into Questions values('"+cmbBranch.Text+"',"+cmbSem.Text+",'"+txtSubCode.Text+"','"+exid+"','"+lblQno.Text+"','"+txtQst.Text+"','"+txtOpt1.Text+"', '"+txtOpt2.Text+"','"+txtOpt3.Text+"','"+txtOpt4.Text+"','"+txtAns.Text+"')"))
                con.Execute(("insert into ExamDetails values('"+cmbBranch.Text+"',"+cmbSem.Text+",'"+txtSubCode.Text+"','"+exid+"','"+txtDate.Text+"','"+txtTime.Text+"')"))
                MsgBox("Question Added Successfully", MsgBoxStyle.Information, " CES")
                txtAns.Text = ""
                txtOpt1.Text = ""
                txtOpt2.Text = ""
                txtOpt3.Text = ""
                txtOpt4.Text = ""
                txtQst.Text = ""
                txtQst.Focus()
                lblQno.Text = lblQno.Text+1
            End If
        End If
#End If	' def_cmdadd_Click
    End Sub

    Private Sub cmdcancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdcancel.Click
'#Const def_cmdcancel_Click = True
#If def_cmdcancel_Click
        Close()
#End If	' def_cmdcancel_Click
    End Sub

    Private Sub cmdNext_Click()
'#Const def_cmdNext_Click = True
#If def_cmdNext_Click
        If (txtSubCode.Text="" Or txtDate.Text="" Or txtTime.Text="") Then
            MsgBox("Please Select and fill all the fields")
            Frame1.Visible = True
            Frame2.Visible = False
        Else

            lblSubName.Text = cmbSubjects.Text
            lblSubCode.Text = txtSubCode.Text
            eID = cmbBranch.Text+txtDate.Text+txtSubCode.Text
            rs = con.Execute("select count(*) from Questions where BranchCode='"+cmbBranch.Text+"' and Sem="+cmbSem.Text+" and SubjectCode='"+txtSubCode.Text+"' and ExamID='"+eID+"'")
            If ( Not rs.EOF) Then
                Qno = rs(0)
                Qno += 1
            Else
                Qno = 1
            End If
            lblQno.Text = Qno
            Frame1.Visible = False
            Frame2.Visible = True
        End If
#End If	' def_cmdNext_Click
    End Sub

    Private Sub frmAddQst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
'#Const def_Form_KeyPress = True
#If def_Form_KeyPress
        If e.KeyAscii=(Asc("'")) Then
            e.KeyAscii = 0
            MsgBox("Character not allowed", MsgBoxStyle.Information)
        End If
        'If KeyAscii = (Asc("-")) Then
        '    KeyAscii = 0
        '    MsgBox "Character not allowed", vbInformation
        'End If
        'If KeyAscii = (Asc("*")) Then
        '    KeyAscii = 0
        '    MsgBox "Character not allowed", vbInformation
        'End If
#End If	' def_Form_KeyPress
    End Sub

    Private Sub frmAddQst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load

        'If KeyAscii = (Asc("-")) Then
        '    KeyAscii = 0
        '    MsgBox "Character not allowed", vbInformation
        'End If
        'If KeyAscii = (Asc("*")) Then
        '    KeyAscii = 0
        '    MsgBox "Character not allowed", vbInformation
        'End If
        Me.BackColor = ColorTranslator.FromOle(color)
        Frame1.Visible = True
        Frame2.Visible = False
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