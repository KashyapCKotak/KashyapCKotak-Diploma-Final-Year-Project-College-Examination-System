Imports VB = Microsoft.VisualBasic

Public Class frmSubjects
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
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmbSem As System.Windows.Forms.ComboBox
    Friend WithEvents txtSubcode As System.Windows.Forms.TextBox
    Friend WithEvents txtSubname As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSubjects))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmbSem = New System.Windows.Forms.ComboBox()
        Me.txtSubcode = New System.Windows.Forms.TextBox()
        Me.txtSubname = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdAdd, Me.cmbSem, Me.txtSubcode, Me.txtSubname, Me.Label5, Me.Label4, Me.Label1})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(105, 65)
        Me.Frame1.Size = New System.Drawing.Size(436, 338)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdAdd
        '
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.TabIndex = 10
        Me.cmdAdd.Location = New System.Drawing.Point(89, 291)
        Me.cmdAdd.Size = New System.Drawing.Size(98, 25)
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmbSem
        '
        Me.cmbSem.Name = "cmbSem"
        Me.cmbSem.TabIndex = 9
        Me.cmbSem.Location = New System.Drawing.Point(185, 143)
        Me.cmbSem.Size = New System.Drawing.Size(199, 23)
        Me.cmbSem.Text = ""
        Me.cmbSem.BackColor = System.Drawing.SystemColors.Window
        Me.cmbSem.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbSem.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtSubcode
        '
        Me.txtSubcode.Name = "txtSubcode"
        Me.txtSubcode.TabIndex = 7
        Me.txtSubcode.Location = New System.Drawing.Point(185, 234)
        Me.txtSubcode.Size = New System.Drawing.Size(199, 25)
        Me.txtSubcode.Text = ""
        Me.txtSubcode.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubcode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubcode.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtSubname
        '
        Me.txtSubname.Name = "txtSubname"
        Me.txtSubname.TabIndex = 6
        Me.txtSubname.Location = New System.Drawing.Point(185, 185)
        Me.txtSubname.Size = New System.Drawing.Size(199, 25)
        Me.txtSubname.Text = ""
        Me.txtSubname.BackColor = System.Drawing.SystemColors.Window
        Me.txtSubname.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSubname.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label5
        '
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 5
        Me.Label5.Location = New System.Drawing.Point(52, 240)
        Me.Label5.Size = New System.Drawing.Size(125, 28)
        Me.Label5.Text = "Subject Code"
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label4
        '
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 4
        Me.Label4.Location = New System.Drawing.Point(52, 194)
        Me.Label4.Size = New System.Drawing.Size(125, 28)
        Me.Label4.Text = "Subject Name"
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Location = New System.Drawing.Point(0, 24)
        Me.Label1.Size = New System.Drawing.Size(446, 33)
        Me.Label1.Text = "Add Subjects"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmSubjects
        '
        Me.ClientSize = New System.Drawing.Size(648, 453)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmSubjects"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmSubjects.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add Subjects"
        Me.cmdAdd.ResumeLayout(False)
        Me.cmbSem.ResumeLayout(False)
        Me.txtSubcode.ResumeLayout(False)
        Me.txtSubname.ResumeLayout(False)
        Me.Label5.ResumeLayout(False)
        Me.Label4.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    ' VBto upgrade warning: x As Short	OnWrite(VBtoRecordSet)
    Dim x As Short
    Dim i As Short
    ' VBto upgrade warning: dec As Object --> As Short
    Dim dec As Short, decname As Object, decsub As Object, deccode As Short



    Private Sub cmbBranch_Click()
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

#End If	' def_cmbBranch_Click
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
'#Const def_cmdAdd_Click = True
#If def_cmdAdd_Click
        dec = 1
        If (cmbBranch.Text="" Or cmbSem.Text="" Or txtSubname.Text="" Or txtSubCode.Text="") Then
            MsgBox("Missing Fields", MsgBoxStyle.Information, " CES")
            dec = 0
        Else
            rs = con.Execute("select * from Subjects where Branchcode='"+cmbBranch.Text+"' AND Subjectname='"+txtSubname.Text+"'")
            If ( Not rs.EOF) Then
                MsgBox("Sorry!! The Specified subject already exists in the given branch. Try another branch name or subject name", MsgBoxStyle.Critical, " CES")
                dec = 0
            Else
                rs = con.Execute("select * from Subjects where Subjectcode='"+txtSubCode.Text+"'")
                If ( Not rs.EOF) Then
                    MsgBox("Sorry!! The Specified branch code already exists. Branch codes must be unique irrespective of their branch/subject", MsgBoxStyle.Critical, " CES")
                    dec = 0
                Else
                    dec = 1
                End If
            End If
        End If
        rs.Close()
        If dec=1 Then
            con.Execute(("insert into Subjects values('"+cmbBranch.Text+"',"+cmbSem.Text+",'"+txtSubname.Text+"','"+txtSubCode.Text+"')"))
            MsgBox("Record added sucessfully", MsgBoxStyle.Information, " CES")
            txtSubname.Text = ""
            txtSubCode.Text = ""
            txtSubname.Focus()
        End If
#End If	' def_cmdAdd_Click
    End Sub

    Private Sub cmdCancel_Click()
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub frmSubjects_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmSubjects_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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