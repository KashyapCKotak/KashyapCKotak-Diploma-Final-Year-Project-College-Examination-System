Imports VB = Microsoft.VisualBasic

Public Class frmBranch
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
    Friend WithEvents txtsem As System.Windows.Forms.TextBox
    Friend WithEvents txtbrcode As System.Windows.Forms.TextBox
    Friend WithEvents txtbrname As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmBranch))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.txtsem = New System.Windows.Forms.TextBox()
        Me.txtbrcode = New System.Windows.Forms.TextBox()
        Me.txtbrname = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdAdd, Me.txtsem, Me.txtbrcode, Me.txtbrname, Me.Label3, Me.Label1})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(105, 57)
        Me.Frame1.Size = New System.Drawing.Size(373, 292)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdAdd
        '
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.TabIndex = 8
        Me.cmdAdd.Location = New System.Drawing.Point(81, 251)
        Me.cmdAdd.Size = New System.Drawing.Size(98, 25)
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtsem
        '
        Me.txtsem.Name = "txtsem"
        Me.txtsem.TabIndex = 7
        Me.txtsem.Location = New System.Drawing.Point(162, 191)
        Me.txtsem.Size = New System.Drawing.Size(180, 25)
        Me.txtsem.Text = ""
        Me.txtsem.BackColor = System.Drawing.SystemColors.Window
        Me.txtsem.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtsem.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtsem.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtbrcode
        '
        Me.txtbrcode.Name = "txtbrcode"
        Me.txtbrcode.TabIndex = 5
        Me.txtbrcode.Location = New System.Drawing.Point(162, 137)
        Me.txtbrcode.Size = New System.Drawing.Size(180, 25)
        Me.txtbrcode.Text = ""
        Me.txtbrcode.BackColor = System.Drawing.SystemColors.Window
        Me.txtbrcode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtbrcode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtbrcode.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtbrname
        '
        Me.txtbrname.Name = "txtbrname"
        Me.txtbrname.TabIndex = 4
        Me.txtbrname.Location = New System.Drawing.Point(162, 82)
        Me.txtbrname.Size = New System.Drawing.Size(180, 25)
        Me.txtbrname.Text = ""
        Me.txtbrname.BackColor = System.Drawing.SystemColors.Window
        Me.txtbrname.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtbrname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtbrname.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 3
        Me.Label3.Location = New System.Drawing.Point(49, 139)
        Me.Label3.Size = New System.Drawing.Size(89, 18)
        Me.Label3.Text = "Branch Code"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Location = New System.Drawing.Point(4, 21)
        Me.Label1.Size = New System.Drawing.Size(365, 22)
        Me.Label1.Text = "Add Branch"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmBranch
        '
        Me.ClientSize = New System.Drawing.Size(571, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmBranch"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmBranch.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Create Branch"
        Me.cmdAdd.ResumeLayout(False)
        Me.txtsem.ResumeLayout(False)
        Me.txtbrcode.ResumeLayout(False)
        Me.txtbrname.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Dim dec As Short
    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
'#Const def_cmdAdd_Click = True
#If def_cmdAdd_Click
        dec = 1
        Dim num As Short
        If (txtbrname.Text="" Or txtbrcode.Text="" Or txtsem.Text="") Then
            MsgBox("Missing Fields", MsgBoxStyle.Information, " CES")
        Else
            rs = con.Execute("select * from Branch where Branchname='"+txtbrname.Text+"'")
            If ( Not rs.EOF) Then
                MsgBox("Sorry!! Branch already exists. Try another branch name", MsgBoxStyle.Critical, " CES")
                dec = 0
            End If
            rs = con.Execute("select * from Branch where Branchcode='"+txtbrcode.Text+"'")
            If ( Not rs.EOF) Then
                MsgBox("Sorry!! Branch code already exists. Try another branch code", MsgBoxStyle.Critical, " CES")
                dec = 0
            End If
        End If
        If dec=1 Then
            rs.Close()
            con.Execute(("insert into Branch values('"+txtbrname.Text+"','"+txtbrcode.Text+"',"+txtsem.Text+")"))
            MsgBox("Record added successfully", MsgBoxStyle.Information, " CES")
            txtbrname.Text = ""
            txtbrcode.Text = ""
            txtsem.Text = ""
            txtbrname.Focus()
        End If
#End If	' def_cmdAdd_Click
    End Sub

    Private Sub cmdCancel_Click()
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub frmBranch_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmBranch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(Color)
        connectdb()
#End If	' def_Form_Load
    End Sub


End Class