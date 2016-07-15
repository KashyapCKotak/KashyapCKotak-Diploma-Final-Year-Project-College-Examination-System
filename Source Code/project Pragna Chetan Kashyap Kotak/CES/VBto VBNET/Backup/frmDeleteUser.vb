Imports VB = Microsoft.VisualBasic

Public Class frmDeleteUser
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
    Friend WithEvents frmLogin As System.Windows.Forms.Panel
    Friend WithEvents cmbUsername As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDeleteUser))
        Me.frmLogin = New System.Windows.Forms.Panel()
        Me.cmbUsername = New System.Windows.Forms.ComboBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'frmLogin
        '
        Me.frmLogin.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmbUsername, Me.cmdCancel, Me.cmdDelete, Me.Label3, Me.Label1, Me.Label2})
        Me.frmLogin.Name = "frmLogin"
        Me.frmLogin.TabIndex = 0
        Me.frmLogin.Location = New System.Drawing.Point(113, 73)
        Me.frmLogin.Size = New System.Drawing.Size(373, 276)
        Me.frmLogin.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.frmLogin.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmbUsername
        '
        Me.cmbUsername.Name = "cmbUsername"
        Me.cmbUsername.TabIndex = 4
        Me.cmbUsername.Location = New System.Drawing.Point(121, 127)
        Me.cmbUsername.Size = New System.Drawing.Size(155, 23)
        Me.cmbUsername.Text = ""
        Me.cmbUsername.BackColor = System.Drawing.SystemColors.Window
        Me.cmbUsername.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbUsername.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdCancel
        '
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Location = New System.Drawing.Point(218, 218)
        Me.cmdCancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdDelete
        '
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.TabIndex = 1
        Me.cmdDelete.Location = New System.Drawing.Point(65, 218)
        Me.cmdDelete.Size = New System.Drawing.Size(98, 25)
        Me.cmdDelete.Text = "Delete"
        Me.cmdDelete.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 6
        Me.Label3.Location = New System.Drawing.Point(283, 126)
        Me.Label3.Size = New System.Drawing.Size(50, 23)
        Me.Label3.Text = "Refresh"
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(255, Byte), CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 5
        Me.Label1.Location = New System.Drawing.Point(0, 24)
        Me.Label1.Size = New System.Drawing.Size(373, 41)
        Me.Label1.Text = "DELETE TEACHER ACCOUNT"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 3
        Me.Label2.Location = New System.Drawing.Point(32, 129)
        Me.Label2.Size = New System.Drawing.Size(98, 17)
        Me.Label2.Text = "User Name"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmDeleteUser
        '
        Me.ClientSize = New System.Drawing.Size(571, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.frmLogin})
        Me.Name = "frmDeleteUser"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmDeleteUser.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Delete User"
        Me.cmbUsername.ResumeLayout(False)
        Me.cmdCancel.ResumeLayout(False)
        Me.cmdDelete.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.frmLogin.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
'#Const def_cmdDelete_Click = True
#If def_cmdDelete_Click
        con.Execute(("delete from Userlogin where Username='"+cmbUsername.Text+"'"))
        MsgBox("User deleted sucessfully!!", MsgBoxStyle.Information, " CES")
        cmbUsername.Text = ""
#End If	' def_cmdDelete_Click
    End Sub

'#Const defUse_Command1_Click = True
#If defUse_Command1_Click
    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click
        Close()
        LoadUnUsed(Me)
        ShowModeless(Me)
        Command1.Visible = True
#End If	' def_Command1_Click
    End Sub
#End If

    Private Sub frmDeleteUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
        Me.Top = 1550
        Me.Left = 5000
        rs = con.Execute("select * from Userlogin")
        While ( Not rs.EOF)
            cmbUsername.Items.Add(rs(0))
            rs.MoveNext()
        End While
#End If	' def_Form_Load
    End Sub

    Private Sub Label3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.Click
'#Const def_Label3_Click = True
#If def_Label3_Click
        Close()
        LoadUnUsed(frmDelSub)
        ShowModeless(frmDelSub)
#End If	' def_Label3_Click
    End Sub

End Class