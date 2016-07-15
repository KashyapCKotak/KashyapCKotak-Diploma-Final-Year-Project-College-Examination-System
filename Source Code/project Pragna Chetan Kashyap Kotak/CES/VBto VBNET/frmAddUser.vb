Imports VB = Microsoft.VisualBasic

Public Class frmAddUser
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
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAddUser))
        Me.frmLogin = New System.Windows.Forms.Panel()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'frmLogin
        '
        Me.frmLogin.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtUsername, Me.txtPassword, Me.cmdAdd, Me.cmdCancel, Me.Label3})
        Me.frmLogin.Name = "frmLogin"
        Me.frmLogin.TabIndex = 0
        Me.frmLogin.Location = New System.Drawing.Point(97, 65)
        Me.frmLogin.Size = New System.Drawing.Size(373, 276)
        Me.frmLogin.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.frmLogin.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'txtUsername
        '
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.TabIndex = 4
        Me.txtUsername.Location = New System.Drawing.Point(146, 97)
        Me.txtUsername.Size = New System.Drawing.Size(195, 28)
        Me.txtUsername.Text = ""
        Me.txtUsername.BackColor = System.Drawing.SystemColors.Window
        Me.txtUsername.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUsername.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtPassword
        '
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.TabIndex = 3
        Me.txtPassword.Location = New System.Drawing.Point(146, 146)
        Me.txtPassword.Size = New System.Drawing.Size(195, 28)
        Me.txtPassword.Text = ""
        Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassword.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdAdd
        '
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.TabIndex = 2
        Me.cmdAdd.Location = New System.Drawing.Point(210, 227)
        Me.cmdAdd.Size = New System.Drawing.Size(98, 25)
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdCancel
        '
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Location = New System.Drawing.Point(65, 227)
        Me.cmdCancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 5
        Me.Label3.Location = New System.Drawing.Point(57, 154)
        Me.Label3.Size = New System.Drawing.Size(96, 20)
        Me.Label3.Text = "Password"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmAddUser
        '
        Me.ClientSize = New System.Drawing.Size(571, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.frmLogin})
        Me.Name = "frmAddUser"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmAddUser.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Add New User"
        Me.txtUsername.ResumeLayout(False)
        Me.txtPassword.ResumeLayout(False)
        Me.cmdAdd.ResumeLayout(False)
        Me.cmdCancel.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.frmLogin.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Private Sub cmdAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
'#Const def_cmdAdd_Click = True
#If def_cmdAdd_Click
        rs = con.Execute("select * from Userlogin where Username='"+txtUsername.Text+"'")
        If ( Not rs.EOF) Then
            MsgBox("Sorry!! User already exists. Try another username", MsgBoxStyle.Critical, " CES")
            txtPassword.Text = ""
            txtUsername.Text = ""
            txtUsername.Focus()
        Else
            con.Execute(("insert into Userlogin values('"+txtUsername.Text+"','"+txtPassword.Text+"')"))
            MsgBox("User added sucessfully", MsgBoxStyle.Information, " CES")
            txtPassword.Text = ""
            txtUsername.Text = ""
            txtUsername.Focus()
        End If
#End If	' def_cmdAdd_Click
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub frmAddUser_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmAddUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load

        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
        Me.Top = 1550
        Me.Left = 5000
#End If	' def_Form_Load
    End Sub

End Class