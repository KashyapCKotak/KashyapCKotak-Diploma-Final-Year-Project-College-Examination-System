Imports VB = Microsoft.VisualBasic

Public Class frmLogin
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
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents cmdAdminLog As System.Windows.Forms.Button
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Image1 As System.Windows.Forms.PictureBox
    Friend WithEvents Image2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogin))
        Me.frmLogin = New System.Windows.Forms.Panel()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.cmdAdminLog = New System.Windows.Forms.Button()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.Image2 = New System.Windows.Forms.PictureBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'frmLogin
        '
        Me.frmLogin.Controls.AddRange(New System.Windows.Forms.Control() {Me.Command1, Me.cmdAdminLog, Me.txtUsername, Me.txtPassword, Me.Image1, Me.Image2, Me.Label3})
        Me.frmLogin.Name = "frmLogin"
        Me.frmLogin.TabIndex = 0
        Me.frmLogin.Location = New System.Drawing.Point(65, 73)
        Me.frmLogin.Size = New System.Drawing.Size(373, 276)
        Me.frmLogin.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.frmLogin.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        '
        'Command1
        '
        Me.Command1.Name = "Command1"
        Me.Command1.Visible = False
        Me.Command1.TabIndex = 8
        Me.Command1.Location = New System.Drawing.Point(16, 227)
        Me.Command1.Size = New System.Drawing.Size(98, 25)
        Me.Command1.Text = "Student Login"
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdAdminLog
        '
        Me.cmdAdminLog.Name = "cmdAdminLog"
        Me.cmdAdminLog.Visible = False
        Me.cmdAdminLog.TabIndex = 5
        Me.cmdAdminLog.Location = New System.Drawing.Point(267, 227)
        Me.cmdAdminLog.Size = New System.Drawing.Size(98, 25)
        Me.cmdAdminLog.Text = "Admin Login"
        Me.cmdAdminLog.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdminLog.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdminLog.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtUsername
        '
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.TabIndex = 3
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
        Me.txtPassword.TabIndex = 4
        Me.txtPassword.Location = New System.Drawing.Point(146, 146)
        Me.txtPassword.Size = New System.Drawing.Size(155, 28)
        Me.txtPassword.Text = ""
        Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
        Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPassword.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Image1
        '
        Me.Image1.Name = "Image1"
        Me.Image1.Location = New System.Drawing.Point(299, 146)
        Me.Image1.Size = New System.Drawing.Size(42, 28)
        Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image1.BackColor = System.Drawing.SystemColors.Control
        Me.Image1.Image = CType(Resources.GetObject("Image1.Image"), System.Drawing.Bitmap)
        '
        'Image2
        '
        Me.Image2.Name = "Image2"
        Me.Image2.Location = New System.Drawing.Point(299, 146)
        Me.Image2.Size = New System.Drawing.Size(42, 28)
        Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image2.BackColor = System.Drawing.SystemColors.Control
        Me.Image2.Image = CType(Resources.GetObject("Image2.Image"), System.Drawing.Bitmap)
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 2
        Me.Label3.Location = New System.Drawing.Point(57, 152)
        Me.Label3.Size = New System.Drawing.Size(100, 19)
        Me.Label3.Text = "Password"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmLogin
        '
        Me.ClientSize = New System.Drawing.Size(513, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.frmLogin})
        Me.Name = "frmLogin"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmLogin.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Login"
        Me.Command1.ResumeLayout(False)
        Me.cmdAdminLog.ResumeLayout(False)
        Me.txtUsername.ResumeLayout(False)
        Me.txtPassword.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.frmLogin.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
'#Const defUse_A_KeyPress = True
#If defUse_A_KeyPress
    Private Sub A_KeyPress(ByVal KeyAscii As Short)
'#Const def_A_KeyPress = True
#If def_A_KeyPress
        txtPassword.PasswordChar = ""
        txtPassword.Refresh()
#End If	' def_A_KeyPress
    End Sub
#End If

'#Const defUse_A_KeyUp = True
#If defUse_A_KeyUp
    Private Sub A_KeyUp(ByVal KeyCode As Short, ByVal Shift As Short)
'#Const def_A_KeyUp = True
#If def_A_KeyUp
        txtPassword.PasswordChar = "*"
        txtPassword.Refresh()
#End If	' def_A_KeyUp
    End Sub
#End If

'#Const defUse_A_MouseDown = True
#If defUse_A_MouseDown
    Private Sub A_MouseDown(ByVal Button As Short, ByVal Shift As Short, ByVal x As Single, ByVal y As Single)
'#Const def_A_MouseDown = True
#If def_A_MouseDown
        txtPassword.PasswordChar = ""
        txtPassword.Refresh()
#End If	' def_A_MouseDown
    End Sub
#End If

'#Const defUse_A_MouseUp = True
#If defUse_A_MouseUp
    Private Sub A_MouseUp(ByVal Button As Short, ByVal Shift As Short, ByVal x As Single, ByVal y As Single)
'#Const def_A_MouseUp = True
#If def_A_MouseUp
        txtPassword.PasswordChar = "*"
        txtPassword.Refresh()
#End If	' def_A_MouseUp
    End Sub
#End If

    Private Sub cmdAdminLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdminLog.Click
'#Const def_cmdAdminLog_Click = True
#If def_cmdAdminLog_Click
        rs = con.Execute("select * from Adminlogin where Username='"+txtUsername.Text+"' and Password='"+txtPassword.Text+"'")
        If ( Not rs.EOF) Then
            MsgBox("Login Success", MsgBoxStyle.Information, " CES")
            MDIForm1.mnuExam.Enabled = True
            MDIForm1.mnuAdminAdSub.Enabled = True
            MDIForm1.mnuAdminchgPass.Enabled = True
            MDIForm1.mnuAdminCrtBrnh.Enabled = True
            MDIForm1.mnuadminCrtUsr.Enabled = True
            MDIForm1.mnuAdminDltUsr.Enabled = True
            MDIForm1.mnuAdminLogout.Enabled = True
            MDIForm1.mnuStudent.Enabled = True
            MDIForm1.mnuAdminLogout.Enabled = True
            MDIForm1.mnuAdminLogin.Enabled = False
            Close()
        Else
            MsgBox("Invalid Username or Password", MsgBoxStyle.Critical, " CES")
        End If
        rs.Close()
#End If	' def_cmdAdminLog_Click
    End Sub

'#Const defUse_cmdCancel_Click = True
#If defUse_cmdCancel_Click
    Private Sub cmdCancel_Click()
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Application.Exit()
#End If	' def_cmdCancel_Click
    End Sub
#End If
    Private Sub cmdTchLog_Click()
'#Const def_cmdTchLog_Click = True
#If def_cmdTchLog_Click
        rs = con.Execute("select * from Userlogin where Username='"+txtUsername.Text+"' and UPassword='"+txtPassword.Text+"'")
        If ( Not rs.EOF) Then
            MsgBox("Login Success", MsgBoxStyle.Information, " CES")
            MDIForm1.mnuStudent.Enabled = False
            MDIForm1.mnuAdminLogout.Enabled = True
            MDIForm1.mnuAdminLogin.Enabled = False
            MDIForm1.mnuExam.Enabled = True
            MDIForm1.mnuAdminchgPass.Enabled = True
            Close()
        Else
            MsgBox("Invalid Username or Password", MsgBoxStyle.Critical, " CES")
        End If
#End If	' def_cmdTchLog_Click
    End Sub

    Private Sub frmLogin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
#End If	' def_Form_Load
    End Sub

    Private Sub Image1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Image1.MouseDown
        Dim Button As Short = e.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(e.X)
        Dim Y As Single = VB6.PixelsToTwipsY(e.Y)
'#Const def_Image1_MouseDown = True
#If def_Image1_MouseDown
        txtPassword.PasswordChar = ""
        txtPassword.Refresh()
        Image1.Visible = False
        Image2.Visible = True
#End If	' def_Image1_MouseDown
    End Sub

    Private Sub Image1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Image1.MouseUp
        Dim Button As Short = e.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(e.X)
        Dim Y As Single = VB6.PixelsToTwipsY(e.Y)
'#Const def_Image1_MouseUp = True
#If def_Image1_MouseUp
        txtPassword.PasswordChar = "*"
        txtPassword.Refresh()
        Image1.Visible = True
        Image2.Visible = False
#End If	' def_Image1_MouseUp
    End Sub


End Class