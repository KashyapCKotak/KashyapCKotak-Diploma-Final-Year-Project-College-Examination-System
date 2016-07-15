Imports VB = Microsoft.VisualBasic

Public Class frmLoginSelect
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
    Friend WithEvents frmLogin As System.Windows.Forms.GroupBox
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdminLog As System.Windows.Forms.Button
    Friend WithEvents cmdTchLog As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents mainFrame As System.Windows.Forms.Panel
    Friend WithEvents Picture2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdStudent As System.Windows.Forms.Button
    Friend WithEvents cmdAdmin As System.Windows.Forms.Button
    Friend WithEvents non3 As System.Windows.Forms.Label
    Friend WithEvents non2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLoginSelect))
        Me.frmLogin = New System.Windows.Forms.GroupBox()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.cmdAdminLog = New System.Windows.Forms.Button()
        Me.cmdTchLog = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.mainFrame = New System.Windows.Forms.Panel()
        Me.Picture2 = New System.Windows.Forms.PictureBox()
        Me.cmdStudent = New System.Windows.Forms.Button()
        Me.cmdAdmin = New System.Windows.Forms.Button()
        Me.non3 = New System.Windows.Forms.Label()
        Me.non2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'frmLogin
        '
        Me.frmLogin.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtUsername, Me.cmdAdminLog, Me.cmdTchLog, Me.Label3, Me.Label2, Me.Label1})
        Me.frmLogin.Name = "frmLogin"
        Me.frmLogin.Enabled = False
        Me.frmLogin.TabIndex = 9
        Me.frmLogin.Location = New System.Drawing.Point(372, 89)
        Me.frmLogin.Size = New System.Drawing.Size(373, 276)
        Me.frmLogin.Text = ""
        Me.frmLogin.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(192, Byte))
        Me.frmLogin.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        '
        'txtUsername
        '
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.TabIndex = 13
        Me.txtUsername.Location = New System.Drawing.Point(121, 97)
        Me.txtUsername.Size = New System.Drawing.Size(195, 28)
        Me.txtUsername.Text = ""
        Me.txtUsername.BackColor = System.Drawing.SystemColors.Window
        Me.txtUsername.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUsername.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdAdminLog
        '
        Me.cmdAdminLog.Name = "cmdAdminLog"
        Me.cmdAdminLog.Visible = False
        Me.cmdAdminLog.TabIndex = 11
        Me.cmdAdminLog.Location = New System.Drawing.Point(210, 210)
        Me.cmdAdminLog.Size = New System.Drawing.Size(98, 25)
        Me.cmdAdminLog.Text = "Admin Login"
        Me.cmdAdminLog.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdAdminLog.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdminLog.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdTchLog
        '
        Me.cmdTchLog.Name = "cmdTchLog"
        Me.cmdTchLog.Visible = False
        Me.cmdTchLog.TabIndex = 12
        Me.cmdTchLog.Location = New System.Drawing.Point(210, 210)
        Me.cmdTchLog.Size = New System.Drawing.Size(98, 25)
        Me.cmdTchLog.Text = "Teacher Login"
        Me.cmdTchLog.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdTchLog.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdTchLog.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 17
        Me.Label3.Location = New System.Drawing.Point(40, 146)
        Me.Label3.Size = New System.Drawing.Size(100, 19)
        Me.Label3.Text = "Password"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(128, Byte))
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 16
        Me.Label2.Location = New System.Drawing.Point(40, 97)
        Me.Label2.Size = New System.Drawing.Size(96, 20)
        Me.Label2.Text = "User Name"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(128, Byte))
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 15
        Me.Label1.Location = New System.Drawing.Point(0, 24)
        Me.Label1.Size = New System.Drawing.Size(373, 33)
        Me.Label1.Text = " LOGIN"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(128, Byte))
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'mainFrame
        '
        Me.mainFrame.Controls.AddRange(New System.Windows.Forms.Control() {Me.Picture2, Me.cmdStudent, Me.cmdAdmin})
        Me.mainFrame.Name = "mainFrame"
        Me.mainFrame.TabIndex = 0
        Me.mainFrame.Location = New System.Drawing.Point(121, 89)
        Me.mainFrame.Size = New System.Drawing.Size(211, 252)
        Me.mainFrame.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.mainFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.mainFrame.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Picture2
        '
        Me.Picture2.Name = "Picture2"
        Me.Picture2.TabStop = False
        Me.Picture2.TabIndex = 21
        Me.Picture2.Location = New System.Drawing.Point(8, 65)
        Me.Picture2.Size = New System.Drawing.Size(50, 50)
        Me.Picture2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture2.BackColor = System.Drawing.SystemColors.Control
        Me.Picture2.Image = CType(Resources.GetObject("Picture2.Image"), System.Drawing.Bitmap)
        '
        'cmdStudent
        '
        Me.cmdStudent.Name = "cmdStudent"
        Me.cmdStudent.TabIndex = 3
        Me.cmdStudent.Location = New System.Drawing.Point(65, 129)
        Me.cmdStudent.Size = New System.Drawing.Size(130, 33)
        Me.cmdStudent.Text = "STUDENT"
        Me.cmdStudent.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.cmdStudent.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStudent.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdAdmin
        '
        Me.cmdAdmin.Name = "cmdAdmin"
        Me.cmdAdmin.TabIndex = 1
        Me.cmdAdmin.Location = New System.Drawing.Point(65, 16)
        Me.cmdAdmin.Size = New System.Drawing.Size(130, 33)
        Me.cmdAdmin.Text = "ADMINISTRATOR"
        Me.cmdAdmin.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.cmdAdmin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdmin.Font = New System.Drawing.Font("MS Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'non3
        '
        Me.non3.Name = "non3"
        Me.non3.TabIndex = 7
        Me.non3.Location = New System.Drawing.Point(81, 332)
        Me.non3.Size = New System.Drawing.Size(292, 33)
        Me.non3.Text = ""
        Me.non3.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.non3.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'non2
        '
        Me.non2.Name = "non2"
        Me.non2.TabIndex = 5
        Me.non2.Location = New System.Drawing.Point(81, 89)
        Me.non2.Size = New System.Drawing.Size(41, 244)
        Me.non2.Text = ""
        Me.non2.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.non2.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Label4
        '
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 19
        Me.Label4.Location = New System.Drawing.Point(81, 8)
        Me.Label4.Size = New System.Drawing.Size(292, 25)
        Me.Label4.Text = ""
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(0, Byte), CType(0, Byte))
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmLoginSelect
        '
        Me.ClientSize = New System.Drawing.Size(829, 407)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.frmLogin, Me.mainFrame, Me.non3, Me.non2, Me.Label4})
        Me.Name = "frmLoginSelect"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmLoginSelect.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Select Login"
        Me.txtUsername.ResumeLayout(False)
        Me.cmdAdminLog.ResumeLayout(False)
        Me.cmdTchLog.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.frmLogin.ResumeLayout(False)
        Me.cmdStudent.ResumeLayout(False)
        Me.cmdAdmin.ResumeLayout(False)
        Me.mainFrame.ResumeLayout(False)
        Me.Label4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================

    Private Sub cmdAdmin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdmin.Click
'#Const def_cmdAdmin_Click = True
#If def_cmdAdmin_Click

        Command2.Enabled = True
        cmdAdmin.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdAdmin.Enabled = False
        cmdTeacher.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdTeacher.Enabled = False
        cmdStudent.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdStudent.Enabled = False
        Image1.Enabled = True
        frmLogin.BackColor = ColorTranslator.FromOle(&HC00000)
        Label1.ForeColor = ColorTranslator.FromOle(&HFF00)
        Label2.ForeColor = ColorTranslator.FromOle(&HFF00)
        Label3.ForeColor = ColorTranslator.FromOle(&HFF00)
        Command2.BackColor = ColorTranslator.FromOle(&H8080FF)
        frmLogin.Enabled = True
        frmLogin.Visible = True
        cmdAdminLog.Visible = True
#End If	' def_cmdAdmin_Click
    End Sub

    Private Sub cmdAdminLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdminLog.Click
'#Const def_cmdAdminLog_Click = True
#If def_cmdAdminLog_Click
        ' VBto upgrade warning: una As Object --> As String
        ' VBto upgrade warning: pwa As Object --> As String
        Dim una As String = "", pwa As String = ""	' - "AutoDim"

        Dim un As String = ""
        Dim pw As String = ""
        Dim decide As Short
        decide = 0
        un = "admin"
        pw = "admin"
        una = txtUsername.Text
        pwa = txtPassword.Text
        sh = 5
        If un=una Then
            decide = 1
        End If
        If pw=pwa Then
            decide = 1
        End If
        'Set rs = con.Execute("select * from Adminlogin where Username='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
        If decide=1 Then
            decide = 0
            MsgBox("Login Success", MsgBoxStyle.Information, " CES")
            MDIForm1.mnuAdmin.Enabled = True
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
            MDIForm1.Command1.Enabled = False
            MDIForm1.Command2.Enabled = True
            MDIForm1.Picture4.Visible = True
            MDIForm1.mnuAdmin.Enabled = True
            MDIForm1.mnuExam.Enabled = True
            MDIForm1.mnuStudent.Enabled = True
            MDIForm1.StatusBar1.Panels.Item(1 - 1).Text = "Status: Logged in as Administrator"
            MDIForm1.StatusBar1.Panels.Item(2 - 1).Text = txtUsername.Text
            MDIForm1.Label4.Text = txtUsername.Text
            Close()
        Else
            MsgBox("Invalid Username or Password", MsgBoxStyle.Critical, " CES")
        End If
        'rs.Close
        MDIForm1.mnuAdminLogin.Enabled = False
#End If	' def_cmdAdminLog_Click
    End Sub

    Private Sub cmdExit_Click()
'#Const def_cmdExit_Click = True
#If def_cmdExit_Click
        MDIForm1.Close()
        'Load frmSplash2
        'frmSplash2.Show
#End If	' def_cmdExit_Click
    End Sub

    Private Sub cmdStudent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdStudent.Click
'#Const def_cmdStudent_Click = True
#If def_cmdStudent_Click

        Command2.Enabled = True
        cmdAdmin.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdAdmin.Enabled = False
        cmdTeacher.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdTeacher.Enabled = False
        cmdStudent.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdStudent.Enabled = False
        frmLogin.BackColor = ColorTranslator.FromOle(&HC00000)
        Image1.Enabled = True
        Label1.ForeColor = ColorTranslator.FromOle(&HFF00)
        Label2.ForeColor = ColorTranslator.FromOle(&HFF00)
        Label3.ForeColor = ColorTranslator.FromOle(&HFF00)
        Command2.BackColor = ColorTranslator.FromOle(&H8080FF)
        frmLogin.Enabled = True
        frmLogin.Visible = True
        Command1.Visible = True
#End If	' def_cmdStudent_Click
    End Sub

    Private Sub cmdTchLog_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdTchLog.Click
'#Const def_cmdTchLog_Click = True
#If def_cmdTchLog_Click
        connectdb()
        sh = 1
        rs = con.Execute("select * from Userlogin where Username='"+txtUsername.Text+"' and UPassword='"+txtPassword.Text+"'")
        If ( Not rs.EOF) Then
            MsgBox("Login Success", MsgBoxStyle.Information, " CES")
            MDIForm1.mnuStudent.Enabled = False
            MDIForm1.mnuAdminLogout.Enabled = True
            MDIForm1.mnuAdminLogin.Enabled = False
            MDIForm1.mnuExam.Enabled = True
            MDIForm1.mnuAdminchgPass.Enabled = True
            MDIForm1.Command1.Enabled = False
            MDIForm1.Command2.Enabled = True
            MDIForm1.Picture5.Visible = True
            MDIForm1.mnuAdmin.Enabled = False
            MDIForm1.mnuExam.Enabled = True
            MDIForm1.mnuStudent.Enabled = False
            MDIForm1.StatusBar1.Panels.Item(1 - 1).Text = "Status: Logged in as Teacher"
            MDIForm1.StatusBar1.Panels.Item(2 - 1).Text = txtUsername.Text
            MDIForm1.Label4.Text = txtUsername.Text
            Close()
        Else
            MsgBox("Invalid Username or Password", MsgBoxStyle.Critical, " CES")
        End If
        MDIForm1.mnuAdminLogin.Enabled = False

#End If	' def_cmdTchLog_Click
    End Sub

    Private Sub cmdTeacher_Click()
'#Const def_cmdTeacher_Click = True
#If def_cmdTeacher_Click

        Command2.Enabled = True
        cmdAdmin.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdAdmin.Enabled = False
        cmdTeacher.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdTeacher.Enabled = False
        cmdStudent.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdStudent.Enabled = False
        frmLogin.BackColor = ColorTranslator.FromOle(&HC00000)
        Image1.Enabled = True
        Label1.ForeColor = ColorTranslator.FromOle(&HFF00)
        Label2.ForeColor = ColorTranslator.FromOle(&HFF00)
        Label3.ForeColor = ColorTranslator.FromOle(&HFF00)
        Command2.BackColor = ColorTranslator.FromOle(&H8080FF)
        frmLogin.Enabled = True
        frmLogin.Visible = True
        cmdTchLog.Visible = True
#End If	' def_cmdTeacher_Click
    End Sub

    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click
        connectdb()
        sh = 1
        rs = con.Execute("select * from Student where Username='"+txtUsername.Text+"' and Password='"+txtPassword.Text+"'")
        If ( Not rs.EOF) Then
            MsgBox("Login Success", MsgBoxStyle.Information, " CES")
            MDIForm1.mnuAdminLogin.Enabled = False
            MDIForm1.mnuAdminLogout.Enabled = True
            MDIForm1.StatusBar1.Panels.Item(1 - 1).Text = "Status: Logged in as Student"
            MDIForm1.mnuExam.Enabled = False
            MDIForm1.StatusBar1.Panels.Item(2 - 1).Text = txtUsername.Text
            MDIForm1.Label4.Text = txtUsername.Text
            MDIForm1.mnuStudent.Visible = True
            MDIForm1.Command1.Enabled = False
            MDIForm1.Command2.Enabled = True
            MDIForm1.Picture6.Visible = True
            MDIForm1.mnuAdmin.Enabled = False
            MDIForm1.mnuExam.Enabled = False
            MDIForm1.mnuStudent.Enabled = True
            Close()
        Else
            MsgBox("Invalid Username or Password", MsgBoxStyle.Critical, " CES")
        End If
        rs.Close()
        MDIForm1.mnuAdminLogin.Enabled = False

#End If	' def_Command1_Click
    End Sub

    Private Sub Command2_Click()
'#Const def_Command2_Click = True
#If def_Command2_Click
        cmdAdmin.BackColor = ColorTranslator.FromOle(&HFF8080)
        cmdAdmin.Enabled = True
        cmdTeacher.BackColor = ColorTranslator.FromOle(&HFF8080)
        cmdTeacher.Enabled = True
        cmdStudent.BackColor = ColorTranslator.FromOle(&HFF8080)
        cmdStudent.Enabled = True
        frmLogin.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        Label1.ForeColor = ColorTranslator.FromOle(&H808080)
        Label2.ForeColor = ColorTranslator.FromOle(&H808080)
        Label3.ForeColor = ColorTranslator.FromOle(&H808080)
        Command2.BackColor = ColorTranslator.FromOle(&HC0C0C0)
        cmdAdminLog.Visible = False
        cmdTchLog.Visible = False
        Command1.Visible = False
        frmLogin.Enabled = False
#End If	' def_Command2_Click
    End Sub



'#Const defUse_Command4_Click = True
#If defUse_Command4_Click
    Private Sub Command4_Click()
'#Const def_Command4_Click = True
#If def_Command4_Click
        Application.Exit()
#End If	' def_Command4_Click
    End Sub
#End If



'#Const defUse_Command3_Click = True
#If defUse_Command3_Click
    Private Sub Command3_Click()
'#Const def_Command3_Click = True
#If def_Command3_Click

        'ShellExecute hWnd, "open", "D:\wallpapers", vbNullString, vbNullString, SW_SHOWNORMAL

#End If	' def_Command3_Click
    End Sub
#End If

    Private Sub frmLoginSelect_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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
        'If KeyAscii = (Asc("'")) Then
        '    KeyAscii = 0
        '    MsgBox "Character not allowed", vbInformation
        'End If
#End If	' def_Form_KeyPress
    End Sub

    Private Sub frmLoginSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        'connectdb
        MDIForm1.mnuAdminLogin.Enabled = False
        Me.Top = 1550
        Me.Left = 3900
#End If	' def_Form_Load
    End Sub

    Private Sub frmLoginSelect_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
'#Const def_Form_Terminate = True
#If def_Form_Terminate
        MDIForm1.mnuAdminLogin.Enabled = True
#End If	' def_Form_Terminate
    End Sub

    Private Sub Form_Unload(ByRef Cancel As Short)
'#Const def_Form_Unload = True
#If def_Form_Unload
        MDIForm1.mnuAdminLogin.Enabled = True
#End If	' def_Form_Unload
    End Sub

    Private Sub frmLoginSelect_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim Cancel As Short = 0

        Form_Unload(Cancel)
        If Cancel <> 0 Then e.Cancel = True
    End Sub

    Private Sub Image1_MouseDown(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image1_MouseDown = True
#If def_Image1_MouseDown
        txtPassword.PasswordChar = ""
        txtPassword.Refresh()
        Image1.Visible = False
        Image2.Visible = True
#End If	' def_Image1_MouseDown
    End Sub

    Private Sub Image1_MouseUp(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image1_MouseUp = True
#If def_Image1_MouseUp
        txtPassword.PasswordChar = "*"
        txtPassword.Refresh()
        Image1.Visible = True
        Image2.Visible = False
#End If	' def_Image1_MouseUp
    End Sub


End Class