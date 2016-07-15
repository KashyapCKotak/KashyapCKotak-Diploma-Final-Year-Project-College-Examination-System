Imports VB = Microsoft.VisualBasic

Public Class frmChangePass
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
    Friend WithEvents txtConfrmpass As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Image4 As System.Windows.Forms.PictureBox
    Friend WithEvents Image3 As System.Windows.Forms.PictureBox
    Friend WithEvents Image6 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmChangePass))
        Me.frmLogin = New System.Windows.Forms.Panel()
        Me.cmbUsername = New System.Windows.Forms.ComboBox()
        Me.txtConfrmpass = New System.Windows.Forms.TextBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Image4 = New System.Windows.Forms.PictureBox()
        Me.Image3 = New System.Windows.Forms.PictureBox()
        Me.Image6 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'frmLogin
        '
        Me.frmLogin.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmbUsername, Me.txtConfrmpass, Me.cmdCancel, Me.Image4, Me.Image3, Me.Image6, Me.Label1, Me.Label5, Me.Label4, Me.Label2})
        Me.frmLogin.Name = "frmLogin"
        Me.frmLogin.TabIndex = 0
        Me.frmLogin.Location = New System.Drawing.Point(105, 49)
        Me.frmLogin.Size = New System.Drawing.Size(356, 356)
        Me.frmLogin.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.frmLogin.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmbUsername
        '
        Me.cmbUsername.Name = "cmbUsername"
        Me.cmbUsername.TabIndex = 10
        Me.cmbUsername.Location = New System.Drawing.Point(158, 100)
        Me.cmbUsername.Size = New System.Drawing.Size(168, 23)
        Me.cmbUsername.Text = ""
        Me.cmbUsername.BackColor = System.Drawing.SystemColors.Window
        Me.cmbUsername.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbUsername.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtConfrmpass
        '
        Me.txtConfrmpass.Name = "txtConfrmpass"
        Me.txtConfrmpass.TabIndex = 8
        Me.txtConfrmpass.Location = New System.Drawing.Point(158, 237)
        Me.txtConfrmpass.Size = New System.Drawing.Size(126, 28)
        Me.txtConfrmpass.Text = ""
        Me.txtConfrmpass.BackColor = System.Drawing.SystemColors.Window
        Me.txtConfrmpass.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtConfrmpass.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdCancel
        '
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Location = New System.Drawing.Point(202, 299)
        Me.cmdCancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Image4
        '
        Me.Image4.Name = "Image4"
        Me.Image4.Location = New System.Drawing.Point(283, 191)
        Me.Image4.Size = New System.Drawing.Size(42, 28)
        Me.Image4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image4.BackColor = System.Drawing.SystemColors.Control
        Me.Image4.Image = CType(Resources.GetObject("Image4.Image"), System.Drawing.Bitmap)
        '
        'Image3
        '
        Me.Image3.Name = "Image3"
        Me.Image3.Location = New System.Drawing.Point(283, 191)
        Me.Image3.Size = New System.Drawing.Size(42, 28)
        Me.Image3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image3.BackColor = System.Drawing.SystemColors.Control
        Me.Image3.Image = CType(Resources.GetObject("Image3.Image"), System.Drawing.Bitmap)
        '
        'Image6
        '
        Me.Image6.Name = "Image6"
        Me.Image6.Location = New System.Drawing.Point(283, 237)
        Me.Image6.Size = New System.Drawing.Size(42, 28)
        Me.Image6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image6.BackColor = System.Drawing.SystemColors.Control
        Me.Image6.Image = CType(Resources.GetObject("Image6.Image"), System.Drawing.Bitmap)
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 11
        Me.Label1.Location = New System.Drawing.Point(0, 24)
        Me.Label1.Size = New System.Drawing.Size(357, 33)
        Me.Label1.Text = "CHANGE TEACHER PASSWORD"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label5
        '
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 9
        Me.Label5.Location = New System.Drawing.Point(30, 243)
        Me.Label5.Size = New System.Drawing.Size(125, 28)
        Me.Label5.Text = "Confirm Password"
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label4
        '
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 7
        Me.Label4.Location = New System.Drawing.Point(30, 197)
        Me.Label4.Size = New System.Drawing.Size(125, 28)
        Me.Label4.Text = "NewPassword"
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 4
        Me.Label2.Location = New System.Drawing.Point(32, 105)
        Me.Label2.Size = New System.Drawing.Size(128, 28)
        Me.Label2.Text = "User Name"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmChangePass
        '
        Me.ClientSize = New System.Drawing.Size(571, 464)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.frmLogin})
        Me.Name = "frmChangePass"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmChangePass.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Change Password"
        Me.cmbUsername.ResumeLayout(False)
        Me.txtConfrmpass.ResumeLayout(False)
        Me.cmdCancel.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Label5.ResumeLayout(False)
        Me.Label4.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.frmLogin.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Dim x As String
    Dim s As Boolean
    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub cmdChange_Click()
'#Const def_cmdChange_Click = True
#If def_cmdChange_Click
        If (txtConfrmpass.Text=txtNewpass.Text) Then
            x = cmbUsername.Text
            rs = con.Execute("select * from Userlogin where Username='"+cmbUsername.Text+"' and UPassword='"+txtcurpassword.Text+"'")
            If ( Not rs.EOF) Then
                s = True
                'con.Execute ("UPDATE Userlogin set Password='anup' where Username='" & cmbUsername.Text & "'")

                'MsgBox "Password successfully updated!!", vbInformation, "Offline Examiner"

            End If
        Else
            MsgBox("Password Mismatch!!", MsgBoxStyle.Information, " CES")
            txtConfrmpass.Text = ""
            txtNewpass.Text = ""
            txtNewpass.Focus()
        End If

        If (s=True) Then
            On Error Resume Next
            con.Execute(("UPDATE Userlogin set UPassword='"+txtNewpass.Text+"' where Username='"+cmbUsername.Text+"'"))

            MsgBox("Password successfully updated!!", MsgBoxStyle.Information, " CES")
            cmdChange.Enabled = False
        Else
            MsgBox("Invalid Password", MsgBoxStyle.Critical, " CES")
            txtcurpassword.Text = ""
            txtcurpassword.Focus()
        End If
#End If	' def_cmdChange_Click
    End Sub

    Private Sub frmChangePass_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmChangePass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
        rs = con.Execute("select * from Userlogin")
        While ( Not rs.EOF)
            cmbUsername.Items.Add(rs(0))
            rs.MoveNext()
        End While
        rs.Close()
        s = False
#End If	' def_Form_Load
    End Sub

    Private Sub Image1_MouseDown(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image1_MouseDown = True
#If def_Image1_MouseDown
        txtcurpassword.PasswordChar = ""
        txtcurpassword.Refresh()
        Image1.Visible = False
        Image2.Visible = True
#End If	' def_Image1_MouseDown
    End Sub

    Private Sub Image1_MouseUp(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image1_MouseUp = True
#If def_Image1_MouseUp
        txtcurpassword.PasswordChar = "*"
        txtcurpassword.Refresh()
        Image1.Visible = True
        Image2.Visible = False
#End If	' def_Image1_MouseUp
    End Sub

    Private Sub Image4_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Image4.MouseDown
        Dim Button As Short = e.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(e.X)
        Dim Y As Single = VB6.PixelsToTwipsY(e.Y)
'#Const def_Image4_MouseDown = True
#If def_Image4_MouseDown
        txtNewpass.PasswordChar = ""
        txtNewpass.Refresh()
        Image4.Visible = False
        Image3.Visible = True
#End If	' def_Image4_MouseDown
    End Sub

    Private Sub Image4_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Image4.MouseUp
        Dim Button As Short = e.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(e.X)
        Dim Y As Single = VB6.PixelsToTwipsY(e.Y)
'#Const def_Image4_MouseUp = True
#If def_Image4_MouseUp
        txtNewpass.PasswordChar = "*"
        txtNewpass.Refresh()
        Image4.Visible = True
        Image3.Visible = False
#End If	' def_Image4_MouseUp
    End Sub

    Private Sub Image5_MouseDown(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image5_MouseDown = True
#If def_Image5_MouseDown
        txtConfrmpass.PasswordChar = ""
        txtConfrmpass.Refresh()
        Image5.Visible = False
        Image6.Visible = True
#End If	' def_Image5_MouseDown
    End Sub

    Private Sub Image5_MouseUp(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image5_MouseUp = True
#If def_Image5_MouseUp
        txtConfrmpass.PasswordChar = "*"
        txtConfrmpass.Refresh()
        Image5.Visible = True
        Image6.Visible = False
#End If	' def_Image5_MouseUp
    End Sub

End Class