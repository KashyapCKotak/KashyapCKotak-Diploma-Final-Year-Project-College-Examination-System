Imports VB = Microsoft.VisualBasic

Public Class frmStartExam
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
    Friend WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents cmdEnter As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Image2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmStartExam))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.txtPass = New System.Windows.Forms.TextBox()
        Me.cmdEnter = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Image2 = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtPass, Me.cmdEnter, Me.cmdCancel, Me.Image2, Me.Label2, Me.Label3})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(95, 65)
        Me.Frame1.Size = New System.Drawing.Size(341, 237)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'txtPass
        '
        Me.txtPass.Name = "txtPass"
        Me.txtPass.TabIndex = 2
        Me.txtPass.Location = New System.Drawing.Point(158, 113)
        Me.txtPass.Size = New System.Drawing.Size(124, 28)
        Me.txtPass.Text = ""
        Me.txtPass.BackColor = System.Drawing.SystemColors.Window
        Me.txtPass.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPass.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdEnter
        '
        Me.cmdEnter.Name = "cmdEnter"
        Me.cmdEnter.TabIndex = 3
        Me.cmdEnter.Location = New System.Drawing.Point(57, 178)
        Me.cmdEnter.Size = New System.Drawing.Size(98, 25)
        Me.cmdEnter.Text = "Enter"
        Me.cmdEnter.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdEnter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnter.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdCancel
        '
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Location = New System.Drawing.Point(186, 178)
        Me.cmdCancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Image2
        '
        Me.Image2.Name = "Image2"
        Me.Image2.Location = New System.Drawing.Point(280, 113)
        Me.Image2.Size = New System.Drawing.Size(42, 28)
        Me.Image2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image2.BackColor = System.Drawing.SystemColors.Control
        Me.Image2.Image = CType(Resources.GetObject("Image2.Image"), System.Drawing.Bitmap)
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 6
        Me.Label2.Location = New System.Drawing.Point(40, 65)
        Me.Label2.Size = New System.Drawing.Size(89, 34)
        Me.Label2.Text = "Username"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 5
        Me.Label3.Location = New System.Drawing.Point(40, 121)
        Me.Label3.Size = New System.Drawing.Size(119, 37)
        Me.Label3.Text = "Password"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmStartExam
        '
        Me.ClientSize = New System.Drawing.Size(531, 339)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmStartExam"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmStartExam.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Start Exam"
        Me.txtPass.ResumeLayout(False)
        Me.cmdEnter.ResumeLayout(False)
        Me.cmdCancel.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
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



    Private Sub cmdEnter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEnter.Click
'#Const def_cmdEnter_Click = True
#If def_cmdEnter_Click
        rs = con.Execute("select ucounter from Student where Username='"+txtUname.Text+"'")
        If ( Not rs.EOF) Then
            cnt = rs(0)
        End If
        rs.Close()
        If (cnt<3) Then
            rs = con.Execute("select * from Student where Username='"+txtUname.Text+"' and Password='"+txtPass.Text+"'")
            If ( Not rs.EOF) Then
                MsgBox("Login Success", MsgBoxStyle.Information, " CES")
                uname = txtUname.Text
                MDIForm1.StatusBar1.Panels.Item(2 - 1).Text = txtUname.Text
                LoadUnUsed(frmSelectExam)
                ShowModeless(frmSelectExam)
                Close()
                'rs.Close
            Else
                MsgBox("Invalid Username or Password", MsgBoxStyle.Critical, " CES")
            End If

        Else
            MsgBox("Your number of attempts over, Please reregister", MsgBoxStyle.Critical, " CES")
        End If
#End If	' def_cmdEnter_Click
    End Sub

    Private Sub frmStartExam_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmStartExam_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
#End If	' def_Form_Load
    End Sub

'#Const defUse_Label1_Click = True
#If defUse_Label1_Click
    Private Sub Label1_Click()
'#Const def_Label1_Click = True
#If def_Label1_Click

#End If	' def_Label1_Click
    End Sub
#End If

    Private Sub Image1_MouseDown(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image1_MouseDown = True
#If def_Image1_MouseDown
        txtPass.PasswordChar = ""
        txtPass.Refresh()
        Image1.Visible = False
        Image2.Visible = True

#End If	' def_Image1_MouseDown
    End Sub

    Private Sub Image1_MouseUp(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
'#Const def_Image1_MouseUp = True
#If def_Image1_MouseUp
        txtPass.PasswordChar = "*"
        txtPass.Refresh()
        Image1.Visible = True
        Image2.Visible = False
#End If	' def_Image1_MouseUp
    End Sub


End Class