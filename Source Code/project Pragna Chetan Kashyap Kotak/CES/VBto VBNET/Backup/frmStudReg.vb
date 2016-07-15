Imports VB = Microsoft.VisualBasic

Public Class frmStudReg
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
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdReg As System.Windows.Forms.Button
    Friend WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents txtUname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmStudReg))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdReg = New System.Windows.Forms.Button()
        Me.txtPass = New System.Windows.Forms.TextBox()
        Me.txtUname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdReg, Me.txtPass, Me.txtUname, Me.Label2})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(97, 65)
        Me.Frame1.Size = New System.Drawing.Size(373, 276)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdCancel
        '
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Location = New System.Drawing.Point(210, 227)
        Me.cmdCancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdReg
        '
        Me.cmdReg.Name = "cmdReg"
        Me.cmdReg.TabIndex = 5
        Me.cmdReg.Location = New System.Drawing.Point(65, 227)
        Me.cmdReg.Size = New System.Drawing.Size(98, 25)
        Me.cmdReg.Text = "Register"
        Me.cmdReg.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdReg.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReg.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtPass
        '
        Me.txtPass.Name = "txtPass"
        Me.txtPass.TabIndex = 4
        Me.txtPass.Location = New System.Drawing.Point(146, 138)
        Me.txtPass.Size = New System.Drawing.Size(195, 28)
        Me.txtPass.Text = ""
        Me.txtPass.BackColor = System.Drawing.SystemColors.Window
        Me.txtPass.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPass.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPass.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'txtUname
        '
        Me.txtUname.Name = "txtUname"
        Me.txtUname.TabIndex = 3
        Me.txtUname.Location = New System.Drawing.Point(146, 89)
        Me.txtUname.Size = New System.Drawing.Size(195, 28)
        Me.txtUname.Text = ""
        Me.txtUname.BackColor = System.Drawing.SystemColors.Window
        Me.txtUname.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUname.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUname.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 1
        Me.Label2.Location = New System.Drawing.Point(57, 91)
        Me.Label2.Size = New System.Drawing.Size(89, 34)
        Me.Label2.Text = "Username"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmStudReg
        '
        Me.ClientSize = New System.Drawing.Size(571, 413)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmStudReg"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmStudReg.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Student Registration"
        Me.cmdCancel.ResumeLayout(False)
        Me.cmdReg.ResumeLayout(False)
        Me.txtPass.ResumeLayout(False)
        Me.txtUname.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    ' VBto upgrade warning: dec As Object --> As Short
    Dim dec As Short, n As Short
    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub cmdReg_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdReg.Click
'#Const def_cmdReg_Click = True
#If def_cmdReg_Click
        dec = 1
        n = 0
        If (txtPass.Text="" Or txtUname.Text="") Then
            MsgBox("Missing Fields", MsgBoxStyle.Information, " CES")
        Else
            rs = con.Execute("select * from Student where Username='"+txtUname.Text+"'")
            If ( Not rs.EOF) Then
                MsgBox("Username already exist! please try another name", MsgBoxStyle.Information, " CES")
                dec = 0
                rs = con.Execute("select ucounter from Student where Username='"+txtUname.Text+"'")
                If ( Not rs.EOF) Then
                    cnt = rs(0)
                End If
                rs.Close()
                If (cnt>=3) Then
                    con.Execute(("update Student set ucounter=" & n & " where Username='" & txtUname.Text & "'"))
                    MsgBox("Re-Registered successfully! Login and attend the exam", MsgBoxStyle.Information, " CES")
                    dec = 0
                End If
            End If
            If dec=1 Then
                con.Execute(("insert into Student values('"+txtUname.Text+"','"+txtPass.Text+"','0',NULL,NULL,NULL)"))
                MsgBox("Registered successfully! Login and attend the exam", MsgBoxStyle.Information, " CES")
            End If
            If dec=0 And n=1 Then
                LoadUnUsed(frmAsk)
                frmAsk.Show()
            End If
        End If
#End If	' def_cmdReg_Click
    End Sub

    Private Sub frmStudReg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmStudReg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
#End If	' def_Form_Load
    End Sub

End Class