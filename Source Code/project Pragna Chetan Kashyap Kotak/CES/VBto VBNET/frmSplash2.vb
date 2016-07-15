Imports VB = Microsoft.VisualBasic

Public Class frmSplash2
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
    Friend WithEvents Command2 As System.Windows.Forms.Button
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents Frame1 As System.Windows.Forms.Panel
    Friend WithEvents lblWarning As System.Windows.Forms.Label
    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents lblProductName As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSplash2))
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.lblProductName = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Command2
        '
        Me.Command2.Name = "Command2"
        Me.Command2.TabIndex = 5
        Me.Command2.Location = New System.Drawing.Point(65, 332)
        Me.Command2.Size = New System.Drawing.Size(130, 33)
        Me.Command2.Text = "GO BACK"
        Me.Command2.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(255, Byte), CType(128, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Command1
        '
        Me.Command1.Name = "Command1"
        Me.Command1.TabIndex = 4
        Me.Command1.Location = New System.Drawing.Point(413, 332)
        Me.Command1.Size = New System.Drawing.Size(130, 33)
        Me.Command1.Text = "EXIT"
        Me.Command1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblWarning, Me.lblVersion, Me.lblProductName})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(65, 49)
        Me.Frame1.Size = New System.Drawing.Size(477, 273)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(255, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'lblWarning
        '
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.TabIndex = 3
        Me.lblWarning.Location = New System.Drawing.Point(0, 259)
        Me.lblWarning.Size = New System.Drawing.Size(357, 13)
        Me.lblWarning.Text = " Warning: Copyright Protected"
        Me.lblWarning.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblWarning.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWarning.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblVersion
        '
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.TabIndex = 2
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(367, 251)
        Me.lblVersion.Size = New System.Drawing.Size(105, 19)
        Me.lblVersion.Text = "Version: 3.0.1"
        Me.lblVersion.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 12.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblProductName
        '
        Me.lblProductName.Name = "lblProductName"
        Me.lblProductName.TabIndex = 1
        Me.lblProductName.AutoSize = True
        Me.lblProductName.Location = New System.Drawing.Point(170, 49)
        Me.lblProductName.Size = New System.Drawing.Size(295, 155)
        Me.lblProductName.Text = "College Examination System"
        Me.lblProductName.BackColor = System.Drawing.Color.Transparent
        Me.lblProductName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProductName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblProductName.Font = New System.Drawing.Font("Arial", 32.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmSplash2
        '
        Me.ClientSize = New System.Drawing.Size(603, 396)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Command2, Me.Command1, Me.Frame1})
        Me.Name = "frmSplash2"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(64, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ShowInTaskbar = False
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ControlBox = False
        Me.Icon = CType(Resources.GetObject("frmSplash2.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = ""
        Me.Command2.ResumeLayout(False)
        Me.Command1.ResumeLayout(False)
        Me.lblWarning.ResumeLayout(False)
        Me.lblVersion.ResumeLayout(False)
        Me.lblProductName.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================


    Private Sub Command1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command1.Click
'#Const def_Command1_Click = True
#If def_Command1_Click
        Application.Exit()
#End If	' def_Command1_Click
    End Sub

    Private Sub Command2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command2.Click
'#Const def_Command2_Click = True
#If def_Command2_Click
        LoadUnUsed(MDIForm1)
        ShowModeless(MDIForm1)
#End If	' def_Command2_Click
    End Sub

    Private Sub frmSplash2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
'#Const def_Form_KeyPress = True
#If def_Form_KeyPress
        Close()
#End If	' def_Form_KeyPress
    End Sub

    Private Sub Frame1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Frame1.Click
'#Const def_Frame1_Click = True
#If def_Frame1_Click
        Close()
#End If	' def_Frame1_Click
    End Sub

'#Const defUse_Timer1_Timer = True
#If defUse_Timer1_Timer
    Private Sub Timer1_Timer()
'#Const def_Timer1_Timer = True
#If def_Timer1_Timer

#End If	' def_Timer1_Timer
    End Sub
#End If


End Class