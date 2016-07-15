Imports VB = Microsoft.VisualBasic

Public Class frmViewRptQst
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
    Friend WithEvents cmdVwRpt As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmViewRptQst))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmdVwRpt = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdVwRpt, Me.Label2, Me.Label1})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(97, 57)
        Me.Frame1.Size = New System.Drawing.Size(308, 189)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdVwRpt
        '
        Me.cmdVwRpt.Name = "cmdVwRpt"
        Me.cmdVwRpt.TabIndex = 2
        Me.cmdVwRpt.Location = New System.Drawing.Point(178, 138)
        Me.cmdVwRpt.Size = New System.Drawing.Size(106, 25)
        Me.cmdVwRpt.Text = "View Report"
        Me.cmdVwRpt.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdVwRpt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdVwRpt.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 4
        Me.Label2.Location = New System.Drawing.Point(0, 16)
        Me.Label2.Size = New System.Drawing.Size(308, 33)
        Me.Label2.Text = "VIEW REPORT"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Location = New System.Drawing.Point(24, 92)
        Me.Label1.Size = New System.Drawing.Size(90, 25)
        Me.Label1.Text = "Select ExamId"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmViewRptQst
        '
        Me.ClientSize = New System.Drawing.Size(499, 324)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmViewRptQst"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmViewRptQst.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "View Questions"
        Me.cmdVwRpt.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Dim x As String
'#Const defUse_cmdNext_Click = True
#If defUse_cmdNext_Click
    Private Sub cmdNext_Click()
'#Const def_cmdNext_Click = True
#If def_cmdNext_Click

#End If	' def_cmdNext_Click
    End Sub
#End If

    Private Sub cmdVwRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdVwRpt.Click
'#Const def_cmdVwRpt_Click = True
#If def_cmdVwRpt_Click
        x = cmbBranch.Text
        If (DataEnvironment1.rsCommand2.State=1) Then
            DataEnvironment1.rsCommand2.Close()
        Else
            DataEnvironment1.Command2((x))
            LoadUnUsed(DataReportQstExID)
            DataReportQstExID.Show()
        End If
#End If	' def_cmdVwRpt_Click
    End Sub

    Private Sub frmViewRptQst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmViewRptQst_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
        rs = con.Execute("select distinct(ExamID) from ExamDetails")
        While ( Not rs.EOF)
            cmbBranch.Items.Add(rs(0))
            rs.MoveNext()
        End While
        rs.Close()
#End If	' def_Form_Load
    End Sub

End Class