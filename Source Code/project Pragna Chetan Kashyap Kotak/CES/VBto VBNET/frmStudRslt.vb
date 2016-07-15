Imports VB = Microsoft.VisualBasic

Public Class frmStudRslt
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
    Friend WithEvents cmdLoad As System.Windows.Forms.Button
    Friend WithEvents cmdVwRpt As System.Windows.Forms.Button
    Friend WithEvents CmbExamID As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmStudRslt))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmdLoad = New System.Windows.Forms.Button()
        Me.cmdVwRpt = New System.Windows.Forms.Button()
        Me.CmbExamID = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdLoad, Me.cmdVwRpt, Me.CmbExamID, Me.Label3, Me.Label2, Me.Label1})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(81, 57)
        Me.Frame1.Size = New System.Drawing.Size(326, 274)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdLoad
        '
        Me.cmdLoad.Name = "cmdLoad"
        Me.cmdLoad.TabIndex = 6
        Me.cmdLoad.Location = New System.Drawing.Point(210, 129)
        Me.cmdLoad.Size = New System.Drawing.Size(98, 25)
        Me.cmdLoad.Text = "Load ExamID"
        Me.cmdLoad.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdLoad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLoad.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdVwRpt
        '
        Me.cmdVwRpt.Name = "cmdVwRpt"
        Me.cmdVwRpt.TabIndex = 5
        Me.cmdVwRpt.Location = New System.Drawing.Point(210, 218)
        Me.cmdVwRpt.Size = New System.Drawing.Size(98, 25)
        Me.cmdVwRpt.Text = "View Report"
        Me.cmdVwRpt.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdVwRpt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdVwRpt.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'CmbExamID
        '
        Me.CmbExamID.Name = "CmbExamID"
        Me.CmbExamID.TabIndex = 2
        Me.CmbExamID.Location = New System.Drawing.Point(127, 178)
        Me.CmbExamID.Size = New System.Drawing.Size(183, 23)
        Me.CmbExamID.Text = ""
        Me.CmbExamID.BackColor = System.Drawing.SystemColors.Window
        Me.CmbExamID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CmbExamID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CmbExamID.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 7
        Me.Label3.Location = New System.Drawing.Point(0, 16)
        Me.Label3.Size = New System.Drawing.Size(322, 27)
        Me.Label3.Text = "VIEW RESULT"
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 4
        Me.Label2.Location = New System.Drawing.Point(16, 181)
        Me.Label2.Size = New System.Drawing.Size(101, 25)
        Me.Label2.Text = "Select ExamID"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Location = New System.Drawing.Point(16, 82)
        Me.Label1.Size = New System.Drawing.Size(110, 28)
        Me.Label1.Text = "Username"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmStudRslt
        '
        Me.ClientSize = New System.Drawing.Size(488, 397)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmStudRslt"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmStudRslt.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Student Result"
        Me.cmdLoad.ResumeLayout(False)
        Me.cmdVwRpt.ResumeLayout(False)
        Me.CmbExamID.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    ' VBto upgrade warning: a As Object --> As String
    Dim a As String, b As String

    Private Sub cmdLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdLoad.Click
'#Const def_cmdLoad_Click = True
#If def_cmdLoad_Click
        If (txtUsername.Text="") Then
            MsgBox("Enter Username", MsgBoxStyle.Information, " CES")
        Else
            rs = con.Execute("select distinct(ExamID) from Result where Username='"+txtUsername.Text+"' ")
            While ( Not rs.EOF)
                CmbExamID.Items.Add(rs(0))
                rs.MoveNext()
            End While
            rs.Close()
        End If
#End If	' def_cmdLoad_Click
    End Sub

    Private Sub cmdVwRpt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdVwRpt.Click
'#Const def_cmdVwRpt_Click = True
#If def_cmdVwRpt_Click
        a = txtUsername.Text
        b = CmbExamID.Text
        If (DataEnvironment1.rsCommand3.State=1) Then
            DataEnvironment1.rsCommand3.Close()
        Else
            DataEnvironment1.Command3((b))
            LoadUnUsed(DataReportStudRslt)
            DataReportStudRslt.Show()

        End If
#End If	' def_cmdVwRpt_Click
    End Sub

    Private Sub frmStudRslt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmStudRslt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        connectdb()
        Me.BackColor = ColorTranslator.FromOle(color)
#End If	' def_Form_Load
    End Sub

End Class