Imports VB = Microsoft.VisualBasic

Public Class frmFeedback
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
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents Frame1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFeedback))
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Command1
        '
        Me.Command1.Name = "Command1"
        Me.Command1.TabIndex = 4
        Me.Command1.Location = New System.Drawing.Point(202, 340)
        Me.Command1.Size = New System.Drawing.Size(139, 33)
        Me.Command1.Text = "OK"
        Me.Command1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.Label1})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(65, 40)
        Me.Frame1.Size = New System.Drawing.Size(422, 292)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.Visible = False
        Me.Label2.TabIndex = 3
        Me.Label2.Location = New System.Drawing.Point(0, 259)
        Me.Label2.Size = New System.Drawing.Size(422, 33)
        Me.Label2.Text = "Please enter your name before feedback text"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label2.Font = New System.Drawing.Font("MS Sans Serif", 12.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Location = New System.Drawing.Point(0, 8)
        Me.Label1.Size = New System.Drawing.Size(422, 41)
        Me.Label1.Text = "FEEDBACK"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmFeedback
        '
        Me.ClientSize = New System.Drawing.Size(540, 385)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Command1, Me.Frame1})
        Me.Name = "frmFeedback"
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = True
        Me.MaximizeBox = True
        Me.Icon = CType(Resources.GetObject("frmFeedback.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "FEEDBACK"
        Me.Command1.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================

    Private Sub Command1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command1.Click
'#Const def_Command1_Click = True
#If def_Command1_Click
        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        'data = text1.Text
        'text1.Tag = data
        iFile = FreeFile
        FileOpen(iFile, str & "\feedback.txt", OpenMode.Output)
        PrintLine(iFile, Str(Text1.Text))
        FileClose(iFile)
        '  strFilename = str & "\timetable.txt"
        '  iFile = FreeFile
        '  Open strFilename For Output As #iFile
        '  'StrConv(InputB(LOF(iFile), iFile), vbUnicode) = strTheData
        '  Close #iFile
        '  text1.Text = strTheData

        Close()
#End If	' def_Command1_Click
    End Sub

    Private Sub frmFeedback_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmFeedback_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.Width = 8250
        Me.BackColor = ColorTranslator.FromOle(color)
        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        If xyz=15 Then
            Label2.Visible = True
            Command1.Enabled = True
            Me.Text1.Enabled = True
            ' data = text1.Tag
            strFilename = str & "\feedback.txt"
            iFile = FreeFile
            FileOpen(iFile, strFilename, OpenMode.Input)
            strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
            FileClose(iFile)
            Text1.Text = strTheData
            'text1.Text = data
        ElseIf xyz=0 Then
            Command1.Enabled = True
            Me.Text1.Enabled = False
            'data = text1.Tag
            'text1.Text = data
            strFilename = str & "\feedback.txt"
            iFile = FreeFile
            FileOpen(iFile, strFilename, OpenMode.Input)
            strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
            FileClose(iFile)
            Text1.Text = strTheData
            'frmTT.text1.Enabled = False
            '  ElseIf xyz = 0 Then
            '  Label2.Visible = False
            '  text1.Enabled = False

        End If
#End If	' def_Form_Load
    End Sub

End Class