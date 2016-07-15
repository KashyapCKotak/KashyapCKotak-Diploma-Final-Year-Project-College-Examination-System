Imports VB = Microsoft.VisualBasic

Public Class frmTT
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
    Friend WithEvents Picture1 As System.Windows.Forms.PictureBox
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents text1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmTT))
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.text1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Picture1
        '
        Me.Picture1.Name = "Picture1"
        Me.Picture1.TabIndex = 4
        Me.Picture1.Location = New System.Drawing.Point(16, 534)
        Me.Picture1.Size = New System.Drawing.Size(251, 79)
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.Image = CType(Resources.GetObject("Picture1.Image"), System.Drawing.Bitmap)
        '
        'Command1
        '
        Me.Command1.Name = "Command1"
        Me.Command1.TabIndex = 3
        Me.Command1.Location = New System.Drawing.Point(639, 550)
        Me.Command1.Size = New System.Drawing.Size(147, 33)
        Me.Command1.Text = "OK"
        Me.Command1.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(255, Byte), CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'text1
        '
        Me.text1.Name = "text1"
        Me.text1.Enabled = False
        Me.text1.TabIndex = 0
        Me.text1.Location = New System.Drawing.Point(16, 32)
        Me.text1.Size = New System.Drawing.Size(770, 503)
        Me.text1.Text = ""
        Me.text1.BackColor = System.Drawing.SystemColors.Window
        Me.text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.text1.Multiline = True
        Me.text1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.text1.Font = New System.Drawing.Font("Times New Roman", 12.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 2
        Me.Label2.Location = New System.Drawing.Point(129, 550)
        Me.Label2.Size = New System.Drawing.Size(648, 33)
        Me.Label2.Text = "ALL THE BEST"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(255, Byte), CType(0, Byte))
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label2.Font = New System.Drawing.Font("Lucida Calligraphy", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Location = New System.Drawing.Point(81, 0)
        Me.Label1.Size = New System.Drawing.Size(640, 41)
        Me.Label1.Text = "TIME TABLE"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(255, Byte), CType(0, Byte))
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 24.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmTT
        '
        Me.ClientSize = New System.Drawing.Size(804, 617)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Picture1, Me.Command1, Me.text1, Me.Label2, Me.Label1})
        Me.Name = "frmTT"
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = True
        Me.MaximizeBox = True
        Me.Icon = CType(Resources.GetObject("frmTT.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Time Table"
        Me.Command1.ResumeLayout(False)
        Me.text1.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
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
        FileOpen(iFile, str & "\timetable.txt", OpenMode.Output)
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

    Private Sub frmTT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmTT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.Width = 12165
        Me.Height = 9690
        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        If sh=1 Then
            Command1.Enabled = True
            Me.Text1.Enabled = True
            ' data = text1.Tag
            strFilename = str & "\timetable.txt"
            iFile = FreeFile
            FileOpen(iFile, strFilename, OpenMode.Input)
            strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
            FileClose(iFile)
            Text1.Text = strTheData
            'text1.Text = data
        Else
            Command1.Enabled = True
            data = Text1.Tag
            Text1.Text = data
            strFilename = str & "\timetable.txt"
            iFile = FreeFile
            FileOpen(iFile, strFilename, OpenMode.Input)
            strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
            FileClose(iFile)
            Text1.Text = strTheData
            Me.Text1.Enabled = False

        End If
        Me.Width = 12960
        Me.Height = 9960
        Me.Top = 0
        Me.Left = 3510
        'Dim iFile As Long
        'Dim strFilename As String
        'Dim strTheData As String

        'strFilename = "C:\Documents and Settings\PragnaChetanKashyap.KOTAK-B43F5C7CD\Desktop\OfflineExaminer\Offline Examiner\timetable.txt"

        'iFile = FreeFile

        'Open strFilename For Input As #iFile
        ' strTheData = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
        'Close #iFile
        'text1.Text = strTheData


#End If	' def_Form_Load
    End Sub


End Class