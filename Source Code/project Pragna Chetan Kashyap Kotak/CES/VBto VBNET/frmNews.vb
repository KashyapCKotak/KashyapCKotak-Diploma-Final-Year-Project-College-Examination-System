Imports VB = Microsoft.VisualBasic

Public Class frmNews
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
    Friend WithEvents Text3 As System.Windows.Forms.TextBox
    Friend WithEvents Text2 As System.Windows.Forms.TextBox
    Friend WithEvents text1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNews))
        Me.Text3 = New System.Windows.Forms.TextBox()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.text1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Text3
        '
        Me.Text3.Name = "Text3"
        Me.Text3.TabIndex = 2
        Me.Text3.Location = New System.Drawing.Point(0, 146)
        Me.Text3.Size = New System.Drawing.Size(1369, 25)
        Me.Text3.Text = ""
        Me.Text3.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.Text3.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(192, Byte))
        Me.Text3.Font = New System.Drawing.Font("MS Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Text2
        '
        Me.Text2.Name = "Text2"
        Me.Text2.TabIndex = 1
        Me.Text2.Location = New System.Drawing.Point(0, 113)
        Me.Text2.Size = New System.Drawing.Size(1369, 25)
        Me.Text2.Text = ""
        Me.Text2.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.Text2.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(192, Byte))
        Me.Text2.Font = New System.Drawing.Font("MS Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'text1
        '
        Me.text1.Name = "text1"
        Me.text1.TabIndex = 0
        Me.text1.Location = New System.Drawing.Point(0, 81)
        Me.text1.Size = New System.Drawing.Size(1369, 25)
        Me.text1.Text = ""
        Me.text1.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(128, Byte), CType(255, Byte))
        Me.text1.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(192, Byte))
        Me.text1.Font = New System.Drawing.Font("MS Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 3
        Me.Label1.Location = New System.Drawing.Point(0, 16)
        Me.Label1.Size = New System.Drawing.Size(1368, 33)
        Me.Label1.Text = "EDIT NEWS-LINES"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmNews
        '
        Me.ClientSize = New System.Drawing.Size(1365, 226)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Text3, Me.Text2, Me.text1, Me.Label1})
        Me.Name = "frmNews"
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = True
        Me.MaximizeBox = True
        Me.Icon = CType(Resources.GetObject("frmNews.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Edit News"
        Me.Text3.ResumeLayout(False)
        Me.Text2.ResumeLayout(False)
        Me.text1.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click
        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        'data = text1.Text
        'text1.Tag = data
        iFile = FreeFile
        FileOpen(iFile, str & "\news1.txt", OpenMode.Output)
        PrintLine(iFile, Str(Text1.Text))
        FileClose(iFile)
        '  strFilename = str & "\timetable.txt"
        '  iFile = FreeFile
        '  Open strFilename For Output As #iFile
        '  'StrConv(InputB(LOF(iFile), iFile), vbUnicode) = strTheData
        '  Close #iFile
        '  text1.Text = strTheData
        iFile = FreeFile
        FileOpen(iFile, str & "\news2.txt", OpenMode.Output)
        PrintLine(iFile, Text2.Text)
        FileClose(iFile)

        iFile = FreeFile
        FileOpen(iFile, str & "\news3.txt", OpenMode.Output)
        PrintLine(iFile, Text3.Text)
        FileClose(iFile)
        MDIForm1.Text1.Refresh()
        MDIForm1.Text2.Refresh()
        MDIForm1.Text3.Refresh()
        Close()
#End If	' def_Command1_Click
    End Sub

    'Private Sub frmNews_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.ValueChanged
    '   If Not sender.Created() Then Exit Sub
    '#Const def_Form_Change = True
    '#If def_Form_Change Then

    '#End If ' def_Form_Change
    '    End Sub

    Private Sub frmNews_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub frmNews_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""

        strFilename = str & "\news1.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        Text1.Text = strTheData

        strFilename = str & "\news2.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        Text2.Text = strTheData

        strFilename = str & "\news3.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        Text3.Text = strTheData

#End If	' def_Form_Load
    End Sub


End Class