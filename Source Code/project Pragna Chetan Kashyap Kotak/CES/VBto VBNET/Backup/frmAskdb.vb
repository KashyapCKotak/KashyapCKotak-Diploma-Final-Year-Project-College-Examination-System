Imports VB = Microsoft.VisualBasic

Public Class frmAskdb
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
    Friend WithEvents Picture1 As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    Friend WithEvents SSTab1 As System.Windows.Forms.TabControl
    Friend WithEvents SSTab1_Page1 As System.Windows.Forms.TabPage
    Friend WithEvents SSTab1_Page2 As System.Windows.Forms.TabPage
    Friend WithEvents Command2 As System.Windows.Forms.Button
    Friend WithEvents Text2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAskdb))
        Me.components = New System.ComponentModel.Container()
        Me.Picture1 = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
        Me.SSTab1 = New System.Windows.Forms.TabControl()
        Me.SSTab1_Page1 = New System.Windows.Forms.TabPage()
        Me.SSTab1_Page2 = New System.Windows.Forms.TabPage()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'SSTab1
        '
        Me.SSTab1.Controls.AddRange(New System.Windows.Forms.Control() {Me.SSTab1_Page1, Me.SSTab1_Page2})
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.TabIndex = 1
        Me.SSTab1.Location = New System.Drawing.Point(40, 105)
        Me.SSTab1.Size = New System.Drawing.Size(310, 228)
        '
        'SSTab1_Page1
        '
        Me.SSTab1_Page1.Name = "SSTab1_Page1"
        Me.SSTab1_Page1.Size = New System.Drawing.Size(310, 208)
        Me.SSTab1_Page1.Text = "Local"
        '
        'SSTab1_Page2
        '
        Me.SSTab1_Page2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Command2, Me.Text2, Me.Label3})
        Me.SSTab1_Page2.Name = "SSTab1_Page2"
        Me.SSTab1_Page2.Size = New System.Drawing.Size(310, 208)
        Me.SSTab1_Page2.Text = "Remote"
        '
        'Command2
        '
        Me.Command2.Name = "Command2"
        Me.Command2.TabIndex = 11
        Me.Command2.Location = New System.Drawing.Point(97, 166)
        Me.Command2.Size = New System.Drawing.Size(122, 25)
        Me.Command2.Text = "ACCEPT"
        Me.Command2.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Font = New System.Drawing.Font("Times New Roman", 12.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Text2
        '
        Me.Text2.Name = "Text2"
        Me.Text2.TabIndex = 8
        Me.Text2.Location = New System.Drawing.Point(162, 45)
        Me.Text2.Size = New System.Drawing.Size(130, 25)
        Me.Text2.Text = ""
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 6
        Me.Label3.Location = New System.Drawing.Point(81, 53)
        Me.Label3.Size = New System.Drawing.Size(90, 17)
        Me.Label3.Text = "Enter IP add.:"
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 9.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 0
        Me.Label1.Location = New System.Drawing.Point(57, 32)
        Me.Label1.Size = New System.Drawing.Size(284, 58)
        Me.Label1.Text = "Select the location of the Database"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Font = New System.Drawing.Font("Arial Narrow", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmAskdb
        '
        Me.ClientSize = New System.Drawing.Size(391, 369)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.SSTab1, Me.Label1})
        Me.Name = "frmAskdb"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmAskdb.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Data Source"
        Me.Picture1.SetIndex(Picture1_0, CType(0, Short))
        CType(Me.Picture1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSTab1_Page1.ResumeLayout(False)
        Me.Command2.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.SSTab1_Page2.ResumeLayout(False)
        Me.SSTab1.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
'#Const defUse_Option1_Click = True
#If defUse_Option1_Click
    Private Sub Option1_Click()
'#Const def_Option1_Click = True
#If def_Option1_Click

#End If	' def_Option1_Click
    End Sub
#End If

'#Const defUse_Option2_Click = True
#If defUse_Option2_Click
    Private Sub Option2_Click()
'#Const def_Option2_Click = True
#If def_Option2_Click

#End If	' def_Option2_Click
    End Sub
#End If

'#Const defUse_Check1_Click = True
#If defUse_Check1_Click
    Private Sub Check1_Click()
'#Const def_Check1_Click = True
#If def_Check1_Click

#End If	' def_Check1_Click
    End Sub
#End If

'#Const defUse_TabStrip1_Click = True
#If defUse_TabStrip1_Click
    Private Sub TabStrip1_Click()
'#Const def_TabStrip1_Click = True
#If def_TabStrip1_Click

#End If	' def_TabStrip1_Click
    End Sub
#End If

    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click
        ' VBto upgrade warning: dbloc As Object --> As String
        Dim dbloc As String = ""	' - "AutoDim"

        str = Text1.Text
        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        'data = text1.Text
        'text1.Tag = data
        iFile = FreeFile
        FileOpen(iFile, Application.StartupPath & "\dbsource.txt", OpenMode.Output)
        PrintLine(iFile, str)
        FileClose(iFile)
        dbloc = "Local"
        iFile = FreeFile
        FileOpen(iFile, Application.StartupPath & "\loc.txt", OpenMode.Output)
        PrintLine(iFile, dbloc)
        FileClose(iFile)
        MDIForm1.StatusBar3.Panels.Item(1 - 1).Text = dbloc & " Data Source"


        strFilename = str & "\news1.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        MDIForm1.Text1.Text = strTheData

        strFilename = str & "\news2.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        MDIForm1.Text2.Text = strTheData

        strFilename = str & "\news3.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        MDIForm1.Text3.Text = strTheData


        Close()
#End If	' def_Command1_Click
    End Sub

    Private Sub Command2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command2.Click
'#Const def_Command2_Click = True
#If def_Command2_Click
        ' VBto upgrade warning: dbloc As Object --> As String
        Dim dbloc As String = ""	' - "AutoDim"

        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        str = "\\" & Text2.Text & Text3.Text
        'data = text1.Text
        'text1.Tag = data
        iFile = FreeFile
        FileOpen(iFile, str & "\dbsource.txt", OpenMode.Output)
        PrintLine(iFile, str)
        FileClose(iFile)
        dbloc = "Remote"
        iFile = FreeFile
        FileOpen(iFile, str & "\loc.txt", OpenMode.Output)
        PrintLine(iFile, dbloc)
        FileClose(iFile)
        MDIForm1.StatusBar3.Panels.Item(1 - 1).Text = dbloc & " Data Source"

        strFilename = str & "\news1.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        MDIForm1.Text1.Text = strTheData

        strFilename = str & "\news2.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        MDIForm1.Text2.Text = strTheData

        strFilename = str & "\news3.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        MDIForm1.Text3.Text = strTheData


        Close()
#End If	' def_Command2_Click
    End Sub


End Class