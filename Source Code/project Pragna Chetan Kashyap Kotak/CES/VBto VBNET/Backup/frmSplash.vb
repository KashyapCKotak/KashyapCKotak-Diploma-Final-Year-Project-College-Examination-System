Imports VB = Microsoft.VisualBasic

Public Class frmSplash
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
    Friend WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Friend WithEvents Picture1 As System.Windows.Forms.PictureBox
    Friend WithEvents Image1 As System.Windows.Forms.PictureBox
    Friend WithEvents Frame1 As System.Windows.Forms.Panel
    Friend WithEvents lblPlatform As System.Windows.Forms.Label
    Friend WithEvents lblProductName As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSplash))
        Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.Image1 = New System.Windows.Forms.PictureBox()
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.lblPlatform = New System.Windows.Forms.Label()
        Me.lblProductName = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'File1
        '
        Me.File1.Name = "File1"
        Me.File1.Visible = False
        Me.File1.TabIndex = 7
        Me.File1.Location = New System.Drawing.Point(275, 283)
        Me.File1.Size = New System.Drawing.Size(176, 19)
        '
        'Picture1
        '
        Me.Picture1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Image1})
        Me.Picture1.Name = "Picture1"
        Me.Picture1.TabIndex = 5
        Me.Picture1.Location = New System.Drawing.Point(15, 312)
        Me.Picture1.Size = New System.Drawing.Size(462, 22)
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Picture1.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(64, Byte), CType(0, Byte))
        '
        'Image1
        '
        Me.Image1.Name = "Image1"
        Me.Image1.Location = New System.Drawing.Point(0, 0)
        Me.Image1.Size = New System.Drawing.Size(25, 19)
        Me.Image1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Image1.BackColor = System.Drawing.SystemColors.Control
        Me.Image1.Image = CType(Resources.GetObject("Image1.Image"), System.Drawing.Bitmap)
        Me.Image1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblPlatform, Me.lblProductName})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(8, 8)
        Me.Frame1.Size = New System.Drawing.Size(477, 269)
        Me.Frame1.BackColor = System.Drawing.SystemColors.HighlightText
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'lblPlatform
        '
        Me.lblPlatform.Name = "lblPlatform"
        Me.lblPlatform.TabIndex = 3
        Me.lblPlatform.AutoSize = True
        Me.lblPlatform.Location = New System.Drawing.Point(146, 218)
        Me.lblPlatform.Size = New System.Drawing.Size(330, 24)
        Me.lblPlatform.Text = "Platform: Windows XP and Later"
        Me.lblPlatform.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblPlatform.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlatform.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblPlatform.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblProductName
        '
        Me.lblProductName.Name = "lblProductName"
        Me.lblProductName.TabIndex = 4
        Me.lblProductName.AutoSize = True
        Me.lblProductName.Location = New System.Drawing.Point(178, 0)
        Me.lblProductName.Size = New System.Drawing.Size(299, 155)
        Me.lblProductName.Text = "College Examination System "
        Me.lblProductName.BackColor = System.Drawing.SystemColors.HighlightText
        Me.lblProductName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProductName.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblProductName.Font = New System.Drawing.Font("Arial", 32.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 6
        Me.Label1.Location = New System.Drawing.Point(8, 283)
        Me.Label1.Size = New System.Drawing.Size(71, 19)
        Me.Label1.Text = "Loading Files:"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        '
        'frmSplash
        '
        Me.ClientSize = New System.Drawing.Size(498, 344)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.File1, Me.Picture1, Me.Frame1, Me.Label1})
        Me.Name = "frmSplash"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(64, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ShowInTaskbar = False
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ControlBox = False
        Me.Icon = CType(Resources.GetObject("frmSplash.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = ""
        Me.Picture1.ResumeLayout(False)
        Me.lblPlatform.ResumeLayout(False)
        Me.lblProductName.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    ' VBto upgrade warning: x As Short	OnWrite(Integer)
    Dim x As Short
    Dim i As Short

    Private Sub frmSplash_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load

        File1.FileName = str
        x = File1.Items.Count
#End If	' def_Form_Load
    End Sub

    Private Sub Frame1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Frame1.Click
'#Const def_Frame1_Click = True
#If def_Frame1_Click

        LoadUnUsed(MDIForm1)
        ShowModeless(MDIForm1)
        Close()
#End If	' def_Frame1_Click
    End Sub

'#Const defUse_Label3_Click = True
#If defUse_Label3_Click
    Private Sub Label3_Click()
'#Const def_Label3_Click = True
#If def_Label3_Click

#End If	' def_Label3_Click
    End Sub
#End If

    Private Sub Timer1_Timer()
'#Const def_Timer1_Timer = True
#If def_Timer1_Timer
        time2 += 1
        If (Image1.Left<=6435) Then
            Image1.Left = Image1.Left+100
        Else
            Image1.Left = 0
        End If
        If (i<=x) Then
            Label2.Text = File1.Items.Item(i).ToString()
            i += 1
        Else
            LoadUnUsed(MDIForm1)
            ShowModeless(MDIForm1)
            Close()
        End If
        If time2=50 Then
            LoadUnUsed(MDIForm1)
            ShowModeless(MDIForm1)
            Close()
        End If

#End If	' def_Timer1_Timer
    End Sub

End Class