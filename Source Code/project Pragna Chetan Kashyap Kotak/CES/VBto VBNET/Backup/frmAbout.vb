Imports VB = Microsoft.VisualBasic

Public Class frmAbout
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
    Friend WithEvents Picture1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.Picture1 = New System.Windows.Forms.PictureBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Picture1, Me.Label5, Me.Label3, Me.Label2, Me.Label4})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(16, 24)
        Me.Frame1.Size = New System.Drawing.Size(826, 527)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Picture1
        '
        Me.Picture1.Name = "Picture1"
        Me.Picture1.TabIndex = 2
        Me.Picture1.Location = New System.Drawing.Point(16, 121)
        Me.Picture1.Size = New System.Drawing.Size(284, 228)
        Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture1.BackColor = System.Drawing.SystemColors.Control
        Me.Picture1.Image = CType(Resources.GetObject("Picture1.Image"), System.Drawing.Bitmap)
        '
        'Label5
        '
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 7
        Me.Label5.Location = New System.Drawing.Point(316, 153)
        Me.Label5.Size = New System.Drawing.Size(438, 25)
        Me.Label5.Text = "COLLEGE EXAMINATION SYSTEM"
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(255, Byte), CType(255, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.Label5.Font = New System.Drawing.Font("Lucida Calligraphy", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 4
        Me.Label3.Location = New System.Drawing.Point(16, 364)
        Me.Label3.Size = New System.Drawing.Size(786, 33)
        Me.Label3.Text = "                  CONTACT AT:"
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(255, Byte), CType(255, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(255, Byte))
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Font = New System.Drawing.Font("Lucida Handwriting", 18.00!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 3
        Me.Label2.Location = New System.Drawing.Point(307, 121)
        Me.Label2.Size = New System.Drawing.Size(494, 228)
        Me.Label2.Text = "                                                          COLLEGE EXAMINATION SYSTEM IS A PROJECT OF COMPUTERISING THE WHOLE PROCESS OF EXAM FROM CREATION OF BRANCHES AND SUBJECTS TO MAKING QUESTION PAPERS TO CONDUCTING EXAMS AND GENERATING RESULTS."
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(255, Byte), CType(255, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(255, Byte))
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Font = New System.Drawing.Font("Lucida Calligraphy", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label4
        '
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 5
        Me.Label4.Location = New System.Drawing.Point(380, 364)
        Me.Label4.Size = New System.Drawing.Size(422, 33)
        Me.Label4.Text = " kckotak99@gmail.com"
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(255, Byte))
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label4.Font = New System.Drawing.Font("Times New Roman", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmAbout
        '
        Me.ClientSize = New System.Drawing.Size(861, 569)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmAbout"
        Me.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(255, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.Icon = CType(Resources.GetObject("frmAbout.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "About"
        Me.Label5.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Label4.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Private Sub cmdStart_Click()
'#Const def_cmdStart_Click = True
#If def_cmdStart_Click
        Close()
#End If	' def_cmdStart_Click
    End Sub

'#Const defUse_Command1_Click = True
#If defUse_Command1_Click
    Private Sub Command1_Click()
'#Const def_Command1_Click = True
#If def_Command1_Click

        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""

        strFilename = "C:\Documents and Settings\PragnaChetanKashyap.KOTAK-B43F5C7CD\Desktop\OfflineExaminer12\Offline CES\kck.txt"

        iFile = FreeFile

        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        Text1.Text = strTheData


#End If	' def_Command1_Click
    End Sub
#End If

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.Width = 12885
        Me.Height = 8955
        Me.Top = 195
        Me.Left = 3765
#End If	' def_Form_Load
    End Sub

End Class