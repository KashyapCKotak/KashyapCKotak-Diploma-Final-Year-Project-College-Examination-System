Imports VB = Microsoft.VisualBasic

Public Class MDIForm1
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
    Friend WithEvents StatusBar3 As System.Windows.Forms.StatusBar
    Friend WithEvents Picture2 As System.Windows.Forms.PictureBox
    Friend WithEvents Command8 As System.Windows.Forms.Button
    Friend WithEvents Command6 As System.Windows.Forms.Button
    Friend WithEvents Command4 As System.Windows.Forms.Button
    Friend WithEvents Command2 As System.Windows.Forms.Button
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents Picture3 As System.Windows.Forms.PictureBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuGeneral As System.Windows.Forms.MenuItem
    Friend WithEvents mnuView As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAdminLogin As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAdmin As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAdminFeedback As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExam As System.Windows.Forms.MenuItem
    Friend WithEvents mnuStudent As System.Windows.Forms.MenuItem
    Friend WithEvents mnuRslt As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MDIForm1))
        Me.StatusBar3 = New System.Windows.Forms.StatusBar
        Me.Picture2 = New System.Windows.Forms.PictureBox
        Me.Command8 = New System.Windows.Forms.Button
        Me.Command6 = New System.Windows.Forms.Button
        Me.Command4 = New System.Windows.Forms.Button
        Me.Command2 = New System.Windows.Forms.Button
        Me.Command1 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.Picture3 = New System.Windows.Forms.PictureBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuGeneral = New System.Windows.Forms.MenuItem
        Me.mnuView = New System.Windows.Forms.MenuItem
        Me.mnuAdminLogin = New System.Windows.Forms.MenuItem
        Me.mnuAdmin = New System.Windows.Forms.MenuItem
        Me.mnuAdminFeedback = New System.Windows.Forms.MenuItem
        Me.mnuExam = New System.Windows.Forms.MenuItem
        Me.mnuStudent = New System.Windows.Forms.MenuItem
        Me.mnuRslt = New System.Windows.Forms.MenuItem
        CType(Me.Picture2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Picture2.SuspendLayout()
        CType(Me.Picture3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusBar3
        '
        Me.StatusBar3.Location = New System.Drawing.Point(0, 623)
        Me.StatusBar3.Name = "StatusBar3"
        Me.StatusBar3.ShowPanels = True
        Me.StatusBar3.Size = New System.Drawing.Size(769, 25)
        Me.StatusBar3.SizingGrip = False
        Me.StatusBar3.TabIndex = 36
        '
        'Picture2
        '
        Me.Picture2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Picture2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Picture2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture2.Controls.Add(Me.Command8)
        Me.Picture2.Controls.Add(Me.Command6)
        Me.Picture2.Controls.Add(Me.Command4)
        Me.Picture2.Controls.Add(Me.Command2)
        Me.Picture2.Controls.Add(Me.Command1)
        Me.Picture2.Controls.Add(Me.Label4)
        Me.Picture2.Controls.Add(Me.Label8)
        Me.Picture2.Controls.Add(Me.Label3)
        Me.Picture2.Location = New System.Drawing.Point(646, 0)
        Me.Picture2.Name = "Picture2"
        Me.Picture2.Size = New System.Drawing.Size(122, 513)
        Me.Picture2.TabIndex = 2
        Me.Picture2.TabStop = False
        '
        'Command8
        '
        Me.Command8.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Command8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command8.Location = New System.Drawing.Point(20, 550)
        Me.Command8.Name = "Command8"
        Me.Command8.Size = New System.Drawing.Size(20, 41)
        Me.Command8.TabIndex = 22
        Me.Command8.UseVisualStyleBackColor = False
        '
        'Command6
        '
        Me.Command6.BackColor = System.Drawing.Color.FromArgb(CType(CType(253, Byte), Integer), CType(CType(79, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.Command6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command6.Image = CType(resources.GetObject("Command6.Image"), System.Drawing.Image)
        Me.Command6.Location = New System.Drawing.Point(0, 227)
        Me.Command6.Name = "Command6"
        Me.Command6.Size = New System.Drawing.Size(122, 74)
        Me.Command6.TabIndex = 20
        Me.Command6.Text = "EXIT"
        Me.Command6.UseVisualStyleBackColor = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.Color.FromArgb(CType(CType(160, Byte), Integer), CType(CType(113, Byte), Integer), CType(CType(219, Byte), Integer))
        Me.Command4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Image = CType(resources.GetObject("Command4.Image"), System.Drawing.Image)
        Me.Command4.Location = New System.Drawing.Point(0, 388)
        Me.Command4.Name = "Command4"
        Me.Command4.Size = New System.Drawing.Size(122, 74)
        Me.Command4.TabIndex = 11
        Me.Command4.Text = "TIME TABLE"
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Command2.Enabled = False
        Me.Command2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Image = CType(resources.GetObject("Command2.Image"), System.Drawing.Image)
        Me.Command2.Location = New System.Drawing.Point(0, 146)
        Me.Command2.Name = "Command2"
        Me.Command2.Size = New System.Drawing.Size(122, 74)
        Me.Command2.TabIndex = 9
        Me.Command2.Text = "LOGOUT"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Command1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Image = CType(resources.GetObject("Command1.Image"), System.Drawing.Image)
        Me.Command1.Location = New System.Drawing.Point(0, 8)
        Me.Command1.Name = "Command1"
        Me.Command1.Size = New System.Drawing.Size(122, 74)
        Me.Command1.TabIndex = 8
        Me.Command1.Text = "LOGIN"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(249, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(130, Byte), Integer))
        Me.Label4.Font = New System.Drawing.Font("Lucida Calligraphy", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(49, 89)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 50)
        Me.Label4.TabIndex = 15
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(0, 275)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(122, 9)
        Me.Label8.TabIndex = 29
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(0, 380)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 9)
        Me.Label3.TabIndex = 14
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 648)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(769, 25)
        Me.StatusBar1.SizingGrip = False
        Me.StatusBar1.TabIndex = 0
        '
        'Picture3
        '
        Me.Picture3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Picture3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Picture3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Picture3.Location = New System.Drawing.Point(0, 513)
        Me.Picture3.Name = "Picture3"
        Me.Picture3.Size = New System.Drawing.Size(769, 111)
        Me.Picture3.TabIndex = 3
        Me.Picture3.TabStop = False
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuGeneral, Me.mnuAdmin, Me.mnuExam, Me.mnuStudent})
        '
        'mnuGeneral
        '
        Me.mnuGeneral.Index = 0
        Me.mnuGeneral.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuView, Me.mnuAdminLogin})
        Me.mnuGeneral.Text = "&General"
        '
        'mnuView
        '
        Me.mnuView.Index = 0
        Me.mnuView.Text = "&View"
        '
        'mnuAdminLogin
        '
        Me.mnuAdminLogin.Index = 1
        Me.mnuAdminLogin.Text = "Log&in"
        '
        'mnuAdmin
        '
        Me.mnuAdmin.Index = 1
        Me.mnuAdmin.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAdminFeedback})
        Me.mnuAdmin.Text = "&Administrator"
        '
        'mnuAdminFeedback
        '
        Me.mnuAdminFeedback.Index = 0
        Me.mnuAdminFeedback.Text = "View &Feedback"
        '
        'mnuExam
        '
        Me.mnuExam.Index = 2
        Me.mnuExam.Text = "&Teacher"
        '
        'mnuStudent
        '
        Me.mnuStudent.Index = 3
        Me.mnuStudent.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuRslt})
        Me.mnuStudent.Text = "&Student"
        '
        'mnuRslt
        '
        Me.mnuRslt.Index = 0
        Me.mnuRslt.Text = "&View Result"
        '
        'MDIForm1
        '
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(769, 673)
        Me.Controls.Add(Me.StatusBar3)
        Me.Controls.Add(Me.Picture2)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.Picture3)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "MDIForm1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "College Examination System"
        CType(Me.Picture2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Picture2.ResumeLayout(False)
        CType(Me.Picture3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Dim num As Short

'#Const defUse_login_Click = True
#If defUse_login_Click
    Private Sub login_Click()
'#Const def_login_Click = True
#If def_login_Click

#End If	' def_login_Click
    End Sub
#End If

    Private Sub Command1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command1.Click
'#Const def_Command1_Click = True
#If def_Command1_Click
        counter = 1
        Me.Picture1.Visible = False
        Me.Picture3.Visible = False
        LoadUnUsed(frmLoginSelect)
        ShowModeless(frmLoginSelect)
        StatusBar1.Panels.Item(1 - 1).Text = "Login or choose to exit"
#End If	' def_Command1_Click
    End Sub

    Private Sub Command10_Click()
'#Const def_Command10_Click = True
#If def_Command10_Click
        Me.BackColor = ColorTranslator.FromOle(&HF486C0)
'color = &HEA1585
#End If	' def_Command10_Click
    End Sub

    Private Sub Command11_Click()
'#Const def_Command11_Click = True
#If def_Command11_Click
        Me.BackColor = ColorTranslator.FromOle(&HFF80FF)
'color = &HFF00FF
#End If	' def_Command11_Click
    End Sub

    Private Sub Command12_Click()
'#Const def_Command12_Click = True
#If def_Command12_Click
        Me.BackColor = ColorTranslator.FromOle(&H8080FF)
'color = &HFF&
#End If	' def_Command12_Click
    End Sub

    Private Sub Command2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command2.Click
'#Const def_Command2_Click = True
#If def_Command2_Click
        Me.Picture1.Visible = True
        Me.Picture3.Visible = True
        mnuExam.Enabled = False
        mnuAdminAdSub.Enabled = False
        mnuAdminchgPass.Enabled = False
        mnuAdminCrtBrnh.Enabled = False
        mnuadminCrtUsr.Enabled = False
        mnuAdminDltUsr.Enabled = False
        mnuAdminLogout.Enabled = False
        mnuStudent.Enabled = False
        Me.mnuAdminLogin.Enabled = True
        Me.Command1.Enabled = True
        Me.Command2.Enabled = False
        Me.Picture4.Visible = False
        Me.Picture5.Visible = False
        Me.Picture6.Visible = False
        Me.mnuAdmin.Enabled = False
        Me.mnuExam.Enabled = False
        Me.mnuStudent.Enabled = False
        LoadUnUsed(frmLoginSelect)
        ShowModeless(frmLoginSelect)
        StatusBar1.Panels.Item(1 - 1).Text = "Status: Logged out successfully. Please login as another user or choose to exit"
        Me.StatusBar1.Panels.Item(2 - 1).Text = ""
        Me.Label4.Text = ""
#End If	' def_Command2_Click
    End Sub

    Private Sub Command3_Click()
'#Const def_Command3_Click = True
#If def_Command3_Click
        Me.Picture1.Visible = False
        Me.Picture3.Visible = False
        xyz = 15
        LoadUnUsed(frmFeedback)
        ShowModeless(frmFeedback)
#End If	' def_Command3_Click
    End Sub

    Private Sub Command4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command4.Click
'#Const def_Command4_Click = True
#If def_Command4_Click
        Me.Picture1.Visible = False
        Me.Picture3.Visible = False
        sh = 5
        LoadUnUsed(frmTT)
        ShowModeless(frmTT)
#End If	' def_Command4_Click
    End Sub

    Private Sub Command5_Click()
'#Const def_Command5_Click = True
#If def_Command5_Click
        If counter=0 Then
            Me.Picture1.Visible = True
            Me.Picture3.Visible = True
            counter = 1
        ElseIf counter=1 Then
            Me.Picture1.Visible = False
            Me.Picture3.Visible = False
            counter = 0
        End If
#End If	' def_Command5_Click
    End Sub

    Private Sub Command6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command6.Click
'#Const def_Command6_Click = True
#If def_Command6_Click
        Me.Close()
        'Load frmSplash2
        'frmSplash2.Show
#End If	' def_Command6_Click
    End Sub

    Private Sub Command7_Click()
'#Const def_Command7_Click = True
#If def_Command7_Click
        Me.BackColor = ColorTranslator.FromOle(&HFF8080)
'color = &HFF0000
#End If	' def_Command7_Click
    End Sub

    Private Sub Command8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command8.Click
'#Const def_Command8_Click = True
#If def_Command8_Click
        Me.BackColor = ColorTranslator.FromOle(&H80FF80)
'color = &HFF00&
#End If	' def_Command8_Click
    End Sub

    Private Sub Command9_Click()
'#Const def_Command9_Click = True
#If def_Command9_Click
        Me.BackColor = ColorTranslator.FromOle(&H80C0FF)
'color = &H80FF&
#End If	' def_Command9_Click
    End Sub

'#Const defUse_MDIForm_Load = True
#If defUse_MDIForm_Load
    Private Sub MDIForm_Load()
'#Const def_MDIForm_Load = True
#If def_MDIForm_Load
        ' VBto upgrade warning: dbloc As Object --> As String
        Dim dbloc As String = ""	' - "AutoDim"


        Dim iFile As Integer
        Dim strFilename As String = ""
        Dim strTheData As String = ""
        strFilename = Application.StartupPath & "\loc.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        dbloc = strTheData
        concat = Len(strTheData)
        dbloc = VB.Left(dbloc, concat-2)

        Me.StatusBar3.Panels.Item(1 - 1).Text = dbloc & " Data Source"

        sh = 1

        strFilename = Application.StartupPath & "\dbsource.txt"
        iFile = FreeFile
        FileOpen(iFile, strFilename, OpenMode.Input)
        strTheData = StrConv(InputB(LOF(iFile), iFile), VbStrConv.None)
        FileClose(iFile)
        str = strTheData
        concat = Len(strTheData)
        str = VB.Left(str, concat-2)



        'Find



        Me.BackColor = ColorTranslator.FromOle(&H80C0FF)
'color = &H80FF&



        counter = 0
        time = 0
        data = ""
        mnuAdmin.Enabled = False
        mnuExam.Enabled = False
        mnuAdminAdSub.Enabled = False
        mnuAdminchgPass.Enabled = False
        mnuAdminCrtBrnh.Enabled = False
        mnuadminCrtUsr.Enabled = False
        mnuAdminDltUsr.Enabled = False
        mnuAdminLogout.Enabled = False
        mnuStudent.Enabled = False
        mnuAdminLogin.Enabled = True
        'Load frmLoginSelect
        'frmLoginSelect.Show
        DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & str & "\OfflineExaminer.mdb;Persist Security Info=False"
        StatusBar1.Panels.Item(1 - 1).Text = "Please Login"

        On Error GoTo localerror
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
         localerror:

#End If	' def_MDIForm_Load
    End Sub
#End If
'#Const defUse_Find = True
#If defUse_Find
    Private Sub Find()
'#Const def_Find = True
#If def_Find
        'Static strData As String * concat
#End If	' def_Find
    End Sub
#End If

'#Const defUse_MDIForm_Terminate = True
#If defUse_MDIForm_Terminate
    Private Sub MDIForm_Terminate()
'#Const def_MDIForm_Terminate = True
#If def_MDIForm_Terminate
        LoadUnUsed(frmSplash2)
        ShowModeless(frmSplash2)
#End If	' def_MDIForm_Terminate
    End Sub
#End If

'#Const defUse_MDIForm_Unload = True
#If defUse_MDIForm_Unload
    Private Sub MDIForm_Unload(ByVal Cancel As Short)
'#Const def_MDIForm_Unload = True
#If def_MDIForm_Unload
        LoadUnUsed(frmSplash2)
        ShowModeless(frmSplash2)
#End If	' def_MDIForm_Unload
    End Sub
#End If

    Private Sub mnuAboutAbout_Click()
'#Const def_mnuAboutAbout_Click = True
#If def_mnuAboutAbout_Click
        LoadUnUsed(frmAbout)
        ShowModeless(frmAbout)
#End If	' def_mnuAboutAbout_Click
    End Sub

    Private Sub mnuAdminAdSub_Click()
'#Const def_mnuAdminAdSub_Click = True
#If def_mnuAdminAdSub_Click
        LoadUnUsed(frmSubjects)
        ShowModeless(frmSubjects)
#End If	' def_mnuAdminAdSub_Click
    End Sub

    Private Sub mnuAdminchgPass_Click()
'#Const def_mnuAdminchgPass_Click = True
#If def_mnuAdminchgPass_Click
        LoadUnUsed(frmChangePass)
        ShowModeless(frmChangePass)
#End If	' def_mnuAdminchgPass_Click
    End Sub

    Private Sub mnuAdminCrtBrnh_Click()
'#Const def_mnuAdminCrtBrnh_Click = True
#If def_mnuAdminCrtBrnh_Click
        LoadUnUsed(frmBranch)
        ShowModeless(frmBranch)
#End If	' def_mnuAdminCrtBrnh_Click
    End Sub

    Private Sub mnuadminCrtUsr_Click()
'#Const def_mnuadminCrtUsr_Click = True
#If def_mnuadminCrtUsr_Click
        LoadUnUsed(frmAddUser)
        ShowModeless(frmAddUser)
#End If	' def_mnuadminCrtUsr_Click
    End Sub

    Private Sub mnuAdminData_Click()
'#Const def_mnuAdminData_Click = True
#If def_mnuAdminData_Click
        LoadUnUsed(frmAskdb)
        ShowModeless(frmAskdb)
#End If	' def_mnuAdminData_Click
    End Sub

    Private Sub mnuAdminDltBranch_Click()
'#Const def_mnuAdminDltBranch_Click = True
#If def_mnuAdminDltBranch_Click
        LoadUnUsed(frmDelBranch)
        ShowModeless(frmDelBranch)
#End If	' def_mnuAdminDltBranch_Click
    End Sub

    Private Sub mnuAdminDltSub_Click()
'#Const def_mnuAdminDltSub_Click = True
#If def_mnuAdminDltSub_Click
        LoadUnUsed(frmDelSub)
        ShowModeless(frmDelSub)
#End If	' def_mnuAdminDltSub_Click
    End Sub

    Private Sub mnuAdminDltUsr_Click()
'#Const def_mnuAdminDltUsr_Click = True
#If def_mnuAdminDltUsr_Click
        LoadUnUsed(frmDeleteUser)
        ShowModeless(frmDeleteUser)
#End If	' def_mnuAdminDltUsr_Click
    End Sub

    Private Sub mnuAdminExit_Click()
'#Const def_mnuAdminExit_Click = True
#If def_mnuAdminExit_Click
        LoadUnUsed(frmSplash2)
        ShowModeless(frmSplash2)
#End If	' def_mnuAdminExit_Click
    End Sub

    Private Sub mnuAdminFeedback_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAdminFeedback.Click
'#Const def_mnuAdminFeedback_Click = True
#If def_mnuAdminFeedback_Click
        xyz = 0
        LoadUnUsed(frmFeedback)
        ShowModeless(frmFeedback)

#End If	' def_mnuAdminFeedback_Click
    End Sub

    Private Sub mnuAdminLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAdminLogin.Click
'#Const def_mnuAdminLogin_Click = True
#If def_mnuAdminLogin_Click
        counter = 1
        Me.Picture1.Visible = False
        Me.Picture3.Visible = False
        LoadUnUsed(frmLoginSelect)
        ShowModeless(frmLoginSelect)
        StatusBar1.Panels.Item(1 - 1).Text = "Login or choose to exit"
#End If	' def_mnuAdminLogin_Click
    End Sub

    Private Sub mnuAdminLogout_Click()
'#Const def_mnuAdminLogout_Click = True
#If def_mnuAdminLogout_Click
        Me.Picture1.Visible = True
        Me.Picture3.Visible = True
        mnuExam.Enabled = False
        mnuAdminAdSub.Enabled = False
        mnuAdminchgPass.Enabled = False
        mnuAdminCrtBrnh.Enabled = False
        mnuadminCrtUsr.Enabled = False
        mnuAdminDltUsr.Enabled = False
        mnuAdminLogout.Enabled = False
        mnuStudent.Enabled = False
        Me.mnuAdminLogin.Enabled = True
        Me.Command1.Enabled = True
        Me.Command2.Enabled = False
        Me.Picture4.Visible = False
        Me.Picture5.Visible = False
        Me.Picture6.Visible = False
        Me.mnuAdmin.Enabled = False
        Me.mnuExam.Enabled = False
        Me.mnuStudent.Enabled = False
        LoadUnUsed(frmLoginSelect)
        ShowModeless(frmLoginSelect)
        StatusBar1.Panels.Item(1 - 1).Text = "Status: Logged out successfully. Please login as another user or choose to exit"
        Me.StatusBar1.Panels.Item(2 - 1).Text = ""
        Me.Label4.Text = ""
#End If	' def_mnuAdminLogout_Click
    End Sub

    Private Sub mnuAdminNews_Click()
'#Const def_mnuAdminNews_Click = True
#If def_mnuAdminNews_Click
        LoadUnUsed(frmNews)
        ShowModeless(frmNews)
#End If	' def_mnuAdminNews_Click
    End Sub

    Private Sub mnuAdminTt_Click()
'#Const def_mnuAdminTt_Click = True
#If def_mnuAdminTt_Click
        sh = 1
        LoadUnUsed(frmTT)
        ShowModeless(frmTT)
#End If	' def_mnuAdminTt_Click
    End Sub

    Private Sub mnuExamAdQst_Click()
'#Const def_mnuExamAdQst_Click = True
#If def_mnuExamAdQst_Click
        LoadUnUsed(frmAddQst)
        ShowModeless(frmAddQst)
#End If	' def_mnuExamAdQst_Click
    End Sub

    Private Sub mnuExamQst_Click()
'#Const def_mnuExamQst_Click = True
#If def_mnuExamQst_Click
        LoadUnUsed(DataReportQst)
        DataReportQst.Show()
#End If	' def_mnuExamQst_Click
    End Sub

    Private Sub mnuExAttEx_Click()
'#Const def_mnuExAttEx_Click = True
#If def_mnuExAttEx_Click
        LoadUnUsed(frmStartExam)
        ShowModeless(frmStartExam)
#End If	' def_mnuExAttEx_Click
    End Sub

    Private Sub mnuExStdReg_Click()
'#Const def_mnuExStdReg_Click = True
#If def_mnuExStdReg_Click
        LoadUnUsed(frmStudReg)
        ShowModeless(frmStudReg)
#End If	' def_mnuExStdReg_Click
    End Sub


    Private Sub mnuExVQBEXID_Click()
'#Const def_mnuExVQBEXID_Click = True
#If def_mnuExVQBEXID_Click
        LoadUnUsed(frmViewRptQst)
        ShowModeless(frmViewRptQst)
#End If	' def_mnuExVQBEXID_Click
    End Sub

    Private Sub mnuRslt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuRslt.Click
'#Const def_mnuRslt_Click = True
#If def_mnuRslt_Click
        LoadUnUsed(frmStudRslt)
        ShowModeless(frmStudRslt)
#End If	' def_mnuRslt_Click
    End Sub

'#Const defUse_mnutchDisable_Click = True
#If defUse_mnutchDisable_Click
    Private Sub mnutchDisable_Click()
'#Const def_mnutchDisable_Click = True
#If def_mnutchDisable_Click
        mnuExAttEx.Enabled = False
        mnutchEnable.Enabled = True
#End If	' def_mnutchDisable_Click
    End Sub
#End If

'#Const defUse_mnutchEnable_Click = True
#If defUse_mnutchEnable_Click
    Private Sub mnutchEnable_Click()
'#Const def_mnutchEnable_Click = True
#If def_mnutchEnable_Click
        mnuExAttEx.Enabled = True
        mnutchDisable.Enabled = True
#End If	' def_mnutchEnable_Click
    End Sub
#End If

    Private Sub mnuViewBl_Click()
'#Const def_mnuViewBl_Click = True
#If def_mnuViewBl_Click
        Me.BackColor = ColorTranslator.FromOle(&HFF8080)
'color = &HFF0000
#End If	' def_mnuViewBl_Click
    End Sub

    Private Sub mnuViewGr_Click()
'#Const def_mnuViewGr_Click = True
#If def_mnuViewGr_Click
        Me.BackColor = ColorTranslator.FromOle(&H80FF80)
'color = &HFF00&
#End If	' def_mnuViewGr_Click
    End Sub

    Private Sub mnuViewOr_Click()
'#Const def_mnuViewOr_Click = True
#If def_mnuViewOr_Click

        Me.BackColor = ColorTranslator.FromOle(&H80C0FF)
'color = &H80FF&
#End If	' def_mnuViewOr_Click
    End Sub

    Private Sub mnuViewPl_Click()
'#Const def_mnuViewPl_Click = True
#If def_mnuViewPl_Click
        Me.BackColor = ColorTranslator.FromOle(&HF486C0)
'color = &HEA1585
#End If	' def_mnuViewPl_Click
    End Sub

    Private Sub mnuViewPnk_Click()
'#Const def_mnuViewPnk_Click = True
#If def_mnuViewPnk_Click
        Me.BackColor = ColorTranslator.FromOle(&HFF80FF)
'color = &HFF00FF
#End If	' def_mnuViewPnk_Click
    End Sub

    Private Sub mnuViewRd_Click()
'#Const def_mnuViewRd_Click = True
#If def_mnuViewRd_Click
        Me.BackColor = ColorTranslator.FromOle(&H8080FF)
'color = &HFF&
#End If	' def_mnuViewRd_Click
    End Sub

    Private Sub mnuViewTt_Click()
'#Const def_mnuViewTt_Click = True
#If def_mnuViewTt_Click
        sh = 5
        LoadUnUsed(frmTT)
        ShowModeless(frmTT)
#End If	' def_mnuViewTt_Click
    End Sub


    Private Sub Timer1_Timer()
'#Const def_Timer1_Timer = True
#If def_Timer1_Timer
        'time = 0
        'For i = 0 To 59
        '    timestr = "10:" & i & "AM"
        'If timestr = MDIForm1.StatusBar1.Panels(4).Text Then
        '    time = 1
        'Else
        '    time = 0
        'End If
        '
        'If time = 0 Then
        '    MDIForm1.mnuExAttEx.Enabled = False
        'Else
        '    MDIForm1.mnuExAttEx.Enabled = True
        'End If
        'Next i


        Text4.Text = (Now).ToString("hh:mm AM/PM")
        'Text4.Text = timestr


        'Text4.Text = time$
        'Text4.Refresh
        time = 0
        For i = 0 To 9
            timestr2 = "10:0" & i & " AM"
            time3 = Text4.Text
            'Text5.Text = time3
            If timestr2=Text4.Text Then
                time = 1
            Else
                time = 0
            End If
            If time=0 Then
                Me.mnuExAttEx.Enabled = False
            Else
                Me.mnuExAttEx.Enabled = True
                goto abcdef
            End If
        Next i

        For i = 10 To 59
            timestr2 = "10:" & i & " AM"
            time3 = Text4.Text
            'Text5.Text = time3
            If timestr2=Text4.Text Then
                time = 1
            Else
                time = 0
            End If
            If time=0 Then
                Me.mnuExAttEx.Enabled = False
            Else
                Me.mnuExAttEx.Enabled = True
                goto abcdef
            End If
        Next i
     abcdef:



#End If	' def_Timer1_Timer
    End Sub

End Class