Imports VB = Microsoft.VisualBasic

Public Class frmDelBranch
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
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDelBranch))
        Me.Frame1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.Label3, Me.Label1, Me.Label2})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(89, 81)
        Me.Frame1.Size = New System.Drawing.Size(373, 276)
        Me.Frame1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(192, Byte), CType(128, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'cmdCancel
        '
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Location = New System.Drawing.Point(218, 218)
        Me.cmdCancel.Size = New System.Drawing.Size(98, 25)
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(128, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label3
        '
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 4
        Me.Label3.Location = New System.Drawing.Point(283, 126)
        Me.Label3.Size = New System.Drawing.Size(50, 23)
        Me.Label3.Text = "Refresh"
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(255, Byte), CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label1
        '
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 6
        Me.Label1.Location = New System.Drawing.Point(0, 16)
        Me.Label1.Size = New System.Drawing.Size(373, 41)
        Me.Label1.Text = "DELETE BRANCH"
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label1.Font = New System.Drawing.Font("Times New Roman", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 5
        Me.Label2.Location = New System.Drawing.Point(32, 129)
        Me.Label2.Size = New System.Drawing.Size(98, 17)
        Me.Label2.Text = "User Name"
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'frmDelBranch
        '
        Me.ClientSize = New System.Drawing.Size(573, 430)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Frame1})
        Me.Name = "frmDelBranch"
        Me.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(0, Byte))
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = True
        Me.MaximizeBox = True
        Me.Icon = CType(Resources.GetObject("frmDelBranch.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Delete Branch"
        Me.cmdCancel.ResumeLayout(False)
        Me.Label3.ResumeLayout(False)
        Me.Label1.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
'#Const def_cmdCancel_Click = True
#If def_cmdCancel_Click
        Close()
#End If	' def_cmdCancel_Click
    End Sub

    Private Sub cmdDelete_Click()
'#Const def_cmdDelete_Click = True
#If def_cmdDelete_Click
        con.Execute(("delete from Branch where Branchcode='"+cmbBranchcode.Text+"'"))
        MsgBox("Branch deleted sucessfully!!", MsgBoxStyle.Information, " CES")
        cmbBranchcode.Text = ""
#End If	' def_cmdDelete_Click
    End Sub

    Private Sub frmDelBranch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        Me.BackColor = ColorTranslator.FromOle(color)
        connectdb()
        Me.Width = 8625
        Me.Height = 6885
        rs = con.Execute("select distinct Branchcode from Branch")
        While ( Not rs.EOF)
            cmbBranchcode.Items.Add(rs(0))
            rs.MoveNext()
        End While
#End If	' def_Form_Load
    End Sub

    Private Sub Label3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.Click
'#Const def_Label3_Click = True
#If def_Label3_Click
        ' VBto upgrade warning: hieght As Object --> As Short
        Dim hieght As Short	' - "AutoDim"

        Close()
        LoadUnUsed(Me)
        Me.Width = 8745
        hieght = 6945
        Me.Top = 0
        Me.Left = -60
        ShowModeless(Me)
#End If	' def_Label3_Click
    End Sub

End Class