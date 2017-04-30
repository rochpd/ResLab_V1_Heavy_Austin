<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_user
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.pnlData = New System.Windows.Forms.Panel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage_details = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtUsername = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.btnResetPassword = New System.Windows.Forms.Button()
        Me.txtUserPassword = New System.Windows.Forms.Label()
        Me.txtMobilePh = New System.Windows.Forms.TextBox()
        Me.txtDOB = New System.Windows.Forms.MaskedTextBox()
        Me.txtSurname = New System.Windows.Forms.TextBox()
        Me.txtFirstname = New System.Windows.Forms.TextBox()
        Me.cmboProfession = New System.Windows.Forms.ComboBox()
        Me.cmboGender = New System.Windows.Forms.ComboBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.txtHomePh = New System.Windows.Forms.TextBox()
        Me.cmboState = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmboDepartment = New System.Windows.Forms.ComboBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmboInstitution = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtAddress1 = New System.Windows.Forms.TextBox()
        Me.txtWorkPh = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSuburb = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtAddress2 = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmboTitle = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtPostcode = New System.Windows.Forms.TextBox()
        Me.TabPage_permissions = New System.Windows.Forms.TabPage()
        Me.grdP_comb = New System.Windows.Forms.DataGridView()
        Me.grdP_chk = New System.Windows.Forms.DataGridView()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel_Main = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsLabel_lastlogin = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.cmboSearch_username = New System.Windows.Forms.ComboBox()
        Me.pnlData.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage_details.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage_permissions.SuspendLayout()
        CType(Me.grdP_comb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdP_chk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlData
        '
        Me.pnlData.BackColor = System.Drawing.SystemColors.Control
        Me.pnlData.Controls.Add(Me.TabControl1)
        Me.pnlData.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlData.Location = New System.Drawing.Point(0, 0)
        Me.pnlData.Name = "pnlData"
        Me.pnlData.Size = New System.Drawing.Size(718, 351)
        Me.pnlData.TabIndex = 79
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_details)
        Me.TabControl1.Controls.Add(Me.TabPage_permissions)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(718, 351)
        Me.TabControl1.TabIndex = 109
        '
        'TabPage_details
        '
        Me.TabPage_details.Controls.Add(Me.GroupBox1)
        Me.TabPage_details.Controls.Add(Me.txtMobilePh)
        Me.TabPage_details.Controls.Add(Me.txtDOB)
        Me.TabPage_details.Controls.Add(Me.txtSurname)
        Me.TabPage_details.Controls.Add(Me.txtFirstname)
        Me.TabPage_details.Controls.Add(Me.cmboProfession)
        Me.TabPage_details.Controls.Add(Me.cmboGender)
        Me.TabPage_details.Controls.Add(Me.Label19)
        Me.TabPage_details.Controls.Add(Me.Label20)
        Me.TabPage_details.Controls.Add(Me.Label21)
        Me.TabPage_details.Controls.Add(Me.Label22)
        Me.TabPage_details.Controls.Add(Me.txtHomePh)
        Me.TabPage_details.Controls.Add(Me.cmboState)
        Me.TabPage_details.Controls.Add(Me.Label13)
        Me.TabPage_details.Controls.Add(Me.Label11)
        Me.TabPage_details.Controls.Add(Me.Label14)
        Me.TabPage_details.Controls.Add(Me.cmboDepartment)
        Me.TabPage_details.Controls.Add(Me.Label23)
        Me.TabPage_details.Controls.Add(Me.Label6)
        Me.TabPage_details.Controls.Add(Me.cmboInstitution)
        Me.TabPage_details.Controls.Add(Me.Label12)
        Me.TabPage_details.Controls.Add(Me.Label1)
        Me.TabPage_details.Controls.Add(Me.Label7)
        Me.TabPage_details.Controls.Add(Me.txtAddress1)
        Me.TabPage_details.Controls.Add(Me.txtWorkPh)
        Me.TabPage_details.Controls.Add(Me.Label2)
        Me.TabPage_details.Controls.Add(Me.txtSuburb)
        Me.TabPage_details.Controls.Add(Me.Label8)
        Me.TabPage_details.Controls.Add(Me.txtEmail)
        Me.TabPage_details.Controls.Add(Me.Label3)
        Me.TabPage_details.Controls.Add(Me.txtAddress2)
        Me.TabPage_details.Controls.Add(Me.Label15)
        Me.TabPage_details.Controls.Add(Me.Label9)
        Me.TabPage_details.Controls.Add(Me.Label4)
        Me.TabPage_details.Controls.Add(Me.cmboTitle)
        Me.TabPage_details.Controls.Add(Me.Label16)
        Me.TabPage_details.Controls.Add(Me.Label10)
        Me.TabPage_details.Controls.Add(Me.Label5)
        Me.TabPage_details.Controls.Add(Me.txtPostcode)
        Me.TabPage_details.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.TabPage_details.Location = New System.Drawing.Point(4, 23)
        Me.TabPage_details.Name = "TabPage_details"
        Me.TabPage_details.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_details.Size = New System.Drawing.Size(710, 324)
        Me.TabPage_details.TabIndex = 0
        Me.TabPage_details.Text = "Details"
        Me.TabPage_details.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label24)
        Me.GroupBox1.Controls.Add(Me.txtUsername)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.btnResetPassword)
        Me.GroupBox1.Controls.Add(Me.txtUserPassword)
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.GroupBox1.Location = New System.Drawing.Point(8, 182)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(268, 117)
        Me.GroupBox1.TabIndex = 109
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Login details"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label24.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label24.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label24.Location = New System.Drawing.Point(16, 33)
        Me.Label24.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(61, 13)
        Me.Label24.TabIndex = 89
        Me.Label24.Text = "User name"
        '
        'txtUsername
        '
        Me.txtUsername.ForeColor = System.Drawing.SystemColors.Desktop
        Me.txtUsername.Location = New System.Drawing.Point(105, 29)
        Me.txtUsername.Margin = New System.Windows.Forms.Padding(2)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(147, 22)
        Me.txtUsername.TabIndex = 7
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label18.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label18.Location = New System.Drawing.Point(16, 56)
        Me.Label18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(83, 13)
        Me.Label18.TabIndex = 90
        Me.Label18.Text = "User password"
        '
        'btnResetPassword
        '
        Me.btnResetPassword.AutoSize = True
        Me.btnResetPassword.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnResetPassword.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnResetPassword.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnResetPassword.Location = New System.Drawing.Point(140, 88)
        Me.btnResetPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.btnResetPassword.Name = "btnResetPassword"
        Me.btnResetPassword.Size = New System.Drawing.Size(112, 24)
        Me.btnResetPassword.TabIndex = 105
        Me.btnResetPassword.Text = "Reset Password"
        Me.btnResetPassword.UseVisualStyleBackColor = False
        '
        'txtUserPassword
        '
        Me.txtUserPassword.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.txtUserPassword.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUserPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserPassword.ForeColor = System.Drawing.SystemColors.Desktop
        Me.txtUserPassword.Location = New System.Drawing.Point(105, 53)
        Me.txtUserPassword.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.txtUserPassword.Name = "txtUserPassword"
        Me.txtUserPassword.Size = New System.Drawing.Size(147, 19)
        Me.txtUserPassword.TabIndex = 91
        '
        'txtMobilePh
        '
        Me.txtMobilePh.Location = New System.Drawing.Point(389, 195)
        Me.txtMobilePh.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMobilePh.Name = "txtMobilePh"
        Me.txtMobilePh.Size = New System.Drawing.Size(245, 22)
        Me.txtMobilePh.TabIndex = 18
        '
        'txtDOB
        '
        Me.txtDOB.Location = New System.Drawing.Point(116, 105)
        Me.txtDOB.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDOB.Mask = "##/##/####"
        Me.txtDOB.Name = "txtDOB"
        Me.txtDOB.Size = New System.Drawing.Size(147, 22)
        Me.txtDOB.TabIndex = 5
        '
        'txtSurname
        '
        Me.txtSurname.Location = New System.Drawing.Point(116, 37)
        Me.txtSurname.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSurname.Name = "txtSurname"
        Me.txtSurname.Size = New System.Drawing.Size(147, 22)
        Me.txtSurname.TabIndex = 2
        '
        'txtFirstname
        '
        Me.txtFirstname.Location = New System.Drawing.Point(116, 60)
        Me.txtFirstname.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFirstname.Name = "txtFirstname"
        Me.txtFirstname.Size = New System.Drawing.Size(147, 22)
        Me.txtFirstname.TabIndex = 3
        '
        'cmboProfession
        '
        Me.cmboProfession.FormattingEnabled = True
        Me.cmboProfession.Location = New System.Drawing.Point(389, 14)
        Me.cmboProfession.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboProfession.Name = "cmboProfession"
        Me.cmboProfession.Size = New System.Drawing.Size(245, 21)
        Me.cmboProfession.TabIndex = 6
        '
        'cmboGender
        '
        Me.cmboGender.FormattingEnabled = True
        Me.cmboGender.Location = New System.Drawing.Point(116, 83)
        Me.cmboGender.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboGender.Name = "cmboGender"
        Me.cmboGender.Size = New System.Drawing.Size(147, 21)
        Me.cmboGender.TabIndex = 4
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Red
        Me.Label19.Location = New System.Drawing.Point(93, 43)
        Me.Label19.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(17, 24)
        Me.Label19.TabIndex = 75
        Me.Label19.Text = "*"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Red
        Me.Label20.Location = New System.Drawing.Point(93, 88)
        Me.Label20.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(17, 24)
        Me.Label20.TabIndex = 76
        Me.Label20.Text = "*"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Red
        Me.Label21.Location = New System.Drawing.Point(93, 112)
        Me.Label21.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(17, 24)
        Me.Label21.TabIndex = 77
        Me.Label21.Text = "*"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Red
        Me.Label22.Location = New System.Drawing.Point(93, 66)
        Me.Label22.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(17, 24)
        Me.Label22.TabIndex = 78
        Me.Label22.Text = "*"
        '
        'txtHomePh
        '
        Me.txtHomePh.Location = New System.Drawing.Point(389, 172)
        Me.txtHomePh.Margin = New System.Windows.Forms.Padding(2)
        Me.txtHomePh.Name = "txtHomePh"
        Me.txtHomePh.Size = New System.Drawing.Size(245, 22)
        Me.txtHomePh.TabIndex = 17
        '
        'cmboState
        '
        Me.cmboState.FormattingEnabled = True
        Me.cmboState.Location = New System.Drawing.Point(551, 150)
        Me.cmboState.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboState.Name = "cmboState"
        Me.cmboState.Size = New System.Drawing.Size(83, 21)
        Me.cmboState.TabIndex = 16
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label13.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label13.Location = New System.Drawing.Point(304, 203)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(78, 13)
        Me.Label13.TabIndex = 59
        Me.Label13.Text = "Phone mobile"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label11.Location = New System.Drawing.Point(515, 155)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(33, 13)
        Me.Label11.TabIndex = 84
        Me.Label11.Text = "State"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label14.Location = New System.Drawing.Point(304, 180)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 13)
        Me.Label14.TabIndex = 57
        Me.Label14.Text = "Phone home"
        '
        'cmboDepartment
        '
        Me.cmboDepartment.FormattingEnabled = True
        Me.cmboDepartment.Location = New System.Drawing.Point(389, 36)
        Me.cmboDepartment.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboDepartment.Name = "cmboDepartment"
        Me.cmboDepartment.Size = New System.Drawing.Size(245, 21)
        Me.cmboDepartment.TabIndex = 10
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label23.Location = New System.Drawing.Point(304, 42)
        Me.Label23.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(68, 13)
        Me.Label23.TabIndex = 82
        Me.Label23.Text = "Department"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label6.Location = New System.Drawing.Point(304, 157)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(54, 13)
        Me.Label6.TabIndex = 52
        Me.Label6.Text = "Postcode"
        '
        'cmboInstitution
        '
        Me.cmboInstitution.FormattingEnabled = True
        Me.cmboInstitution.Location = New System.Drawing.Point(389, 58)
        Me.cmboInstitution.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboInstitution.Name = "cmboInstitution"
        Me.cmboInstitution.Size = New System.Drawing.Size(245, 21)
        Me.cmboInstitution.TabIndex = 11
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label12.Location = New System.Drawing.Point(304, 226)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(69, 13)
        Me.Label12.TabIndex = 61
        Me.Label12.Text = "Phone work"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label1.Location = New System.Drawing.Point(304, 65)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 80
        Me.Label1.Text = "Institution"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label7.Location = New System.Drawing.Point(304, 249)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(34, 13)
        Me.Label7.TabIndex = 51
        Me.Label7.Text = "Email"
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(389, 80)
        Me.txtAddress1.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(245, 22)
        Me.txtAddress1.TabIndex = 12
        '
        'txtWorkPh
        '
        Me.txtWorkPh.Location = New System.Drawing.Point(389, 218)
        Me.txtWorkPh.Margin = New System.Windows.Forms.Padding(2)
        Me.txtWorkPh.Name = "txtWorkPh"
        Me.txtWorkPh.Size = New System.Drawing.Size(245, 22)
        Me.txtWorkPh.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(27, 44)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "Surname"
        '
        'txtSuburb
        '
        Me.txtSuburb.Location = New System.Drawing.Point(389, 126)
        Me.txtSuburb.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSuburb.Name = "txtSuburb"
        Me.txtSuburb.Size = New System.Drawing.Size(245, 22)
        Me.txtSuburb.TabIndex = 14
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label8.Location = New System.Drawing.Point(304, 134)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 13)
        Me.Label8.TabIndex = 49
        Me.Label8.Text = "Suburb"
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(389, 241)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(245, 22)
        Me.txtEmail.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(27, 67)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "First name"
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(389, 103)
        Me.txtAddress2.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(245, 22)
        Me.txtAddress2.TabIndex = 13
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label15.Location = New System.Drawing.Point(27, 21)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(28, 13)
        Me.Label15.TabIndex = 66
        Me.Label15.Text = "Title"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label9.Location = New System.Drawing.Point(304, 111)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(54, 13)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "Address2"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(27, 113)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(31, 13)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "DOB"
        '
        'cmboTitle
        '
        Me.cmboTitle.FormattingEnabled = True
        Me.cmboTitle.Location = New System.Drawing.Point(116, 15)
        Me.cmboTitle.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboTitle.Name = "cmboTitle"
        Me.cmboTitle.Size = New System.Drawing.Size(147, 21)
        Me.cmboTitle.TabIndex = 1
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label16.Location = New System.Drawing.Point(304, 20)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(61, 13)
        Me.Label16.TabIndex = 70
        Me.Label16.Text = "Profession"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label10.Location = New System.Drawing.Point(304, 88)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(54, 13)
        Me.Label10.TabIndex = 45
        Me.Label10.Text = "Address1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 8.150944!)
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(27, 90)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "Gender"
        '
        'txtPostcode
        '
        Me.txtPostcode.Location = New System.Drawing.Point(389, 149)
        Me.txtPostcode.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPostcode.Name = "txtPostcode"
        Me.txtPostcode.Size = New System.Drawing.Size(83, 22)
        Me.txtPostcode.TabIndex = 15
        '
        'TabPage_permissions
        '
        Me.TabPage_permissions.Controls.Add(Me.grdP_comb)
        Me.TabPage_permissions.Controls.Add(Me.grdP_chk)
        Me.TabPage_permissions.Location = New System.Drawing.Point(4, 23)
        Me.TabPage_permissions.Name = "TabPage_permissions"
        Me.TabPage_permissions.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_permissions.Size = New System.Drawing.Size(710, 324)
        Me.TabPage_permissions.TabIndex = 1
        Me.TabPage_permissions.Text = "Permissions"
        Me.TabPage_permissions.UseVisualStyleBackColor = True
        '
        'grdP_comb
        '
        Me.grdP_comb.AllowUserToAddRows = False
        Me.grdP_comb.AllowUserToDeleteRows = False
        Me.grdP_comb.AllowUserToResizeRows = False
        Me.grdP_comb.BackgroundColor = System.Drawing.Color.White
        Me.grdP_comb.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdP_comb.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.grdP_comb.EnableHeadersVisualStyles = False
        Me.grdP_comb.Location = New System.Drawing.Point(356, 32)
        Me.grdP_comb.MultiSelect = False
        Me.grdP_comb.Name = "grdP_comb"
        Me.grdP_comb.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.grdP_comb.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdP_comb.Size = New System.Drawing.Size(299, 286)
        Me.grdP_comb.TabIndex = 1
        '
        'grdP_chk
        '
        Me.grdP_chk.AllowUserToAddRows = False
        Me.grdP_chk.AllowUserToDeleteRows = False
        Me.grdP_chk.AllowUserToResizeRows = False
        Me.grdP_chk.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.grdP_chk.BackgroundColor = System.Drawing.Color.White
        Me.grdP_chk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdP_chk.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.grdP_chk.EnableHeadersVisualStyles = False
        Me.grdP_chk.Location = New System.Drawing.Point(23, 32)
        Me.grdP_chk.MultiSelect = False
        Me.grdP_chk.Name = "grdP_chk"
        Me.grdP_chk.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.grdP_chk.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdP_chk.ShowEditingIcon = False
        Me.grdP_chk.Size = New System.Drawing.Size(311, 286)
        Me.grdP_chk.TabIndex = 0
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel_Main, Me.ToolStripStatusLabel2, Me.tsLabel_lastlogin})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 479)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(718, 22)
        Me.StatusStrip1.TabIndex = 81
        Me.StatusStrip1.Tag = ""
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel_Main
        '
        Me.ToolStripStatusLabel_Main.Name = "ToolStripStatusLabel_Main"
        Me.ToolStripStatusLabel_Main.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.ForeColor = System.Drawing.Color.Red
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(94, 17)
        Me.ToolStripStatusLabel2.Text = "Mandatory fields"
        '
        'tsLabel_lastlogin
        '
        Me.tsLabel_lastlogin.AutoSize = False
        Me.tsLabel_lastlogin.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsLabel_lastlogin.Name = "tsLabel_lastlogin"
        Me.tsLabel_lastlogin.Size = New System.Drawing.Size(200, 17)
        Me.tsLabel_lastlogin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnFind
        '
        Me.btnFind.AutoSize = True
        Me.btnFind.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnFind.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnFind.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFind.Location = New System.Drawing.Point(185, 385)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(36, 24)
        Me.btnFind.TabIndex = 104
        Me.btnFind.Text = "Go"
        Me.btnFind.UseVisualStyleBackColor = False
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label26.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.Label26.Location = New System.Drawing.Point(31, 371)
        Me.Label26.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(127, 14)
        Me.Label26.TabIndex = 103
        Me.Label26.Text = "Search for user name"
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnDelete.Location = New System.Drawing.Point(477, 433)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(2)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(71, 27)
        Me.btnDelete.TabIndex = 26
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = False
        Me.btnDelete.Visible = False
        '
        'btnNew
        '
        Me.btnNew.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnNew.Location = New System.Drawing.Point(402, 371)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(2)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(71, 27)
        Me.btnNew.TabIndex = 22
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnClose.Location = New System.Drawing.Point(636, 433)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(71, 27)
        Me.btnClose.TabIndex = 27
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'btnEdit
        '
        Me.btnEdit.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnEdit.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnEdit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnEdit.Location = New System.Drawing.Point(402, 402)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(71, 27)
        Me.btnEdit.TabIndex = 23
        Me.btnEdit.Text = "Edit"
        Me.btnEdit.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnCancel.Location = New System.Drawing.Point(477, 402)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(71, 27)
        Me.btnCancel.TabIndex = 25
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnSave.Location = New System.Drawing.Point(477, 371)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(71, 27)
        Me.btnSave.TabIndex = 24
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'cmboSearch_username
        '
        Me.cmboSearch_username.FormattingEnabled = True
        Me.cmboSearch_username.Location = New System.Drawing.Point(30, 388)
        Me.cmboSearch_username.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboSearch_username.Name = "cmboSearch_username"
        Me.cmboSearch_username.Size = New System.Drawing.Size(147, 21)
        Me.cmboSearch_username.Sorted = True
        Me.cmboSearch_username.TabIndex = 21
        '
        'form_user
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Info
        Me.ClientSize = New System.Drawing.Size(718, 501)
        Me.Controls.Add(Me.cmboSearch_username)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.pnlData)
        Me.Name = "form_user"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "User details"
        Me.pnlData.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage_details.ResumeLayout(False)
        Me.TabPage_details.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabPage_permissions.ResumeLayout(False)
        CType(Me.grdP_comb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdP_chk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pnlData As System.Windows.Forms.Panel
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDOB As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtSurname As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtFirstname As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmboGender As System.Windows.Forms.ComboBox
    Friend WithEvents txtPostcode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmboTitle As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtSuburb As System.Windows.Forms.TextBox
    Friend WithEvents txtWorkPh As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtMobilePh As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtHomePh As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents cmboState As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmboDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents cmboInstitution As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmboProfession As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel_Main As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents tsLabel_lastlogin As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtUserPassword As System.Windows.Forms.Label
    Friend WithEvents btnResetPassword As System.Windows.Forms.Button
    Friend WithEvents cmboSearch_username As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_details As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_permissions As System.Windows.Forms.TabPage
    Friend WithEvents grdP_chk As System.Windows.Forms.DataGridView
    Friend WithEvents grdP_comb As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
