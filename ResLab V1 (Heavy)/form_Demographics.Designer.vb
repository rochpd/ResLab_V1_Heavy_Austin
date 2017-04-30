<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_Demographics
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
        Me.btnSave = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel_Main = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel_source = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.cmboCountryOfBirth = New System.Windows.Forms.ComboBox()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.cmboRaceForRfts = New System.Windows.Forms.ComboBox()
        Me.txtAddress1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDOB = New System.Windows.Forms.MaskedTextBox()
        Me.txtSurname = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtFirstname = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmboGender = New System.Windows.Forms.ComboBox()
        Me.txtPostcode = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.cmboTitle = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtAddress2 = New System.Windows.Forms.TextBox()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtSuburb = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.pnlDemo = New System.Windows.Forms.Panel()
        Me.txtDeathDate = New System.Windows.Forms.MaskedTextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.cmbpPreferredLanguage = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmboAboriginalStatus = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnUR = New System.Windows.Forms.Button()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.lvAllURs = New System.Windows.Forms.ListView()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtDeathStatus = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtMedicareExpires = New System.Windows.Forms.MaskedTextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtMedicareNo = New System.Windows.Forms.MaskedTextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtWorkPh = New System.Windows.Forms.TextBox()
        Me.txtMobilePh = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.txtHomePh = New System.Windows.Forms.TextBox()
        Me.StatusStrip1.SuspendLayout()
        Me.pnlDemo.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnSave.Location = New System.Drawing.Point(635, 388)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(103, 27)
        Me.btnSave.TabIndex = 17
        Me.btnSave.Text = "Save and Close"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel_Main, Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel_source})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 422)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(762, 22)
        Me.StatusStrip1.TabIndex = 73
        Me.StatusStrip1.Tag = ""
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel_Main
        '
        Me.ToolStripStatusLabel_Main.Name = "ToolStripStatusLabel_Main"
        Me.ToolStripStatusLabel_Main.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.AutoSize = False
        Me.ToolStripStatusLabel2.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ToolStripStatusLabel2.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripStatusLabel2.ForeColor = System.Drawing.Color.Red
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(150, 17)
        Me.ToolStripStatusLabel2.Text = "Mandatory fields *"
        Me.ToolStripStatusLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripStatusLabel_source
        '
        Me.ToolStripStatusLabel_source.AutoSize = False
        Me.ToolStripStatusLabel_source.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ToolStripStatusLabel_source.Name = "ToolStripStatusLabel_source"
        Me.ToolStripStatusLabel_source.Size = New System.Drawing.Size(200, 17)
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnClose.Location = New System.Drawing.Point(526, 388)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(103, 27)
        Me.btnClose.TabIndex = 104
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'cmboCountryOfBirth
        '
        Me.cmboCountryOfBirth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboCountryOfBirth.FormattingEnabled = True
        Me.cmboCountryOfBirth.Location = New System.Drawing.Point(476, 154)
        Me.cmboCountryOfBirth.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboCountryOfBirth.Name = "cmboCountryOfBirth"
        Me.cmboCountryOfBirth.Size = New System.Drawing.Size(210, 21)
        Me.cmboCountryOfBirth.Sorted = True
        Me.cmboCountryOfBirth.TabIndex = 136
        '
        'btnSelect
        '
        Me.btnSelect.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSelect.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnSelect.Location = New System.Drawing.Point(419, 388)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(103, 27)
        Me.btnSelect.TabIndex = 127
        Me.btnSelect.Text = "Select Patient"
        Me.btnSelect.UseVisualStyleBackColor = False
        '
        'cmboRaceForRfts
        '
        Me.cmboRaceForRfts.BackColor = System.Drawing.Color.White
        Me.cmboRaceForRfts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboRaceForRfts.FormattingEnabled = True
        Me.cmboRaceForRfts.Location = New System.Drawing.Point(475, 205)
        Me.cmboRaceForRfts.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboRaceForRfts.Name = "cmboRaceForRfts"
        Me.cmboRaceForRfts.Size = New System.Drawing.Size(211, 21)
        Me.cmboRaceForRfts.TabIndex = 128
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(110, 156)
        Me.txtAddress1.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(210, 20)
        Me.txtAddress1.TabIndex = 94
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(18, 52)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 106
        Me.Label2.Text = "Surname"
        '
        'txtDOB
        '
        Me.txtDOB.Location = New System.Drawing.Point(111, 119)
        Me.txtDOB.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDOB.Mask = "##/##/####"
        Me.txtDOB.Name = "txtDOB"
        Me.txtDOB.Size = New System.Drawing.Size(210, 20)
        Me.txtDOB.TabIndex = 93
        '
        'txtSurname
        '
        Me.txtSurname.Location = New System.Drawing.Point(111, 48)
        Me.txtSurname.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSurname.Name = "txtSurname"
        Me.txtSurname.Size = New System.Drawing.Size(210, 20)
        Me.txtSurname.TabIndex = 90
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(18, 75)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 107
        Me.Label3.Text = "First name"
        '
        'txtFirstname
        '
        Me.txtFirstname.Location = New System.Drawing.Point(111, 71)
        Me.txtFirstname.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFirstname.Name = "txtFirstname"
        Me.txtFirstname.Size = New System.Drawing.Size(210, 20)
        Me.txtFirstname.TabIndex = 91
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(18, 123)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 108
        Me.Label4.Text = "DOB"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(375, 159)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(78, 13)
        Me.Label16.TabIndex = 120
        Me.Label16.Text = "Country of birth"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(18, 99)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(42, 13)
        Me.Label5.TabIndex = 109
        Me.Label5.Text = "Gender"
        '
        'cmboGender
        '
        Me.cmboGender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboGender.FormattingEnabled = True
        Me.cmboGender.Location = New System.Drawing.Point(111, 94)
        Me.cmboGender.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboGender.Name = "cmboGender"
        Me.cmboGender.Size = New System.Drawing.Size(210, 21)
        Me.cmboGender.TabIndex = 92
        '
        'txtPostcode
        '
        Me.txtPostcode.Location = New System.Drawing.Point(110, 227)
        Me.txtPostcode.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPostcode.Name = "txtPostcode"
        Me.txtPostcode.Size = New System.Drawing.Size(210, 20)
        Me.txtPostcode.TabIndex = 97
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(17, 160)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(51, 13)
        Me.Label10.TabIndex = 110
        Me.Label10.Text = "Address1"
        '
        'cmboTitle
        '
        Me.cmboTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboTitle.FormattingEnabled = True
        Me.cmboTitle.Location = New System.Drawing.Point(111, 23)
        Me.cmboTitle.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboTitle.Name = "cmboTitle"
        Me.cmboTitle.Size = New System.Drawing.Size(210, 21)
        Me.cmboTitle.TabIndex = 89
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label9.Location = New System.Drawing.Point(17, 184)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(51, 13)
        Me.Label9.TabIndex = 111
        Me.Label9.Text = "Address2"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(18, 29)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(27, 13)
        Me.Label15.TabIndex = 119
        Me.Label15.Text = "Title"
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(110, 180)
        Me.txtAddress2.Margin = New System.Windows.Forms.Padding(2)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(210, 20)
        Me.txtAddress2.TabIndex = 95
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(110, 251)
        Me.txtEmail.Margin = New System.Windows.Forms.Padding(2)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(210, 20)
        Me.txtEmail.TabIndex = 98
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(17, 207)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(41, 13)
        Me.Label8.TabIndex = 112
        Me.Label8.Text = "Suburb"
        '
        'txtSuburb
        '
        Me.txtSuburb.Location = New System.Drawing.Point(110, 203)
        Me.txtSuburb.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSuburb.Name = "txtSuburb"
        Me.txtSuburb.Size = New System.Drawing.Size(210, 20)
        Me.txtSuburb.TabIndex = 96
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(17, 254)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 113
        Me.Label7.Text = "Email"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(17, 230)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(52, 13)
        Me.Label6.TabIndex = 114
        Me.Label6.Text = "Postcode"
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Red
        Me.Label22.Location = New System.Drawing.Point(96, 74)
        Me.Label22.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(17, 24)
        Me.Label22.TabIndex = 126
        Me.Label22.Text = "*"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Red
        Me.Label21.Location = New System.Drawing.Point(96, 121)
        Me.Label21.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(17, 24)
        Me.Label21.TabIndex = 125
        Me.Label21.Text = "*"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Red
        Me.Label20.Location = New System.Drawing.Point(96, 97)
        Me.Label20.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(17, 24)
        Me.Label20.TabIndex = 124
        Me.Label20.Text = "*"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Red
        Me.Label19.Location = New System.Drawing.Point(96, 51)
        Me.Label19.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(17, 24)
        Me.Label19.TabIndex = 123
        Me.Label19.Text = "*"
        '
        'pnlDemo
        '
        Me.pnlDemo.Controls.Add(Me.txtDeathDate)
        Me.pnlDemo.Controls.Add(Me.Label24)
        Me.pnlDemo.Controls.Add(Me.Label13)
        Me.pnlDemo.Controls.Add(Me.cmbpPreferredLanguage)
        Me.pnlDemo.Controls.Add(Me.Label12)
        Me.pnlDemo.Controls.Add(Me.cmboRaceForRfts)
        Me.pnlDemo.Controls.Add(Me.cmboAboriginalStatus)
        Me.pnlDemo.Controls.Add(Me.Label1)
        Me.pnlDemo.Controls.Add(Me.btnUR)
        Me.pnlDemo.Controls.Add(Me.Label26)
        Me.pnlDemo.Controls.Add(Me.lvAllURs)
        Me.pnlDemo.Controls.Add(Me.Label25)
        Me.pnlDemo.Controls.Add(Me.Label23)
        Me.pnlDemo.Controls.Add(Me.cmboCountryOfBirth)
        Me.pnlDemo.Controls.Add(Me.txtAddress1)
        Me.pnlDemo.Controls.Add(Me.Label2)
        Me.pnlDemo.Controls.Add(Me.txtDOB)
        Me.pnlDemo.Controls.Add(Me.txtSurname)
        Me.pnlDemo.Controls.Add(Me.txtDeathStatus)
        Me.pnlDemo.Controls.Add(Me.Label17)
        Me.pnlDemo.Controls.Add(Me.Label3)
        Me.pnlDemo.Controls.Add(Me.txtMedicareExpires)
        Me.pnlDemo.Controls.Add(Me.txtFirstname)
        Me.pnlDemo.Controls.Add(Me.Label18)
        Me.pnlDemo.Controls.Add(Me.Label4)
        Me.pnlDemo.Controls.Add(Me.txtMedicareNo)
        Me.pnlDemo.Controls.Add(Me.Label16)
        Me.pnlDemo.Controls.Add(Me.Label11)
        Me.pnlDemo.Controls.Add(Me.Label5)
        Me.pnlDemo.Controls.Add(Me.txtWorkPh)
        Me.pnlDemo.Controls.Add(Me.cmboGender)
        Me.pnlDemo.Controls.Add(Me.txtPostcode)
        Me.pnlDemo.Controls.Add(Me.txtMobilePh)
        Me.pnlDemo.Controls.Add(Me.Label10)
        Me.pnlDemo.Controls.Add(Me.Label14)
        Me.pnlDemo.Controls.Add(Me.cmboTitle)
        Me.pnlDemo.Controls.Add(Me.Label9)
        Me.pnlDemo.Controls.Add(Me.txtHomePh)
        Me.pnlDemo.Controls.Add(Me.Label15)
        Me.pnlDemo.Controls.Add(Me.txtAddress2)
        Me.pnlDemo.Controls.Add(Me.txtEmail)
        Me.pnlDemo.Controls.Add(Me.Label8)
        Me.pnlDemo.Controls.Add(Me.txtSuburb)
        Me.pnlDemo.Controls.Add(Me.Label7)
        Me.pnlDemo.Controls.Add(Me.Label6)
        Me.pnlDemo.Controls.Add(Me.Label22)
        Me.pnlDemo.Controls.Add(Me.Label21)
        Me.pnlDemo.Controls.Add(Me.Label20)
        Me.pnlDemo.Controls.Add(Me.Label19)
        Me.pnlDemo.Location = New System.Drawing.Point(12, 12)
        Me.pnlDemo.Name = "pnlDemo"
        Me.pnlDemo.Size = New System.Drawing.Size(726, 359)
        Me.pnlDemo.TabIndex = 137
        '
        'txtDeathDate
        '
        Me.txtDeathDate.Location = New System.Drawing.Point(631, 296)
        Me.txtDeathDate.Name = "txtDeathDate"
        Me.txtDeathDate.Size = New System.Drawing.Size(54, 20)
        Me.txtDeathDate.TabIndex = 160
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(586, 299)
        Me.Label24.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(30, 13)
        Me.Label24.TabIndex = 161
        Me.Label24.Text = "Date"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(373, 233)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(97, 13)
        Me.Label13.TabIndex = 159
        Me.Label13.Text = "Preferred language"
        '
        'cmbpPreferredLanguage
        '
        Me.cmbpPreferredLanguage.BackColor = System.Drawing.Color.White
        Me.cmbpPreferredLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbpPreferredLanguage.FormattingEnabled = True
        Me.cmbpPreferredLanguage.Location = New System.Drawing.Point(475, 230)
        Me.cmbpPreferredLanguage.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbpPreferredLanguage.Name = "cmbpPreferredLanguage"
        Me.cmbpPreferredLanguage.Size = New System.Drawing.Size(211, 21)
        Me.cmbpPreferredLanguage.TabIndex = 158
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(374, 208)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(83, 13)
        Me.Label12.TabIndex = 157
        Me.Label12.Text = "Race (for RFTs)"
        '
        'cmboAboriginalStatus
        '
        Me.cmboAboriginalStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboAboriginalStatus.FormattingEnabled = True
        Me.cmboAboriginalStatus.Location = New System.Drawing.Point(475, 179)
        Me.cmboAboriginalStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboAboriginalStatus.Name = "cmboAboriginalStatus"
        Me.cmboAboriginalStatus.Size = New System.Drawing.Size(210, 21)
        Me.cmboAboriginalStatus.Sorted = True
        Me.cmboAboriginalStatus.TabIndex = 156
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(374, 184)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(84, 13)
        Me.Label1.TabIndex = 155
        Me.Label1.Text = "Aboriginal status"
        '
        'btnUR
        '
        Me.btnUR.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnUR.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.goto_anywhere
        Me.btnUR.Location = New System.Drawing.Point(684, 23)
        Me.btnUR.Name = "btnUR"
        Me.btnUR.Size = New System.Drawing.Size(23, 22)
        Me.btnUR.TabIndex = 154
        Me.btnUR.UseVisualStyleBackColor = True
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(217, 278)
        Me.Label26.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(18, 13)
        Me.Label26.TabIndex = 152
        Me.Label26.Text = "W"
        '
        'lvAllURs
        '
        Me.lvAllURs.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvAllURs.Location = New System.Drawing.Point(378, 23)
        Me.lvAllURs.Name = "lvAllURs"
        Me.lvAllURs.Size = New System.Drawing.Size(300, 83)
        Me.lvAllURs.TabIndex = 150
        Me.lvAllURs.UseCompatibleStateImageBehavior = False
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(111, 302)
        Me.Label25.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(16, 13)
        Me.Label25.TabIndex = 151
        Me.Label25.Text = "M"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(111, 278)
        Me.Label23.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(15, 13)
        Me.Label23.TabIndex = 150
        Me.Label23.Text = "H"
        '
        'txtDeathStatus
        '
        Me.txtDeathStatus.Location = New System.Drawing.Point(475, 296)
        Me.txtDeathStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDeathStatus.Name = "txtDeathStatus"
        Me.txtDeathStatus.Size = New System.Drawing.Size(103, 20)
        Me.txtDeathStatus.TabIndex = 148
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(374, 299)
        Me.Label17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(67, 13)
        Me.Label17.TabIndex = 149
        Me.Label17.Text = "Death status"
        '
        'txtMedicareExpires
        '
        Me.txtMedicareExpires.Location = New System.Drawing.Point(631, 273)
        Me.txtMedicareExpires.Name = "txtMedicareExpires"
        Me.txtMedicareExpires.Size = New System.Drawing.Size(54, 20)
        Me.txtMedicareExpires.TabIndex = 146
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(586, 276)
        Me.Label18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(40, 13)
        Me.Label18.TabIndex = 147
        Me.Label18.Text = "expires"
        '
        'txtMedicareNo
        '
        Me.txtMedicareNo.Location = New System.Drawing.Point(474, 273)
        Me.txtMedicareNo.Mask = "0000-00000-0-0"
        Me.txtMedicareNo.Name = "txtMedicareNo"
        Me.txtMedicareNo.Size = New System.Drawing.Size(104, 20)
        Me.txtMedicareNo.TabIndex = 141
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(374, 275)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(69, 13)
        Me.Label11.TabIndex = 145
        Me.Label11.Text = "Medicare no."
        '
        'txtWorkPh
        '
        Me.txtWorkPh.Location = New System.Drawing.Point(127, 299)
        Me.txtWorkPh.Margin = New System.Windows.Forms.Padding(2)
        Me.txtWorkPh.Name = "txtWorkPh"
        Me.txtWorkPh.Size = New System.Drawing.Size(86, 20)
        Me.txtWorkPh.TabIndex = 140
        '
        'txtMobilePh
        '
        Me.txtMobilePh.Location = New System.Drawing.Point(234, 275)
        Me.txtMobilePh.Margin = New System.Windows.Forms.Padding(2)
        Me.txtMobilePh.Name = "txtMobilePh"
        Me.txtMobilePh.Size = New System.Drawing.Size(86, 20)
        Me.txtMobilePh.TabIndex = 139
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(17, 277)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(81, 13)
        Me.Label14.TabIndex = 142
        Me.Label14.Text = "Phone numbers"
        '
        'txtHomePh
        '
        Me.txtHomePh.Location = New System.Drawing.Point(127, 275)
        Me.txtHomePh.Margin = New System.Windows.Forms.Padding(2)
        Me.txtHomePh.Name = "txtHomePh"
        Me.txtHomePh.Size = New System.Drawing.Size(86, 20)
        Me.txtHomePh.TabIndex = 138
        '
        'form_Demographics
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(762, 444)
        Me.Controls.Add(Me.pnlDemo)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnClose)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "form_Demographics"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Demographic data for "
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.pnlDemo.ResumeLayout(False)
        Me.pnlDemo.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel_source As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel_Main As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cmboCountryOfBirth As System.Windows.Forms.ComboBox
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents cmboRaceForRfts As System.Windows.Forms.ComboBox
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDOB As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtSurname As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtFirstname As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
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
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents pnlDemo As System.Windows.Forms.Panel
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtDeathStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtMedicareExpires As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtMedicareNo As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtWorkPh As System.Windows.Forms.TextBox
    Friend WithEvents txtMobilePh As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtHomePh As System.Windows.Forms.TextBox
    Friend WithEvents btnUR As System.Windows.Forms.Button
    Friend WithEvents lvAllURs As System.Windows.Forms.ListView
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cmboAboriginalStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cmbpPreferredLanguage As System.Windows.Forms.ComboBox
    Friend WithEvents txtDeathDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
End Class
