<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_Find
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
        Me.cmboGender = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtFirstname = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtSurname = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.txtUR = New System.Windows.Forms.TextBox()
        Me.lblUR = New System.Windows.Forms.Label()
        Me.grdPts = New System.Windows.Forms.DataGridView()
        Me.PatientID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Counter = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Surname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Firstname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.UR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DOB = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Gender = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cmboSurnameBy = New System.Windows.Forms.ComboBox()
        Me.cmboFirstnameBy = New System.Windows.Forms.ComboBox()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.txtDOB = New System.Windows.Forms.MaskedTextBox()
        CType(Me.grdPts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmboGender
        '
        Me.cmboGender.FormattingEnabled = True
        Me.cmboGender.Location = New System.Drawing.Point(108, 95)
        Me.cmboGender.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboGender.Name = "cmboGender"
        Me.cmboGender.Size = New System.Drawing.Size(106, 21)
        Me.cmboGender.TabIndex = 23
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(43, 98)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(42, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Gender"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(43, 124)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "DOB"
        '
        'txtFirstname
        '
        Me.txtFirstname.Location = New System.Drawing.Point(108, 70)
        Me.txtFirstname.Margin = New System.Windows.Forms.Padding(2)
        Me.txtFirstname.Name = "txtFirstname"
        Me.txtFirstname.Size = New System.Drawing.Size(106, 20)
        Me.txtFirstname.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(43, 74)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "First name"
        '
        'txtSurname
        '
        Me.txtSurname.Location = New System.Drawing.Point(108, 47)
        Me.txtSurname.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSurname.Name = "txtSurname"
        Me.txtSurname.Size = New System.Drawing.Size(106, 20)
        Me.txtSurname.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label2.Location = New System.Drawing.Point(43, 51)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Surname"
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(375, 110)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(74, 26)
        Me.btnClose.TabIndex = 27
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnClear.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClear.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(375, 81)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(74, 26)
        Me.btnClear.TabIndex = 26
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = False
        '
        'btnFind
        '
        Me.btnFind.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnFind.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnFind.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFind.Location = New System.Drawing.Point(375, 25)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(74, 26)
        Me.btnFind.TabIndex = 25
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = False
        '
        'txtUR
        '
        Me.txtUR.Location = New System.Drawing.Point(108, 25)
        Me.txtUR.Margin = New System.Windows.Forms.Padding(2)
        Me.txtUR.Name = "txtUR"
        Me.txtUR.Size = New System.Drawing.Size(106, 20)
        Me.txtUR.TabIndex = 16
        '
        'lblUR
        '
        Me.lblUR.AutoSize = True
        Me.lblUR.Location = New System.Drawing.Point(43, 29)
        Me.lblUR.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblUR.Name = "lblUR"
        Me.lblUR.Size = New System.Drawing.Size(23, 13)
        Me.lblUR.TabIndex = 15
        Me.lblUR.Text = "UR"
        '
        'grdPts
        '
        Me.grdPts.AllowUserToAddRows = False
        Me.grdPts.AllowUserToDeleteRows = False
        Me.grdPts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdPts.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PatientID, Me.Counter, Me.Surname, Me.Firstname, Me.UR, Me.DOB, Me.Gender})
        Me.grdPts.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grdPts.Location = New System.Drawing.Point(0, 166)
        Me.grdPts.Margin = New System.Windows.Forms.Padding(2)
        Me.grdPts.MultiSelect = False
        Me.grdPts.Name = "grdPts"
        Me.grdPts.ReadOnly = True
        Me.grdPts.RowTemplate.Height = 24
        Me.grdPts.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.grdPts.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdPts.Size = New System.Drawing.Size(627, 214)
        Me.grdPts.TabIndex = 30
        '
        'PatientID
        '
        Me.PatientID.HeaderText = "PatientID"
        Me.PatientID.Name = "PatientID"
        Me.PatientID.ReadOnly = True
        Me.PatientID.Visible = False
        '
        'Counter
        '
        Me.Counter.HeaderText = "#"
        Me.Counter.Name = "Counter"
        Me.Counter.ReadOnly = True
        Me.Counter.Width = 40
        '
        'Surname
        '
        Me.Surname.HeaderText = "Surname"
        Me.Surname.Name = "Surname"
        Me.Surname.ReadOnly = True
        '
        'Firstname
        '
        Me.Firstname.HeaderText = "Firstname"
        Me.Firstname.Name = "Firstname"
        Me.Firstname.ReadOnly = True
        '
        'UR
        '
        Me.UR.HeaderText = "UR"
        Me.UR.Name = "UR"
        Me.UR.ReadOnly = True
        '
        'DOB
        '
        Me.DOB.HeaderText = "DOB"
        Me.DOB.Name = "DOB"
        Me.DOB.ReadOnly = True
        '
        'Gender
        '
        Me.Gender.HeaderText = "Gender"
        Me.Gender.Name = "Gender"
        Me.Gender.ReadOnly = True
        '
        'cmboSurnameBy
        '
        Me.cmboSurnameBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboSurnameBy.FormattingEnabled = True
        Me.cmboSurnameBy.Location = New System.Drawing.Point(221, 46)
        Me.cmboSurnameBy.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboSurnameBy.Name = "cmboSurnameBy"
        Me.cmboSurnameBy.Size = New System.Drawing.Size(80, 21)
        Me.cmboSurnameBy.TabIndex = 31
        '
        'cmboFirstnameBy
        '
        Me.cmboFirstnameBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmboFirstnameBy.FormattingEnabled = True
        Me.cmboFirstnameBy.Location = New System.Drawing.Point(221, 68)
        Me.cmboFirstnameBy.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboFirstnameBy.Name = "cmboFirstnameBy"
        Me.cmboFirstnameBy.Size = New System.Drawing.Size(80, 21)
        Me.cmboFirstnameBy.TabIndex = 32
        '
        'btnSelect
        '
        Me.btnSelect.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSelect.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(375, 53)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(74, 26)
        Me.btnSelect.TabIndex = 33
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = False
        '
        'txtDOB
        '
        Me.txtDOB.Location = New System.Drawing.Point(108, 121)
        Me.txtDOB.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDOB.Mask = "##/##/####"
        Me.txtDOB.Name = "txtDOB"
        Me.txtDOB.Size = New System.Drawing.Size(106, 20)
        Me.txtDOB.TabIndex = 75
        '
        'form_Find
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(627, 380)
        Me.Controls.Add(Me.txtDOB)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.cmboFirstnameBy)
        Me.Controls.Add(Me.cmboSurnameBy)
        Me.Controls.Add(Me.grdPts)
        Me.Controls.Add(Me.cmboGender)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtFirstname)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtSurname)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.txtUR)
        Me.Controls.Add(Me.lblUR)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "form_Find"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Find patient"
        CType(Me.grdPts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmboGender As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtFirstname As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSurname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents txtUR As System.Windows.Forms.TextBox
    Friend WithEvents lblUR As System.Windows.Forms.Label
    Friend WithEvents grdPts As System.Windows.Forms.DataGridView
    Friend WithEvents cmboSurnameBy As System.Windows.Forms.ComboBox
    Friend WithEvents cmboFirstnameBy As System.Windows.Forms.ComboBox
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents txtDOB As System.Windows.Forms.MaskedTextBox
    Friend WithEvents PatientID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Counter As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Surname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Firstname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents UR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Gender As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
