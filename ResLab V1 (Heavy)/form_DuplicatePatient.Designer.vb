<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_DuplicatePatients
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
        Me.grd = New System.Windows.Forms.DataGridView()
        Me.PatientID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Surname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DOB = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Gender = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Address = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Mobile_ph = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Home_phone = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnContinue = New System.Windows.Forms.Button()
        CType(Me.grd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grd
        '
        Me.grd.AllowUserToAddRows = False
        Me.grd.AllowUserToDeleteRows = False
        Me.grd.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grd.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.PatientID, Me.Surname, Me.DOB, Me.Gender, Me.Address, Me.Mobile_ph, Me.Home_phone})
        Me.grd.Dock = System.Windows.Forms.DockStyle.Top
        Me.grd.Location = New System.Drawing.Point(0, 0)
        Me.grd.MultiSelect = False
        Me.grd.Name = "grd"
        Me.grd.ReadOnly = True
        Me.grd.RowHeadersVisible = False
        Me.grd.RowTemplate.Height = 24
        Me.grd.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.grd.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grd.Size = New System.Drawing.Size(938, 344)
        Me.grd.TabIndex = 0
        '
        'PatientID
        '
        Me.PatientID.HeaderText = "PatientID"
        Me.PatientID.Name = "PatientID"
        Me.PatientID.ReadOnly = True
        Me.PatientID.Visible = False
        '
        'Surname
        '
        Me.Surname.HeaderText = "Name"
        Me.Surname.Name = "Surname"
        Me.Surname.ReadOnly = True
        Me.Surname.Width = 250
        '
        'DOB
        '
        Me.DOB.HeaderText = "DOB"
        Me.DOB.Name = "DOB"
        Me.DOB.ReadOnly = True
        Me.DOB.Width = 80
        '
        'Gender
        '
        Me.Gender.HeaderText = "Gender"
        Me.Gender.Name = "Gender"
        Me.Gender.ReadOnly = True
        Me.Gender.Width = 80
        '
        'Address
        '
        Me.Address.HeaderText = "Address"
        Me.Address.Name = "Address"
        Me.Address.ReadOnly = True
        Me.Address.Width = 300
        '
        'Mobile_ph
        '
        Me.Mobile_ph.HeaderText = "Mobile ph"
        Me.Mobile_ph.Name = "Mobile_ph"
        Me.Mobile_ph.ReadOnly = True
        '
        'Home_phone
        '
        Me.Home_phone.HeaderText = "Home ph"
        Me.Home_phone.Name = "Home_phone"
        Me.Home_phone.ReadOnly = True
        '
        'btnContinue
        '
        Me.btnContinue.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnContinue.Location = New System.Drawing.Point(813, 351)
        Me.btnContinue.Name = "btnContinue"
        Me.btnContinue.Size = New System.Drawing.Size(83, 31)
        Me.btnContinue.TabIndex = 1
        Me.btnContinue.Text = "Continue"
        Me.btnContinue.UseVisualStyleBackColor = True
        '
        'form_DuplicatePatients
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(938, 385)
        Me.Controls.Add(Me.btnContinue)
        Me.Controls.Add(Me.grd)
        Me.Name = "form_DuplicatePatients"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Possible duplicate patients found"
        CType(Me.grd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grd As System.Windows.Forms.DataGridView
    Friend WithEvents btnContinue As System.Windows.Forms.Button
    Friend WithEvents PatientID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Surname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Gender As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Address As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Mobile_ph As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Home_phone As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
