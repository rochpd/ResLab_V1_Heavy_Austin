<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_PredSourceEditor
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ToolStrip5 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel4 = New System.Windows.Forms.ToolStripLabel()
        Me.grdEqs = New System.Windows.Forms.DataGridView()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.pnlSource = New System.Windows.Forms.Panel()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.lstParams = New System.Windows.Forms.CheckedListBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.cmboTest = New System.Windows.Forms.ComboBox()
        Me.cmboEquationClass = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtSource = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtReference = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.pnlEquationDetails = New System.Windows.Forms.Panel()
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel()
        Me.cmboParameter = New System.Windows.Forms.ComboBox()
        Me.cmboWt_clipmethod = New System.Windows.Forms.ComboBox()
        Me.txtWtUpper = New System.Windows.Forms.TextBox()
        Me.txtWtLower = New System.Windows.Forms.TextBox()
        Me.cmboHt_clipmethod = New System.Windows.Forms.ComboBox()
        Me.txtHtUpper = New System.Windows.Forms.TextBox()
        Me.txtHtLower = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cmboAge_clipmethod = New System.Windows.Forms.ComboBox()
        Me.chkApplyNonCaucasianCorrection = New System.Windows.Forms.CheckBox()
        Me.cmboAgeGroup = New System.Windows.Forms.ComboBox()
        Me.cmboStat = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtAgeUpper = New System.Windows.Forms.TextBox()
        Me.cmboGender = New System.Windows.Forms.ComboBox()
        Me.cmboEthnicity = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtEquation = New System.Windows.Forms.TextBox()
        Me.txtAgeLower = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btnCheckEquation = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.ToolStrip5.SuspendLayout()
        CType(Me.grdEqs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlSource.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.pnlEquationDetails.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ToolStrip5)
        Me.Panel1.Controls.Add(Me.grdEqs)
        Me.Panel1.Location = New System.Drawing.Point(9, 198)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(575, 378)
        Me.Panel1.TabIndex = 190
        '
        'ToolStrip5
        '
        Me.ToolStrip5.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ToolStrip5.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel4})
        Me.ToolStrip5.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip5.Name = "ToolStrip5"
        Me.ToolStrip5.Size = New System.Drawing.Size(573, 25)
        Me.ToolStrip5.TabIndex = 41
        '
        'ToolStripLabel4
        '
        Me.ToolStripLabel4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel4.Name = "ToolStripLabel4"
        Me.ToolStripLabel4.Size = New System.Drawing.Size(59, 22)
        Me.ToolStripLabel4.Text = "Equations"
        '
        'grdEqs
        '
        Me.grdEqs.AllowUserToAddRows = False
        Me.grdEqs.AllowUserToDeleteRows = False
        Me.grdEqs.AllowUserToResizeRows = False
        Me.grdEqs.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdEqs.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.grdEqs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.grdEqs.DefaultCellStyle = DataGridViewCellStyle6
        Me.grdEqs.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.grdEqs.Location = New System.Drawing.Point(0, 27)
        Me.grdEqs.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.grdEqs.Name = "grdEqs"
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdEqs.RowHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.grdEqs.RowHeadersVisible = False
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grdEqs.RowsDefaultCellStyle = DataGridViewCellStyle8
        Me.grdEqs.RowTemplate.Height = 24
        Me.grdEqs.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.grdEqs.Size = New System.Drawing.Size(573, 349)
        Me.grdEqs.TabIndex = 40
        '
        'btnNew
        '
        Me.btnNew.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnNew.Enabled = False
        Me.btnNew.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.Location = New System.Drawing.Point(918, 74)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 25)
        Me.btnNew.TabIndex = 189
        Me.btnNew.Text = "New equation"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'pnlSource
        '
        Me.pnlSource.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSource.Controls.Add(Me.ToolStrip1)
        Me.pnlSource.Controls.Add(Me.lstParams)
        Me.pnlSource.Controls.Add(Me.Label20)
        Me.pnlSource.Controls.Add(Me.cmboTest)
        Me.pnlSource.Controls.Add(Me.cmboEquationClass)
        Me.pnlSource.Controls.Add(Me.Label8)
        Me.pnlSource.Controls.Add(Me.txtSource)
        Me.pnlSource.Controls.Add(Me.Label6)
        Me.pnlSource.Controls.Add(Me.txtReference)
        Me.pnlSource.Controls.Add(Me.Label2)
        Me.pnlSource.Controls.Add(Me.Label1)
        Me.pnlSource.Location = New System.Drawing.Point(9, 10)
        Me.pnlSource.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.pnlSource.Name = "pnlSource"
        Me.pnlSource.Size = New System.Drawing.Size(845, 167)
        Me.pnlSource.TabIndex = 188
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(843, 25)
        Me.ToolStrip1.TabIndex = 185
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(81, 22)
        Me.ToolStripLabel1.Text = "Source details"
        '
        'lstParams
        '
        Me.lstParams.CheckOnClick = True
        Me.lstParams.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstParams.FormattingEnabled = True
        Me.lstParams.Location = New System.Drawing.Point(640, 58)
        Me.lstParams.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.lstParams.Name = "lstParams"
        Me.lstParams.Size = New System.Drawing.Size(180, 100)
        Me.lstParams.Sorted = True
        Me.lstParams.TabIndex = 184
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(562, 88)
        Me.Label20.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(66, 15)
        Me.Label20.TabIndex = 183
        Me.Label20.Text = "Parameters"
        '
        'cmboTest
        '
        Me.cmboTest.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboTest.FormattingEnabled = True
        Me.cmboTest.Location = New System.Drawing.Point(640, 33)
        Me.cmboTest.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboTest.Name = "cmboTest"
        Me.cmboTest.Size = New System.Drawing.Size(180, 23)
        Me.cmboTest.TabIndex = 181
        '
        'cmboEquationClass
        '
        Me.cmboEquationClass.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboEquationClass.FormattingEnabled = True
        Me.cmboEquationClass.Location = New System.Drawing.Point(125, 121)
        Me.cmboEquationClass.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboEquationClass.Name = "cmboEquationClass"
        Me.cmboEquationClass.Size = New System.Drawing.Size(352, 23)
        Me.cmboEquationClass.TabIndex = 3
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(19, 124)
        Me.Label8.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(103, 15)
        Me.Label8.TabIndex = 137
        Me.Label8.Text = "Class of equations"
        '
        'txtSource
        '
        Me.txtSource.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSource.Location = New System.Drawing.Point(125, 33)
        Me.txtSource.Name = "txtSource"
        Me.txtSource.Size = New System.Drawing.Size(352, 23)
        Me.txtSource.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(562, 38)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 15)
        Me.Label6.TabIndex = 127
        Me.Label6.Text = "Test"
        '
        'txtReference
        '
        Me.txtReference.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReference.Location = New System.Drawing.Point(125, 58)
        Me.txtReference.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtReference.Multiline = True
        Me.txtReference.Name = "txtReference"
        Me.txtReference.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtReference.Size = New System.Drawing.Size(352, 61)
        Me.txtReference.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(19, 74)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 15)
        Me.Label2.TabIndex = 125
        Me.Label2.Text = "Pub. reference"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(19, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 15)
        Me.Label1.TabIndex = 124
        Me.Label1.Text = "Source description"
        '
        'pnlEquationDetails
        '
        Me.pnlEquationDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlEquationDetails.Controls.Add(Me.ToolStrip2)
        Me.pnlEquationDetails.Controls.Add(Me.cmboParameter)
        Me.pnlEquationDetails.Controls.Add(Me.cmboWt_clipmethod)
        Me.pnlEquationDetails.Controls.Add(Me.txtWtUpper)
        Me.pnlEquationDetails.Controls.Add(Me.txtWtLower)
        Me.pnlEquationDetails.Controls.Add(Me.cmboHt_clipmethod)
        Me.pnlEquationDetails.Controls.Add(Me.txtHtUpper)
        Me.pnlEquationDetails.Controls.Add(Me.txtHtLower)
        Me.pnlEquationDetails.Controls.Add(Me.Label19)
        Me.pnlEquationDetails.Controls.Add(Me.Label18)
        Me.pnlEquationDetails.Controls.Add(Me.Label17)
        Me.pnlEquationDetails.Controls.Add(Me.Label16)
        Me.pnlEquationDetails.Controls.Add(Me.cmboAge_clipmethod)
        Me.pnlEquationDetails.Controls.Add(Me.chkApplyNonCaucasianCorrection)
        Me.pnlEquationDetails.Controls.Add(Me.cmboAgeGroup)
        Me.pnlEquationDetails.Controls.Add(Me.cmboStat)
        Me.pnlEquationDetails.Controls.Add(Me.Label10)
        Me.pnlEquationDetails.Controls.Add(Me.Label9)
        Me.pnlEquationDetails.Controls.Add(Me.Label13)
        Me.pnlEquationDetails.Controls.Add(Me.txtAgeUpper)
        Me.pnlEquationDetails.Controls.Add(Me.cmboGender)
        Me.pnlEquationDetails.Controls.Add(Me.cmboEthnicity)
        Me.pnlEquationDetails.Controls.Add(Me.Label12)
        Me.pnlEquationDetails.Controls.Add(Me.Label11)
        Me.pnlEquationDetails.Controls.Add(Me.Label7)
        Me.pnlEquationDetails.Controls.Add(Me.Label3)
        Me.pnlEquationDetails.Controls.Add(Me.txtEquation)
        Me.pnlEquationDetails.Controls.Add(Me.txtAgeLower)
        Me.pnlEquationDetails.Controls.Add(Me.Label5)
        Me.pnlEquationDetails.Controls.Add(Me.btnCheckEquation)
        Me.pnlEquationDetails.Location = New System.Drawing.Point(601, 198)
        Me.pnlEquationDetails.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.pnlEquationDetails.Name = "pnlEquationDetails"
        Me.pnlEquationDetails.Size = New System.Drawing.Size(540, 378)
        Me.pnlEquationDetails.TabIndex = 187
        '
        'ToolStrip2
        '
        Me.ToolStrip2.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel2})
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(538, 25)
        Me.ToolStrip2.TabIndex = 193
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(157, 22)
        Me.ToolStripLabel2.Text = "Details for selected equation"
        '
        'cmboParameter
        '
        Me.cmboParameter.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboParameter.FormattingEnabled = True
        Me.cmboParameter.Location = New System.Drawing.Point(72, 40)
        Me.cmboParameter.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboParameter.Name = "cmboParameter"
        Me.cmboParameter.Size = New System.Drawing.Size(156, 23)
        Me.cmboParameter.Sorted = True
        Me.cmboParameter.TabIndex = 192
        '
        'cmboWt_clipmethod
        '
        Me.cmboWt_clipmethod.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboWt_clipmethod.FormattingEnabled = True
        Me.cmboWt_clipmethod.Location = New System.Drawing.Point(173, 141)
        Me.cmboWt_clipmethod.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboWt_clipmethod.Name = "cmboWt_clipmethod"
        Me.cmboWt_clipmethod.Size = New System.Drawing.Size(344, 23)
        Me.cmboWt_clipmethod.TabIndex = 191
        '
        'txtWtUpper
        '
        Me.txtWtUpper.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWtUpper.Location = New System.Drawing.Point(124, 142)
        Me.txtWtUpper.Name = "txtWtUpper"
        Me.txtWtUpper.Size = New System.Drawing.Size(45, 23)
        Me.txtWtUpper.TabIndex = 190
        Me.txtWtUpper.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtWtLower
        '
        Me.txtWtLower.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWtLower.Location = New System.Drawing.Point(74, 142)
        Me.txtWtLower.Name = "txtWtLower"
        Me.txtWtLower.Size = New System.Drawing.Size(45, 23)
        Me.txtWtLower.TabIndex = 189
        Me.txtWtLower.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cmboHt_clipmethod
        '
        Me.cmboHt_clipmethod.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboHt_clipmethod.FormattingEnabled = True
        Me.cmboHt_clipmethod.Location = New System.Drawing.Point(173, 123)
        Me.cmboHt_clipmethod.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboHt_clipmethod.Name = "cmboHt_clipmethod"
        Me.cmboHt_clipmethod.Size = New System.Drawing.Size(344, 23)
        Me.cmboHt_clipmethod.TabIndex = 188
        '
        'txtHtUpper
        '
        Me.txtHtUpper.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHtUpper.Location = New System.Drawing.Point(124, 124)
        Me.txtHtUpper.Name = "txtHtUpper"
        Me.txtHtUpper.Size = New System.Drawing.Size(45, 23)
        Me.txtHtUpper.TabIndex = 187
        Me.txtHtUpper.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtHtLower
        '
        Me.txtHtLower.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHtLower.Location = New System.Drawing.Point(74, 124)
        Me.txtHtLower.Name = "txtHtLower"
        Me.txtHtLower.Size = New System.Drawing.Size(45, 23)
        Me.txtHtLower.TabIndex = 186
        Me.txtHtLower.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(14, 148)
        Me.Label19.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(45, 15)
        Me.Label19.TabIndex = 185
        Me.Label19.Text = "Weight"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(178, 85)
        Me.Label18.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(223, 15)
        Me.Label18.TabIndex = 184
        Me.Label18.Text = "Clip method (when value <min or >max)"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(134, 85)
        Me.Label17.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(29, 15)
        Me.Label17.TabIndex = 183
        Me.Label17.Text = "Max"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(83, 86)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(28, 15)
        Me.Label16.TabIndex = 182
        Me.Label16.Text = "Min"
        '
        'cmboAge_clipmethod
        '
        Me.cmboAge_clipmethod.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboAge_clipmethod.FormattingEnabled = True
        Me.cmboAge_clipmethod.Location = New System.Drawing.Point(173, 104)
        Me.cmboAge_clipmethod.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboAge_clipmethod.Name = "cmboAge_clipmethod"
        Me.cmboAge_clipmethod.Size = New System.Drawing.Size(344, 23)
        Me.cmboAge_clipmethod.TabIndex = 181
        '
        'chkApplyNonCaucasianCorrection
        '
        Me.chkApplyNonCaucasianCorrection.AutoSize = True
        Me.chkApplyNonCaucasianCorrection.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkApplyNonCaucasianCorrection.Location = New System.Drawing.Point(21, 345)
        Me.chkApplyNonCaucasianCorrection.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkApplyNonCaucasianCorrection.Name = "chkApplyNonCaucasianCorrection"
        Me.chkApplyNonCaucasianCorrection.Size = New System.Drawing.Size(254, 19)
        Me.chkApplyNonCaucasianCorrection.TabIndex = 180
        Me.chkApplyNonCaucasianCorrection.Text = "Apply ATS (1991) non-caucasian correction"
        Me.chkApplyNonCaucasianCorrection.UseVisualStyleBackColor = True
        '
        'cmboAgeGroup
        '
        Me.cmboAgeGroup.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboAgeGroup.FormattingEnabled = True
        Me.cmboAgeGroup.Location = New System.Drawing.Point(218, 183)
        Me.cmboAgeGroup.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboAgeGroup.Name = "cmboAgeGroup"
        Me.cmboAgeGroup.Size = New System.Drawing.Size(299, 23)
        Me.cmboAgeGroup.TabIndex = 165
        '
        'cmboStat
        '
        Me.cmboStat.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboStat.FormattingEnabled = True
        Me.cmboStat.Location = New System.Drawing.Point(90, 315)
        Me.cmboStat.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboStat.Name = "cmboStat"
        Me.cmboStat.Size = New System.Drawing.Size(81, 23)
        Me.cmboStat.TabIndex = 169
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(14, 127)
        Me.Label10.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(43, 15)
        Me.Label10.TabIndex = 174
        Me.Label10.Text = "Height"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(11, 185)
        Me.Label9.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(63, 15)
        Me.Label9.TabIndex = 173
        Me.Label9.Text = "Age group"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(13, 317)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(74, 15)
        Me.Label13.TabIndex = 178
        Me.Label13.Text = "Statistic type"
        '
        'txtAgeUpper
        '
        Me.txtAgeUpper.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAgeUpper.Location = New System.Drawing.Point(124, 104)
        Me.txtAgeUpper.Name = "txtAgeUpper"
        Me.txtAgeUpper.Size = New System.Drawing.Size(45, 23)
        Me.txtAgeUpper.TabIndex = 167
        Me.txtAgeUpper.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'cmboGender
        '
        Me.cmboGender.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboGender.FormattingEnabled = True
        Me.cmboGender.Location = New System.Drawing.Point(218, 206)
        Me.cmboGender.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboGender.Name = "cmboGender"
        Me.cmboGender.Size = New System.Drawing.Size(299, 23)
        Me.cmboGender.TabIndex = 163
        '
        'cmboEthnicity
        '
        Me.cmboEthnicity.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboEthnicity.FormattingEnabled = True
        Me.cmboEthnicity.Location = New System.Drawing.Point(218, 229)
        Me.cmboEthnicity.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cmboEthnicity.Name = "cmboEthnicity"
        Me.cmboEthnicity.Size = New System.Drawing.Size(299, 23)
        Me.cmboEthnicity.TabIndex = 164
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(10, 43)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(61, 15)
        Me.Label12.TabIndex = 177
        Me.Label12.Text = "Parameter"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(14, 108)
        Me.Label11.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(28, 15)
        Me.Label11.TabIndex = 175
        Me.Label11.Text = "Age"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(14, 282)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(61, 15)
        Me.Label7.TabIndex = 170
        Me.Label7.Text = "Equation"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 231)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(200, 15)
        Me.Label3.TabIndex = 172
        Me.Label3.Text = "Ethnicity(s) handled by this equation"
        '
        'txtEquation
        '
        Me.txtEquation.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEquation.Location = New System.Drawing.Point(90, 269)
        Me.txtEquation.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtEquation.Multiline = True
        Me.txtEquation.Name = "txtEquation"
        Me.txtEquation.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEquation.Size = New System.Drawing.Size(428, 43)
        Me.txtEquation.TabIndex = 168
        '
        'txtAgeLower
        '
        Me.txtAgeLower.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAgeLower.Location = New System.Drawing.Point(74, 104)
        Me.txtAgeLower.Name = "txtAgeLower"
        Me.txtAgeLower.Size = New System.Drawing.Size(45, 23)
        Me.txtAgeLower.TabIndex = 166
        Me.txtAgeLower.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AccessibleRole = System.Windows.Forms.AccessibleRole.OutlineButton
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(11, 208)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(192, 15)
        Me.Label5.TabIndex = 176
        Me.Label5.Text = "Gender(s) handled by this equation"
        '
        'btnCheckEquation
        '
        Me.btnCheckEquation.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheckEquation.Location = New System.Drawing.Point(389, 329)
        Me.btnCheckEquation.Name = "btnCheckEquation"
        Me.btnCheckEquation.Size = New System.Drawing.Size(128, 23)
        Me.btnCheckEquation.TabIndex = 171
        Me.btnCheckEquation.Text = "Check equation syntax"
        Me.btnCheckEquation.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnCancel.Enabled = False
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(918, 130)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 25)
        Me.btnCancel.TabIndex = 185
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnEdit.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEdit.Location = New System.Drawing.Point(1030, 101)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 25)
        Me.btnEdit.TabIndex = 183
        Me.btnEdit.Text = "Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnClose.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(1030, 130)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 25)
        Me.btnClose.TabIndex = 186
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(918, 102)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 25)
        Me.btnSave.TabIndex = 184
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'form_PredSourceEditor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1152, 587)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.pnlSource)
        Me.Controls.Add(Me.pnlEquationDetails)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "form_PredSourceEditor"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Predicted values editor"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ToolStrip5.ResumeLayout(False)
        Me.ToolStrip5.PerformLayout()
        CType(Me.grdEqs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlSource.ResumeLayout(False)
        Me.pnlSource.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.pnlEquationDetails.ResumeLayout(False)
        Me.pnlEquationDetails.PerformLayout()
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip5 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel4 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents grdEqs As System.Windows.Forms.DataGridView
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents pnlSource As System.Windows.Forms.Panel
    Friend WithEvents lstParams As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cmboTest As System.Windows.Forms.ComboBox
    Friend WithEvents cmboEquationClass As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtSource As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtReference As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlEquationDetails As System.Windows.Forms.Panel
    Friend WithEvents cmboParameter As System.Windows.Forms.ComboBox
    Friend WithEvents cmboWt_clipmethod As System.Windows.Forms.ComboBox
    Friend WithEvents txtWtUpper As System.Windows.Forms.TextBox
    Friend WithEvents txtWtLower As System.Windows.Forms.TextBox
    Friend WithEvents cmboHt_clipmethod As System.Windows.Forms.ComboBox
    Friend WithEvents txtHtUpper As System.Windows.Forms.TextBox
    Friend WithEvents txtHtLower As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cmboAge_clipmethod As System.Windows.Forms.ComboBox
    Friend WithEvents chkApplyNonCaucasianCorrection As System.Windows.Forms.CheckBox
    Friend WithEvents cmboAgeGroup As System.Windows.Forms.ComboBox
    Friend WithEvents cmboStat As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtAgeUpper As System.Windows.Forms.TextBox
    Friend WithEvents cmboGender As System.Windows.Forms.ComboBox
    Friend WithEvents cmboEthnicity As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtEquation As System.Windows.Forms.TextBox
    Friend WithEvents txtAgeLower As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnCheckEquation As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
End Class
