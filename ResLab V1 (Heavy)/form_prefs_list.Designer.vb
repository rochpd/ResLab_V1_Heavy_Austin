<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_prefs_list
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(form_prefs_list))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.tabPrefs = New System.Windows.Forms.TabControl()
        Me.tabPage_Lists = New System.Windows.Forms.TabPage()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lstFields = New System.Windows.Forms.ListBox()
        Me.ToolStrip4 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.tsbField_edit = New System.Windows.Forms.ToolStripButton()
        Me.tsbField_new = New System.Windows.Forms.ToolStripButton()
        Me.tsbField_delete = New System.Windows.Forms.ToolStripButton()
        Me.lstFieldOptions = New System.Windows.Forms.ListBox()
        Me.ToolStrip_fieldoptions = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel()
        Me.tsbFieldOptions_delete = New System.Windows.Forms.ToolStripButton()
        Me.tsbFieldOptions_makedefault = New System.Windows.Forms.ToolStripButton()
        Me.tsbFieldOptions_edit = New System.Windows.Forms.ToolStripButton()
        Me.tsbFieldOptions_new = New System.Windows.Forms.ToolStripButton()
        Me.tabpage_ReportHeader = New System.Windows.Forms.TabPage()
        Me.grdReportHeader = New System.Windows.Forms.DataGridView()
        Me.recordID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Item = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Content = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.tabPrefs.SuspendLayout()
        Me.tabPage_Lists.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.ToolStrip4.SuspendLayout()
        Me.ToolStrip_fieldoptions.SuspendLayout()
        Me.tabpage_ReportHeader.SuspendLayout()
        CType(Me.grdReportHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tabPrefs
        '
        Me.tabPrefs.Controls.Add(Me.tabPage_Lists)
        Me.tabPrefs.Controls.Add(Me.tabpage_ReportHeader)
        Me.tabPrefs.Dock = System.Windows.Forms.DockStyle.Top
        Me.tabPrefs.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.tabPrefs.Location = New System.Drawing.Point(0, 0)
        Me.tabPrefs.Margin = New System.Windows.Forms.Padding(2)
        Me.tabPrefs.Name = "tabPrefs"
        Me.tabPrefs.SelectedIndex = 0
        Me.tabPrefs.Size = New System.Drawing.Size(887, 400)
        Me.tabPrefs.TabIndex = 1
        '
        'tabPage_Lists
        '
        Me.tabPage_Lists.Controls.Add(Me.SplitContainer1)
        Me.tabPage_Lists.Location = New System.Drawing.Point(4, 24)
        Me.tabPage_Lists.Margin = New System.Windows.Forms.Padding(2)
        Me.tabPage_Lists.Name = "tabPage_Lists"
        Me.tabPage_Lists.Padding = New System.Windows.Forms.Padding(2)
        Me.tabPage_Lists.Size = New System.Drawing.Size(879, 372)
        Me.tabPage_Lists.TabIndex = 1
        Me.tabPage_Lists.Text = "Option lists"
        Me.tabPage_Lists.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lstFields)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ToolStrip4)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lstFieldOptions)
        Me.SplitContainer1.Panel2.Controls.Add(Me.ToolStrip_fieldoptions)
        Me.SplitContainer1.Size = New System.Drawing.Size(875, 368)
        Me.SplitContainer1.SplitterDistance = 291
        Me.SplitContainer1.TabIndex = 0
        '
        'lstFields
        '
        Me.lstFields.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.lstFields.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstFields.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lstFields.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstFields.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstFields.ItemHeight = 15
        Me.lstFields.Location = New System.Drawing.Point(0, 23)
        Me.lstFields.Margin = New System.Windows.Forms.Padding(2)
        Me.lstFields.Name = "lstFields"
        Me.lstFields.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstFields.Size = New System.Drawing.Size(291, 345)
        Me.lstFields.Sorted = True
        Me.lstFields.TabIndex = 34
        '
        'ToolStrip4
        '
        Me.ToolStrip4.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ToolStrip4.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1, Me.tsbField_edit, Me.tsbField_new, Me.tsbField_delete})
        Me.ToolStrip4.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip4.Name = "ToolStrip4"
        Me.ToolStrip4.Size = New System.Drawing.Size(291, 25)
        Me.ToolStrip4.TabIndex = 35
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.ActiveLinkColor = System.Drawing.Color.RoyalBlue
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.ForeColor = System.Drawing.Color.Blue
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(77, 22)
        Me.ToolStripLabel1.Text = "Fields             "
        '
        'tsbField_edit
        '
        Me.tsbField_edit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbField_edit.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbField_edit.Image = CType(resources.GetObject("tsbField_edit.Image"), System.Drawing.Image)
        Me.tsbField_edit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbField_edit.Name = "tsbField_edit"
        Me.tsbField_edit.Size = New System.Drawing.Size(47, 22)
        Me.tsbField_edit.Text = "Edit"
        Me.tsbField_edit.ToolTipText = "Edit selected field"
        Me.tsbField_edit.Visible = False
        '
        'tsbField_new
        '
        Me.tsbField_new.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbField_new.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbField_new.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.new_pro32
        Me.tsbField_new.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbField_new.Name = "tsbField_new"
        Me.tsbField_new.Size = New System.Drawing.Size(50, 22)
        Me.tsbField_new.Text = "New"
        Me.tsbField_new.ToolTipText = "New field"
        Me.tsbField_new.Visible = False
        '
        'tsbField_delete
        '
        Me.tsbField_delete.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbField_delete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbField_delete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbField_delete.Name = "tsbField_delete"
        Me.tsbField_delete.Size = New System.Drawing.Size(23, 22)
        Me.tsbField_delete.Text = "ToolStripButton1"
        Me.tsbField_delete.ToolTipText = "Delete selected field"
        '
        'lstFieldOptions
        '
        Me.lstFieldOptions.BackColor = System.Drawing.SystemColors.Window
        Me.lstFieldOptions.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstFieldOptions.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lstFieldOptions.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lstFieldOptions.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstFieldOptions.ItemHeight = 15
        Me.lstFieldOptions.Location = New System.Drawing.Point(0, 23)
        Me.lstFieldOptions.Margin = New System.Windows.Forms.Padding(2)
        Me.lstFieldOptions.Name = "lstFieldOptions"
        Me.lstFieldOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstFieldOptions.Size = New System.Drawing.Size(580, 345)
        Me.lstFieldOptions.Sorted = True
        Me.lstFieldOptions.TabIndex = 34
        '
        'ToolStrip_fieldoptions
        '
        Me.ToolStrip_fieldoptions.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.ToolStrip_fieldoptions.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel2, Me.tsbFieldOptions_delete, Me.tsbFieldOptions_makedefault, Me.tsbFieldOptions_edit, Me.tsbFieldOptions_new})
        Me.ToolStrip_fieldoptions.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip_fieldoptions.Name = "ToolStrip_fieldoptions"
        Me.ToolStrip_fieldoptions.Size = New System.Drawing.Size(580, 26)
        Me.ToolStrip_fieldoptions.TabIndex = 36
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.ActiveLinkColor = System.Drawing.Color.RoyalBlue
        Me.ToolStripLabel2.AutoSize = False
        Me.ToolStripLabel2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel2.ForeColor = System.Drawing.Color.Blue
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(150, 23)
        Me.ToolStripLabel2.Text = " Field options                                                                "
        Me.ToolStripLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tsbFieldOptions_delete
        '
        Me.tsbFieldOptions_delete.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbFieldOptions_delete.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbFieldOptions_delete.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.delete
        Me.tsbFieldOptions_delete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFieldOptions_delete.Name = "tsbFieldOptions_delete"
        Me.tsbFieldOptions_delete.Size = New System.Drawing.Size(60, 23)
        Me.tsbFieldOptions_delete.Text = "Delete"
        Me.tsbFieldOptions_delete.ToolTipText = "Delete selected field option"
        '
        'tsbFieldOptions_makedefault
        '
        Me.tsbFieldOptions_makedefault.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbFieldOptions_makedefault.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbFieldOptions_makedefault.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.CHECKER
        Me.tsbFieldOptions_makedefault.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFieldOptions_makedefault.Name = "tsbFieldOptions_makedefault"
        Me.tsbFieldOptions_makedefault.Size = New System.Drawing.Size(83, 23)
        Me.tsbFieldOptions_makedefault.Text = "Set default"
        Me.tsbFieldOptions_makedefault.ToolTipText = "Set as default"
        '
        'tsbFieldOptions_edit
        '
        Me.tsbFieldOptions_edit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbFieldOptions_edit.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbFieldOptions_edit.Image = CType(resources.GetObject("tsbFieldOptions_edit.Image"), System.Drawing.Image)
        Me.tsbFieldOptions_edit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFieldOptions_edit.Name = "tsbFieldOptions_edit"
        Me.tsbFieldOptions_edit.Size = New System.Drawing.Size(47, 23)
        Me.tsbFieldOptions_edit.Text = "Edit"
        Me.tsbFieldOptions_edit.ToolTipText = "Edit selected field option"
        '
        'tsbFieldOptions_new
        '
        Me.tsbFieldOptions_new.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbFieldOptions_new.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbFieldOptions_new.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.new_pro32
        Me.tsbFieldOptions_new.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbFieldOptions_new.Name = "tsbFieldOptions_new"
        Me.tsbFieldOptions_new.Size = New System.Drawing.Size(48, 23)
        Me.tsbFieldOptions_new.Text = "Add"
        Me.tsbFieldOptions_new.ToolTipText = "New field option"
        '
        'tabpage_ReportHeader
        '
        Me.tabpage_ReportHeader.Controls.Add(Me.grdReportHeader)
        Me.tabpage_ReportHeader.Location = New System.Drawing.Point(4, 24)
        Me.tabpage_ReportHeader.Name = "tabpage_ReportHeader"
        Me.tabpage_ReportHeader.Padding = New System.Windows.Forms.Padding(3)
        Me.tabpage_ReportHeader.Size = New System.Drawing.Size(879, 372)
        Me.tabpage_ReportHeader.TabIndex = 2
        Me.tabpage_ReportHeader.Text = "Report header"
        Me.tabpage_ReportHeader.UseVisualStyleBackColor = True
        '
        'grdReportHeader
        '
        Me.grdReportHeader.AllowUserToAddRows = False
        Me.grdReportHeader.AllowUserToDeleteRows = False
        Me.grdReportHeader.AllowUserToResizeRows = False
        Me.grdReportHeader.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.grdReportHeader.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.grdReportHeader.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Blue
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdReportHeader.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdReportHeader.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.recordID, Me.Item, Me.Content})
        Me.grdReportHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grdReportHeader.EnableHeadersVisualStyles = False
        Me.grdReportHeader.Location = New System.Drawing.Point(3, 3)
        Me.grdReportHeader.Margin = New System.Windows.Forms.Padding(2)
        Me.grdReportHeader.MultiSelect = False
        Me.grdReportHeader.Name = "grdReportHeader"
        Me.grdReportHeader.RowHeadersVisible = False
        Me.grdReportHeader.RowTemplate.Height = 18
        Me.grdReportHeader.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdReportHeader.Size = New System.Drawing.Size(873, 366)
        Me.grdReportHeader.TabIndex = 3
        '
        'recordID
        '
        Me.recordID.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.recordID.HeaderText = "recordID"
        Me.recordID.Name = "recordID"
        Me.recordID.ReadOnly = True
        Me.recordID.Visible = False
        '
        'Item
        '
        Me.Item.HeaderText = "Item"
        Me.Item.Name = "Item"
        Me.Item.ReadOnly = True
        Me.Item.Width = 220
        '
        'Content
        '
        Me.Content.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Content.HeaderText = "Text"
        Me.Content.Name = "Content"
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(801, 414)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 22)
        Me.btnClose.TabIndex = 55
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'form_prefs_list
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(887, 444)
        Me.Controls.Add(Me.tabPrefs)
        Me.Controls.Add(Me.btnClose)
        Me.DoubleBuffered = True
        Me.Name = "form_prefs_list"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Preferences: Fields and field options"
        Me.tabPrefs.ResumeLayout(False)
        Me.tabPage_Lists.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ToolStrip4.ResumeLayout(False)
        Me.ToolStrip4.PerformLayout()
        Me.ToolStrip_fieldoptions.ResumeLayout(False)
        Me.ToolStrip_fieldoptions.PerformLayout()
        Me.tabpage_ReportHeader.ResumeLayout(False)
        CType(Me.grdReportHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabPrefs As System.Windows.Forms.TabControl
    Friend WithEvents tabPage_Lists As System.Windows.Forms.TabPage
    Public WithEvents lstFieldOptions As System.Windows.Forms.ListBox
    Friend WithEvents ToolStrip4 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbField_new As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbField_edit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbField_delete As System.Windows.Forms.ToolStripButton
    Public WithEvents lstFields As System.Windows.Forms.ListBox
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStrip_fieldoptions As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tsbFieldOptions_new As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFieldOptions_edit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbFieldOptions_delete As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents tsbFieldOptions_makedefault As System.Windows.Forms.ToolStripButton
    Friend WithEvents tabpage_ReportHeader As System.Windows.Forms.TabPage
    Friend WithEvents grdReportHeader As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents recordID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Content As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
