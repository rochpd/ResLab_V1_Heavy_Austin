<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_Reporting
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
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lv = New System.Windows.Forms.ListView()
        Me.ss = New System.Windows.Forms.StatusStrip()
        Me.tsLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FilterOnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportStatusToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsCmbo_reportstatus = New System.Windows.Forms.ToolStripComboBox()
        Me.ReportingMOToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsCmbo_requestingmo = New System.Windows.Forms.ToolStripComboBox()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsSave_all = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsSave_selected = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsSave_list = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsPrint_all = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsPrint_selected = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsPrint_list = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmboDefaultReporter = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDefaultDate = New System.Windows.Forms.MaskedTextBox()
        Me.ss.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSelect
        '
        Me.btnSelect.AutoSize = True
        Me.btnSelect.BackColor = System.Drawing.Color.PowderBlue
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnSelect.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnSelect.ForeColor = System.Drawing.Color.Blue
        Me.btnSelect.Location = New System.Drawing.Point(811, 485)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(71, 24)
        Me.btnSelect.TabIndex = 82
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.PowderBlue
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnClose.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnClose.ForeColor = System.Drawing.Color.Blue
        Me.btnClose.Location = New System.Drawing.Point(887, 485)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(71, 24)
        Me.btnClose.TabIndex = 81
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'lv
        '
        Me.lv.Dock = System.Windows.Forms.DockStyle.Top
        Me.lv.FullRowSelect = True
        Me.lv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lv.HideSelection = False
        Me.lv.Location = New System.Drawing.Point(0, 37)
        Me.lv.MultiSelect = False
        Me.lv.Name = "lv"
        Me.lv.Size = New System.Drawing.Size(989, 431)
        Me.lv.TabIndex = 85
        Me.lv.UseCompatibleStateImageBehavior = False
        Me.lv.View = System.Windows.Forms.View.List
        '
        'ss
        '
        Me.ss.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsLabel})
        Me.ss.Location = New System.Drawing.Point(0, 531)
        Me.ss.Name = "ss"
        Me.ss.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.ss.Size = New System.Drawing.Size(989, 22)
        Me.ss.TabIndex = 86
        Me.ss.Text = "StatusStrip1"
        '
        'tsLabel
        '
        Me.tsLabel.Name = "tsLabel"
        Me.tsLabel.Size = New System.Drawing.Size(0, 17)
        '
        'MenuStrip1
        '
        Me.MenuStrip1.AutoSize = False
        Me.MenuStrip1.BackColor = System.Drawing.Color.Silver
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FilterOnToolStripMenuItem, Me.ReportStatusToolStripMenuItem, Me.tsCmbo_reportstatus, Me.ReportingMOToolStripMenuItem, Me.tsCmbo_requestingmo, Me.SaveToolStripMenuItem, Me.PrintToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(989, 37)
        Me.MenuStrip1.TabIndex = 89
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FilterOnToolStripMenuItem
        '
        Me.FilterOnToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterOnToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.FilterOnToolStripMenuItem.Name = "FilterOnToolStripMenuItem"
        Me.FilterOnToolStripMenuItem.Size = New System.Drawing.Size(87, 33)
        Me.FilterOnToolStripMenuItem.Text = "Filter list on:"
        '
        'ReportStatusToolStripMenuItem
        '
        Me.ReportStatusToolStripMenuItem.Name = "ReportStatusToolStripMenuItem"
        Me.ReportStatusToolStripMenuItem.Size = New System.Drawing.Size(88, 33)
        Me.ReportStatusToolStripMenuItem.Text = "Report status"
        '
        'tsCmbo_reportstatus
        '
        Me.tsCmbo_reportstatus.BackColor = System.Drawing.SystemColors.Control
        Me.tsCmbo_reportstatus.Name = "tsCmbo_reportstatus"
        Me.tsCmbo_reportstatus.Size = New System.Drawing.Size(160, 33)
        '
        'ReportingMOToolStripMenuItem
        '
        Me.ReportingMOToolStripMenuItem.Name = "ReportingMOToolStripMenuItem"
        Me.ReportingMOToolStripMenuItem.Size = New System.Drawing.Size(101, 33)
        Me.ReportingMOToolStripMenuItem.Text = "Requesting MO"
        '
        'tsCmbo_requestingmo
        '
        Me.tsCmbo_requestingmo.BackColor = System.Drawing.SystemColors.Control
        Me.tsCmbo_requestingmo.Name = "tsCmbo_requestingmo"
        Me.tsCmbo_requestingmo.Size = New System.Drawing.Size(146, 33)
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.SaveToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsSave_all, Me.tsSave_selected, Me.tsSave_list})
        Me.SaveToolStripMenuItem.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.save_32x32
        Me.SaveToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(115, 33)
        Me.SaveToolStripMenuItem.Text = "Save reports"
        '
        'tsSave_all
        '
        Me.tsSave_all.Name = "tsSave_all"
        Me.tsSave_all.Size = New System.Drawing.Size(159, 22)
        Me.tsSave_all.Text = "All reports in list"
        '
        'tsSave_selected
        '
        Me.tsSave_selected.Enabled = False
        Me.tsSave_selected.Name = "tsSave_selected"
        Me.tsSave_selected.Size = New System.Drawing.Size(159, 22)
        Me.tsSave_selected.Text = "Selected reports"
        Me.tsSave_selected.Visible = False
        '
        'tsSave_list
        '
        Me.tsSave_list.Enabled = False
        Me.tsSave_list.Name = "tsSave_list"
        Me.tsSave_list.Size = New System.Drawing.Size(159, 22)
        Me.tsSave_list.Text = "Report list"
        Me.tsSave_list.Visible = False
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.PrintToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsPrint_all, Me.tsPrint_selected, Me.tsPrint_list})
        Me.PrintToolStripMenuItem.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.print_32x32
        Me.PrintToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(116, 33)
        Me.PrintToolStripMenuItem.Text = "Print reports"
        '
        'tsPrint_all
        '
        Me.tsPrint_all.Name = "tsPrint_all"
        Me.tsPrint_all.Size = New System.Drawing.Size(159, 22)
        Me.tsPrint_all.Text = "All reports in list"
        '
        'tsPrint_selected
        '
        Me.tsPrint_selected.Enabled = False
        Me.tsPrint_selected.Name = "tsPrint_selected"
        Me.tsPrint_selected.Size = New System.Drawing.Size(159, 22)
        Me.tsPrint_selected.Text = "Selected reports"
        Me.tsPrint_selected.Visible = False
        '
        'tsPrint_list
        '
        Me.tsPrint_list.Enabled = False
        Me.tsPrint_list.Name = "tsPrint_list"
        Me.tsPrint_list.Size = New System.Drawing.Size(159, 22)
        Me.tsPrint_list.Text = "Report list"
        Me.tsPrint_list.Visible = False
        '
        'cmboDefaultReporter
        '
        Me.cmboDefaultReporter.FormattingEnabled = True
        Me.cmboDefaultReporter.Location = New System.Drawing.Point(357, 482)
        Me.cmboDefaultReporter.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboDefaultReporter.Name = "cmboDefaultReporter"
        Me.cmboDefaultReporter.Size = New System.Drawing.Size(150, 21)
        Me.cmboDefaultReporter.TabIndex = 90
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(272, 485)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 91
        Me.Label1.Text = "Default reporter"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(272, 507)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 13)
        Me.Label2.TabIndex = 92
        Me.Label2.Text = "Default date"
        '
        'txtDefaultDate
        '
        Me.txtDefaultDate.Location = New System.Drawing.Point(357, 504)
        Me.txtDefaultDate.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDefaultDate.Mask = "##/##/####"
        Me.txtDefaultDate.Name = "txtDefaultDate"
        Me.txtDefaultDate.Size = New System.Drawing.Size(150, 20)
        Me.txtDefaultDate.TabIndex = 93
        '
        'form_Reporting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 553)
        Me.Controls.Add(Me.txtDefaultDate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmboDefaultReporter)
        Me.Controls.Add(Me.ss)
        Me.Controls.Add(Me.lv)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "form_Reporting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reporting"
        Me.ss.ResumeLayout(False)
        Me.ss.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lv As System.Windows.Forms.ListView
    Friend WithEvents ss As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FilterOnToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportStatusToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsCmbo_reportstatus As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents ReportingMOToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsCmbo_requestingmo As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents tsLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsPrint_all As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsPrint_selected As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsPrint_list As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmboDefaultReporter As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDefaultDate As System.Windows.Forms.MaskedTextBox
    Friend WithEvents SaveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsSave_all As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsSave_selected As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsSave_list As System.Windows.Forms.ToolStripMenuItem
End Class
