<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_reportsdb
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
        Me.split = New System.Windows.Forms.SplitContainer()
        Me.txtEnd = New System.Windows.Forms.MaskedTextBox()
        Me.txtStart = New System.Windows.Forms.MaskedTextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.optPatientList = New System.Windows.Forms.RadioButton()
        Me.optStatsByMonth = New System.Windows.Forms.RadioButton()
        Me.ToolStrip_ActionPanel = New System.Windows.Forms.ToolStrip()
        Me.tslbl_ReportTitle = New System.Windows.Forms.ToolStripLabel()
        Me.sBar = New System.Windows.Forms.StatusStrip()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lblEnd = New System.Windows.Forms.Label()
        Me.lblStart = New System.Windows.Forms.Label()
        Me.dtp = New System.Windows.Forms.DateTimePicker()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.pnlDates = New System.Windows.Forms.Panel()
        CType(Me.split, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.split.Panel1.SuspendLayout()
        Me.split.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.ToolStrip_ActionPanel.SuspendLayout()
        Me.pnlDates.SuspendLayout()
        Me.SuspendLayout()
        '
        'split
        '
        Me.split.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.split.Dock = System.Windows.Forms.DockStyle.Fill
        Me.split.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.split.Location = New System.Drawing.Point(0, 0)
        Me.split.Margin = New System.Windows.Forms.Padding(2)
        Me.split.Name = "split"
        '
        'split.Panel1
        '
        Me.split.Panel1.Controls.Add(Me.pnlDates)
        Me.split.Panel1.Controls.Add(Me.Panel1)
        Me.split.Panel1.Controls.Add(Me.ToolStrip_ActionPanel)
        Me.split.Panel1.Controls.Add(Me.sBar)
        Me.split.Panel1.Controls.Add(Me.btnClose)
        Me.split.Panel1.Controls.Add(Me.btnGenerate)
        '
        'split.Panel2
        '
        Me.split.Panel2.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.split.Size = New System.Drawing.Size(827, 505)
        Me.split.SplitterDistance = 210
        Me.split.SplitterWidth = 3
        Me.split.TabIndex = 18
        '
        'txtEnd
        '
        Me.txtEnd.Location = New System.Drawing.Point(80, 37)
        Me.txtEnd.Name = "txtEnd"
        Me.txtEnd.Size = New System.Drawing.Size(83, 20)
        Me.txtEnd.TabIndex = 29
        '
        'txtStart
        '
        Me.txtStart.Location = New System.Drawing.Point(80, 14)
        Me.txtStart.Name = "txtStart"
        Me.txtStart.Size = New System.Drawing.Size(83, 20)
        Me.txtStart.TabIndex = 28
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.optPatientList)
        Me.Panel1.Controls.Add(Me.optStatsByMonth)
        Me.Panel1.Location = New System.Drawing.Point(26, 39)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(153, 65)
        Me.Panel1.TabIndex = 27
        '
        'optPatientList
        '
        Me.optPatientList.AutoSize = True
        Me.optPatientList.Location = New System.Drawing.Point(19, 31)
        Me.optPatientList.Name = "optPatientList"
        Me.optPatientList.Size = New System.Drawing.Size(73, 17)
        Me.optPatientList.TabIndex = 1
        Me.optPatientList.Text = "Patient list"
        Me.optPatientList.UseVisualStyleBackColor = True
        '
        'optStatsByMonth
        '
        Me.optStatsByMonth.AutoSize = True
        Me.optStatsByMonth.Checked = True
        Me.optStatsByMonth.Location = New System.Drawing.Point(19, 8)
        Me.optStatsByMonth.Name = "optStatsByMonth"
        Me.optStatsByMonth.Size = New System.Drawing.Size(95, 17)
        Me.optStatsByMonth.TabIndex = 0
        Me.optStatsByMonth.TabStop = True
        Me.optStatsByMonth.Text = "Stats by month"
        Me.optStatsByMonth.UseVisualStyleBackColor = True
        '
        'ToolStrip_ActionPanel
        '
        Me.ToolStrip_ActionPanel.BackColor = System.Drawing.SystemColors.Info
        Me.ToolStrip_ActionPanel.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslbl_ReportTitle})
        Me.ToolStrip_ActionPanel.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip_ActionPanel.Name = "ToolStrip_ActionPanel"
        Me.ToolStrip_ActionPanel.Size = New System.Drawing.Size(206, 25)
        Me.ToolStrip_ActionPanel.TabIndex = 26
        Me.ToolStrip_ActionPanel.Text = "ToolStrip1"
        '
        'tslbl_ReportTitle
        '
        Me.tslbl_ReportTitle.Name = "tslbl_ReportTitle"
        Me.tslbl_ReportTitle.Size = New System.Drawing.Size(0, 22)
        '
        'sBar
        '
        Me.sBar.Location = New System.Drawing.Point(0, 479)
        Me.sBar.Name = "sBar"
        Me.sBar.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.sBar.Size = New System.Drawing.Size(206, 22)
        Me.sBar.TabIndex = 25
        Me.sBar.Text = "StatusStrip1"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(37, 221)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 22)
        Me.btnClose.TabIndex = 24
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'lblEnd
        '
        Me.lblEnd.AutoSize = True
        Me.lblEnd.Location = New System.Drawing.Point(26, 37)
        Me.lblEnd.Name = "lblEnd"
        Me.lblEnd.Size = New System.Drawing.Size(50, 13)
        Me.lblEnd.TabIndex = 21
        Me.lblEnd.Text = "End date"
        '
        'lblStart
        '
        Me.lblStart.AutoSize = True
        Me.lblStart.Location = New System.Drawing.Point(26, 17)
        Me.lblStart.Name = "lblStart"
        Me.lblStart.Size = New System.Drawing.Size(53, 13)
        Me.lblStart.TabIndex = 19
        Me.lblStart.Text = "Start date"
        '
        'dtp
        '
        Me.dtp.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp.Location = New System.Drawing.Point(65, 64)
        Me.dtp.Name = "dtp"
        Me.dtp.Size = New System.Drawing.Size(98, 20)
        Me.dtp.TabIndex = 20
        Me.dtp.Visible = False
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(114, 221)
        Me.btnGenerate.Margin = New System.Windows.Forms.Padding(2)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(64, 22)
        Me.btnGenerate.TabIndex = 18
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'pnlDates
        '
        Me.pnlDates.Controls.Add(Me.txtStart)
        Me.pnlDates.Controls.Add(Me.txtEnd)
        Me.pnlDates.Controls.Add(Me.lblStart)
        Me.pnlDates.Controls.Add(Me.lblEnd)
        Me.pnlDates.Controls.Add(Me.dtp)
        Me.pnlDates.Location = New System.Drawing.Point(10, 110)
        Me.pnlDates.Name = "pnlDates"
        Me.pnlDates.Size = New System.Drawing.Size(181, 87)
        Me.pnlDates.TabIndex = 30
        '
        'form_reportsdb
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(827, 505)
        Me.Controls.Add(Me.split)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "form_reportsdb"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Database reports"
        Me.split.Panel1.ResumeLayout(False)
        Me.split.Panel1.PerformLayout()
        CType(Me.split, System.ComponentModel.ISupportInitialize).EndInit()
        Me.split.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ToolStrip_ActionPanel.ResumeLayout(False)
        Me.ToolStrip_ActionPanel.PerformLayout()
        Me.pnlDates.ResumeLayout(False)
        Me.pnlDates.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents split As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblEnd As System.Windows.Forms.Label
    Friend WithEvents lblStart As System.Windows.Forms.Label
    Friend WithEvents dtp As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents ToolStrip_ActionPanel As System.Windows.Forms.ToolStrip
    Friend WithEvents tslbl_ReportTitle As System.Windows.Forms.ToolStripLabel
    Friend WithEvents sBar As System.Windows.Forms.StatusStrip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents optPatientList As System.Windows.Forms.RadioButton
    Friend WithEvents optStatsByMonth As System.Windows.Forms.RadioButton
    Friend WithEvents txtEnd As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtStart As System.Windows.Forms.MaskedTextBox
    Friend WithEvents pnlDates As System.Windows.Forms.Panel
End Class
