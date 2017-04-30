<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_customize
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.tabCustom = New System.Windows.Forms.TabControl()
        Me.tabpageReportHeader = New System.Windows.Forms.TabPage()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.grdReportHeader = New System.Windows.Forms.DataGridView()
        Me.recordID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Item = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Content = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.tsBtn_Cancel = New System.Windows.Forms.ToolStripButton()
        Me.tsBtn_SaveAndClose = New System.Windows.Forms.ToolStripButton()
        Me.tabCustom.SuspendLayout()
        Me.tabpageReportHeader.SuspendLayout()
        CType(Me.grdReportHeader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabCustom
        '
        Me.tabCustom.Controls.Add(Me.tabpageReportHeader)
        Me.tabCustom.Controls.Add(Me.TabPage2)
        Me.tabCustom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabCustom.Location = New System.Drawing.Point(0, 0)
        Me.tabCustom.Name = "tabCustom"
        Me.tabCustom.SelectedIndex = 0
        Me.tabCustom.Size = New System.Drawing.Size(655, 395)
        Me.tabCustom.TabIndex = 0
        '
        'tabpageReportHeader
        '
        Me.tabpageReportHeader.Controls.Add(Me.Label1)
        Me.tabpageReportHeader.Controls.Add(Me.grdReportHeader)
        Me.tabpageReportHeader.Location = New System.Drawing.Point(4, 25)
        Me.tabpageReportHeader.Name = "tabpageReportHeader"
        Me.tabpageReportHeader.Padding = New System.Windows.Forms.Padding(3)
        Me.tabpageReportHeader.Size = New System.Drawing.Size(647, 366)
        Me.tabpageReportHeader.TabIndex = 0
        Me.tabpageReportHeader.Text = "Reports"
        Me.tabpageReportHeader.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(280, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Report header - service name and address"
        '
        'grdReportHeader
        '
        Me.grdReportHeader.AllowUserToAddRows = False
        Me.grdReportHeader.AllowUserToDeleteRows = False
        Me.grdReportHeader.AllowUserToResizeRows = False
        Me.grdReportHeader.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Linen
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.grdReportHeader.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.grdReportHeader.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.grdReportHeader.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.recordID, Me.Item, Me.Content})
        Me.grdReportHeader.EnableHeadersVisualStyles = False
        Me.grdReportHeader.Location = New System.Drawing.Point(8, 45)
        Me.grdReportHeader.MultiSelect = False
        Me.grdReportHeader.Name = "grdReportHeader"
        Me.grdReportHeader.RowHeadersVisible = False
        Me.grdReportHeader.RowTemplate.Height = 24
        Me.grdReportHeader.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.grdReportHeader.Size = New System.Drawing.Size(619, 268)
        Me.grdReportHeader.TabIndex = 2
        '
        'recordID
        '
        Me.recordID.HeaderText = "recordID"
        Me.recordID.Name = "recordID"
        Me.recordID.ReadOnly = True
        Me.recordID.Visible = False
        Me.recordID.Width = 50
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
        Me.Content.HeaderText = "Contents"
        Me.Content.Name = "Content"
        Me.Content.Width = 390
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(647, 366)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Other"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsBtn_Cancel, Me.tsBtn_SaveAndClose})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 368)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(655, 27)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsBtn_Cancel
        '
        Me.tsBtn_Cancel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsBtn_Cancel.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.cancel_32x32
        Me.tsBtn_Cancel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsBtn_Cancel.Name = "tsBtn_Cancel"
        Me.tsBtn_Cancel.Size = New System.Drawing.Size(73, 24)
        Me.tsBtn_Cancel.Text = "Cancel"
        '
        'tsBtn_SaveAndClose
        '
        Me.tsBtn_SaveAndClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsBtn_SaveAndClose.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.save_32x32
        Me.tsBtn_SaveAndClose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsBtn_SaveAndClose.Name = "tsBtn_SaveAndClose"
        Me.tsBtn_SaveAndClose.Size = New System.Drawing.Size(129, 24)
        Me.tsBtn_SaveAndClose.Text = "Save and Close"
        '
        'form_customize
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(655, 395)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.tabCustom)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "form_customize"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customize"
        Me.tabCustom.ResumeLayout(False)
        Me.tabpageReportHeader.ResumeLayout(False)
        Me.tabpageReportHeader.PerformLayout()
        CType(Me.grdReportHeader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tabCustom As System.Windows.Forms.TabControl
    Friend WithEvents tabpageReportHeader As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsBtn_Cancel As System.Windows.Forms.ToolStripButton
    Protected Friend WithEvents tsBtn_SaveAndClose As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents grdReportHeader As System.Windows.Forms.DataGridView
    Friend WithEvents recordID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Item As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Content As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
