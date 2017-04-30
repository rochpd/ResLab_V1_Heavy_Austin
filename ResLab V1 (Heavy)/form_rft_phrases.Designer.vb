<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_rft_phrases
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
        Me.components = New System.ComponentModel.Container()
        Me.split = New System.Windows.Forms.SplitContainer()
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.lstGroups = New System.Windows.Forms.ListBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel()
        Me.tsbtn_clear = New System.Windows.Forms.ToolStripButton()
        Me.tsbtn_Undo = New System.Windows.Forms.ToolStripButton()
        Me.tsbtnAutoReport = New System.Windows.Forms.ToolStripButton()
        Me.lstItems = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip_edititems = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.NewItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.split, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.split.Panel1.SuspendLayout()
        Me.split.Panel2.SuspendLayout()
        Me.split.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.ContextMenuStrip_edititems.SuspendLayout()
        Me.SuspendLayout()
        '
        'split
        '
        Me.split.Dock = System.Windows.Forms.DockStyle.Fill
        Me.split.Location = New System.Drawing.Point(0, 0)
        Me.split.Margin = New System.Windows.Forms.Padding(2)
        Me.split.Name = "split"
        '
        'split.Panel1
        '
        Me.split.Panel1.Controls.Add(Me.ToolStrip2)
        Me.split.Panel1.Controls.Add(Me.lstGroups)
        '
        'split.Panel2
        '
        Me.split.Panel2.Controls.Add(Me.ToolStrip1)
        Me.split.Panel2.Controls.Add(Me.lstItems)
        Me.split.Size = New System.Drawing.Size(532, 285)
        Me.split.SplitterDistance = 93
        Me.split.SplitterWidth = 3
        Me.split.TabIndex = 0
        '
        'ToolStrip2
        '
        Me.ToolStrip2.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStrip2.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1})
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(93, 25)
        Me.ToolStrip2.TabIndex = 3
        Me.ToolStrip2.Text = "ToolStrip2"
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Segoe UI Semibold", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(79, 22)
        Me.ToolStripLabel1.Text = "Phrase groups"
        '
        'lstGroups
        '
        Me.lstGroups.BackColor = System.Drawing.Color.LightYellow
        Me.lstGroups.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstGroups.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lstGroups.FormattingEnabled = True
        Me.lstGroups.Location = New System.Drawing.Point(0, 25)
        Me.lstGroups.Margin = New System.Windows.Forms.Padding(2)
        Me.lstGroups.Name = "lstGroups"
        Me.lstGroups.Size = New System.Drawing.Size(93, 260)
        Me.lstGroups.TabIndex = 0
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel2, Me.tsbtnAutoReport, Me.tsbtn_clear, Me.tsbtn_Undo})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(436, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.Font = New System.Drawing.Font("Segoe UI Semibold", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(57, 22)
        Me.ToolStripLabel2.Text = "Phrases    "
        '
        'tsbtn_clear
        '
        Me.tsbtn_clear.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtn_clear.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.cancel_32x32
        Me.tsbtn_clear.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtn_clear.Name = "tsbtn_clear"
        Me.tsbtn_clear.Size = New System.Drawing.Size(88, 22)
        Me.tsbtn_clear.Text = "Clear report"
        '
        'tsbtn_Undo
        '
        Me.tsbtn_Undo.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtn_Undo.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.undo_32x32
        Me.tsbtn_Undo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtn_Undo.Name = "tsbtn_Undo"
        Me.tsbtn_Undo.Size = New System.Drawing.Size(56, 22)
        Me.tsbtn_Undo.Text = "Undo"
        '
        'tsbtnAutoReport
        '
        Me.tsbtnAutoReport.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtnAutoReport.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.round_add_32x321
        Me.tsbtnAutoReport.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtnAutoReport.Name = "tsbtnAutoReport"
        Me.tsbtnAutoReport.Size = New System.Drawing.Size(87, 22)
        Me.tsbtnAutoReport.Text = "Auto report"
        '
        'lstItems
        '
        Me.lstItems.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstItems.ContextMenuStrip = Me.ContextMenuStrip_edititems
        Me.lstItems.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lstItems.FormattingEnabled = True
        Me.lstItems.Location = New System.Drawing.Point(0, 25)
        Me.lstItems.Margin = New System.Windows.Forms.Padding(2)
        Me.lstItems.Name = "lstItems"
        Me.lstItems.Size = New System.Drawing.Size(436, 260)
        Me.lstItems.TabIndex = 0
        '
        'ContextMenuStrip_edititems
        '
        Me.ContextMenuStrip_edititems.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewItemToolStripMenuItem, Me.EditItemToolStripMenuItem, Me.DeleteToolStripMenuItem})
        Me.ContextMenuStrip_edititems.Name = "ContextMenuStrip_edititems"
        Me.ContextMenuStrip_edititems.Size = New System.Drawing.Size(108, 70)
        '
        'NewItemToolStripMenuItem
        '
        Me.NewItemToolStripMenuItem.Name = "NewItemToolStripMenuItem"
        Me.NewItemToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.NewItemToolStripMenuItem.Text = "New"
        '
        'EditItemToolStripMenuItem
        '
        Me.EditItemToolStripMenuItem.Name = "EditItemToolStripMenuItem"
        Me.EditItemToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.EditItemToolStripMenuItem.Text = "Edit"
        '
        'DeleteToolStripMenuItem
        '
        Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
        Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.DeleteToolStripMenuItem.Text = "Delete"
        '
        'form_rft_phrases
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(532, 285)
        Me.Controls.Add(Me.split)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "form_rft_phrases"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Reporting phrases"
        Me.TopMost = True
        Me.split.Panel1.ResumeLayout(False)
        Me.split.Panel1.PerformLayout()
        Me.split.Panel2.ResumeLayout(False)
        Me.split.Panel2.PerformLayout()
        CType(Me.split, System.ComponentModel.ISupportInitialize).EndInit()
        Me.split.ResumeLayout(False)
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ContextMenuStrip_edititems.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents split As System.Windows.Forms.SplitContainer
    Friend WithEvents lstGroups As System.Windows.Forms.ListBox
    Friend WithEvents lstItems As System.Windows.Forms.ListBox
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbtn_clear As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbtn_Undo As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripLabel2 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tsbtnAutoReport As System.Windows.Forms.ToolStripButton
    Friend WithEvents ContextMenuStrip_edititems As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents NewItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
