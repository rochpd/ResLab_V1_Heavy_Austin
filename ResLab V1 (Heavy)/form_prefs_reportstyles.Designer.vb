<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_prefs_reportstyles
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lstStyles = New System.Windows.Forms.ListView()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lstStyles)
        Me.SplitContainer1.Size = New System.Drawing.Size(769, 475)
        Me.SplitContainer1.SplitterDistance = 256
        Me.SplitContainer1.TabIndex = 1
        '
        'lstStyles
        '
        Me.lstStyles.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lstStyles.Location = New System.Drawing.Point(0, 25)
        Me.lstStyles.Name = "lstStyles"
        Me.lstStyles.Size = New System.Drawing.Size(256, 450)
        Me.lstStyles.TabIndex = 1
        Me.lstStyles.UseCompatibleStateImageBehavior = False
        '
        'form_prefs_reportstyles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 475)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "form_prefs_reportstyles"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Preferences"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents lstStyles As System.Windows.Forms.ListView
End Class
