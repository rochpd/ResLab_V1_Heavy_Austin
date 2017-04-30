<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_UpdateReportStatus
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
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.cmboUpdate = New System.Windows.Forms.ComboBox()
        Me.Label82 = New System.Windows.Forms.Label()
        Me.Label80 = New System.Windows.Forms.Label()
        Me.lblCurrentReportStatus = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.PowderBlue
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.Color.Blue
        Me.btnCancel.Location = New System.Drawing.Point(271, 86)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(85, 28)
        Me.btnCancel.TabIndex = 82
        Me.btnCancel.Text = "Don't update"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnUpdate
        '
        Me.btnUpdate.BackColor = System.Drawing.Color.PowderBlue
        Me.btnUpdate.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnUpdate.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.ForeColor = System.Drawing.Color.Blue
        Me.btnUpdate.Location = New System.Drawing.Point(185, 86)
        Me.btnUpdate.Margin = New System.Windows.Forms.Padding(2)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(85, 28)
        Me.btnUpdate.TabIndex = 83
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        'cmboUpdate
        '
        Me.cmboUpdate.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboUpdate.FormattingEnabled = True
        Me.cmboUpdate.Location = New System.Drawing.Point(155, 46)
        Me.cmboUpdate.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboUpdate.Name = "cmboUpdate"
        Me.cmboUpdate.Size = New System.Drawing.Size(198, 23)
        Me.cmboUpdate.TabIndex = 87
        '
        'Label82
        '
        Me.Label82.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label82.Location = New System.Drawing.Point(23, 49)
        Me.Label82.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(120, 18)
        Me.Label82.TabIndex = 85
        Me.Label82.Text = "Update report status to"
        '
        'Label80
        '
        Me.Label80.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label80.Location = New System.Drawing.Point(23, 24)
        Me.Label80.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label80.Name = "Label80"
        Me.Label80.Size = New System.Drawing.Size(108, 18)
        Me.Label80.TabIndex = 84
        Me.Label80.Text = "Current report status"
        '
        'lblCurrentReportStatus
        '
        Me.lblCurrentReportStatus.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCurrentReportStatus.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentReportStatus.Location = New System.Drawing.Point(155, 19)
        Me.lblCurrentReportStatus.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblCurrentReportStatus.Name = "lblCurrentReportStatus"
        Me.lblCurrentReportStatus.Size = New System.Drawing.Size(198, 24)
        Me.lblCurrentReportStatus.TabIndex = 88
        Me.lblCurrentReportStatus.Text = "Current report status"
        Me.lblCurrentReportStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'form_UpdateReportStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 125)
        Me.Controls.Add(Me.lblCurrentReportStatus)
        Me.Controls.Add(Me.cmboUpdate)
        Me.Controls.Add(Me.Label82)
        Me.Controls.Add(Me.Label80)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "form_UpdateReportStatus"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update report status"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents cmboUpdate As System.Windows.Forms.ComboBox
    Friend WithEvents Label82 As System.Windows.Forms.Label
    Friend WithEvents Label80 As System.Windows.Forms.Label
    Friend WithEvents lblCurrentReportStatus As System.Windows.Forms.Label
End Class
