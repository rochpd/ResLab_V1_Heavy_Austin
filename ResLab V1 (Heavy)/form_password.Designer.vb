<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_password
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
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.tsbtn_resetpassword = New System.Windows.Forms.ToolStripButton()
        Me.pnlReset = New System.Windows.Forms.Panel()
        Me.pnlPassword = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPassword_new2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPassword_new1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPassword_old = New System.Windows.Forms.TextBox()
        Me.tsbtn_Cancel = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip1.SuspendLayout()
        Me.pnlReset.SuspendLayout()
        Me.pnlPassword.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbtn_resetpassword, Me.tsbtn_Cancel})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 150)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(384, 26)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsbtn_resetpassword
        '
        Me.tsbtn_resetpassword.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtn_resetpassword.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.redo_32x32
        Me.tsbtn_resetpassword.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtn_resetpassword.Name = "tsbtn_resetpassword"
        Me.tsbtn_resetpassword.Size = New System.Drawing.Size(124, 23)
        Me.tsbtn_resetpassword.Text = "Reset password"
        '
        'pnlReset
        '
        Me.pnlReset.Controls.Add(Me.Label3)
        Me.pnlReset.Controls.Add(Me.txtPassword_new2)
        Me.pnlReset.Controls.Add(Me.Label2)
        Me.pnlReset.Controls.Add(Me.txtPassword_new1)
        Me.pnlReset.Controls.Add(Me.Label1)
        Me.pnlReset.Controls.Add(Me.txtPassword_old)
        Me.pnlReset.Location = New System.Drawing.Point(16, 22)
        Me.pnlReset.Name = "pnlReset"
        Me.pnlReset.Size = New System.Drawing.Size(348, 110)
        Me.pnlReset.TabIndex = 7
        Me.pnlReset.Visible = False
        '
        'pnlPassword
        '
        Me.pnlPassword.Controls.Add(Me.Label6)
        Me.pnlPassword.Controls.Add(Me.txtPassword)
        Me.pnlPassword.Location = New System.Drawing.Point(88, 12)
        Me.pnlPassword.Name = "pnlPassword"
        Me.pnlPassword.Size = New System.Drawing.Size(208, 135)
        Me.pnlPassword.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(60, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(103, 19)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Enter password"
        '
        'txtPassword
        '
        Me.txtPassword.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower
        Me.txtPassword.Location = New System.Drawing.Point(43, 46)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(133, 22)
        Me.txtPassword.TabIndex = 7
        Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(25, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(153, 19)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Re-enter new password"
        '
        'txtPassword_new2
        '
        Me.txtPassword_new2.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower
        Me.txtPassword_new2.Location = New System.Drawing.Point(191, 75)
        Me.txtPassword_new2.Name = "txtPassword_new2"
        Me.txtPassword_new2.Size = New System.Drawing.Size(133, 22)
        Me.txtPassword_new2.TabIndex = 11
        Me.txtPassword_new2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPassword_new2.UseSystemPasswordChar = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(25, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 19)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "New password"
        '
        'txtPassword_new1
        '
        Me.txtPassword_new1.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower
        Me.txtPassword_new1.Location = New System.Drawing.Point(191, 45)
        Me.txtPassword_new1.Name = "txtPassword_new1"
        Me.txtPassword_new1.Size = New System.Drawing.Size(133, 22)
        Me.txtPassword_new1.TabIndex = 9
        Me.txtPassword_new1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPassword_new1.UseSystemPasswordChar = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(25, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 19)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Old password"
        '
        'txtPassword_old
        '
        Me.txtPassword_old.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower
        Me.txtPassword_old.Location = New System.Drawing.Point(191, 15)
        Me.txtPassword_old.Name = "txtPassword_old"
        Me.txtPassword_old.Size = New System.Drawing.Size(133, 22)
        Me.txtPassword_old.TabIndex = 7
        Me.txtPassword_old.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPassword_old.UseSystemPasswordChar = True
        '
        'tsbtn_Cancel
        '
        Me.tsbtn_Cancel.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtn_Cancel.Image = Global.ResLab_V1_Heavy.My.Resources.Resources.cancel_32x32
        Me.tsbtn_Cancel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtn_Cancel.Name = "tsbtn_Cancel"
        Me.tsbtn_Cancel.Size = New System.Drawing.Size(69, 23)
        Me.tsbtn_Cancel.Text = "Cancel"
        '
        'form_password
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 176)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlPassword)
        Me.Controls.Add(Me.pnlReset)
        Me.Controls.Add(Me.ToolStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "form_password"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Password"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.pnlReset.ResumeLayout(False)
        Me.pnlReset.PerformLayout()
        Me.pnlPassword.ResumeLayout(False)
        Me.pnlPassword.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbtn_resetpassword As System.Windows.Forms.ToolStripButton
    Friend WithEvents pnlReset As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPassword_new2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPassword_new1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPassword_old As System.Windows.Forms.TextBox
    Friend WithEvents pnlPassword As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents tsbtn_Cancel As System.Windows.Forms.ToolStripButton
End Class
