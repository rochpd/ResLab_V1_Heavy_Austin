<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_password_create
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
        Me.pnlReset = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPassword_new2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPassword_new1 = New System.Windows.Forms.TextBox()
        Me.btnPasswordOK = New System.Windows.Forms.Button()
        Me.btnCancelPassword = New System.Windows.Forms.Button()
        Me.pnlReset.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlReset
        '
        Me.pnlReset.Controls.Add(Me.Label3)
        Me.pnlReset.Controls.Add(Me.txtPassword_new2)
        Me.pnlReset.Controls.Add(Me.Label2)
        Me.pnlReset.Controls.Add(Me.txtPassword_new1)
        Me.pnlReset.Location = New System.Drawing.Point(2, 25)
        Me.pnlReset.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlReset.Name = "pnlReset"
        Me.pnlReset.Size = New System.Drawing.Size(251, 69)
        Me.pnlReset.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(19, 38)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Re-enter password"
        '
        'txtPassword_new2
        '
        Me.txtPassword_new2.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower
        Me.txtPassword_new2.Location = New System.Drawing.Point(136, 34)
        Me.txtPassword_new2.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPassword_new2.Name = "txtPassword_new2"
        Me.txtPassword_new2.Size = New System.Drawing.Size(101, 20)
        Me.txtPassword_new2.TabIndex = 2
        Me.txtPassword_new2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPassword_new2.UseSystemPasswordChar = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(19, 14)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Enter a password"
        '
        'txtPassword_new1
        '
        Me.txtPassword_new1.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower
        Me.txtPassword_new1.Location = New System.Drawing.Point(136, 10)
        Me.txtPassword_new1.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPassword_new1.Name = "txtPassword_new1"
        Me.txtPassword_new1.Size = New System.Drawing.Size(101, 20)
        Me.txtPassword_new1.TabIndex = 1
        Me.txtPassword_new1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtPassword_new1.UseSystemPasswordChar = True
        '
        'btnPasswordOK
        '
        Me.btnPasswordOK.AutoSize = True
        Me.btnPasswordOK.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnPasswordOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPasswordOK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPasswordOK.Location = New System.Drawing.Point(75, 109)
        Me.btnPasswordOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnPasswordOK.Name = "btnPasswordOK"
        Me.btnPasswordOK.Size = New System.Drawing.Size(59, 24)
        Me.btnPasswordOK.TabIndex = 3
        Me.btnPasswordOK.Text = "OK"
        Me.btnPasswordOK.UseVisualStyleBackColor = False
        '
        'btnCancelPassword
        '
        Me.btnCancelPassword.AutoSize = True
        Me.btnCancelPassword.BackColor = System.Drawing.Color.LightSteelBlue
        Me.btnCancelPassword.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnCancelPassword.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancelPassword.Location = New System.Drawing.Point(137, 109)
        Me.btnCancelPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancelPassword.Name = "btnCancelPassword"
        Me.btnCancelPassword.Size = New System.Drawing.Size(59, 24)
        Me.btnCancelPassword.TabIndex = 4
        Me.btnCancelPassword.Text = "Cancel"
        Me.btnCancelPassword.UseVisualStyleBackColor = False
        '
        'form_password_create
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(270, 145)
        Me.Controls.Add(Me.btnPasswordOK)
        Me.Controls.Add(Me.btnCancelPassword)
        Me.Controls.Add(Me.pnlReset)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "form_password_create"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New password"
        Me.pnlReset.ResumeLayout(False)
        Me.pnlReset.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pnlReset As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPassword_new2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtPassword_new1 As System.Windows.Forms.TextBox
    Friend WithEvents btnPasswordOK As System.Windows.Forms.Button
    Friend WithEvents btnCancelPassword As System.Windows.Forms.Button
End Class
