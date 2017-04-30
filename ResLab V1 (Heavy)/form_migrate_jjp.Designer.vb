<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_migrate_jjp
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
        Me.btnMigrate = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.txtSourceConn = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtDestinationConn = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmboLab = New System.Windows.Forms.ComboBox()
        Me.cmboDBType_source = New System.Windows.Forms.ComboBox()
        Me.cmboDBType_destination = New System.Windows.Forms.ComboBox()
        Me.btnTestSource = New System.Windows.Forms.Button()
        Me.btnTestDestination = New System.Windows.Forms.Button()
        Me.btnMigrateFlowVols = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnMigrate
        '
        Me.btnMigrate.Location = New System.Drawing.Point(763, 477)
        Me.btnMigrate.Name = "btnMigrate"
        Me.btnMigrate.Size = New System.Drawing.Size(155, 29)
        Me.btnMigrate.TabIndex = 0
        Me.btnMigrate.Text = "Start prov migration"
        Me.btnMigrate.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(763, 512)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(155, 29)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'txtLog
        '
        Me.txtLog.Dock = System.Windows.Forms.DockStyle.Top
        Me.txtLog.Location = New System.Drawing.Point(0, 0)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLog.Size = New System.Drawing.Size(1135, 277)
        Me.txtLog.TabIndex = 2
        '
        'txtSourceConn
        '
        Me.txtSourceConn.Location = New System.Drawing.Point(271, 329)
        Me.txtSourceConn.Name = "txtSourceConn"
        Me.txtSourceConn.Size = New System.Drawing.Size(801, 22)
        Me.txtSourceConn.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 334)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Source connection"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 362)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Destination connection"
        '
        'txtDestinationConn
        '
        Me.txtDestinationConn.Location = New System.Drawing.Point(271, 357)
        Me.txtDestinationConn.Name = "txtDestinationConn"
        Me.txtDestinationConn.Size = New System.Drawing.Size(801, 22)
        Me.txtDestinationConn.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 301)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Lab"
        '
        'cmboLab
        '
        Me.cmboLab.FormattingEnabled = True
        Me.cmboLab.Items.AddRange(New Object() {"MRSS", "NRG", "John Hunter", "Monash", "Alfred", "Box Hill", "MyLapTop"})
        Me.cmboLab.Location = New System.Drawing.Point(171, 297)
        Me.cmboLab.Name = "cmboLab"
        Me.cmboLab.Size = New System.Drawing.Size(94, 24)
        Me.cmboLab.TabIndex = 8
        '
        'cmboDBType_source
        '
        Me.cmboDBType_source.FormattingEnabled = True
        Me.cmboDBType_source.Items.AddRange(New Object() {"MySQL", "SQLServer"})
        Me.cmboDBType_source.Location = New System.Drawing.Point(171, 327)
        Me.cmboDBType_source.Name = "cmboDBType_source"
        Me.cmboDBType_source.Size = New System.Drawing.Size(94, 24)
        Me.cmboDBType_source.TabIndex = 9
        '
        'cmboDBType_destination
        '
        Me.cmboDBType_destination.FormattingEnabled = True
        Me.cmboDBType_destination.Items.AddRange(New Object() {"MySQL", "SQLServer"})
        Me.cmboDBType_destination.Location = New System.Drawing.Point(171, 357)
        Me.cmboDBType_destination.Name = "cmboDBType_destination"
        Me.cmboDBType_destination.Size = New System.Drawing.Size(94, 24)
        Me.cmboDBType_destination.TabIndex = 10
        '
        'btnTestSource
        '
        Me.btnTestSource.Location = New System.Drawing.Point(1078, 327)
        Me.btnTestSource.Name = "btnTestSource"
        Me.btnTestSource.Size = New System.Drawing.Size(47, 23)
        Me.btnTestSource.TabIndex = 11
        Me.btnTestSource.Text = "Test"
        Me.btnTestSource.UseVisualStyleBackColor = True
        '
        'btnTestDestination
        '
        Me.btnTestDestination.Location = New System.Drawing.Point(1078, 356)
        Me.btnTestDestination.Name = "btnTestDestination"
        Me.btnTestDestination.Size = New System.Drawing.Size(47, 23)
        Me.btnTestDestination.TabIndex = 12
        Me.btnTestDestination.Text = "Test"
        Me.btnTestDestination.UseVisualStyleBackColor = True
        '
        'btnMigrateFlowVols
        '
        Me.btnMigrateFlowVols.Location = New System.Drawing.Point(763, 433)
        Me.btnMigrateFlowVols.Name = "btnMigrateFlowVols"
        Me.btnMigrateFlowVols.Size = New System.Drawing.Size(155, 29)
        Me.btnMigrateFlowVols.TabIndex = 13
        Me.btnMigrateFlowVols.Text = "Start flow vol migration"
        Me.btnMigrateFlowVols.UseVisualStyleBackColor = True
        '
        'form_migrate_jjp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1135, 672)
        Me.Controls.Add(Me.btnMigrateFlowVols)
        Me.Controls.Add(Me.btnTestDestination)
        Me.Controls.Add(Me.btnTestSource)
        Me.Controls.Add(Me.cmboDBType_destination)
        Me.Controls.Add(Me.cmboDBType_source)
        Me.Controls.Add(Me.cmboLab)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtDestinationConn)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSourceConn)
        Me.Controls.Add(Me.txtLog)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnMigrate)
        Me.Name = "form_migrate_jjp"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "form_migrate_jjp"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnMigrate As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents txtSourceConn As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDestinationConn As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmboLab As System.Windows.Forms.ComboBox
    Friend WithEvents cmboDBType_source As System.Windows.Forms.ComboBox
    Friend WithEvents cmboDBType_destination As System.Windows.Forms.ComboBox
    Friend WithEvents btnTestSource As System.Windows.Forms.Button
    Friend WithEvents btnTestDestination As System.Windows.Forms.Button
    Friend WithEvents btnMigrateFlowVols As System.Windows.Forms.Button
End Class
