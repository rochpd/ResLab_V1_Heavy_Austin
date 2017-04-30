Public Class form_password_create

    Private Sub txtPassword_new2_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtPassword_new2.KeyUp

        Select Case e.KeyCode
            Case Keys.Enter
                btnPasswordOK.PerformClick()
            Case Keys.Escape
                
        End Select

    End Sub


    Private Sub btnCancelPassword_Click(sender As System.Object, e As System.EventArgs) Handles btnCancelPassword.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnPasswordOK_Click(sender As System.Object, e As System.EventArgs) Handles btnPasswordOK.Click

        If txtPassword_new1.Text = txtPassword_new2.Text Then
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        ElseIf txtPassword_new1.Text = "" Then
            MsgBox("Password field is empty.", vbOKOnly, "New password")
            txtPassword_new1.Focus()
        ElseIf txtPassword_new2.Text = "" Then
            MsgBox("Re-entered password field is empty.", vbOKOnly, "New password")
            txtPassword_new2.Focus()
        Else
            MsgBox("Passwords don't match.", vbOKOnly, "New password")
            txtPassword_new1.Text = ""
            txtPassword_new2.Text = ""
            txtPassword_new1.Focus()
        End If

    End Sub

    Private Sub form_password_create_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp

        If e.KeyCode = Keys.Escape Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If

    End Sub
End Class