Public Class form_password

    Private Sub txtPassword_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyUp

        Select Case e.KeyCode
            Case Keys.Enter
                If txtPassword.Text = My.Settings.Password_hidden Or txtPassword.Text = My.Settings.Password_userset Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                Else
                    MsgBox("Password is incorrect.", vbOKOnly, "Password")
                    txtPassword.Text = ""
                    txtPassword.Focus()
                End If

            Case Keys.Escape
                Me.DialogResult = Windows.Forms.DialogResult.None
                Me.Close()
        End Select
    End Sub

   
    Private Sub form_password_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        txtPassword_new1.Text = ""
        txtPassword_new2.Text = ""
        txtPassword_old.Text = ""
        txtPassword.Text = ""
        My.Settings.Reload()

    End Sub

    Private Sub tsbtn_resetpassword_Click(sender As System.Object, e As System.EventArgs) Handles tsbtn_resetpassword.Click

        txtPassword_new1.Text = ""
        txtPassword_new2.Text = ""
        txtPassword_old.Text = ""
        txtPassword.Text = ""
        pnlPassword.Visible = False
        pnlReset.Visible = True

    End Sub

    Private Sub txtPassword_new2_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtPassword_new1.KeyUp, txtPassword_new2.KeyUp

        If txtPassword_old.Text = My.Settings.Password_userset Then
            Select Case e.KeyCode
                Case Keys.Enter
                    If txtPassword_new1.Text = txtPassword_new2.Text Then
                        My.Settings.Password_userset = txtPassword_new2.Text
                        My.Settings.Save()
                        MsgBox("Password updated successfully.", vbOKOnly, "Password reset")
                        txtPassword_new1.Text = ""
                        txtPassword_new2.Text = ""
                        txtPassword_old.Text = ""
                        txtPassword.Text = ""
                        pnlPassword.Visible = True
                        pnlReset.Visible = False
                        Me.Close()
                    Else
                        MsgBox("New password copies do not match.", vbOKOnly, "Password")
                        txtPassword_new1.Text = ""
                        txtPassword_new2.Text = ""
                        txtPassword_new1.Focus()
                    End If

                Case Keys.Escape
                    Me.DialogResult = Windows.Forms.DialogResult.None
                    Me.Close()
            End Select
        Else
            MsgBox("Old password is incorrect.", vbOKOnly, "Password")
            txtPassword_old.Text = ""
            txtPassword_old.Focus()
        End If



    End Sub

  
    Private Sub tsbtn_Cancel_Click(sender As System.Object, e As System.EventArgs) Handles tsbtn_Cancel.Click

        pnlPassword.Visible = True
        pnlReset.Visible = False
        Me.Close()

    End Sub
End Class