Imports System.Security.Policy
Imports System.Text

Public Class form_login

    Private Sub form_login_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Environment.UserName = "rochpd" Then
            txtUsername.Text = "rochpd"
            txtPassword.Text = "beer"
        End If
        btnOKsalt.Focus()

    End Sub

    Private Sub btnOKsalt_Click(sender As System.Object, e As System.EventArgs) Handles btnOKsalt.Click

        Dim cA As New class_authenticate

        If txtUsername.Text = "" Then
            MsgBox("Please enter your user name", vbOKOnly, "Login")
            txtUsername.Focus()
        Else
            If txtPassword.Text = "" Then
                MsgBox("Please enter your password", vbOKOnly, "Login")
                txtPassword.Focus()
            Else
                If cmboLocation.Text = "" Then
                    MsgBox("Please select your location", vbOKOnly, "Login")
                    cmboLocation.Focus()
                Else
                    Select Case cA.authenticate_login(txtUsername.Text, txtPassword.Text)
                        Case True
                            cUser = New class_user(txtUsername.Text)
                            cHs.Current_UserLocation_HSID_code = cHs.get_healthservice_HSID_fromName(cmboLocation.Text)
                            Form_MainNew.Show()
                            Me.Visible = False
                            txtUsername.Text = ""
                            txtPassword.Text = ""
                            cA = Nothing
                        Case False
                            MsgBox("Username or password is incorrect.", vbOKOnly, "Login")
                            txtUsername.SelectAll()
                            txtUsername.Focus()
                    End Select
                End If
            End If
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
        End
    End Sub

    Private Sub txtUsername_LostFocus(sender As Object, e As System.EventArgs) Handles txtUsername.LostFocus

        'Check that username is valid
        Dim sql As String = "SELECT count(personid) AS count_of_user FROM persons WHERE user_name='" & txtUsername.Text & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows(0)("count_of_user") = 1 Then
            'Set locations combo options to user's default
            Dim cUser As New class_user(txtUsername.Text)
            cmboLocation.Items.AddRange(cHs.get_list_Healthservices(, , True).ToArray)
            cmboLocation.SelectedIndex = cmboLocation.FindString(cHs.get_healthservice_name_fromHSID(cUser.CurrentUser_login_healthservice_code_default))
        Else
            MsgBox("Invalid user name", vbOKOnly, "ResLab")
            txtUsername.Focus()
        End If



        

    End Sub

End Class
