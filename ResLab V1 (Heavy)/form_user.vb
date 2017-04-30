
Public Class form_user

    Private _Calling_Form As Form
    Private FormIsLoading As Boolean = False
    Private _IsNewPerson As Boolean = False
    Private _personID As Long = 0
    Private _cA As New class_authenticate


    Public Sub New(personID As Long, Calling_Form As Form)

        InitializeComponent()
        _Calling_Form = Calling_Form
        _personID = personID

        Select Case personID
            Case 0 : _IsNewPerson = True
            Case Else : _IsNewPerson = False
        End Select

    End Sub

    Private Sub form_user_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        'Load combos
        Me.Setup_ComboBoxes(_IsNewPerson)

        'Set edit mode
        btnCancel.Enabled = False
        btnEdit.Enabled = Not _IsNewPerson
        btnSave.Enabled = False
        btnClose.Enabled = True
        btnDelete.Enabled = False
        btnNew.Enabled = True
        btnFind.Enabled = True
        btnResetPassword.Enabled = False
        Me.set_textboxes_toReadonly(True)

        Select Case _IsNewPerson
            Case True
                Me.Text = "Data for <new user>"
                txtSurname.Focus()
            Case False
                Dim d As New class_DemographicFields
                Dim dicD As Dictionary(Of String, String) = cPt.Get_Demographics(Me.Tag)
                Me.Text = "Data for " & dicD(d.Surname) & ", " & dicD(d.Firstname)
                Load_ScreenFieldsFromMem(_IsNewPerson, _personID)
                btnClose.Focus()
        End Select

        Me.get_users()
        Me.setup_grid_permissions()

    End Sub

    Private Sub get_users()

        'Get all users
        cmboSearch_username.Items.Clear()

        Dim i As Integer = 0
        Dim users As Dictionary(Of Long, String) = cUser.Get_all_users(False)
        If Not IsNothing(users) Then
            For Each kv As KeyValuePair(Of Long, String) In users
                cmboSearch_username.Items.Add(kv)
            Next
            cmboSearch_username.DisplayMember = "value"
            cmboSearch_username.ValueMember = "key"
        End If

    End Sub

    Private Function grid_finditem(g As DataGridView, findstring As String, colname As String) As Integer
        'returns the first matching row index, otherwise nothing

        For Each row As DataGridViewRow In g.Rows
            If row.Cells(colname).Value = findstring Then
                Return row.Index
                Exit Function
            End If
        Next

        Return Nothing

    End Function

    Private Function Load_ScreenFieldsFromMem(IsNewPerson As Boolean, personID As Long) As Boolean


        Try
            Select Case IsNewPerson
                Case True


                    Return True
                Case False
                    'Do details
                    Dim u As Dictionary(Of String, String) = cUser.Get_user_details(personID)
                    If IsNothing(u) Then
                        'should never happen
                    Else
                        'Load the user data
                        Dim flds As New class_Fields_User
                        cmboTitle.Text = u(flds.Title) & ""
                        txtSurname.Text = u(flds.Surname) & ""
                        txtFirstname.Text = u(flds.Firstname) & ""
                        If IsDate(u(flds.DOB)) Then txtDOB.Text = Format(CDate(u(flds.DOB)), "dd/MM/yyyy") Else txtDOB.Text = Nothing
                        cmboGender.Text = u(flds.Gender) & ""
                        txtAddress1.Text = u(flds.Address_1) & ""
                        txtAddress2.Text = u(flds.Address_2) & ""
                        txtSuburb.Text = u(flds.Suburb) & ""
                        cmboState.Text = u(flds.State) & ""
                        cmboProfession.Text = u(flds.Profession_category) & ""
                        cmboDepartment.Text = u(flds.Department) & ""
                        cmboInstitution.Text = u(flds.Institution) & ""
                        txtPostcode.Text = u(flds.PostCode) & ""
                        txtHomePh.Text = u(flds.Phone_home) & ""
                        txtMobilePh.Text = u(flds.Phone_mobile) & ""
                        txtWorkPh.Text = u(flds.Phone_work) & ""
                        txtEmail.Text = u(flds.Email) & ""
                        txtUsername.Text = u(flds.User_name) & ""
                        If u(flds.User_name) <> "" Then
                            txtUserPassword.Text = Strings.StrDup(u(flds.User_name).Length, "*")
                            txtUserPassword.Tag = u(flds.User_password)
                        Else
                            txtUserPassword.Text = ""
                            txtUserPassword.Tag = ""
                        End If
                        tsLabel_lastlogin.Text = "Last login: " & u(flds.Last_login)
                    End If

                    'Do permissions
                    Dim p() As Dictionary(Of String, String) = cUser.get_user_permissions_dics(personID, True)
                    If IsNothing(p) Then
                        'should never happen
                        Return False
                    Else
                        'Load the user permissions
                        Dim row As Integer = 0
                        For Each d As Dictionary(Of String, String) In p
                            Select Case d("datatype")
                                Case "bit"
                                    row = Me.grid_finditem(grdP_chk, d("description"), "description")
                                    grdP_chk("value", row).Value = CBool(d("value"))
                                    grdP_chk("permissionid", row).Value = d("permissionid")
                                    grdP_chk("person_permissionid", row).Value = d("person_permissionid")
                                Case "text"
                                    row = Me.grid_finditem(grdP_comb, d("description"), "description")
                                    grdP_comb("value", row).Value = d("value")
                                    grdP_comb("permissionid", row).Value = d("permissionid")
                                    grdP_comb("person_permissionid", row).Value = d("person_permissionid")
                            End Select
                        Next
                    End If
            End Select
            Return True

        Catch ex As Exception
            MsgBox("Error in form_user.Load_ScreenFieldsFromMem" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Private Function Load_MemFromScreenFields_details(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicUser
        Dim flds As New class_Fields_User

        dicR(flds.personID) = Me._personID
        dicR(flds.Title) = q & cmboTitle.Text & q
        dicR(flds.Surname) = q & txtSurname.Text & q
        dicR(flds.Firstname) = q & txtFirstname.Text & q
        If IsDate(txtDOB.Text) Then dicR(flds.DOB) = q & CDate(txtDOB.Text).ToString("yyyy/MM/dd") & q Else dicR(flds.DOB) = "NULL"
        dicR(flds.Gender) = q & cmboGender.Text & q
        dicR(flds.Address_1) = q & txtAddress1.Text & q
        dicR(flds.Address_2) = q & txtAddress2.Text & q
        dicR(flds.Suburb) = q & txtSuburb.Text & q
        dicR(flds.State) = q & cmboState.Text & q
        dicR(flds.Profession_category) = q & cmboProfession.Text & q
        dicR(flds.Department) = q & cmboDepartment.Text & q
        dicR(flds.Institution) = q & cmboInstitution.Text & q
        dicR(flds.PostCode) = q & txtPostcode.Text & q
        dicR(flds.Phone_home) = q & txtHomePh.Text & q
        dicR(flds.Phone_mobile) = q & txtMobilePh.Text & q
        dicR(flds.Phone_work) = q & txtWorkPh.Text & q
        dicR(flds.Email) = q & txtEmail.Text & q
        dicR(flds.User_name) = q & txtUsername.Text & q
        dicR(flds.User_password) = q & txtUserPassword.Tag & q
        dicR(flds.Lastupdated) = cMyRoutines.FormatDBDateTime(Date.Now)

        Return dicR

    End Function

    Private Function Load_MemFromScreenFields_permissions(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)()

        Try

            Dim q As String = ""
            If AddQuotesAroundStrings Then q = "'"

            Dim i As Integer = 0
            Dim s As String = ""
            Dim d() As Dictionary(Of String, String) = Nothing

            For Each row As DataGridViewRow In grdP_chk.Rows
                ReDim Preserve d(i)
                d(i) = New Dictionary(Of String, String)
                s = class_user.ePermission_fields.person_permissionid.ToString : d(i).Add(s, row.Cells(s).Value)
                s = class_user.ePermission_fields.personID.ToString : d(i).Add(s, _personID)
                s = class_user.ePermission_fields.lastupdated.ToString : d(i).Add(s, cMyRoutines.FormatDBDateTime(Date.Now))
                s = class_user.ePermission_fields.permissionid.ToString : d(i).Add(s, row.Cells(s).Value)
                s = class_user.ePermission_fields.value.ToString : d(i).Add(s, q & CBool(row.Cells(s).Value) & q)
                i = i + 1
            Next

            For Each row In grdP_comb.Rows
                ReDim Preserve d(i)
                d(i) = New Dictionary(Of String, String)
                s = class_user.ePermission_fields.person_permissionid.ToString : d(i).Add(s, row.Cells(s).Value)
                s = class_user.ePermission_fields.personID.ToString : d(i).Add(s, _personID)
                s = class_user.ePermission_fields.lastupdated.ToString : d(i).Add(s, cMyRoutines.FormatDBDateTime(Date.Now))
                s = class_user.ePermission_fields.permissionid.ToString : d(i).Add(s, row.Cells(s).Value)
                s = class_user.ePermission_fields.value.ToString : d(i).Add(s, q & row.Cells(s).Value & q)
                i = i + 1
            Next

            Return d

        Catch ex As Exception
            MsgBox("Error in form_user.Load_MemFromScreenFields_permissions", vbOKOnly)
            Return Nothing
        End Try

    End Function

    Private Sub Setup_ComboBoxes(ByVal IsNewTest As Boolean)

        'Load combo box options
        cMyRoutines.Combo_LoadItems(cmboGender, cPred.Get_RefItems(RefItems.Genders))
        cmboGender.Items.RemoveAt(cmboGender.FindString("Male, Female"))

        cMyRoutines.Combo_LoadItemsFromPrefs(cmboProfession, cPrefs.get_fieldID_fromFieldName("Professional_groups"))
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboDepartment, cPrefs.get_fieldID_fromFieldName("Departments"))
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboInstitution, cPrefs.get_fieldID_fromFieldName("Institutions"))
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboState, cPrefs.get_fieldID_fromFieldName("States"))
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboTitle, cPrefs.get_fieldID_fromFieldName("Titles"))

    End Sub


    Private Sub clear_textboxes()

        'Clear details
        For Each c As Control In TabPage_details.Controls
            If TypeOf (c) Is TextBox Then
                CType(c, TextBox).Text = ""
            ElseIf TypeOf (c) Is ComboBox Then
                CType(c, ComboBox).Text = ""
            ElseIf TypeOf (c) Is MaskedTextBox Then
                CType(c, MaskedTextBox).Text = ""
            End If
        Next
        txtUsername.Text = ""
        txtUserPassword.Text = ""
        cmboSearch_username.Text = ""
        tsLabel_lastlogin.Text = ""

        'Clear permissions
        Me.setup_grid_permissions()

    End Sub

    Private Sub set_textboxes_toReadonly(readonlystatus As Boolean)

        For Each c As Control In TabPage_details.Controls
            If TypeOf (c) Is TextBox Then
                CType(c, TextBox).ReadOnly = readonlystatus
            ElseIf TypeOf (c) Is ComboBox Then
                CType(c, ComboBox).Enabled = Not readonlystatus
            ElseIf TypeOf (c) Is MaskedTextBox Then
                CType(c, MaskedTextBox).ReadOnly = readonlystatus
            End If
        Next


    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click

        'Set edit mode
        btnCancel.Enabled = True
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnClose.Enabled = False
        btnDelete.Enabled = False
        btnNew.Enabled = False
        btnFind.Enabled = False
        btnResetPassword.Enabled = True
        Me.set_textboxes_toReadonly(False)
        Me.clear_textboxes()
        cmboTitle.Focus()

        _IsNewPerson = True
        _personID = 0


    End Sub

    Private Sub btnEdit_Click(sender As Object, e As System.EventArgs) Handles btnEdit.Click

        'Set edit mode
        btnCancel.Enabled = True
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnClose.Enabled = False
        btnDelete.Enabled = False
        btnNew.Enabled = False
        btnFind.Enabled = False
        btnResetPassword.Enabled = True
        Me.set_textboxes_toReadonly(False)
        cmboTitle.Focus()
        cmboSearch_username.Enabled = False

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click

        'Set edit mode
        btnCancel.Enabled = False
        btnEdit.Enabled = Not _IsNewPerson
        btnSave.Enabled = False
        btnClose.Enabled = True
        btnDelete.Enabled = False
        btnNew.Enabled = True
        btnFind.Enabled = True
        btnResetPassword.Enabled = False
        Me.set_textboxes_toReadonly(True)
        cmboSearch_username.Focus()
        Me.clear_textboxes()
        Me.setup_grid_permissions()
        cmboSearch_username.Enabled = True

    End Sub

    Private Sub btnFind_Click(sender As Object, e As System.EventArgs) Handles btnFind.Click

        If cmboSearch_username.Text = "" Then Exit Sub

        Dim personID As Long = cUser.Find_user(cmboSearch_username.Text)
        If personID > 0 Then
            Me.Load_ScreenFieldsFromMem(False, personID)
            btnEdit.Enabled = True
            _IsNewPerson = False
            _personID = personID
        Else
            MsgBox("Username not found", vbOKOnly, "Find Username")
        End If

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        'Check for mandatory fields
        If Not IsDate(txtDOB.Text) Then
            MsgBox("Invalid DOB (DOB is a mandatory field).", vbOKOnly, "Data entry error")
            txtDOB.Focus()
            Exit Sub
        End If
        If cmboGender.Text = "" Then
            MsgBox("Invalid Gender (Gender is a mandatory field).", vbOKOnly, "Data entry error")
            cmboGender.Focus()
            Exit Sub
        End If
        If txtSurname.Text = "" Then
            MsgBox("Invalid surname (Surname is a mandatory field).", vbOKOnly, "Data entry error")
            txtSurname.Focus()
            Exit Sub
        End If
        If txtFirstname.Text = "" Then
            MsgBox("Invalid first name (First name is a mandatory field).", vbOKOnly, "Data entry error")
            txtFirstname.Focus()
            Exit Sub
        End If

        Dim user As Dictionary(Of String, String) = Load_MemFromScreenFields_details(True)
        Dim p() As Dictionary(Of String, String) = Load_MemFromScreenFields_permissions(True)

        Select Case _IsNewPerson
            Case True
                _personID = cUser.Insert_user_details(user)
                cUser.insert_user_permissions(_personID, p)
            Case False
                cUser.Update_user_details(user)
                cUser.update_user_permissions(p)              
        End Select

        'Set edit mode
        btnCancel.Enabled = False
        btnEdit.Enabled = Not _IsNewPerson
        btnSave.Enabled = False
        btnClose.Enabled = True
        btnDelete.Enabled = False
        btnNew.Enabled = True
        btnFind.Enabled = True
        btnResetPassword.Enabled = False
        Me.set_textboxes_toReadonly(True)
        cmboSearch_username.Enabled = True

        Me.get_users()

        'Refresh permissions for current user
        Dim currentuser As String = cUser.CurrentUser_name
        cUser = Nothing
        cUser = New class_user(currentuser)
        gRefreshMainForm = True

    End Sub

    Private Sub txtSearch_username_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            e.Handled = True
            btnFind.PerformClick()
        End If

    End Sub

    Private Sub btnResetPassword_Click(sender As System.Object, e As System.EventArgs) Handles btnResetPassword.Click

        Dim f As New form_password_create
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim pw_encrypted As String = _cA.encrypt_password(f.txtPassword_new1.Text)
            txtUserPassword.Tag = pw_encrypted
            txtUserPassword.Text = Strings.StrDup(f.txtPassword_new1.Text.Length, "*")
        End If

        f.Dispose()
        
    End Sub

    Private Sub setup_grid_permissions()

        Try

            Dim i As Integer = 0
            Dim wid As Integer = 45
            Dim c_comb, c_chk, c, c_copy As DataGridViewColumn
            Dim Align_mc As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleCenter
            Dim Align_ml As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft

            Dim header_text() As String = {"person_permissionID", "permissionid", "Description", "Description", "Value"}
            Dim col_name() As String = {"person_permissionid", "permissionid", "description", "displaytext", "value"}
            Dim col_width() As Integer = {wid, wid, wid, wid * 2, wid * 2}
            Dim col_visible() As Boolean = {False, False, False, False, True}
            Dim col_readonly() As Boolean = {True, True, True, True, False}
            Dim col_textalign() As DataGridViewContentAlignment = {Align_mc, Align_mc, Align_mc, Align_mc, Align_mc}

            grdP_chk.RowHeadersWidth = wid * 3
            grdP_comb.RowHeadersWidth = wid * 3
            grdP_chk.ColumnHeadersVisible = False
            grdP_comb.ColumnHeadersVisible = False

            'Set column headers
            Dim perms As List(Of class_user.permission_type) = cUser.get_permissions_types(True)
            grdP_chk.Columns.Clear()
            grdP_comb.Columns.Clear()

            For i = 0 To UBound(header_text)
                Select Case col_name(i)
                    Case "value"
                        'Checkbox perm grid for settings
                        c_chk = New DataGridViewCheckBoxColumn
                        c_chk.CellTemplate = New DataGridViewCheckBoxCell
                        c_chk.DefaultCellStyle.BackColor = Color.White
                        c_chk.DefaultCellStyle.SelectionBackColor = Color.White
                        c_chk.DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Regular)
                        c_chk.HeaderText = header_text(i)
                        c_chk.HeaderCell.Style.Alignment = col_textalign(i)
                        c_chk.HeaderCell.Style.BackColor = SystemColors.Control
                        c_chk.HeaderCell.Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                        c_chk.Name = col_name(i)
                        c_chk.Width = col_width(i)
                        c_chk.Visible = col_visible(i)
                        c_chk.ReadOnly = col_readonly(i)
                        c_chk.HeaderCell.Style.SelectionBackColor = c_chk.HeaderCell.Style.BackColor
                        Me.grdP_chk.Columns.Add(c_chk)

                        'Combobox perm grid for settings
                        c_comb = New DataGridViewComboBoxColumn
                        c_comb.CellTemplate = New DataGridViewComboBoxCell
                        c_comb.DefaultCellStyle.BackColor = Color.White
                        c_comb.DefaultCellStyle.SelectionBackColor = Color.White
                        c_comb.DefaultCellStyle.Font = New Font("Segoe UI", 8, FontStyle.Regular)
                        c_comb.HeaderText = header_text(i)
                        c_comb.HeaderCell.Style.Alignment = col_textalign(i)
                        c_comb.HeaderCell.Style.BackColor = SystemColors.Control
                        c_comb.HeaderCell.Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                        c_comb.Name = col_name(i)
                        c_comb.Width = col_width(i)
                        c_comb.Visible = col_visible(i)
                        c_comb.ReadOnly = col_readonly(i)
                        c_comb.HeaderCell.Style.SelectionBackColor = c_comb.HeaderCell.Style.BackColor
                        Me.grdP_comb.Columns.Add(c_comb)

                    Case Else
                        c = New DataGridViewTextBoxColumn
                        c.CellTemplate = New DataGridViewTextBoxCell
                        c.DefaultCellStyle.BackColor = Color.White
                        c.HeaderText = header_text(i)
                        c.HeaderCell.Style.Alignment = col_textalign(i)
                        c.HeaderCell.Style.BackColor = SystemColors.Control
                        c.HeaderCell.Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
                        c.Name = col_name(i)
                        c.Width = col_width(i)
                        c.Visible = col_visible(i)
                        c.ReadOnly = col_readonly(i)
                        c.HeaderCell.Style.SelectionBackColor = c.HeaderCell.Style.BackColor
                        Me.grdP_comb.Columns.Add(c)
                        c_copy = c.Clone
                        c_copy.CellTemplate = New DataGridViewTextBoxCell
                        Me.grdP_chk.Columns.Add(c_copy)

                End Select
            Next

            'Add the permission type rows
            Dim r As DataGridViewRow = Nothing
            For Each p In perms
                Select Case p.datatype
                    Case "text"
                        grdP_comb.Rows.Add()
                        r = grdP_comb.Rows(grdP_comb.Rows.Count - 1)
                        CType(grdP_comb("value", grdP_comb.Rows.Count - 1), DataGridViewComboBoxCell).Items.AddRange(cUser.get_permissions_listofvalues("list_accesstypes", "description", True))
                        CType(grdP_comb("value", grdP_comb.Rows.Count - 1), DataGridViewComboBoxCell).Style.BackColor = Color.White
                        CType(grdP_comb("value", grdP_comb.Rows.Count - 1), DataGridViewComboBoxCell).Value = ""
                    Case "bit"
                        grdP_chk.Rows.Add()
                        r = grdP_chk.Rows(grdP_chk.Rows.Count - 1)
                End Select

                r.Cells("person_permissionID").Value = "0"
                r.Cells("permissionid").Value = p.permissionid
                r.Cells("description").Value = p.description
                r.Cells("displaytext").Value = p.displaytext
                r.HeaderCell.Value = p.displaytext

            Next

        Catch ex As Exception
            MsgBox("Error in form_user.setup_grid_permissions" & vbCrLf & Err.Description)
        End Try

    End Sub






    Private Sub grdP_comb_DataError(sender As Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles grdP_comb.DataError
        MsgBox("grid error combo" & vbCrLf & Err.Description)
    End Sub

    Private Sub cmboSearch_username_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles cmboSearch_username.KeyPress

        Select Case e.KeyChar
            Case ChrW(Keys.Return) : btnFind.PerformClick()
            Case Else
        End Select

    End Sub

    Private Sub cmboSearch_username_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmboSearch_username.SelectedIndexChanged

        btnFind.PerformClick()

    End Sub

    Private Sub txtSurname_LostFocus(sender As Object, e As System.EventArgs) Handles txtSurname.LostFocus
        txtSurname.Text = txtSurname.Text.ToUpper
    End Sub

    Private Sub txtFirstname_LostFocus(sender As Object, e As System.EventArgs) Handles txtFirstname.LostFocus
        txtFirstname.Text = Strings.Left(txtFirstname.Text.ToUpper, 1) & Mid(txtFirstname.Text.ToLower, 2)
    End Sub

    Private Sub cmboGender_DropDownClosed(sender As Object, e As System.EventArgs) Handles cmboGender.DropDownClosed
        txtDOB.Focus()
    End Sub

    Private Sub cmboTitle_DropDownClosed(sender As Object, e As System.EventArgs) Handles cmboTitle.DropDownClosed
        txtSurname.Focus()
    End Sub

    Private Sub cmboState_HandleCreated(sender As Object, e As System.EventArgs) Handles cmboState.HandleCreated
        txtHomePh.Focus()
    End Sub

    Private Sub cmboInstitution_DropDownClosed(sender As Object, e As System.EventArgs) Handles cmboInstitution.DropDownClosed
        txtAddress1.Focus()
    End Sub

    Private Sub cmboProfession_DropDownClosed(sender As Object, e As System.EventArgs) Handles cmboProfession.DropDownClosed
        txtUsername.Focus()
    End Sub
End Class