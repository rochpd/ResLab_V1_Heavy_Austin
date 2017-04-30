Imports Microsoft.VisualBasic.Strings
Imports System.Text

Public Class form_Demographics

    Private _Calling_Form As Form
    Public cPtData As New class_pas_rfts
    Dim FormIsLoading As Boolean = False

    Public Sub New(Calling_Form As Form)

        InitializeComponent()
        _Calling_Form = Calling_Form

        If _Calling_Form.Name.ToLower = "form_find" Then
            btnSelect.Visible = True
            btnClose.Visible = True
            btnSave.Visible = False

        Else
            btnSelect.Visible = False
            btnClose.Visible = True
            btnSave.Visible = True
        End If

    End Sub

    Private Sub form_Demographics_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If FormIsLoading Then
            Select Case Me.Tag
                Case 0
                    Me.Text = "Demographic data for <new patient>"
                    txtSurname.Focus()

                Case Else
                    Dim d As New class_DemographicFields
                    Dim dicD As Dictionary(Of String, String) = cPt.get_demographics(Me.Tag)
                    Me.Text = "Demographic data for " & dicD(d.Surname) & ", " & dicD(d.Firstname)
                    Load_ScreenFieldsFromMem()
                    btnSelect.Focus()   'convenient default
            End Select
            FormIsLoading = False
        End If

    End Sub

    Private Sub form_Demographics_Deactivate(sender As Object, e As System.EventArgs) Handles Me.Deactivate
        gRefreshMainForm = True
    End Sub

    Private Sub form_Demographics_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        FormIsLoading = True

        Me.configure_for_demosource()

        cUser.set_access(Me)

    End Sub

    Private Sub configure_for_demosource()

        If cAppConfig.pas_mode_local Then ToolStripStatusLabel_source.Text = "Demographics source: Local" Else ToolStripStatusLabel_source.Text = "Demographics source: External PAS"
        'pnlDemo.Enabled = cAppConfig.pas_mode_local

        cmboTitle.Enabled = cAppConfig.pas_mode_local
        txtSurname.ReadOnly = Not cAppConfig.pas_mode_local
        txtFirstname.ReadOnly = Not cAppConfig.pas_mode_local
        cmboGender.Enabled = cAppConfig.pas_mode_local
        txtDOB.ReadOnly = Not cAppConfig.pas_mode_local
        cmboCountryOfBirth.Enabled = cAppConfig.pas_mode_local
        cmboRaceForRfts.Enabled = True
        txtAddress1.ReadOnly = Not cAppConfig.pas_mode_local
        txtAddress2.ReadOnly = Not cAppConfig.pas_mode_local
        txtSuburb.ReadOnly = Not cAppConfig.pas_mode_local
        txtPostcode.ReadOnly = Not cAppConfig.pas_mode_local
        txtEmail.ReadOnly = Not cAppConfig.pas_mode_local
        txtDeathStatus.ReadOnly = Not cAppConfig.pas_mode_local
        txtHomePh.ReadOnly = Not cAppConfig.pas_mode_local
        txtMobilePh.ReadOnly = Not cAppConfig.pas_mode_local
        txtWorkPh.ReadOnly = Not cAppConfig.pas_mode_local
        txtMedicareNo.ReadOnly = Not cAppConfig.pas_mode_local
        txtMedicareExpires.ReadOnly = Not cAppConfig.pas_mode_local
        lvAllURs.Enabled = cAppConfig.pas_mode_local

        cMyRoutines.Combo_LoadItems(cmboGender, cPred.Get_RefItems(RefItems.Genders))
        cmboGender.Items.RemoveAt(cmboGender.FindString("Male, Female"))
        cMyRoutines.Combo_LoadItems(cmboCountryOfBirth, cMyRoutines.get_items_listTable(cDatabaseInfo.eTables.list_Nationality))
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboTitle, cPrefs.get_fieldID_fromFieldName("Titles"), True)
        cMyRoutines.Combo_LoadItemsFromList(cmboRaceForRfts, cDatabaseInfo.eTables.List_RacesForRFTs, True)

    End Sub

    Private Sub Load_ScreenFieldsFromMem()

        Dim dicD As Dictionary(Of String, String) = cPt.Get_Demographics(Convert.ToInt16(Me.Tag))
        If IsNothing(dicD) Then
            'should never happen
        Else
            Dim d As New class_DemographicFields

            cmboTitle.SelectedIndex = cmboTitle.FindStringExact(Trim(dicD(d.Title)))
            txtSurname.Text = dicD(d.Surname)
            txtFirstname.Text = dicD(d.Firstname)

            If cMyRoutines.IsRealDate(dicD(d.DOB)) Then txtDOB.Text = Format(CDate(dicD(d.DOB)), "dd/MM/yyyy") Else txtDOB.Text = "__/__/____"
            cmboGender.SelectedIndex = cmboGender.FindStringExact(Trim(dicD(d.Gender)))
            cmboCountryOfBirth.Text = dicD(d.CountryOfBirth)
            txtAddress1.Text = dicD(d.Address_1)
            txtAddress2.Text = dicD(d.Address_2)
            txtSuburb.Text = dicD(d.Suburb)
            txtPostcode.Text = dicD(d.PostCode)
            txtHomePh.Text = dicD(d.Phone_home)
            txtMobilePh.Text = dicD(d.Phone_mobile)
            txtWorkPh.Text = dicD(d.Phone_work)
            txtEmail.Text = dicD(d.Email)
            txtMedicareNo.Text = dicD(d.Medicare_No)
            txtMedicareExpires.Text = dicD(d.Medicare_expirydate)
            cmboCountryOfBirth.Text = dicD(d.Race)

            'Do all URs
            Dim f As New class_fields_ur
            Dim urs = cPt.get_ur_all(dicD(d.PatientID))
            Dim lvi As ListViewItem
            lvAllURs.Clear()
            Me.Setup_lvAllURs()
            For Each ur As Dictionary(Of String, String) In urs
                lvi = New ListViewItem
                lvi.SubItems(0).Text = (ur(f.UR_id))
                lvi.SubItems.Add(ur(f.UR_hsid))
                lvi.SubItems.Add(ur(f.UR))
                lvi.SubItems.Add(cHs.get_healthservice_name_fromHSID(ur(f.UR_hsid)))
                lvi.SubItems.Add(ur(f.UR_mergeto))

                lvAllURs.Items.Add(lvi)
            Next

            dicD = Nothing
            d = Nothing
        End If
    End Sub


    Private Function Load_MemFromScreenFields() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)
        D = cMyRoutines.MakeEmpty_dicDemo()
        Dim c = New class_DemographicFields

        Select Case cAppConfig.pas_mode_local
            Case True
                'URs

                'Demographics 
                D(c.PatientID) = Me.Tag
                D(c.Title) = "'" & cmboTitle.Text & "'"
                D(c.Surname) = "'" & txtSurname.Text & "'"
                D(c.Firstname) = "'" & txtFirstname.Text & "'"
                If IsDate(txtDOB.Text) Then D(c.DOB) = "'" & CDate(txtDOB.Text).ToString("yyyy/MM/dd") & "'" Else D(c.DOB) = "NULL"
                D(c.Gender_code) = "'" & cMyRoutines.Lookup_list_ByDescription(cmboGender.Text, cDatabaseInfo.eTables.Pred_ref_genders) & "'"

                D(c.Address_1) = "'" & txtAddress1.Text & "'"
                D(c.Address_2) = "'" & txtAddress2.Text & "'"
                D(c.Suburb) = "'" & txtSuburb.Text & "'"
                D(c.PostCode) = "'" & txtPostcode.Text & "'"

                D(c.Phone_work) = "'" & txtWorkPh.Text & "'"
                D(c.Phone_mobile) = "'" & txtMobilePh.Text & "'"
                D(c.Phone_home) = "'" & txtHomePh.Text & "'"
                D(c.Email) = "'" & txtEmail.Text & "'"

                D(c.Medicare_No) = "'" & txtMedicareNo.Text & "'"
                D(c.Medicare_expirydate) = "'" & txtMedicareExpires.Text & "'"

                D(c.Race_code) = "'" & cMyRoutines.Lookup_list_ByDescription(cmboCountryOfBirth.Text, cDatabaseInfo.eTables.list_Nationality) & "'"
                D(c.Race_forRfts_code) = "'" & cMyRoutines.Lookup_list_ByDescription(cmboRaceForRfts.Text, cDatabaseInfo.eTables.Pred_ref_ethnicities) & "'"
                D(c.AboriginalStatus_code) = "'" & cMyRoutines.Lookup_list_ByDescription(cmboAboriginalStatus.Text, cDatabaseInfo.eTables.list_AboriginalStatus) & "'"

                D(c.Death_indicator) = "'" & txtDeathStatus.Text & "'"
                D(c.Death_date) = "'" & txtDeathDate.Text & "'"     'text field




            Case False



        End Select


        
        

        

        Return D

    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim PatientID As Long = 0

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

        'Check for invalid chars
        If InStr(txtSurname.Text, """") > 0 Then
            MsgBox("Quotation mark is an invalid character", vbOKOnly, "Data entry error")
            txtSurname.Focus()
            Exit Sub
        End If
        If InStr(txtFirstname.Text, """") > 0 Then
            MsgBox("Quotation mark is an invalid character", vbOKOnly, "Data entry error")
            txtFirstname.Focus()
            Exit Sub
        End If
       

        If Me.Tag = 0 Then
            'Check for possible matches already in database
            Dim IDs() As Long = cMyRoutines.Check_ForDuplicatePatient(txtSurname.Text, txtFirstname.Text, txtDOB.Text, cmboGender.Text)
            If IsNothing(IDs) Then   'no matches
                PatientID = cPt.insert_Demographics(Load_MemFromScreenFields())
                cPt.PatientID = PatientID
                gRefreshMainForm = True
                Me.Close()
            Else
                Dim f As New form_DuplicatePatients(IDs, Me)
                Dim Response = f.ShowDialog()
                
                Select Case CLng(Me.Tag)
                    Case Is > 0 'Use selected patient - ID has been copied to form tag
                        Load_ScreenFieldsFromMem()
                        cPt.PatientID = CLng(Me.Tag)
                        gRefreshMainForm = True
                        Me.Close()
                    Case 0  'Create new patient
                        PatientID = cPt.insert_Demographics(Load_MemFromScreenFields())
                        cPt.PatientID = PatientID
                        gRefreshMainForm = True
                        Me.Close()
                End Select
            End If
        Else
            Dim ReturnValue As Boolean = cPt.Update_Demographics(Load_MemFromScreenFields())
            cPt.PatientID = CLng(Me.Tag)
            gRefreshMainForm = True
            Me.Close()
        End If
       
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    Private Sub btnSelect_Click(sender As Object, e As System.EventArgs) Handles btnSelect.Click

        'Log this patient
        cPt.PatientID = CLng(Me.Tag)
        gRefreshMainForm = True
        If _Calling_Form.Name = "form_Find" Then _Calling_Form.Tag = "Close"
        Me.Close()

    End Sub

    Private Sub txtSurname_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurname.LostFocus

        txtSurname.Text = txtSurname.Text.ToUpper

    End Sub

    Private Sub txtFirstname_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFirstname.LostFocus

        txtFirstname.Text = StrConv(txtFirstname.Text, VbStrConv.ProperCase)

    End Sub

    Private Sub txtAddress1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddress1.LostFocus

        txtAddress1.Text = StrConv(txtAddress1.Text, VbStrConv.ProperCase)

    End Sub

    Private Sub txtAddress2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddress2.LostFocus

        txtAddress2.Text = StrConv(txtAddress2.Text, VbStrConv.ProperCase)

    End Sub

    Private Sub txtSuburb_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuburb.LostFocus

        txtSuburb.Text = StrConv(txtSuburb.Text, VbStrConv.ProperCase)

    End Sub

    Private Sub txtMedicareNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMedicareNo.LostFocus

        Dim ErrorVar As String = ""

        If Not cMyRoutines.Validate_MedicareNumber(txtMedicareNo.Text, ErrorVar) Then
            MsgBox("Invalid medicare number format - " & vbCrLf & ErrorVar, vbOKOnly, "Data entry error")
            txtMedicareNo.Focus()
        End If

    End Sub

    Private Sub SetEnabled(ByVal IsEditable As Boolean)

        Dim ctl

        For Each c As Control In Me.Controls
            If TypeOf c Is TextBox Then
                ctl = CType(c, TextBox)
                ctl.ReadOnly = Not IsEditable
            ElseIf TypeOf c Is MaskedTextBox Then
                ctl = CType(c, MaskedTextBox)
                ctl.ReadOnly = Not IsEditable
            ElseIf TypeOf c Is ComboBox Then
                ctl = CType(c, ComboBox)
                ctl.enabled = IsEditable
            End If
        Next

    End Sub


    Private Sub txtDOB_LostFocus(sender As Object, e As System.EventArgs) Handles txtDOB.LostFocus

        If txtDOB.Text = "  /  /" Or IsDate(txtDOB.Text) Then

        Else
            MsgBox("Invalid DOB", vbOKOnly, "Data entry error")
            txtDOB.Focus()
        End If

    End Sub

    Private Sub Setup_lvAllURs()

        lvAllURs.View = View.Details
        lvAllURs.FullRowSelect = True
        lvAllURs.GridLines = True
        lvAllURs.MultiSelect = False
        lvAllURs.HeaderStyle = ColumnHeaderStyle.Nonclickable
        lvAllURs.Sorting = SortOrder.None

        ' Create columns for the items and subitems.
        lvAllURs.Columns.Add("ur_id", 0, HorizontalAlignment.Left)
        lvAllURs.Columns.Add("hs_id", 0, HorizontalAlignment.Left)
        lvAllURs.Columns.Add(gURlabel, 70, HorizontalAlignment.Left)
        lvAllURs.Columns.Add("Health Service", 140, HorizontalAlignment.Left)
        lvAllURs.Columns.Add("Merge to", 70, HorizontalAlignment.Left)

    End Sub



End Class