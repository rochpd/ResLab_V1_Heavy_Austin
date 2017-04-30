Imports System.Text
Imports System.Globalization
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_spt

    Private flagFormLoad As Boolean = False
    Private gdicPreds As New Dictionary(Of String, String)     'ParameterID|StatTypeID, result

    Private _CallingFormName As String
    Private _NewReportstatus As String

    Private _RecordID As Long = 0
    Private _PatientID As Long = 0
    Private _panelID As Long = 0
    Private _IsNewTest As Boolean = False

    Private _formPhrases As form_rft_phrases
    Private _form_session As form_rft_session

    Public Sub UpdateReportStatusTo(ByVal NewReportStatus As String)
        _NewReportstatus = NewReportStatus
    End Sub

    Public Sub New(ByVal RecordID As Long, ByVal CallingForm As Form, Optional ByVal protocolID As Long = 0)

        InitializeComponent()
        _CallingFormName = CallingForm.Name
        _RecordID = RecordID

        Select Case RecordID
            Case 0 : _IsNewTest = True
            Case Else : _IsNewTest = False
        End Select

        'use a local patientid because can come from reporting list or logged pt
        Select Case _RecordID
            Case 0 : _PatientID = cPt.PatientID
            Case Is > 0 : _PatientID = cPt.Get_PatientIDFromPkID(_RecordID, eTables.r_spt)
        End Select

    End Sub

    Private Sub form_chall_generic_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Me.RefreshFromSession()

    End Sub

    Private Sub RefreshFromSession()

        'Refresh relevant info from session form - check if form is open first
        'Dim frmCollection = System.Windows.Forms.Application.OpenForms
        'If frmCollection.OfType(Of form_rft_session).Any Then
        With _form_session
            txtTestDate.Text = .txtTestDate.Text
            txtHeight.Text = .txtHt.Text
            txtWeight.Text = .txtWt.Text
            txtBMI.Text = cMyRoutines.calc_BMI(.txtHt.Text, .txtWt.Text)
            txtEthnicity.Text = .txtRace.Text
            txtSmokingHx.Text = .cmboSmoke_hx.Text
            txtPackYears.Text = .txtSmoke_packyears.Text
            txtYearsSmoked.Text = .txtSmoke_yrs.Text
            txtCigsPerDay.Text = .txtSmoke_cigsperday.Text
            txtDOB.Text = .txtDOB.Text
            txtGender.Text = .txtGender.Text
            txtRequestedBy.Text = .cmboReqMO_name.Text
            txtRequestedByAddress.Text = .cmboReqMO_address.Text
            txtCopyTo.Text = .cmboReportCopyTo.Text
            txtClinicalNote.Text = .txtClinicalNote.Text
        End With
        'End If

    End Sub

    Private Sub form_chall_generic_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not IsNothing(_formPhrases) Then
            If _formPhrases.Visible Then _formPhrases.Close()
        End If

    End Sub

    Private Sub form_spt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        flagFormLoad = True

        'Do session
        _form_session = New form_rft_session(_PatientID, _RecordID, eTables.r_spt, Me, _IsNewTest)
        If _IsNewTest Then
            Dim response As Integer = _form_session.ShowDialog()
            If response = DialogResult.Cancel Then
                Me.Close()
                Exit Sub
            End If
        End If
        Me.RefreshFromSession()
        Me.Text = cPt.Get_PtNameString(_PatientID, eNameStringFormats.Name_UR)

        'Load combos
        Me.Setup_ComboBoxes(_IsNewTest)
        Me.setup_datagrid()

        If _CallingFormName = "form_Reporting" Then
            If cmboReportedBy.Text = "" Then cmboReportedBy.Text = form_Reporting.cmboDefaultReporter.Text
            If txtReportedDate.Text = "" Then txtReportedDate.Text = form_Reporting.txtDefaultDate.Text
        End If

        'Load test 
        Me.Load_ScreenFieldsFromMem(Me._IsNewTest, Me._RecordID)

        btnSelectPanel.Enabled = Me._IsNewTest
        cUser.set_access(Me)

    End Sub

    Private Sub Setup_ComboBoxes(ByVal IsNewTest As Boolean)

        'Load combo box options
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLab, "Labs", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboReportedBy, "ReportingMO names", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromList(cmboReportStatus, eTables.List_ReportStatuses)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboTechnicalNote, "Technical comments", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboVerifiedBy, "ReportingMO names", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboScientist, "Scientist", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLastBD, "Last_BD", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboSite, "spt_sites", IsNewTest)
        cMyRoutines.Listbox_LoadItemsFromPrefs(lstMedications, "spt_medications", IsNewTest)

    End Sub

    Private Function Load_ScreenFieldsFromMem(IsNewTest As Boolean, sptID As Long) As Boolean

        Dim i As Integer = 0

        Try
            Select Case IsNewTest
                Case True
                    cmboReportStatus.Text = "Unreported"
                    Return True
                Case False
                    Dim flds As New class_fields_SptAndSessionFields
                    Dim p As Dictionary(Of String, String) = cRfts.Get_rft_spt_test_session(sptID)
                    If IsNothing(p) Then
                        'should never happen
                        Return False
                    Else
                        'Load the test data
                        txtReport.Text = p(flds.Report_text)
                        If IsDate(p(flds.Report_reporteddate)) Then txtReportedDate.Text = p(flds.Report_reporteddate)
                        cmboReportedBy.Text = p(flds.Report_reportedby)
                        cmboVerifiedBy.Text = p(flds.Report_verifiedby)
                        If IsDate(p(flds.Report_verifieddate)) Then txtVerifiedDate.Text = p(flds.Report_verifieddate)
                        cmboReportStatus.Text = p(flds.Report_status)
                        cmboScientist.Text = p(flds.Scientist)
                        If IsDate(p(flds.TestTime)) Then txtTestTime.Text = cMyRoutines.FormatDBTime(p(flds.TestTime))
                        cmboLastBD.Text = p(flds.BDStatus)
                        cmboLab.Text = p(flds.Lab)
                        txtTechnicalNote.Text = p(flds.TechnicalNotes)

                        txtPanelName.Text = p(flds.panel_name)
                        cmboSite.Text = p(flds.site)
                        Dim meds() As String
                        meds = Split(p(flds.medications), vbCrLf)
                        For i = 0 To meds.Count - 1
                            If meds(i) <> "" Then
                                If lstMedications.FindString(meds(i)) = -1 Then
                                    Dim kv As New KeyValuePair(Of String, Integer)(meds(i), 1000000 + i)
                                    lstMedications.Items.Add(kv)
                                End If
                                lstMedications.SetItemChecked(lstMedications.FindString(meds(i)), True)
                            End If
                        Next


                        'Load the allergen data
                        Dim TD() As Dictionary(Of String, String) = cRfts.Get_rft_spt_allergenresults(sptID)
                        If IsNothing(TD) Then

                        Else
                            Dim f As New class_fields_Spt_Allergens
                            Dim d As Dictionary(Of String, String)
                            grd.Rows.Clear()
                            For Each d In TD
                                grd.Rows.Add()
                                i = grd.Rows.Count - 1
                                grd("allergenid", i).Value = d(f.allergenID)
                                grd("panelmemberID", i).Value = d(f.panelmemberID)
                                grd("groupID", i).Value = d(f.allergen_category_id)
                                grd("#", i).Value = i + 1
                                grd("group", i).Value = d(f.allergen_category_name)
                                grd("allergen", i).Value = d(f.allergen_name)
                                grd("wheal_mm", i).Value = d(f.wheal_mm)
                                grd("note", i).Value = d(f.note)
                                grd("allergen_category_colour", i).Value = d(f.allergen_category_colour)

                                grd("group", i).Style.ForeColor = Color.FromArgb(d(f.allergen_category_colour))
                                grd("allergen", i).Style.ForeColor = Color.FromArgb(d(f.allergen_category_colour))
                                grd.Rows(i).Height = 21
                            Next
                        End If
                        TD = Nothing
                        Return True
                    End If
            End Select
            Return True

        Catch ex As Exception
            MsgBox("Error in form_spt.Load_ScreenFieldsFromMem" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Private Sub LockForm(ByVal LockState As Boolean)
        'True=locked, false=unlocked

        txtReport.ReadOnly = LockState
        pnlReporters.Enabled = Not LockState
        pnlOther.Enabled = Not LockState
        pnlPreds.Enabled = Not LockState

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    Private Sub setup_datagrid()

        Dim i As Integer = 0
        Dim c As DataGridViewColumn
        Dim Align_mc As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleCenter
        Dim Align_ml As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft

        Dim header_text() As String = {"allergenID", "panelmemberID", "groupID", "#", "Allergen group", "Allergen", "Wheal (mm)", "Note", "allergen_category_colour"}
        Dim col_name() As String = {"allergenid", "panelmemberID", "groupID", "#", "group", "allergen", "wheal_mm", "note", "allergen_category_colour"}
        Dim col_width() As Integer = {20, 20, 20, 30, 100, 100, 80, 230, 20}
        Dim col_visible() As Integer = {False, False, False, True, True, True, True, True, False}
        Dim col_readonly() As Integer = {True, True, True, True, True, True, False, False, False}
        Dim col_textalign() As DataGridViewContentAlignment = {Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_ml, Align_ml}

        Dim cell As DataGridViewCell = New DataGridViewTextBoxCell

        grd.Columns.Clear()
        For i = 0 To UBound(header_text)
            c = New DataGridViewColumn
            c.HeaderText = header_text(i)
            c.Name = col_name(i)
            c.Width = col_width(i)
            c.Visible = col_visible(i)
            c.ReadOnly = col_readonly(i)
            c.DefaultCellStyle.Alignment = col_textalign(i)
            c.CellTemplate = cell
            Me.grd.Columns.Add(c)
        Next

        grd.RowsDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8)
        grd.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8)
        grd.ColumnHeadersDefaultCellStyle.Alignment = Align_mc

    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim sptID As Long = 0
        Dim dTest As Dictionary(Of String, String) = Nothing
        Dim dTestData() As Dictionary(Of String, String) = Nothing

        dTest = Load_MemFromScreenFields_test(True)
        dTestData = Load_MemFromScreenFields_testdata(True)

        'Save session
        Dim sessionID As Long = _form_session._sessionID
        Select Case sessionID
            Case 0 : sessionID = cRfts.Insert_rft_session(Me.Load_MemFromScreenFields_session(True))
            Case Else : cRfts.Update_rft_session(Me.Load_MemFromScreenFields_session(True))
        End Select

        If Me._RecordID = 0 Then
            sptID = cRfts.insert_rft_spt(sessionID, dTest, dTestData)
        Else
            Dim ReturnValue As Boolean = cRfts.Update_rft_spt(sessionID, dTest, dTestData)
            sptID = CLng(Me._RecordID)
        End If

        'Refresh the list of tests and pdf to reflect any changes
        If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

    End Sub

    Private Function Load_MemFromScreenFields_session(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Session
        Dim R As New class_fields_SptAndSessionFields

        d(R.patientID) = Me._PatientID
        d(R.sessionID) = _form_session._sessionID
        d(R.TestDate) = cMyRoutines.FormatDBDate(txtTestDate.Text)
        d(R.Height) = q & _form_session.txtHt.Text & q
        d(R.Weight) = q & _form_session.txtWt.Text & q
        d(R.Smoke_hx) = q & Trim(_form_session.cmboSmoke_hx.Text) & q
        d(R.Smoke_yearssmoked) = q & Trim(_form_session.txtSmoke_yrs.Text) & q
        d(R.Smoke_cigsperday) = q & Trim(_form_session.txtSmoke_cigsperday.Text) & q
        d(R.Smoke_packyears) = q & Trim(_form_session.txtSmoke_packyears.Text) & q
        d(R.Req_date) = cMyRoutines.FormatDBDate(_form_session.txtRequestDate.Text)
        d(R.Req_name) = q & Trim(_form_session.cmboReqMO_name.Text) & q
        d(R.Req_phone) = q & Trim(_form_session.txtReqMO_phone.Text) & q
        d(R.Req_fax) = q & Trim(_form_session.txtReqMO_fax.Text) & q
        d(R.Req_email) = q & Trim(_form_session.txtReqMO_email.Text) & q
        d(R.Req_address) = q & Trim(_form_session.cmboReqMO_address.Text) & q
        d(R.Req_providernumber) = q & Trim(_form_session.txtReqMO_pn.Text) & q
        d(R.Report_copyto) = q & Trim(_form_session.cmboReportCopyTo.Text) & q
        d(R.Req_clinicalnotes) = q & Trim(_form_session.txtClinicalNote.Text) & q
        d(R.Req_healthservice) = q & Trim(_form_session.cmboReqMO_HealthService.Text) & q
        d(R.AdmissionStatus) = q & Trim(_form_session.cmboAdmission.Text) & q
        d(R.Billing_billedto) = q & Trim(_form_session.cmboBilledTo.Text) & q
        d(R.Billing_billingMO) = q & Trim(_form_session.cmboBillingMO_name.Text) & q
        d(R.Billing_billingMOproviderno) = q & Trim(_form_session.cmboBillingMO_pn.Text) & q
        d(R.Pred_SourceIDs) = q & txtPredsRaw.Text & q
        d(R.LastUpdated_session) = cMyRoutines.FormatDBDateTime(Date.Now)
        d(R.LastUpdatedBy_session) = q & My.User.Name & q

        Return d

    End Function

    Private Function Load_MemFromScreenFields_test(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)


        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicSPT
        Dim R As New class_fields_SptAndSessionFields
        Dim i As Integer = 0

        dicR(R.patientID) = Me._PatientID
        dicR(R.sessionID) = Nothing     'set in save
        dicR(R.sptID) = Me._RecordID
        dicR(R.TestTime) = cMyRoutines.FormatDBTime(txtTestTime.Text)
        dicR(R.TestType) = q & StrConv(lblTestType.Text, VbStrConv.ProperCase) & q
        dicR(R.Lab) = q & cmboLab.Text & q
        dicR(R.Scientist) = q & cmboScientist.Text & q
        dicR(R.BDStatus) = q & cmboLastBD.Text & q
        dicR(R.TechnicalNotes) = q & txtTechnicalNote.Text & q
        dicR(R.Report_text) = q & txtReport.Text & q
        dicR(R.Report_reportedby) = q & cmboReportedBy.Text & q
        dicR(R.Report_reporteddate) = cMyRoutines.FormatDBDate(txtReportedDate.Text)
        dicR(R.Report_verifiedby) = q & cmboVerifiedBy.Text & q
        dicR(R.Report_verifieddate) = cMyRoutines.FormatDBDate(txtVerifiedDate.Text)
        dicR(R.Report_status) = q & cmboReportStatus.Text & q
        dicR(R.LastUpdated_spt) = cMyRoutines.FormatDBDateTime(Now)
        dicR(R.LastUpdatedBy_spt) = q & My.User.Name & q

        If _CallingFormName = "form_Reporting" And cmboReportStatus.SelectedIndex < cmboReportStatus.Items.Count - 1 Then
            Dim f As New form_UpdateReportStatus(cmboReportStatus.Text)
            Dim Response As String = f.ShowDialog(Me)
            If Response = Windows.Forms.DialogResult.OK Then
                dicR(R.Report_status) = q & f.cmboUpdate.Text & q
            Else
                dicR(R.Report_status) = q & cmboReportStatus.Text & q
            End If
            f = Nothing
        Else
            dicR(R.Report_status) = q & cmboReportStatus.Text & q
        End If

        dicR(R.site) = q & cmboSite.Text & q
        dicR(R.panel_name) = q & txtPanelName.Text & q
        dicR(R.panelID) = _panelID
        dicR(R.medications) = q
        For i = 0 To lstMedications.CheckedItems.Count - 1
            dicR(R.medications) = dicR(R.medications) & lstMedications.CheckedItems(i).key & vbCrLf
        Next i
        dicR(R.medications) = dicR(R.medications) & q

        Return dicR

    End Function

    Private Function Load_MemFromScreenFields_testdata(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)()


        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d() As Dictionary(Of String, String) = Nothing
        Dim i As Integer = -1
        Dim f As New class_fields_Spt_Allergens

        For Each r As DataGridViewRow In grd.Rows
            i = i + 1
            ReDim Preserve d(i)
            d(i) = cMyRoutines.MakeEmpty_dicSpt_allergen
            d(i)(f.sptID) = Me._RecordID
            d(i)(f.allergenID) = r.Cells.Item("allergenid").Value
            d(i)(f.allergen_category_id) = r.Cells.Item("groupID").Value
            d(i)(f.allergen_category_name) = q & r.Cells.Item("group").Value & q
            d(i)(f.allergen_category_colour) = q & r.Cells.Item("allergen_category_colour").Value & q
            d(i)(f.allergen_name) = q & r.Cells.Item("allergen").Value & q
            d(i)(f.wheal_mm) = q & r.Cells.Item("wheal_mm").Value & q
            d(i)(f.note) = q & r.Cells.Item("note").Value & q
            d(i)(f.panelmemberID) = r.Cells.Item("panelmemberID").Value
        Next

        Return d

    End Function


    Private Sub btnSelectPanel_Click(sender As Object, e As System.EventArgs) Handles btnSelectPanel.Click

        Dim f As New form_spt_panelbuilder(_IsNewTest)
        If f.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim panelID As Integer = CLng(f.lvPanels.Items(f.lvPanels.SelectedIndices(0)).SubItems(1).Text)
            Me.load_blankpaneltogrid(panelID)
            txtPanelName.Text = cSpt.Get_PanelnameFromID(_panelID)
        End If

    End Sub

    Private Sub load_blankpaneltogrid(panelid As Long)

        'Check if grid is empty
        If grd.Rows.Count > 0 Then
            If MsgBox("Existing panel data in the grid will be lost. Continue?", vbOKCancel, "Copy panel") = vbCancel Then
                Exit Sub
            End If
        End If

        'Clear and load the grid
        grd.Rows.Clear()
        Dim a As List(Of PanelMember_AllData) = cSpt.Get_AllergensForPanel(True, panelid)
        If Not IsNothing(a) Then
            _panelID = panelid
            For i As Integer = 0 To a.Count - 1
                grd.Rows.Add()
                grd("allergenID", i).Value = a(i).allergenid
                grd("panelmemberID", i).Value = a(i).memberid
                grd("groupID", i).Value = a(i).allergengroupID
                grd("#", i).Value = i + 1
                grd("group", i).Value = a(i).allergengroup
                grd("allergen", i).Value = a(i).allergenname
                grd("allergen_category_colour", i).Value = a(i).displayColour

                grd("group", i).Style.ForeColor = Color.FromArgb(a(i).displayColour)
                grd("allergen", i).Style.ForeColor = Color.FromArgb(a(i).displayColour)
            Next i
        Else

        End If

    End Sub


    Private Sub btnSession_Click(sender As Object, e As System.EventArgs) Handles btnSession.Click
        _form_session._cancelButton_enable = False
        _form_session.Visible = True
        _form_session.BringToFront()
    End Sub

    Private Sub cmboTechnicalNote_DropDownClosed(sender As Object, e As System.EventArgs) Handles cmboTechnicalNote.DropDownClosed

        Dim kv As KeyValuePair(Of String, Integer) = cmboTechnicalNote.SelectedItem
        If Len(txtTechnicalNote.Text) = 0 Then txtTechnicalNote.Text = kv.Key Else txtTechnicalNote.Text = txtTechnicalNote.Text & " " & kv.Key
        cmboTechnicalNote.SelectedIndex = -1
        txtTechnicalNote.Focus()
        txtTechnicalNote.SelectionStart = txtTechnicalNote.Text.Length
        txtTechnicalNote.SelectionLength = 0

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

        If cRfts.Delete_rft(Me._RecordID, eTables.r_spt) Then
            'Refresh the list of tests and pdf to reflect any changes
            If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True
            Me.Close()
        Else
            'Keep form open if delete cancelled
        End If

    End Sub

    Private Sub btnPhrases_Click(sender As System.Object, e As System.EventArgs) Handles btnPhrases.Click

        Dim demo As New Pred_demo
        demo.Age = cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text))
        demo.DOB = CDate(txtDOB.Text)
        demo.Htcm = Val(txtHeight.Text)
        demo.Wtkg = Val(txtWeight.Text)
        demo.GenderID = cMyRoutines.Lookup_list_ByDescription(txtGender.Text, eTables.Pred_ref_genders)
        demo.Gender = txtGender.Text
        demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(txtEthnicity.Text, eTables.Pred_ref_ethnicities)
        demo.Ethnicity = txtEthnicity.Text
        demo.TestDate = txtTestDate.Text

        Dim dicR As New Dictionary(Of String, String)

        _formPhrases = New form_rft_phrases(eAutoreport_testgroups.Spt, Me)
        _formPhrases.ReportTextBox = Me.txtReport
        _formPhrases.Show()

    End Sub

End Class