
Imports System.Text
Imports System.Globalization
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_hast

    Private flagFormLoad As Boolean = False
    Private gdicPreds As New Dictionary(Of String, String)     'ParameterID|StatTypeID, result
    Private _protocol_data As New class_hast_protocoldata
    Private _CallingFormName As String
    Private _NewReportstatus As String

    Private _RecordID As Long = 0
    Private _PatientID As Long = 0
    Private _ProtocolID As Integer = 0
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
        _ProtocolID = protocolID

        Select Case RecordID
            Case 0 : _IsNewTest = True
            Case Else : _IsNewTest = False
        End Select

        'use a local patientid because can come from reporting list or logged pt
        Select Case _RecordID
            Case 0 : _PatientID = cPt.PatientID
            Case Is > 0 : _PatientID = cPt.Get_PatientIDFromPkID(_RecordID, eTables.r_hast)
        End Select


    End Sub

    Private Sub form_hast_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

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

    Private Sub form_hast_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not IsNothing(_formPhrases) Then
            If _formPhrases.Visible Then _formPhrases.Close()
        End If

    End Sub

    Private Sub form_hast_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        flagFormLoad = True

        'Do session
        _form_session = New form_rft_session(_PatientID, _RecordID, eTables.r_hast, Me, _IsNewTest)
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

        If _CallingFormName = "form_Reporting" Then
            If cmboReportedBy.Text = "" Then cmboReportedBy.Text = form_Reporting.cmboDefaultReporter.Text
            If txtReportedDate.Text = "" Then txtReportedDate.Text = form_Reporting.txtDefaultDate.Text
        End If

        'Load test 
        Me.Load_ScreenFieldsFromMem(Me._IsNewTest, Me._RecordID)

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
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboDeliverymethod_fio2, "AST_deliverymethod_FiO2", IsNewTest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboDeliverymethod_suppo2, "AST_deliverymethod_Suppl_O2", IsNewTest)

    End Sub

    Private Sub Setup_datagrid(p As class_hast_protocoldata, isnewtest As Boolean, Optional TD() As Dictionary(Of String, String) = Nothing)

        Dim i As Integer = 0
        Dim wid As Integer = 45
        Dim c As DataGridViewColumn
        Dim Align_mc As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleCenter
        Dim Align_ml As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleLeft

        Dim header_text() As String = {"levelID", "#", "FiO2", "Altitude", "Supp O2", "SpO2", "pH", "PaCO2", "PaO2", "BE", "HCO3", "SaO2", "Note"}
        Dim header_units() As String = {"", "", "(%)", "(ft)", "(L/min)", "(%)", "", "(mmHg)", "(mmHg)", "(mmol/L)", "(mmol/L)", "(%)", ""}
        Dim col_name() As String = {"levelid", "#", "fio2", "altitude", "suppo2", "spo2", "ph", "paco2", "pao2", "be", "hco3", "sao2", "note"}
        Dim col_width() As Integer = {20, 25, wid, wid + 10, wid + 13, wid, wid, wid, wid, wid, wid, wid, 290}
        Dim col_visible() As Boolean = Nothing
        Select Case p.abg_enabled
            Case True
                Select Case p.abg_part_enabled
                    Case True : col_visible = {False, True, True, True, True, True, True, True, True, False, False, True, True}
                    Case False : col_visible = {False, True, True, True, True, True, True, True, True, True, True, True, True}
                End Select
            Case False : col_visible = {False, True, True, True, True, True, False, False, False, False, False, False, True}
        End Select

        Dim col_readonly() As Boolean = {True, True, True, True, False, False, False, False, False, False, False, False, False}
        Dim col_textalign() As DataGridViewContentAlignment = {Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_ml}

        'Set column headers
        grd.Columns.Clear()
        For i = 0 To UBound(header_text)
            c = New DataGridViewColumn
            c.CellTemplate = New DataGridViewTextBoxCell
            c.HeaderText = header_text(i)
            c.HeaderCell.Style.Alignment = col_textalign(i)
            c.HeaderCell.Style.BackColor = SystemColors.Control
            c.HeaderCell.Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
            c.Name = col_name(i)
            c.Width = col_width(i)
            c.Visible = col_visible(i)
            c.ReadOnly = col_readonly(i)
            c.HeaderCell.Style.SelectionBackColor = c.HeaderCell.Style.BackColor
            Me.grd.Columns.Add(c)
        Next

        'Set units row
        grd.Rows.Add(header_units)
        grd.Rows(0).DefaultCellStyle.BackColor = SystemColors.Control
        grd.Rows(0).DefaultCellStyle.SelectionBackColor = SystemColors.Control
        grd.Rows(0).DefaultCellStyle.SelectionForeColor = Color.Black
        grd.Rows(0).ReadOnly = True
        For i = 0 To UBound(header_text)
            grd.Rows(0).Cells(i).Style.Font = New Font("Segoe UI", 8, FontStyle.Bold)
            grd.Rows(0).Cells(i).Style.Alignment = col_textalign(i)
        Next i

        'Set cell style
        Dim st As New DataGridViewCellStyle
        st.BackColor = Color.White
        st.Font = New Font("Microsoft Sans Serif", 8)
        st.SelectionBackColor = Color.White
        st.SelectionForeColor = Color.Black
        st.Alignment = DataGridViewContentAlignment.MiddleCenter

        'Add the empty results rows
        Select Case isnewtest
            Case True
                i = 1
                If p.fio2_1_enabled Then grd.Rows.Add("0", i, p.fio2_1, p.fio2_1_altitude) : grd.Rows(i).DefaultCellStyle = st : i = i + 1
                If p.fio2_2_enabled Then grd.Rows.Add("0", i, p.fio2_2, p.fio2_2_altitude) : grd.Rows(i).DefaultCellStyle = st : i = i + 1
                If p.fio2_3_enabled Then grd.Rows.Add("0", i, p.fio2_3, p.fio2_3_altitude) : grd.Rows(i).DefaultCellStyle = st : i = i + 1
                If p.fio2_4_enabled Then grd.Rows.Add("0", i, p.fio2_4, p.fio2_4_altitude) : grd.Rows(i).DefaultCellStyle = st : i = i + 1
                If p.fio2_5_enabled Then grd.Rows.Add("0", i, p.fio2_5, p.fio2_5_altitude) : grd.Rows(i).DefaultCellStyle = st : i = i + 1
                If p.fio2_6_enabled Then grd.Rows.Add("0", i, p.fio2_6, p.fio2_6_altitude) : grd.Rows(i).DefaultCellStyle = st
            Case False
                Dim f As New class_fields_Hast_Levels
                Dim d As Dictionary(Of String, String)
                i = 1
                For Each d In TD
                    grd.Rows.Add(d(f.levelID), i, d(f.altitude_fio2), d(f.altitude_ft), d(f.suppO2_flow), d(f.r_spo2), d(f.r_ph), d(f.r_paco2), d(f.r_pao2), d(f.r_be), d(f.r_hco3), d(f.r_sao2), d(f.r_note))
                    grd.Rows(i).DefaultCellStyle = st
                    grd.Rows(i).Height = 21
                    i = i + 1
                Next
        End Select


        For i = 1 To grd.Rows.Count - 1
            grd("note", i).Style.Alignment = Align_ml
            grd("#", i).Style.BackColor = SystemColors.Control
            grd("fio2", i).Style.BackColor = SystemColors.Control
            grd("altitude", i).Style.BackColor = SystemColors.Control
            grd("#", i).Style.SelectionBackColor = SystemColors.Control
            grd("fio2", i).Style.SelectionBackColor = SystemColors.Control
            grd("altitude", i).Style.SelectionBackColor = SystemColors.Control
            grd("#", i).Style.SelectionForeColor = Color.Black
            grd("fio2", i).Style.SelectionForeColor = Color.Black
            grd("altitude", i).Style.SelectionForeColor = Color.Black
        Next

    End Sub

    Private Sub btnRow_add_Click(sender As System.Object, e As System.EventArgs) Handles btnRow_add.Click

        Dim response As String = InputBox("Enter row # to use as a template", "Add a new row")
        If response = "" Then

        ElseIf Not IsNumeric(response) Then
            MsgBox("Please enter a valid row number")
        ElseIf CInt(response) < 1 Or CInt(response) > grd.Rows.Count - 1 Then
            MsgBox("Row number out of range - please enter a valid row number")
        Else
            grd.Rows.AddCopy(response)
            grd.Rows(grd.Rows.Count - 1).SetValues("0", "", grd(2, CInt(response)).Value, grd(3, CInt(response)).Value)

            'Re-number visible rows
            Dim i As Integer = 1
            For j = 1 To grd.Rows.Count - 1
                If grd.Rows(j).Visible Then
                    grd("#", j).Value = i.ToString
                    i = i + 1
                End If
            Next
        End If

    End Sub

    Private Sub btnRow_delete_Click(sender As Object, e As System.EventArgs) Handles btnRow_delete.Click

        Dim response As String = InputBox("Enter row # to delete", "Delete row")
        If response = "" Then

        ElseIf Not IsNumeric(response) Then
            MsgBox("Please enter a valid row number")
        ElseIf CInt(response) < 1 Or CInt(response) > grd.Rows.Count - 1 Then
            MsgBox("Row number out of range - please enter a valid row number")
        Else
            Dim r As Integer = CInt(response)
            grd(0, r).Value = (Val(grd(0, r).Value) * -1).ToString
            grd.Rows(r).Visible = False

            'Re-number visible rows
            Dim i As Integer = 1
            For j = 1 To grd.Rows.Count - 1
                If grd.Rows(j).Visible Then
                    grd("#", j).Value = i.ToString
                    i = i + 1
                End If
            Next
        End If

    End Sub

    Private Sub btnRow_insert_Click(sender As Object, e As System.EventArgs) Handles btnRow_insert.Click

        Dim response As String = InputBox("Enter row # to use as a template", "Insert a new row")
        If response = "" Then

        ElseIf Not IsNumeric(response) Then
            MsgBox("Please enter a valid row number")
        ElseIf CInt(response) < 1 Or CInt(response) > grd.Rows.Count - 1 Then
            MsgBox("Row number out of range - please enter a valid row number")
        Else
            grd.Rows.InsertCopy(CInt(response), CInt(response))
            grd.Rows(CInt(response)).SetValues("0", "", grd(2, CInt(response) + 1).Value, grd(3, CInt(response) + 1).Value)

            'Re-number visible rows
            Dim i As Integer = 1
            For j = 1 To grd.Rows.Count - 1
                If grd.Rows(j).Visible Then
                    grd("#", j).Value = i.ToString
                    i = i + 1
                End If
            Next
        End If

    End Sub

    Private Function Load_ScreenFieldsFromMem(IsNewTest As Boolean, hastID As Long) As Boolean

        Dim i As Integer = 0

        Try
            Select Case IsNewTest
                Case True
                    _protocol_data = cMyRoutines.Get_Hast_Protocol
                    _ProtocolID = _protocol_data.protocolID
                    Me.Setup_datagrid(_protocol_data, IsNewTest)
                    cmboReportStatus.Text = "Unreported"
                    Return True
                Case False
                    Dim flds As New class_fields_HastAndSessionFields
                    Dim p As Dictionary(Of String, String) = cRfts.Get_rft_hast_test_session(hastID)
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

                        cmboDeliverymethod_fio2.Text = p(flds.deliverymethod_fio2)
                        cmboDeliverymethod_suppo2.Text = p(flds.deliverymethod_suppo2)

                        'Load the level data
                        Dim TD() As Dictionary(Of String, String) = cRfts.Get_rft_hast_levelresults(hastID)
                        If IsNothing(TD) Then

                        Else
                            _protocol_data = cMyRoutines.Get_Hast_Protocol(p(flds.protocol_id))
                            _ProtocolID = _protocol_data.protocolID
                            Me.Setup_datagrid(_protocol_data, IsNewTest, TD)
                        End If
                        TD = Nothing
                        Return True
                    End If
            End Select
            Return True

        Catch ex As Exception
            MsgBox("Error in form_hast.Load_ScreenFieldsFromMem" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Private Function Load_MemFromScreenFields_session(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Session
        Dim R As New class_fields_HastAndSessionFields

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

        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicHAST
        Dim R As New class_fields_HastAndSessionFields
        Dim i As Integer = 0

        dicR(R.patientID) = Me._PatientID
        dicR(R.sessionID) = Nothing     'set in save
        dicR(R.hastID) = Me._RecordID
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
        dicR(R.LastUpdated_hast) = cMyRoutines.FormatDBDateTime(Now)
        dicR(R.LastUpdatedBy_hast) = q & My.User.Name & q

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

        dicR(R.deliverymethod_fio2) = q & cmboDeliverymethod_fio2.Text & q
        dicR(R.deliverymethod_suppo2) = q & cmboDeliverymethod_suppo2.Text & q
        dicR(R.protocol_id) = _ProtocolID
        dicR(R.protocol_name) = q & _protocol_data.description & q

        Return dicR

    End Function

    Private Function Load_MemFromScreenFields_testdata(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)()

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d() As Dictionary(Of String, String) = Nothing
        Dim i As Integer = 0
        Dim f As New class_fields_Hast_Levels

        For i = 1 To grd.Rows.Count - 1
            ReDim Preserve d(i - 1)
            d(i - 1) = cMyRoutines.MakeEmpty_dicHAST_level
            d(i - 1)(f.hastID) = Me._RecordID
            d(i - 1)(f.levelID) = grd("levelid", i).Value
            d(i - 1)(f.row_order) = grd("#", i).Value
            d(i - 1)(f.altitude_fio2) = q & grd("fio2", i).Value & q
            d(i - 1)(f.altitude_ft) = q & grd("altitude", i).Value & q
            d(i - 1)(f.r_be) = q & grd("be", i).Value & q
            d(i - 1)(f.r_hco3) = q & grd("hco3", i).Value & q
            d(i - 1)(f.r_note) = q & grd("note", i).Value & q
            d(i - 1)(f.r_paco2) = q & grd("paco2", i).Value & q
            d(i - 1)(f.r_pao2) = q & grd("pao2", i).Value & q
            d(i - 1)(f.r_ph) = q & grd("ph", i).Value & q
            d(i - 1)(f.r_sao2) = q & grd("sao2", i).Value & q
            d(i - 1)(f.r_spo2) = q & grd("spo2", i).Value & q
            d(i - 1)(f.suppO2_flow) = q & grd("suppo2", i).Value & q
        Next

        Return d

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

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click

        Dim hastID As Long = 0
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
            hastID = cRfts.insert_rft_hast(sessionID, dTest, dTestData)
        Else
            Dim ReturnValue As Boolean = cRfts.Update_rft_hast(sessionID, dTest, dTestData)
            hastID = CLng(Me._RecordID)
        End If

        'Refresh the list of tests and pdf to reflect any changes
        If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

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

        If cRfts.Delete_rft(Me._RecordID, eTables.r_hast) Then
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

        _formPhrases = New form_rft_phrases(eAutoreport_testgroups.Hast, Me)
        _formPhrases.ReportTextBox = Me.txtReport
        _formPhrases.Show()

    End Sub

End Class