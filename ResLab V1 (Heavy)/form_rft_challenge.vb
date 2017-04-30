Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_rft_challenge

    Private flagFormLoad As Boolean = False
    Private gdicPreds As New Dictionary(Of String, String)     'ParameterID|StatTypeID, result
    Private _pData As class_challenge.ProtocolData

    Private _CallingFormName As String
    Private _NewReportstatus As String

    Private _RecordID As Long = 0
    Private _ProtocolID As Integer = 0
    Private _PatientID As Long = 0
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
            Case Is > 0 : _PatientID = cPt.Get_PatientIDFromPkID(_RecordID, eTables.Prov_test)
        End Select

    End Sub

    Private Sub form_chall_generic_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Me.RefreshFromSession()
        Me.Do_calculations()

    End Sub

    Private Sub RefreshFromSession()

        'Refresh relevant info from session form - check if form is open first
        'Dim frmCollection = System.Windows.Forms.Application.OpenForms
        'If frmCollection.OfType(Of form_rft_session).Any Then
        With _form_session
            txtTestDate.Text = .txtTestDate.Text
            txtHeight.Text = .txtHt.Text
            txtWeight.Text = .txtWt.Text
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

    Private Sub form_chall_generic_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        flagFormLoad = True

        'Do session
        _form_session = New form_rft_session(_PatientID, _RecordID, eTables.Prov_test, Me, _IsNewTest)
        If _IsNewTest Then
            Dim response As Integer = _form_session.ShowDialog()
            If response = DialogResult.Cancel Then
                Me.Close()
                Exit Sub
            End If
        End If
        Me.RefreshFromSession()
        Me.Text = cPt.Get_PtNameString(_PatientID, eNameStringFormats.Name_UR)

        'Load protocol 
        Select Case _IsNewTest
            Case True 'passed in
            Case False : _ProtocolID = cRfts.Get_ProtocolID_FromProvID(Me._RecordID)
        End Select
        Me._pData = cChall.Get_ProtocolProperties(_ProtocolID)  'Get current protocol info
        tabResults.TabPages(0).Text = _pData.title
        Me.Setup_ProtocolGrid(_IsNewTest)

        'Load combos
        Me.Setup_ComboBoxes(_IsNewTest)
        If _CallingFormName = "form_Reporting" Then
            If cmboReportedBy.Text = "" Then cmboReportedBy.Text = form_Reporting.cmboDefaultReporter.Text
            If txtReportedDate.Text = "" Then txtReportedDate.Text = form_Reporting.txtDefaultDate.Text
        End If

        'Load test 
        Me.Load_ScreenFieldsFromMem(Me._IsNewTest, Me._RecordID)
        Me.Do_calculations()

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

    End Sub

    Private Sub Setup_ProtocolGrid(ByVal IsNewProv As Boolean)
        'If new prov get protocol data for grid from protocols table
        'If existing, get the protocol from data stored with the results

        Dim headerFont As Font = New Font("Segoe UI", 8)
        grdD.ColumnHeadersDefaultCellStyle.Font = headerFont
        grdD.ColumnHeadersDefaultCellStyle.ForeColor = Color.Blue
        grdD.Rows.Clear()

        Select Case IsNewProv
            Case True
                For i As Integer = 0 To UBound(_pData.doses)
                    If Val(_pData.doses(i).xaxislabel) > 0 Then
                        grdD.Rows.Add(_pData.doses(i).doseID, _pData.doses(i).xaxislabel & " " & _pData.agent_units, "", "", "", _pData.doses(i).dosenumber, "0")
                    Else
                        grdD.Rows.Add(_pData.doses(i).doseID, _pData.doses(i).xaxislabel, "", "", "", _pData.doses(i).dosenumber, "0")
                    End If
                    grdD.Rows(grdD.Rows.Count - 1).Height = 21
                Next
                lblPD.Text = "PD" & _pData.pd_thresh
                lblTable.Text = "Dose response to " & _pData.agent
                grdD.Columns(2).HeaderText = _pData.parameter
            Case Else
                'Leave to the retrieve routine
        End Select

    End Sub

    Private Function Load_MemFromScreenFields_session(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Session
        Dim R As New class_fields_ProvAndSession

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

        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicProv_test
        Dim R As New class_fields_ProvAndSession

        dicR(R.patientID) = Me._PatientID
        dicR(R.sessionID) = Nothing     'set in save
        dicR(R.provID) = Me._RecordID
        dicR(R.TestTime) = cMyRoutines.FormatDBTime(txtTestTime.Text)
        dicR(R.TestType) = q & tabResults.TabPages(0).Text & q
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
        dicR(R.ProtocolID) = Me._ProtocolID
        dicR(R.LastUpdated_prov) = cMyRoutines.FormatDBDateTime(Now)
        dicR(R.LastUpdatedBy_prov) = q & My.User.Name & q

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

        dicR(R.R_bl_Fev1) = q & txtR_bl_Fev1.Text & q
        dicR(R.R_bl_Fvc) = q & txtR_bl_Fvc.Text & q
        dicR(R.R_bl_Vc) = q & txtR_bl_Vc.Text & q
        dicR(R.R_bl_Fef2575) = q & txtR_bl_Fef2575.Text & q
        dicR(R.R_bl_Pef) = q & txtR_bl_Pef.Text & q
        dicR(R.ProtocolID) = _pData.protocolID
        dicR(R.Protocol_threshold) = q & _pData.pd_thresh & q
        dicR(R.Protocol_pd_decimalplaces) = q & _pData.pd_decimalplaces & q
        dicR(R.Protocol_method) = q & _pData.pd_method & q
        dicR(R.Protocol_doseunits) = q & _pData.agent_units & q
        dicR(R.Protocol_drug) = q & _pData.agent & q
        dicR(R.Protocol_parameter) = q & _pData.parameter & q
        dicR(R.Protocol_parameter_units) = q & _pData.parameter_units & q
        dicR(R.Protocol_parameter_response) = q & _pData.parameter_response & q
        dicR(R.Protocol_dose_effect) = q & _pData.pd_dose_effect & q
        dicR(R.Protocol_method_reference) = q & _pData.pd_method_reference & q
        dicR(R.Protocol_post_drug) = q & _pData.post_drug & q

        dicR(R.Protocol_title) = q & _pData.title & q
        dicR(R.plot_title) = q & _pData.plot_title & q
        dicR(R.plot_xscaling_type) = q & _pData.plot_xscaling_type & q
        dicR(R.plot_xtitle) = q & _pData.plot_xtitle & q
        dicR(R.plot_ymax) = q & _pData.plot_ymax & q
        dicR(R.plot_ymin) = q & _pData.plot_ymin & q
        dicR(R.plot_ystep) = q & _pData.plot_ystep & q

        dicR(R.LastUpdated_prov) = cMyRoutines.FormatDBDate(Now)
        dicR(R.LastUpdatedBy_prov) = q & My.User.Name & q

        Return dicR

    End Function

    Private Function Load_MemFromScreenFields_testdata(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)()

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d() As Dictionary(Of String, String) = Nothing
        Dim i As Integer = -1
        Dim f As New class_ProvTestDataFields

        For Each r As DataGridViewRow In grdD.Rows
            i = i + 1
            ReDim Preserve d(i)
            d(i) = cMyRoutines.MakeEmpty_dicProv_testdata
            d(i)(f.testdataid) = r.Cells.Item(6).Value
            d(i)(f.provid) = Me._RecordID
            d(i)(f.doseid) = r.Cells.Item(0).Value
            d(i)(f.dose_number) = q & r.Cells.Item(5).Value & q
            If _pData.pd_dose_effect = "Cumulative" Then
                d(i)(f.dose_cumulative) = q & _pData.doses(i).dose_cumulative & q
                d(i)(f.dose_discrete) = "NULL"
            ElseIf _pData.pd_dose_effect = "Discrete" Then
                d(i)(f.dose_cumulative) = "NULL"
                d(i)(f.dose_discrete) = q & _pData.doses(i).dose_discrete & q
            Else
                d(i)(f.dose_cumulative) = "NULL"
                d(i)(f.dose_discrete) = "NULL"
            End If
            d(i)(f.result) = q & r.Cells.Item(2).Value & q
            d(i)(f.response) = q & r.Cells.Item(3).Value & q
            d(i)(f.xaxis_label) = q & _pData.doses(i).xaxislabel & q
            d(i)(f.dose_time_min) = q & _pData.doses(i).time_min & q
            'd(i)(f.dose_canskip) = q & CInt(_pData.doses(i).canskip) & q
        Next

        Return d

    End Function



    Private Function Load_ScreenFieldsFromMem(IsNewTest As Boolean, testID As Long) As Boolean

        Try
            Select Case IsNewTest
                Case True
                    cmboReportStatus.Text = "Unreported"
                    Return True
                Case False
                    Dim flds As New class_fields_ProvAndSession
                    Dim p As Dictionary(Of String, String) = cRfts.Get_prov_test_session(testID)
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

                        Dim T As Dictionary(Of String, String) = cRfts.Get_Prov_VisitDataByProvID(testID)
                        Dim TD() As Dictionary(Of String, String) = cRfts.Get_Prov_TestDataByProvID(testID)
                        Me._pData = cChall.load_provdataStructure_from_dictionaries(T, TD)

                        If IsNothing(T) Or IsNothing(TD) Then
                            'should never happen
                            Return False
                        Else
                            Dim f1 As New class_fields_ProvAndSession
                            Dim f2 As New class_ProvTestDataFields

                            txtR_bl_Fev1.Text = fmt(T(f1.R_bl_Fev1), 2)
                            txtR_bl_Fvc.Text = fmt(T(f1.R_bl_Fvc), 2)
                            txtR_bl_Vc.Text = fmt(T(f1.R_bl_Vc), 2)
                            txtR_bl_Fef2575.Text = fmt(T(f1.R_bl_Fef2575), 1)
                            txtR_bl_Pef.Text = fmt(T(f1.R_bl_Pef), 1)
                            lblPD.Text = "PD" & T(f1.Protocol_threshold)
                            lblTable.Text = "Dose response to " & T(f1.Protocol_drug)
                            grdD.Columns(2).HeaderText = T(f1.Protocol_parameter)

                            Dim d As Dictionary(Of String, String)
                            grdD.Rows.Clear()
                            For Each d In TD
                                If Val(d(f2.xaxis_label)) > 0 Then
                                    grdD.Rows.Add(d(f2.doseid), d(f2.xaxis_label) & " " & T(f1.Protocol_doseunits), fmt(d(f2.result), 2), fmt(d(f2.response), 1), d(f2.dose_canskip), d(f2.dose_number), d(f2.testdataid))
                                Else
                                    grdD.Rows.Add(d(f2.doseid), d(f2.xaxis_label), fmt(d(f2.result), 2), fmt(d(f2.response), 1), d(f2.dose_canskip), d(f2.dose_number), d(f2.testdataid))
                                End If
                                grdD.Rows(grdD.Rows.Count - 1).Height = 21
                            Next

                        End If

                        'Continue and load flow vol loop
                        cDAL.Get_image(testID, picFV, eTables.Prov_test, "flowvolloop")

                        T = Nothing
                        TD = Nothing
                        Return True
                    End If
            End Select
            Return True

        Catch ex As Exception
            MsgBox("Error in form_chall_generic.Load_ScreenFieldsFromMem" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Private Function fmt(ByVal s As String, ByVal Places As Integer) As String
        'Takes a number as string and formats it to x decimal places and returns as string

        If s <> "" Then
            Dim p As String = "0." & Strings.StrDup(Places, "0")
            Return Format(Val(s), p)
        Else
            Return ""
        End If

    End Function

    Private Sub Do_calculations() Handles grdD.CellEndEdit

        Me.Load_Normals_Spirometry(class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)

        'Response
        Dim Ref As Single = 0
        Select Case _pData.pd_method_reference
            Case "Control" : Ref = Val(grdD.Item(2, 1).Value)
            Case "Baseline" : Ref = Val(grdD.Item(2, 0).Value)
        End Select
        If Ref > 0 Then
            For Each r As DataGridViewRow In grdD.Rows
                If r.Cells.Item(2).Value = "" Then
                    r.Cells.Item(3).Value = ""
                Else
                    r.Cells.Item(3).Value = fmt(100 * r.Cells.Item(2).Value / Ref, 1)
                End If
            Next
            Me.DrawGraph()
        End If

        If Not IsNothing(_pData.doses) Then txtPD.Text = cChall.Calculate_PDx(_pData)

        'Baseline FER - use the larger of VC and FVC
        If Val(txtR_bl_Fev1.Text) > 0 Then
            If Val(txtR_bl_Fvc.Text) > Val(txtR_bl_Vc.Text) Then
                If Val(txtR_bl_Fvc.Text) > 0 Then
                    txtR_bl_Fer.Text = fmt(100 * Val(txtR_bl_Fev1.Text) / Val(txtR_bl_Fvc.Text), 0)
                Else
                    txtR_bl_Fer.Text = ""
                End If
            Else
                If Val(txtR_bl_Vc.Text) > 0 Then
                    txtR_bl_Fer.Text = fmt(100 * Val(txtR_bl_Fev1.Text) / Val(txtR_bl_Vc.Text), 0)
                Else
                    txtR_bl_Fer.Text = ""
                End If
            End If
        End If

        'BMI
        txtBMI.Text = cMyRoutines.calc_BMI(txtHeight.Text, txtWeight.Text)

        'Age
        txtAge.Text = "(" & cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text)) & ")"

        '% predicteds
        cPred.LoadPPN(txtR_blPpn_Fev1, txtR_bl_Fev1, "FEV1|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Fvc, txtR_bl_Fvc, "FVC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Vc, txtR_bl_Vc, "FVC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Fer, txtR_bl_Fer, "FER|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Pef, txtR_bl_Pef, "PEF|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Fef2575, txtR_bl_Fef2575, "FEF2575|MPV", gdicPreds)

    End Sub

    Private Sub Load_Normals_Spirometry(ByVal Method As class_Pred.eLoadNormalsMode)

        Dim demo As Pred_demo = Nothing

        demo.Age = cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text))
        demo.Htcm = Val(txtHeight.Text)
        demo.Wtkg = Val(txtWeight.Text)
        demo.GenderID = cMyRoutines.Lookup_list_ByDescription(txtGender.Text, eTables.Pred_ref_genders)
        demo.Gender = txtGender.Text
        demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(txtEthnicity.Text, eTables.Pred_ref_ethnicities)
        demo.Ethnicity = txtEthnicity.Text
        demo.TestDate = txtTestDate.Text
        demo.SourcesString = txtPredsRaw.Text

        gdicPreds = cPred.Get_PredValues(demo, Method)
        txtPredsRaw.Text = cPred.Get_PrefSources_AsCodedString(class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate, demo)
        txtPreds.Text = cPred.Get_PrefSources_AsFormattedString(CDate(txtTestDate.Text), demo)

        cPred.LoadLbl(lblPn_Fev1, "FEV1|LLN", gdicPreds, 2)
        cPred.LoadLbl(lblPn_Fvc, "FVC|LLN", gdicPreds, 2)
        cPred.LoadLbl(lblPn_Vc, "VC|LLN", gdicPreds, 2)
        cPred.LoadLbl(lblPn_Fer, "FER|LLN", gdicPreds, 1)
        cPred.LoadLbl(lblPn_Pef, "PEF|LLN", gdicPreds, 2)
        cPred.LoadLbl(lblPn_Fvc, "FVC|LLN", gdicPreds, 2)
        cPred.LoadLbl(lblPn_FEF2575, "FEF2575|LLN", gdicPreds, 2)

    End Sub

    Private Sub DrawGraph()

        'grdD.CurrentCell.Value = fmt(grdD.CurrentCell.Value, 2)

        'Load results from table into passing array
        Dim i As Integer = 0
        For Each r As DataGridViewRow In grdD.Rows
            _pData.doses(i).result = r.Cells.Item(2).Value
            _pData.doses(i).response = r.Cells.Item(3).Value
            i = i + 1
        Next

        'Draw the graph
        Dim bmp As Bitmap = cChall.Draw_ProvocationGraph(_pData, pic.Height, pic.Width, New Font("Arial", 8))
        pic.Image = bmp

    End Sub

    Private Sub LockForm(ByVal LockState As Boolean)
        'True=locked, false=unlocked

        txtReport.ReadOnly = LockState
        pnlSpirometry.Enabled = Not LockState
        pnlReporters.Enabled = Not LockState
        pnlOther.Enabled = Not LockState
        pnlPreds.Enabled = Not LockState

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    Private Sub txtR_bl_Fev1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Fev1.LostFocus
        txtR_bl_Fev1.Text = fmt(txtR_bl_Fev1.Text, 2)
        grdD.Item(2, 0).Value = txtR_bl_Fev1.Text
        Do_calculations()
    End Sub
    Private Sub txtR_bl_Fvc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Fvc.LostFocus
        txtR_bl_Fvc.Text = fmt(txtR_bl_Fvc.Text, 2)
        Do_calculations()
    End Sub
    Private Sub txtR_bl_vc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Vc.LostFocus
        txtR_bl_Vc.Text = fmt(txtR_bl_Vc.Text, 2)
        Do_calculations()
    End Sub
    Private Sub txtR_bl_Fef2575_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Fef2575.LostFocus
        txtR_bl_Fef2575.Text = fmt(txtR_bl_Fef2575.Text, 2)
        Do_calculations()
    End Sub
    Private Sub txtR_bl_Pef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Pef.LostFocus
        txtR_bl_Pef.Text = fmt(txtR_bl_Pef.Text, 2)
        Do_calculations()
    End Sub

    Private Sub cmboTechnicalNote_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim kv As KeyValuePair(Of String, Integer) = cmboTechnicalNote.SelectedItem
        If Len(txtTechnicalNote.Text) = 0 Then txtTechnicalNote.Text = kv.Key Else txtTechnicalNote.Text = txtTechnicalNote.Text & " " & kv.Key
        cmboTechnicalNote.SelectedIndex = -1
        txtTechnicalNote.Focus()
        txtTechnicalNote.SelectionStart = txtTechnicalNote.Text.Length
        txtTechnicalNote.SelectionLength = 0

    End Sub


    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim provID As Long = 0
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
            provID = cRfts.insert_rft_genericProv(sessionID, dTest, dTestData)
        Else
            Dim ReturnValue As Boolean = cRfts.Update_rft_genericProv(sessionID, dTest, dTestData)
            provID = CLng(Me._RecordID)
        End If

        'Save flow vol
        cDAL.Update_image(provID, picFV, eTables.Prov_test, "flowvolloop")

        'Refresh the list of tests and pdf to reflect any changes
        If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

    End Sub

    Private Sub btnLoadfromfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadfromfile.Click

        Dim f As String = ""
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "JPG image files|*.jpg|GIF image files|*.gif"
        openFileDialog1.Title = "Select a flow volume loop image file"

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            picFV.Load(openFileDialog1.FileName)
        End If

    End Sub

    Private Sub btnClearFlowVol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearFlowVol.Click

        picFV.Image = Nothing

    End Sub

    Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click

        cRfts.Delete_rft(Me._RecordID, eTables.Prov_test)

        'Refresh the list of tests and pdf to reflect any changes
        If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

    End Sub

    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked

        Select Case UCase(e.ClickedItem.Text)
            Case "PASTE IMAGE" : If Clipboard.ContainsImage() Then picFV.Image = Clipboard.GetImage()
            Case "CLEAR IMAGE" : picFV.Image = Nothing
        End Select

    End Sub

    Private Sub btnSession_Click(sender As System.Object, e As System.EventArgs) Handles btnSession.Click
        _form_session._cancelButton_enable = False
        _form_session.Visible = True
        _form_session.BringToFront()
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

        Dim dicR As Dictionary(Of String, String) = Me.Load_MemFromScreenFields_test(False)

        _formPhrases = New form_rft_phrases(eAutoreport_testgroups.Provocation, Me, demo, dicR)
        _formPhrases.ReportTextBox = Me.txtReport
        _formPhrases.Show()

    End Sub
End Class