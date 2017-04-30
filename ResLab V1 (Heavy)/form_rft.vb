Imports System.Windows.Forms.DataVisualization.Charting
Imports ResLab_V1_Heavy.cDatabaseInfo


Public Class form_Rft
    Dim i As Integer = 0
    Dim _suppressFormRefresh As Boolean = False
    Dim flagFormLoad As Boolean = False
    Dim gdicPreds As New Dictionary(Of String, String)     'ParameterID|StatTypeID, result
    Dim chrt As Chart = Nothing

    Private _CallingFormName As String
    Private _NewReportstatus As String

    Private _RecordID As Long = 0
    Private _PatientID As Long = 0
    Private _isNewRft As Boolean
    Private _form_session As form_rft_session
    Private _formPhrases As form_rft_phrases

    Public Sub UpdateReportStatusTo(ByVal NewReportStatus As String)
        _NewReportstatus = NewReportStatus
    End Sub


    Public Sub New(ByVal RecordID As Long, ByVal CallingForm As Form)

        InitializeComponent()
        _CallingFormName = CallingForm.Name
        _RecordID = RecordID

        Select Case _RecordID
            Case 0 : _isNewRft = True
            Case Else : _isNewRft = False
        End Select

        'use a local patientid because can come from reporting list or logged pt
        Select Case _RecordID
            Case 0 : _PatientID = cPt.PatientID
            Case Is > 0 : _PatientID = cPt.Get_PatientIDFromPkID(_RecordID, eTables.rft_routine)
        End Select

    End Sub

    Private Sub RefreshFromSession()

        'Refresh relevant info from session form 

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
        Me.Do_calculations()

    End Sub

    Private Sub Load_Normals(ByVal Method As class_Pred.eLoadNormalsMode)

        Dim demo As Pred_demo = Nothing

        demo.Age = Val(cMyRoutines.Calc_Age(txtDOB.Text, txtTestDate.Text))
        demo.DOB = txtDOB.Text
        demo.Htcm = Val(txtHeight.Text)
        demo.Wtkg = Val(txtWeight.Text)
        demo.GenderID = cMyRoutines.Lookup_list_ByDescription(txtGender.Text, eTables.Pred_ref_genders)
        demo.Gender = txtGender.Text
        demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(txtEthnicity.Text, eTables.Pred_ref_ethnicities)
        demo.Ethnicity = txtEthnicity.Text
        If IsDate(txtTestDate.Text) Then demo.TestDate = CDate(txtTestDate.Text)
        demo.SourcesString = txtPredsRaw.Text

        gdicPreds = cPred.Get_PredValues(demo, Method)
        txtPredsRaw.Text = cPred.Get_PrefSources_AsCodedString(class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate, demo)
        txtPreds.Text = cPred.Get_PrefSources_AsFormattedString(CDate(txtTestDate.Text), demo)

        If Not (gdicPreds Is Nothing) Then
            cPred.LoadLbl(lblPn_Fev1, "FEV1|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_Fvc, "FVC|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_Vc, "VC|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_Fer, "FER|LLN", gdicPreds, 1)
            cPred.LoadLbl(lblPn_Pef, "PEF|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_Fvc, "FVC|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_FEF2575, "FEF2575|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_tlco, "TLCO|LLN", gdicPreds, 1)
            cPred.LoadLbl(lblPn_tlcohb, "TLCO|LLN", gdicPreds, 1)
            cPred.LoadLbl(lblPn_kco, "KCO|Range", gdicPreds, 1)
            cPred.LoadLbl(lblPn_kcohb, "KCO|Range", gdicPreds, 1)
            cPred.LoadLbl(lblPn_va, "VA|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_tlc, "TLC|Range", gdicPreds, 1)
            cPred.LoadLbl(lblPn_frc, "FRC|Range", gdicPreds, 1)
            cPred.LoadLbl(lblPn_rv, "RV|Range", gdicPreds, 1)
            cPred.LoadLbl(lblPn_rvtlc, "RV/TLC|ULN", gdicPreds, 0)
            cPred.LoadLbl(lblPn_lvvc, "VC|LLN", gdicPreds, 2)
            cPred.LoadLbl(lblPn_mip, "MIP|LLN", gdicPreds, 0)
            cPred.LoadLbl(lblPn_mep, "MEP|LLN", gdicPreds, 0)
            cPred.LoadLbl(lblPn_feno, "FENO|LLN", gdicPreds, 0)
            cPred.LoadLbl(lblPn_spo2, "SpO2|Range", gdicPreds, 0)
        End If

    End Sub


    Private Sub Do_calculations()

        Me.Load_Normals(class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)

        'Baseline FER - use the larger of VC and FVC
        If Val(txtR_bl_Fev1.Text) > 0 Then
            If Val(txtR_bl_Fvc.Text) > Val(txtR_bl_Vc.Text) Then
                If Val(txtR_bl_Fvc.Text) > 0 Then
                    txtR_bl_Fer.Text = f(100 * Val(txtR_bl_Fev1.Text) / Val(txtR_bl_Fvc.Text), 0)
                Else
                    txtR_bl_Fer.Text = ""
                End If
            Else
                If Val(txtR_bl_Vc.Text) > 0 Then
                    txtR_bl_Fer.Text = f(100 * Val(txtR_bl_Fev1.Text) / Val(txtR_bl_Vc.Text), 0)
                Else
                    txtR_bl_Fer.Text = ""
                End If
            End If
        End If

        'Post FER - use the larger of VC and FVC
        If Val(txtR_Post_Fev1.Text) > 0 Then
            If Val(txtR_Post_Fvc.Text) > Val(txtR_Post_Vc.Text) Then
                If Val(txtR_Post_Fvc.Text) > 0 Then
                    txtR_Post_Fer.Text = f(100 * Val(txtR_Post_Fev1.Text) / Val(txtR_Post_Fvc.Text), 0)
                Else
                    txtR_Post_Fer.Text = ""
                End If
            Else
                If Val(txtR_Post_Vc.Text) > 0 Then
                    txtR_Post_Fer.Text = f(100 * Val(txtR_Post_Fev1.Text) / Val(txtR_Post_Vc.Text), 0)
                Else
                    txtR_Post_Fer.Text = ""
                End If
            End If
        End If

        '% change
        If Val(txtR_bl_Fev1.Text) > 0 And Val(txtR_Post_Fev1.Text) > 0 Then
            txtR_PostCh_Fev1.Text = fChange(Val(txtR_Post_Fev1.Text) / Val(txtR_bl_Fev1.Text) - 1, 0)
        End If
        If Val(txtR_bl_Fvc.Text) > 0 And Val(txtR_Post_Fvc.Text) > 0 Then
            txtR_PostCh_Fvc.Text = fChange(Val(txtR_Post_Fvc.Text) / Val(txtR_bl_Fvc.Text) - 1, 0)
        End If
        If Val(txtR_bl_Vc.Text) > 0 And Val(txtR_Post_Vc.Text) > 0 Then
            txtR_PostCh_Vc.Text = fChange(Val(txtR_Post_Vc.Text) / Val(txtR_bl_Vc.Text) - 1, 0)
        End If
        If Val(txtR_bl_Fer.Text) > 0 And Val(txtR_Post_Fer.Text) > 0 Then
            txtR_PostCh_Fer.Text = fChange(Val(txtR_Post_Fer.Text) / Val(txtR_bl_Fer.Text) - 1, 0)
        End If
        If Val(txtR_bl_Fef2575z.Text) > 0 And Val(txtR_Post_Fef2575.Text) > 0 Then
            txtR_PostCh_Fef2575.Text = fChange(Val(txtR_Post_Fef2575.Text) / Val(txtR_bl_Fef2575z.Text) - 1, 0)
        End If
        If Val(txtR_bl_Pef.Text) > 0 And Val(txtR_Post_Pef.Text) > 0 Then
            txtR_PostCh_Pef.Text = fChange(Val(txtR_Post_Pef.Text) / Val(txtR_bl_Pef.Text) - 1, 0)
        End If

        'KCO 
        If Val(txtR_Bl_Tlco.Text) > 0 And Val(txtR_Bl_Va.Text) > 0 Then
            txtR_Bl_Kco.Text = f(Val(txtR_Bl_Tlco.Text) / Val(txtR_Bl_Va.Text), 2)
        End If

        'Hb corrections for TLCO
        Dim HbFactor As Single = cMyRoutines.calc_HbFac(txtR_Bl_Hb.Text, txtDOB.Text, txtGender.Text, txtTestDate.Text)
        If HbFactor > 0 Then
            If Val(txtR_Bl_Tlco.Text) > 0 Then
                txtR_Bl_TlcoHb.Text = f(HbFactor * Val(txtR_Bl_Tlco.Text), 1)
            Else
                txtR_Bl_TlcoHb.Text = ""
            End If
            If Val(txtR_Bl_Kco.Text) > 0 Then
                txtR_Bl_KcoHb.Text = f(HbFactor * Val(txtR_Bl_Kco.Text), 1)
            Else
                txtR_Bl_KcoHb.Text = ""
            End If
        Else
            txtR_Bl_TlcoHb.Text = ""
            txtR_Bl_KcoHb.Text = ""
        End If

        'RV/TLC
        If Val(txtR_Bl_Rv.Text) > 0 And Val(txtR_Bl_Tlc.Text) > 0 Then
            txtR_Bl_RvTlc.Text = f(100 * Val(txtR_Bl_Rv.Text) / Val(txtR_Bl_Tlc.Text), 0)
        Else
            txtR_Bl_RvTlc.Text = ""
        End If

        'BMI
        txtBMI.Text = cMyRoutines.calc_BMI(txtHeight.Text, txtWeight.Text)

        'Age
        txtAge.Text = "(" & cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text)) & ")"

        '% predicteds
        cPred.LoadPPN(txtR_blPpn_Fev1, txtR_bl_Fev1, "FEV1|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Fvc, txtR_bl_Fvc, "FVC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Vc, txtR_bl_Vc, "VC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Fer, txtR_bl_Fer, "FER|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Pef, txtR_bl_Pef, "PEF|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Fef2575, txtR_bl_Fef2575z, "FEF2575|MPV", gdicPreds)
        cPred.LoadPPN(txtR_PostPpn_Fev1, txtR_Post_Fev1, "FEV1|MPV", gdicPreds)
        cPred.LoadPPN(txtR_PostPpn_Fvc, txtR_Post_Fvc, "FVC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_PostPpn_Vc, txtR_Post_Vc, "FVC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_PostPpn_Fer, txtR_Post_Fer, "FER|MPV", gdicPreds)
        cPred.LoadPPN(txtR_PostPpn_Pef, txtR_Post_Pef, "PEF|MPV", gdicPreds)
        cPred.LoadPPN(txtR_PostPpn_Fef2575, txtR_Post_Fef2575, "FEF2575|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Tlco, txtR_Bl_Tlco, "TLCO|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_TlcoHb, txtR_Bl_TlcoHb, "TLCO|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Kco, txtR_Bl_Kco, "KCO|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_KcoHb, txtR_Bl_KcoHb, "KCO|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Va, txtR_Bl_Va, "VA|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Tlc, txtR_Bl_Tlc, "TLC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Frc, txtR_Bl_Frc, "FRC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Rv, txtR_Bl_Rv, "RV|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_RvTlc, txtR_Bl_RvTlc, "RV/TLC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_LvVc, txtR_Bl_LvVc, "VC|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Mip, txtR_Bl_Mip, "MIP|MPV", gdicPreds)
        cPred.LoadPPN(txtR_blPpn_Mep, txtR_Bl_Mep, "MEP|MPV", gdicPreds)

        'Pack tears
        txtPackYears.Text = cMyRoutines.calc_PackYears(txtCigsPerDay.Text, txtYearsSmoked.Text)

    End Sub

    Private Function f(ByVal Num As Single, ByVal DecPlaces As Integer) As String
        'takes a number as string and returns as string to selected decimal places

        Select Case DecPlaces
            Case 0 : Return Format(Num, "###")
            Case 1 : Return Format(Num, "###.0")
            Case 2 : Return Format(Num, "###.00")
            Case Else : Return Format(Num, "###.00")
        End Select

    End Function

    Private Function fChange(ByVal Num As Single, ByVal DecPlaces As Integer) As String
        'takes a number as string and returns as string to selected decimal places

        Dim Plus As String

        If Num >= 0 Then Plus = "+" Else Plus = ""
        Select Case DecPlaces
            Case 0 : Return Format(Num, Plus & "0%")
            Case 1 : Return Format(Num, Plus & "0.0%")
            Case 2 : Return Format(Num, Plus & "0.00%")
            Case Else : Return Format(Num, Plus & "0.00%")
        End Select

    End Function

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub form_Rft_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Me.RefreshFromSession()

    End Sub

    Private Sub form_Rft_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not IsNothing(_formPhrases) Then
            If _formPhrases.Visible Then _formPhrases.Close()
        End If

    End Sub

    Private Sub form_Rft_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        flagFormLoad = True

        'Do session
        _form_session = New form_rft_session(_PatientID, _RecordID, eTables.rft_routine, Me, _isNewRft)
        If _isNewRft Then
            Dim response As Integer = _form_session.ShowDialog()
            If response = DialogResult.Cancel Then
                Me.Close()
                Exit Sub
            End If
        End If
        Me.RefreshFromSession()
        Me.Text = cPt.Get_PtNameString(_PatientID, eNameStringFormats.Name_UR)

        'Load combo box options
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLab, "Labs", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromList(cmboReportStatus, eTables.List_ReportStatuses)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboTechnicalNote, "Technical comments", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboScientist, "Scientist", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboPost_Condition, "Post spirometry condition", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLVmethod, "LungVolumeMethods", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboR_fio2_1, "FiO2", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboR_fio2_2, "FiO2", _isNewRft)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLastBD, "Last_BD", _isNewRft)

        If _CallingFormName = "form_Reporting" Then
            If cmboReportedBy.Text = "" Then cmboReportedBy.Text = form_Reporting.cmboDefaultReporter.Text
            If txtReportedDate.Text = "" Then txtReportedDate.Text = form_Reporting.txtDefaultDate.Text
        End If

        'Load test 
        Me.Load_ScreenFieldsFromMem(Me._isNewRft, Me._RecordID)
        Me.Do_calculations()

        txtR_bl_Fev1.Focus()

        cUser.set_access(Me)

        flagFormLoad = False

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        'Save session
        Dim sessionID As Long = _form_session._sessionID
        Select Case sessionID
            Case 0 : sessionID = cRfts.Insert_rft_session(Me.Load_MemFromScreenFields_session(True))
            Case Else : cRfts.Update_rft_session(Me.Load_MemFromScreenFields_session(True))
        End Select

        'Save Rft
        Select Case Me._isNewRft
            Case True : Me._RecordID = cRfts.insert_rft_routine(sessionID, Me.Load_MemFromScreenFields_rft(True))
            Case False : cRfts.update_rft_routine(sessionID, Me.Load_MemFromScreenFields_rft(True))
        End Select

        'Save flow vol
        cDAL.Update_image(Me._RecordID, picFV, eTables.rft_routine, "flowvolloop")

        'Refresh the list of tests and pdf to reflect any changes
        If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

    End Sub

    Private Sub LockForm(ByVal LockState As Boolean)
        'True=locked, false=unlocked

        txtReport.ReadOnly = LockState
        pnlReporters.Enabled = Not LockState
        pnlOther.Enabled = Not LockState
        pnlPreds.Enabled = Not LockState

    End Sub


    Private Function Load_MemFromScreenFields_session(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Session
        Dim R As New class_Rft_RoutineAndSessionFields

        d(R.PatientID) = Me._PatientID
        d(R.SessionID) = _form_session._sessionID
        d(R.TestDate) = cMyRoutines.FormatDBDate(txtTestDate.Text)
        d(R.Height) = q & _form_session.txtHt.Text & q
        d(R.Weight) = q & _form_session.txtWt.Text & q
        d(R.Smoke_hx) = q & Trim(_form_session.cmboSmoke_hx.Text) & q
        d(R.Smoke_yearssmoked) = q & Trim(_form_session.txtSmoke_yrs.Text) & q
        d(R.Smoke_cigsperday) = q & Trim(_form_session.txtSmoke_cigsperday.Text) & q
        d(R.Smoke_packyears) = q & Trim(_form_session.txtSmoke_packyears.Text) & q
        d(R.Req_date) = cMyRoutines.FormatDBDate(_form_session.txtRequestDate.Text)
        d(R.Req_name) = q & Trim(_form_session.cmboReqMO_name.Text) & q
        d(R.Req_address) = q & Trim(_form_session.cmboReqMO_address.Text) & q
        d(R.Req_phone) = q & Trim(_form_session.txtReqMO_phone.Text) & q
        d(R.Req_fax) = q & Trim(_form_session.txtReqMO_fax.Text) & q
        d(R.Req_email) = q & Trim(_form_session.txtReqMO_email.Text) & q
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

    Private Function Load_MemFromScreenFields_rft(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Routine
        Dim R As New class_Rft_RoutineAndSessionFields

        dicR(R.PatientID) = Me._PatientID
        dicR(R.SessionID) = Nothing     'Set in btn_save_click routine
        dicR(R.RftID) = Me._RecordID
        'dicR(R.TestDate) = txtTestDate.Text
        dicR(R.TestTime) = cMyRoutines.FormatDBTime(txtTestTime.Text)
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

        dicR(R.LungVolumes_method) = q & cmboLVmethod.Text & q
        dicR(R.R_bl_Fev1) = q & txtR_bl_Fev1.Text & q
        dicR(R.R_bl_Fvc) = q & txtR_bl_Fvc.Text & q
        dicR(R.R_bl_Vc) = q & txtR_bl_Vc.Text & q
        dicR(R.R_Bl_Fer) = q & txtR_bl_Fer.Text & q
        dicR(R.R_bl_Fef2575) = q & txtR_bl_Fef2575z.Text & q
        dicR(R.R_bl_Pef) = q & txtR_bl_Pef.Text & q
        dicR(R.R_Bl_Tlco) = q & txtR_Bl_Tlco.Text & q
        dicR(R.R_Bl_Kco) = q & txtR_Bl_Kco.Text & q
        dicR(R.R_Bl_Va) = q & txtR_Bl_Va.Text & q
        dicR(R.R_Bl_Hb) = q & txtR_Bl_Hb.Text & q
        dicR(R.R_Bl_Ivc) = q & txtR_Bl_Ivc.Text & q
        dicR(R.R_Bl_Tlc) = q & txtR_Bl_Tlc.Text & q
        dicR(R.R_Bl_Frc) = q & txtR_Bl_Frc.Text & q
        dicR(R.R_Bl_Rv) = q & txtR_Bl_Rv.Text & q
        dicR(R.R_Bl_RvTlc) = q & txtR_Bl_RvTlc.Text & q
        dicR(R.R_Bl_LvVc) = q & txtR_Bl_LvVc.Text & q
        dicR(R.R_Bl_Mip) = q & txtR_Bl_Mip.Text & q
        dicR(R.R_Bl_Mep) = q & txtR_Bl_Mep.Text & q

        dicR(R.R_Bl_FeNO) = q & txtR_Bl_feno.Text & q
        dicR(R.R_post_Fev1) = q & txtR_Post_Fev1.Text & q
        dicR(R.R_post_Fvc) = q & txtR_Post_Fvc.Text & q
        dicR(R.R_post_Vc) = q & txtR_Post_Vc.Text & q
        dicR(R.R_Post_Fer) = q & txtR_Post_Fer.Text & q
        dicR(R.R_post_Fef2575) = q & txtR_Post_Fef2575.Text & q
        dicR(R.R_post_Pef) = q & txtR_Post_Pef.Text & q
        dicR(R.R_post_Condition) = q & cmboPost_Condition.Text & q

        dicR(R.R_SpO2_1) = q & txtR_spo2_1.Text & q
        dicR(R.R_SpO2_2) = q & txtR_spo2_2.Text & q

        dicR(R.R_abg1_be) = q & txtR_be_1.Text & q
        dicR(R.R_abg1_fio2) = q & cmboR_fio2_1.Text & q
        dicR(R.R_abg1_hco3) = q & txtR_hco3_1.Text & q
        dicR(R.R_abg1_paco2) = q & txtR_paco2_1.Text & q
        dicR(R.R_abg1_pao2) = q & txtR_pao2_1.Text & q
        dicR(R.R_abg1_ph) = q & txtR_ph_1.Text & q
        dicR(R.R_abg1_sao2) = q & txtR_sao2_1.Text & q
        dicR(R.R_abg1_shunt) = q & txtR_shunt_1.Text & q
        dicR(R.R_abg2_be) = q & txtR_be_2.Text & q
        dicR(R.R_abg2_fio2) = q & cmboR_fio2_2.Text & q
        dicR(R.R_abg2_hco3) = q & txtR_hco3_2.Text & q
        dicR(R.R_abg2_paco2) = q & txtR_paco2_2.Text & q
        dicR(R.R_abg2_pao2) = q & txtR_pao2_2.Text & q
        dicR(R.R_abg2_ph) = q & txtR_ph_2.Text & q
        dicR(R.R_abg2_sao2) = q & txtR_sao2_2.Text & q
        dicR(R.R_abg2_shunt) = q & txtR_shunt_2.Text & q

        dicR(R.TestType) = q & cMyRoutines.Get_TestsDone(dicR).TestsDoneString & q

        dicR(R.LastUpdated_rft) = cMyRoutines.FormatDBDateTime(Date.Now)
        dicR(R.LastUpdatedBy_rft) = q & My.User.Name & q

        Return dicR

    End Function

    Private Function Load_ScreenFieldsFromMem(IsNewRft As Boolean, RftID As Long) As Boolean

        Select Case IsNewRft
            Case True
                cmboReportStatus.Text = "Unreported"


                Return True
            Case False
                Dim dic As Dictionary(Of String, String) = cRfts.Get_rft_byRftID(RftID)
                If IsNothing(dic) Then
                    'should never happen
                    Return False
                Else
                    'Load the test data
                    Dim R As New class_Rft_RoutineAndSessionFields
                    txtReport.Text = dic(R.Report_text)
                    If IsDate(dic(R.Report_reporteddate)) Then txtReportedDate.Text = dic(R.Report_reporteddate)
                    cmboReportedBy.Text = dic(R.Report_reportedby)
                    cmboVerifiedBy.Text = dic(R.Report_verifiedby)
                    If IsDate(dic(R.Report_verifieddate)) Then txtVerifiedDate.Text = dic(R.Report_verifieddate)
                    cmboReportStatus.Text = dic(R.Report_status)
                    cmboScientist.Text = dic(R.Scientist)
                    If IsDate(dic(R.TestTime)) Then txtTestTime.Text = dic(R.TestTime)
                    cmboLastBD.Text = dic(R.BDStatus)
                    cmboLab.Text = dic(R.Lab)
                    txtTechnicalNote.Text = dic(R.TechnicalNotes)

                    cmboLVmethod.Text = dic(R.LungVolumes_method)
                    txtR_bl_Fev1.Text = dic(R.R_bl_Fev1)
                    txtR_bl_Fvc.Text = dic(R.R_bl_Fvc)
                    txtR_bl_Vc.Text = dic(R.R_bl_Vc)
                    txtR_bl_Fer.Text = dic(R.R_Bl_Fer)
                    txtR_bl_Fef2575z.Text = dic(R.R_bl_Fef2575)
                    txtR_bl_Pef.Text = dic(R.R_bl_Pef)
                    txtR_Bl_Tlco.Text = dic(R.R_Bl_Tlco)
                    txtR_Bl_Kco.Text = dic(R.R_Bl_Kco)
                    txtR_Bl_Va.Text = dic(R.R_Bl_Va)
                    txtR_Bl_Hb.Text = dic(R.R_Bl_Hb)
                    txtR_Bl_Ivc.Text = dic(R.R_Bl_Ivc)
                    txtR_Bl_Tlc.Text = dic(R.R_Bl_Tlc)
                    txtR_Bl_Frc.Text = dic(R.R_Bl_Frc)
                    txtR_Bl_Rv.Text = dic(R.R_Bl_Rv)
                    txtR_Bl_LvVc.Text = dic(R.R_Bl_LvVc)
                    txtR_Bl_RvTlc.Text = dic(R.R_Bl_RvTlc)
                    txtR_Bl_Mip.Text = dic(R.R_Bl_Mip)
                    txtR_Bl_Mep.Text = dic(R.R_Bl_Mep)
                    txtR_Bl_feno.Text = dic(R.R_Bl_FeNO)

                    txtR_Post_Fev1.Text = dic(R.R_post_Fev1)
                    txtR_Post_Fvc.Text = dic(R.R_post_Fvc)
                    txtR_Post_Vc.Text = dic(R.R_post_Vc)
                    txtR_Post_Fer.Text = dic(R.R_Post_Fer)
                    txtR_Post_Fef2575.Text = dic(R.R_post_Fef2575)
                    txtR_Post_Pef.Text = dic(R.R_post_Pef)
                    cmboPost_Condition.Text = dic(R.R_post_Condition)

                    txtR_spo2_1.Text = dic(R.R_SpO2_1)
                    txtR_spo2_2.Text = dic(R.R_SpO2_2)

                    cmboR_fio2_1.Text = dic(R.R_abg1_fio2)
                    txtR_be_1.Text = dic(R.R_abg1_be)
                    txtR_hco3_1.Text = dic(R.R_abg1_hco3)
                    txtR_paco2_1.Text = dic(R.R_abg1_paco2)
                    txtR_pao2_1.Text = dic(R.R_abg1_pao2)
                    txtR_ph_1.Text = dic(R.R_abg1_ph)
                    txtR_sao2_1.Text = dic(R.R_abg1_sao2)
                    txtR_shunt_1.Text = dic(R.R_abg1_shunt)

                    cmboR_fio2_2.Text = dic(R.R_abg2_fio2)
                    txtR_be_2.Text = dic(R.R_abg2_be)
                    txtR_hco3_2.Text = dic(R.R_abg2_hco3)
                    txtR_paco2_2.Text = dic(R.R_abg2_paco2)
                    txtR_pao2_2.Text = dic(R.R_abg2_pao2)
                    txtR_ph_2.Text = dic(R.R_abg2_ph)
                    txtR_sao2_2.Text = dic(R.R_abg2_sao2)
                    txtR_shunt_2.Text = dic(R.R_abg2_shunt)

                    'Load predicteds viewer
                    txtPreds.Text = cPred.Decode_SourcesStringForDisplay(dic("pred_sourceids"))
                    txtPredsRaw.Text = dic("pred_sourceids")

                    'Load flow vol loop
                    cDAL.Get_image(CLng(Me._RecordID), picFV, eTables.rft_routine, "flowvolloop")

                    dic = Nothing
                    R = Nothing
                    Return True
                End If
            Case Else
                Return False
        End Select

    End Function

    Private Sub Label67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtTestDate.Text = Format(Now, "dd/MM/yyyy")
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        If cRfts.Delete_rft(Me._RecordID, eTables.rft_routine) Then
            'Refresh the list of tests and pdf to reflect any changes
            If Me._CallingFormName = "Form_MainNew" Then gRefreshMainForm = True
            Me.Close()
        Else
            'Keep the form open if delete is cancelled
        End If

    End Sub

    Private Sub Label40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Not IsDate(txtTestTime.Text) Then txtTestTime.Text = Format(Now, "HH:mm")

    End Sub

    Private Sub Label85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Not IsDate(txtReportedDate.Text) Then
            txtReportedDate.Text = Format(Now, "dd/MM/yyyy")
        End If

    End Sub

    Private Sub Label61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not IsDate(txtVerifiedDate.Text) Then
            txtVerifiedDate.Text = Format(Now, "dd/MM/yyyy")
        End If
    End Sub

    Private Sub btnLoadfromfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadfromfile.Click

        Dim defaultpath As String = cMyRoutines.Get_AppString("flowvolimage_filepath")
        Dim openFileDialog1 As New OpenFileDialog()

        If defaultpath = "" Then
            openFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Temp
        Else
            openFileDialog1.InitialDirectory = defaultpath
        End If

        openFileDialog1.Filter = "JPG image files|*.jpg|GIF image files|*.gif"
        openFileDialog1.Title = "Select a flow volume loop image file"

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            picFV.Load(openFileDialog1.FileName)
        End If

        If openFileDialog1.FileName <> "" Then cMyRoutines.Update_AppString("flowvolimage_filepath", IO.Path.GetDirectoryName(openFileDialog1.FileName))
        openFileDialog1.Dispose()

    End Sub

    Private Sub btnClearFlowVol_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearFlowVol.Click

        picFV.Image = Nothing

    End Sub

    Private Sub txtR_bl_Fev1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Fev1.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_bl_Fvc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Fvc.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_bl_vc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Vc.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_bl_Fef2575_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Fef2575z.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_bl_Pef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_bl_Pef.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Post_Fev1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Post_Fev1.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Post_Fvc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Post_Fvc.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Post_vc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Post_Vc.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Post_Fef2575_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Post_Fef2575.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_post_Pef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Post_Pef.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Tlco_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Tlco.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Va_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Va.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Hb_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Hb.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Tlc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Tlc.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Frc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Frc.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Rv_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Rv.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Mip_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Mip.LostFocus
        Do_calculations()
    End Sub
    Private Sub txtR_Bl_Mep_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_Mep.LostFocus
        Do_calculations()
    End Sub

    Private Sub txtHeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Do_calculations()
    End Sub

    Private Sub txtWeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Do_calculations()
    End Sub

    Private Sub txtTestDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Do_calculations()
    End Sub

    Private Sub txtR_Bl_LvVc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtR_Bl_LvVc.LostFocus
        Do_calculations()
    End Sub

    Private Sub tabResults_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles tabResults.SelectedIndexChanged

        Select Case tabResults.SelectedIndex
            Case 2
                Dim p As class_plot_trend.plotproperties = Form_MainNew.get_trendplotproperties(Me.grdTrend)
                If cPt.PatientID > 0 Then
                    'Build table 
                    splitTrend.Panel1.Controls.Clear()
                    grdTrend = cTrend.Create_trend_table(Me._PatientID, p.GridCanMultiselect)
                    splitTrend.Panel1.Controls.Add(grdTrend)


                    'Kill off any existing graph
                    splitTrend.Panel2.Controls.Clear()
                    chrt = cTrend.Create_trend_plot(p)
                    splitTrend.Panel2.Controls.Add(chrt)

                    tabResults.Dock = DockStyle.Top
                    tabResults.BringToFront()

                Else
                    tabResults.Dock = DockStyle.None
                End If
            Case Else
                tabResults.Dock = DockStyle.None
        End Select

    End Sub


    Private Sub ContextMenuStrip1_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked

        Select Case UCase(e.ClickedItem.Text)
            Case "PASTE IMAGE" : If Clipboard.ContainsImage() Then picFV.Image = Clipboard.GetImage()
            Case "CLEAR IMAGE" : picFV.Image = Nothing
        End Select

    End Sub

    Private Sub ContextMenuStrip_phrases_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip_phrases.ItemClicked

        Select Case UCase(e.ClickedItem.Text)
            Case ""
            Case Else
        End Select

    End Sub

    Private Sub ContextMenuStrip_phrases_LostFocus(sender As Object, e As System.EventArgs) Handles ContextMenuStrip_phrases.LostFocus

        txtReport.Cursor = Cursors.IBeam

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

        Dim dicR As Dictionary(Of String, String) = Me.Load_MemFromScreenFields_rft(False)

        _formPhrases = New form_rft_phrases(eAutoreport_testgroups.RoutineRft, Me, demo, dicR)
        _formPhrases.ReportTextBox = Me.txtReport
        _formPhrases.Show()


    End Sub

    Private Sub txtCigsPerDay_LostFocus(sender As Object, e As System.EventArgs)

        Me.Do_calculations()

    End Sub

    Private Sub txtYearsSmoked_LostFocus(sender As Object, e As System.EventArgs)

        Me.Do_calculations()

    End Sub

    Private Sub grdTrend_Click(sender As Object, e As System.EventArgs) Handles grdTrend.Click

        Dim r As Integer = 0
        Dim c As Integer = 0

        _suppressFormRefresh = True

        'Display report text
        If grdTrend.SelectedRows.Count > 0 And grdTrend.CurrentCell.ColumnIndex >= 0 Then
            r = grdTrend.SelectedRows(0).Index
            c = grdTrend.CurrentCell.ColumnIndex
            If r > grdTrend.RowCount - 3 Then
                Select Case grdTrend.SelectedRows(0).HeaderCell.Value
                    Case "Technical note"
                        Dim Msg As String = "TECHNICAL NOTE:" & vbCrLf & grdTrend(c, r).Value
                        Msg = Msg & vbCrLf & vbCrLf & "REPORT:" & vbCrLf & grdTrend(c, r + 1).Value
                        MsgBox(Msg, vbOKOnly, "Report")
                    Case "Report"
                        Dim Msg As String = "TECHNICAL NOTE:" & vbCrLf & grdTrend(c, r - 1).Value
                        Msg = Msg & vbCrLf & vbCrLf & "REPORT:" & vbCrLf & grdTrend(c, r).Value
                        MsgBox(Msg, vbOKOnly, "Report")
                End Select
            End If
        End If

        _suppressFormRefresh = False

    End Sub

    Private Sub grdTrend_RowHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles grdTrend.RowHeaderMouseClick

        Dim r As Integer = grdTrend.SelectedRows(0).Index
        If r > 2 And r < grdTrend.RowCount - 2 Then
            'Kill off any existing graph
            splitTrend.Panel2.Controls.Clear()
            'Generate and display graph
            Dim p As class_plot_trend.plotproperties = Form_MainNew.get_trendplotproperties(Me.grdTrend)
            chrt = cTrend.Create_trend_plot(p)
            grdTrend.MultiSelect = p.GridCanMultiselect
            splitTrend.Panel2.Controls.Add(chrt)
        End If

    End Sub



    Private Sub cmboLab_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)

    End Sub
    Private Sub txtLastBD_MaskInputRejected(sender As System.Object, e As System.Windows.Forms.MaskInputRejectedEventArgs)

    End Sub

    Private Sub btnSession_Click(sender As System.Object, e As System.EventArgs) Handles btnSession.Click
        _form_session._cancelButton_enable = False
        _form_session.Visible = True
        _form_session.BringToFront()
    End Sub

    Private Sub btnCalcShunt1_Click(sender As System.Object, e As System.EventArgs) Handles btnCalcShunt1.Click

        txtR_shunt_1.Text = cMyRoutines.Calculate_Shunt(txtR_pao2_1.Text, txtR_paco2_1.Text, cmboR_fio2_1.Text)

    End Sub

    Private Sub btnCalcShunt2_Click(sender As System.Object, e As System.EventArgs) Handles btnCalcShunt2.Click
        txtR_shunt_2.Text = cMyRoutines.Calculate_Shunt(txtR_pao2_2.Text, txtR_paco2_2.Text, cmboR_fio2_2.Text)
    End Sub

    Private Sub cmboTechnicalNote_DropDownClosed(sender As Object, e As System.EventArgs) Handles cmboTechnicalNote.DropDownClosed

        Dim kv As KeyValuePair(Of String, Integer) = cmboTechnicalNote.SelectedItem
        If Len(txtTechnicalNote.Text) = 0 Then txtTechnicalNote.Text = kv.Key Else txtTechnicalNote.Text = txtTechnicalNote.Text & " " & kv.Key
        cmboTechnicalNote.SelectedIndex = -1
        txtTechnicalNote.Focus()
        txtTechnicalNote.SelectionStart = txtTechnicalNote.Text.Length
        txtTechnicalNote.SelectionLength = 0

    End Sub

    Private Sub Label40_DoubleClick(sender As Object, e As System.EventArgs) Handles Label40.DoubleClick
        If Not IsDate(txtTestTime.Text) Then
            txtTestTime.Text = Format(Now, "HH:mm")
        End If
    End Sub

    Private Sub cmboTechnicalNote_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmboTechnicalNote.SelectedIndexChanged

    End Sub
End Class