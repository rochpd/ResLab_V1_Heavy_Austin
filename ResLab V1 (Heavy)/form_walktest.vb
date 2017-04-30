
Imports System.Windows.Forms.DataVisualization.Charting
Imports ResLab_V1_Heavy.class_walktest_plot
Imports ResLab_V1_Heavy.cDatabaseInfo


Public Class form_walktest

    Dim flagFormLoad As Boolean = False
    Dim grd1 As DataGridView
    Dim grd2 As DataGridView
    Private _protocoldata As walk_protocoldata

    Private _callingFormName As String
    Private _newReportstatus As String

    Private _RecordID As Long = 0
    Private _protocolID As Long = 0
    Private _patientID As Long = 0
    Private _IsNewTest As Boolean = False

    Private _formPhrases As form_rft_phrases
    Private _form_session As form_rft_session


    Public Sub UpdateReportStatusTo(ByVal NewReportStatus As String)
        _newReportstatus = NewReportStatus
    End Sub

    Public Sub New(ByVal RecordID As Long, ByVal CallingForm As Form, Optional protocolID As Long = 0)
        'If recordID=0 then new test record and protocolID must be supplied
        'Otherwise, call back existing test and use its protocolID

        InitializeComponent()
        _callingFormName = CallingForm.Name
        _RecordID = RecordID
        _protocolID = protocolID

        Select Case _RecordID
            Case 0 : _IsNewTest = True
            Case Else : _IsNewTest = False
        End Select

        'use a local patientid because can come from reporting list or logged pt
        Select Case _RecordID
            Case 0 : _patientID = cPt.PatientID
            Case Is > 0 : _patientID = cPt.Get_PatientIDFromPkID(_RecordID, eTables.r_walktests_v1heavy)
        End Select

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub form_walktest_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated

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

    Private Sub form_walktest_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not IsNothing(_formPhrases) Then
            If _formPhrases.Visible Then _formPhrases.Close()
        End If

    End Sub

    Private Sub form_walktest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        flagFormLoad = True
        Me.Text = cPt.Get_PtNameString(_patientID, eNameStringFormats.Name_UR)

        'Do session
        _form_session = New form_rft_session(_patientID, _RecordID, eTables.r_walktests_v1heavy, Me, _IsNewTest)
        If _IsNewTest Then
            Dim response As Integer = _form_session.ShowDialog()
            If response = DialogResult.Cancel Then
                Me.Close()
                Exit Sub
            End If
        End If
        Me.RefreshFromSession()

        'Load protocol 
        If _IsNewTest Then
            'protocolID passed in
        Else
            _protocolID = cRfts.Get_ProtocolID_FromWalkID(_RecordID)  'get protocolID from test record
        End If
        _protocoldata = cWalk.get_ProtocolData(_protocolID)
        tabResults.TabPages(0).Text = _protocoldata.pMenuLabel

        'Setup empty raw data grids
        grd1 = cWalk.create_walkgrid(_protocoldata, _IsNewTest)
        grd1.Name = "grd1"
        split_trial_1.Panel1.Controls.Add(grd1)
        grd2 = cWalk.create_walkgrid(_protocoldata, _IsNewTest)
        grd2.Name = "grd2"
        split_trial_2.Panel1.Controls.Add(grd2)

        split_trial_1.SplitterDistance = split_trial_1.Height - txtDistance_trial_1.Height * 2
        split_trial_2.SplitterDistance = split_trial_2.Height - txtDistance_trial_2.Height * 2

        'Setup empty results grid
        Me.Setup_resultsgrid()

        'Load combo options
        Me.Load_combos(_IsNewTest)
        If _callingFormName = "form_Reporting" Then
            If cmboReportedBy.Text = "" Then cmboReportedBy.Text = form_Reporting.cmboDefaultReporter.Text
            If txtReportedDate.Text = "" Then txtReportedDate.Text = form_Reporting.txtDefaultDate.Text
        End If

        'Load test 
        Me.Load_ScreenFieldsFromMem(Me._IsNewTest, Me._RecordID)
        Me.Do_calculations()

        If _protocoldata.walkmode.ToLower = "treadmill" Then btnCalc_Trial1_Distance.Enabled = True Else btnCalc_Trial1_Distance.Enabled = False

        cUser.set_access(Me)

        flagFormLoad = False

    End Sub

    Private Sub Load_combos(_isnewtest As Boolean)

        'Load combo box options
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLab, "Labs", _isnewtest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboReportedBy, "ReportingMO names", _isnewtest)
        cMyRoutines.Combo_LoadItemsFromList(cmboReportStatus, eTables.List_ReportStatuses)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboTechnicalNote, "Technical comments", _isnewtest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboVerifiedBy, "ReportingMO names", _isnewtest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboScientist, "Scientist", _isnewtest)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLastBD, "Last_BD", _isnewtest)

    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim sessionID As Long = _form_session._sessionID
        cRfts.Save_rft_walk(Me._IsNewTest, sessionID, Me.Load_MemFromScreenFields_session(True), Me.Load_MemFromScreenFields_walk(True), Me.Load_MemFromScreenFields_trials(True), Me.Load_MemFromScreenFields_levels(True))

        'Refresh the list of tests and pdf to reflect any changes
        If Me._callingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

    End Sub

    Private Sub LockForm(ByVal LockState As Boolean)
        'True=locked, false=unlocked

        txtReport.ReadOnly = LockState
        pnlReporters.Enabled = Not LockState
        pnlOther.Enabled = Not LockState

    End Sub

    Private Function Load_MemFromScreenFields_session(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Session
        Dim R As New class_Rft_RoutineAndSessionFields

        d(R.PatientID) = Me._patientID
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

    Private Function Load_MemFromScreenFields_trials(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)()

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        'Do trial record(s)
        '    tab control pages containing trials must be named 'tabpage_trial_x' where x=1,2,3 ...etc

        Dim d(0) As Dictionary(Of String, String)
        Dim f As New class_fields_Walk_Trial
        Dim grd As DataGridView = Nothing
        Dim sp As SplitterPanel = Nothing
        Dim n As Integer = 0

        For i = 0 To tabTrials.TabPages.Count - 1
            If Strings.InStr(tabTrials.TabPages(i).Name, "trial") > 0 Then
                ReDim Preserve d(n)
                d(n) = cMyRoutines.MakeEmpty_dicWalk_trial

                'Get a reference to the grid on this tabpage (assume only ever one grid on page)
                For Each c As Control In tabTrials.TabPages(i).Controls
                    If c.GetType = GetType(SplitContainer) Then
                        For Each c1 As Control In c.Controls
                            If c1.GetType = GetType(SplitterPanel) Then
                                For Each c2 As Control In c1.Controls
                                    If c2.GetType = GetType(DataGridView) Then
                                        grd = c2
                                        sp = c1
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next

                If Not IsNothing(grd) Then
                    d(n)(f.trialID) = grd.Tag
                    d(n)(f.walkID) = Me._RecordID
                    d(n)(f.trial_number) = q & Strings.Right(tabTrials.TabPages(i).Name, 1) & q
                    d(n)(f.trial_label) = q & tabTrials.TabPages(i).Text & q
                    Select Case Strings.Right(tabTrials.TabPages(i).Name, 1)
                        Case 1
                            d(n)(f.trial_distance) = q & txtDistance_trial_1.Text & q
                            If IsDate(txtTime_trial_1.Text) Then d(n)(f.trial_timeofday) = q & txtTime_trial_1.Text & q Else d(n)(f.trial_timeofday) = "NULL"
                        Case 2
                            d(n)(f.trial_distance) = q & txtDistance_trial_2.Text & q
                            If IsDate(txtTime_trial_2.Text) Then d(n)(f.trial_timeofday) = q & txtTime_trial_2.Text & q Else d(n)(f.trial_timeofday) = "NULL"
                    End Select
                    d(n)(f.LastUpdated_trial) = cMyRoutines.FormatDBDateTime(Date.Now)
                    d(n)(f.LastUpdatedBy_trial) = q & My.User.Name & q
                    n = n + 1
                End If
            End If
        Next i

        Return d

    End Function

    Private Function Load_MemFromScreenFields_levels(AddQuotesAroundStrings As Boolean) As List(Of Dictionary(Of String, String)())

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        'Do level record(s)
        '    tab control pages containing trials must be named 'tabpage_trial_x' where x=1,2,3 ...etc

        Dim d() As Dictionary(Of String, String) = Nothing
        Dim t As New List(Of Dictionary(Of String, String)())
        Dim f As New class_fields_Walk_TrialLevel
        Dim grd As DataGridView = Nothing
        Dim sp As SplitterPanel = Nothing

        Dim level As Integer = 0

        For i = 0 To tabTrials.TabPages.Count - 1
            If Strings.InStr(tabTrials.TabPages(i).Name, "trial") > 0 Then

                'Get a reference to the grid on this tabpage (assume only ever one grid on page)
                For Each c As Control In tabTrials.TabPages(i).Controls
                    If c.GetType = GetType(SplitContainer) Then
                        For Each c1 As Control In c.Controls
                            If c1.GetType = GetType(SplitterPanel) Then
                                For Each c2 As Control In c1.Controls
                                    If c2.GetType = GetType(DataGridView) Then
                                        grd = c2
                                        sp = c1
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next

                If Not IsNothing(grd) Then
                    For level = 0 To grd.Rows.Count - 1
                        ReDim Preserve d(level)
                        d(level) = cMyRoutines.MakeEmpty_dicWalk_level
                        d(level)(f.trialID) = grd.Tag
                        d(level)(f.levelID) = grd(0, level).Value
                        d(level)(f.time_minute) = q & grd(1, level).Value & q
                        d(level)(f.time_label) = q & grd(2, level).Value & q
                        d(level)(f.time_speed_kph) = q & grd(3, level).Value & q
                        d(level)(f.time_gradient) = q & grd(4, level).Value & q
                        d(level)(f.time_o2flow) = q & grd(5, level).Value & q
                        d(level)(f.time_spo2) = q & grd(6, level).Value & q
                        d(level)(f.time_hr) = q & grd(7, level).Value & q
                        d(level)(f.time_dyspnoea) = q & grd(8, level).Value & q
                        d(level)(f.LastUpdated_level) = cMyRoutines.FormatDBDateTime(Date.Now)
                        d(level)(f.LastUpdatedBy_level) = q & My.User.Name & q
                    Next level
                    t.Add(d)
                End If
            End If
        Next i

        Return t

    End Function

    Private Function Load_MemFromScreenFields_walk(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicWalk_test
        Dim R As New class_fields_WalkAndSession

        dicR(R.patientID) = Me._patientID
        dicR(R.sessionID) = Nothing     'set in save
        dicR(R.walkID) = Me._RecordID
        dicR(R.TestTime) = cMyRoutines.FormatDBTime(txtTestTime.Text)
        dicR(R.TestType) = "'Walk'"
        dicR(R.WalkType) = q & tabResults.TabPages(0).Text & q
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
        dicR(R.ProtocolID) = Me._protocolID
        dicR(R.LastUpdated_walk) = cMyRoutines.FormatDBDateTime(Now)
        dicR(R.LastUpdatedBy_walk) = q & My.User.Name & q

        If _callingFormName = "form_Reporting" And cmboReportStatus.SelectedIndex < cmboReportStatus.Items.Count - 1 Then
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

        Return dicR

    End Function

    Private Sub Setup_resultsgrid()

        Dim i As Integer = 0, j As Integer = 0
        Dim CellStyle_body As New DataGridViewCellStyle

        CellStyle_body.Alignment = DataGridViewContentAlignment.MiddleLeft
        CellStyle_body.BackColor = Drawing.Color.White
        CellStyle_body.Font = New Font("Arial", 8, FontStyle.Bold, GraphicsUnit.Point, CType(0, Byte))
        CellStyle_body.ForeColor = Drawing.Color.Black

        grdResults.AllowUserToAddRows = False
        grdResults.AllowUserToDeleteRows = False
        grdResults.AllowUserToResizeRows = False
        grdResults.BorderStyle = BorderStyle.None
        grdResults.DefaultCellStyle = CellStyle_body
        grdResults.Dock = DockStyle.None
        grdResults.MultiSelect = False
        grdResults.ReadOnly = True
        grdResults.RowHeadersVisible = False
        grdResults.ScrollBars = ScrollBars.None
        grdResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grdResults.ColumnHeadersVisible = False
        grdResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        grdResults.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        'Set columns
        Dim col_names As New List(Of String) From {"parameter", "units", "rest_1", "walk_1", "rest_2", "walk_2"}
        Dim col_headertext1 As New List(Of String) From {"", "", "Trial 1", "", "Trial 2", ""}
        Dim col_headertext2 As New List(Of String) From {"", "", "Rest", "Walk", "Rest", "Walk"}
        Dim col_width As New List(Of Integer) From {70, 90, 60, 60, 60, 60}
        Dim col_align As New List(Of Integer) From {DataGridViewContentAlignment.MiddleLeft, DataGridViewContentAlignment.MiddleLeft, DataGridViewContentAlignment.MiddleCenter, DataGridViewContentAlignment.MiddleCenter, DataGridViewContentAlignment.MiddleCenter, DataGridViewContentAlignment.MiddleCenter}

        Dim c As DataGridViewTextBoxColumn
        grdResults.Columns.Clear()
        For i = 0 To col_names.Count - 1
            c = New DataGridViewTextBoxColumn
            c.Name = col_names(i)
            c.Width = col_width(i)
            c.DefaultCellStyle = CellStyle_body
            c.DefaultCellStyle.Alignment = col_align(i)
            c.ReadOnly = True
            c.SortMode = DataGridViewColumnSortMode.NotSortable
            If i < 2 Then c.DefaultCellStyle.BackColor = SystemColors.ButtonFace
            grdResults.Columns.Add(c)
        Next

        grdResults.Rows.Add(2)
        For i = 0 To 1
            For j = 0 To col_headertext1.Count - 1
                If i = 0 Then grdResults(j, i).Value = col_headertext1(j)
                If i = 1 Then grdResults(j, i).Value = col_headertext2(j)
            Next
        Next
        grdResults.Rows(0).DefaultCellStyle.BackColor = SystemColors.ButtonFace
        grdResults.Rows(1).DefaultCellStyle.BackColor = SystemColors.ButtonFace

        'Set rows
        Dim row_parameters As New List(Of String) From {"Suppl. O2", "SpO2", "HR", "Dyspnoea", "Speed", "Grade", "Distance"}
        Dim row_units As New List(Of String) From {"L/min", "%", "bpm", "Borg", "km/hr", "%", "m"}
        Dim row_enabled As New List(Of Boolean) From {False, False, False, False, False, False, False}
        For i = 0 To row_parameters.Count - 1
            Select Case i
                Case 0 : row_enabled(i) = _protocoldata.var_fio2
                Case 1 : row_enabled(i) = _protocoldata.var_spo2
                Case 2 : row_enabled(i) = _protocoldata.var_hr
                Case 3 : row_enabled(i) = _protocoldata.var_dyspnoea
                Case 4 : row_enabled(i) = _protocoldata.var_speed
                Case 5 : row_enabled(i) = _protocoldata.var_grade
                Case 6 : row_enabled(i) = True
            End Select
        Next

        For i = 0 To row_parameters.Count - 1
            grdResults.Rows.Add(1)
            grdResults(0, i + 2).Value = row_parameters(i)
            grdResults(1, i + 2).Value = row_units(i)
            grdResults.Rows(grdResults.Rows.Count - 1).Visible = row_enabled(i)

            For j = 2 To grdResults.Columns.Count - 1
                grdResults(j, grdResults.Rows.Count - 1).Style.BackColor = Color.LightYellow
            Next
        Next



    End Sub

    Private Function Load_MemfromTestDataGrids() As walk_trialdata()
        'Map 'air' entries to zero flow values

        Dim i As Integer
        Dim s() As walk_trialdata
        ReDim s(1)
        ReDim s(0).timepoints(_protocoldata.nTimePoints_rest + _protocoldata.nTimePoints_walk + _protocoldata.nTimePoints_post - 1)
        ReDim s(1).timepoints(_protocoldata.nTimePoints_rest + _protocoldata.nTimePoints_walk + _protocoldata.nTimePoints_post - 1)

        With s(0)
            .trial_distance = txtDistance_trial_1.Text
            .trial_label = tabTrials.TabPages(1).Text
            .trial_number = 1
            For i = 0 To grd1.Rows.Count - 1
                .timepoints(i).time_minute = grd1(1, i).Value
                .timepoints(i).time_label = grd1(2, i).Value
                .timepoints(i).time_speed_kph = grd1(3, i).Value
                .timepoints(i).time_gradient = grd1(4, i).Value
                If LCase(grd1(5, i).Value) = "air" Then .timepoints(i).time_o2flow = "0" Else .timepoints(i).time_o2flow = grd1(5, i).Value
                .timepoints(i).time_spo2 = grd1(6, i).Value
                .timepoints(i).time_hr = grd1(7, i).Value
                .timepoints(i).time_dyspnoea = grd1(8, i).Value
            Next
        End With
        With s(1)
            .trial_distance = txtDistance_trial_2.Text
            .trial_label = tabTrials.TabPages(2).Text
            .trial_number = 2
            For i = 0 To grd2.Rows.Count - 1
                .timepoints(i).time_minute = grd2(1, i).Value
                .timepoints(i).time_label = grd2(2, i).Value
                .timepoints(i).time_speed_kph = grd2(3, i).Value
                .timepoints(i).time_gradient = grd2(4, i).Value
                If LCase(grd2(5, i).Value) = "air" Then .timepoints(i).time_o2flow = "0" Else .timepoints(i).time_o2flow = grd2(5, i).Value
                .timepoints(i).time_spo2 = grd2(6, i).Value
                .timepoints(i).time_hr = grd2(7, i).Value
                .timepoints(i).time_dyspnoea = grd2(8, i).Value
            Next
        End With

        Return s

    End Function



    Private Sub btnRefreshSummary_Click(sender As System.Object, e As System.EventArgs)

        Me.Load_SummaryResultsToGrid()

    End Sub

    Private Sub plot_walkdata()

        For Each c In pnlChart.Controls
            pnlChart.Controls.Remove(c)
            If TypeOf (c) Is Chart Then c = Nothing
        Next

        Dim p = cWalkPlot.Get_plotproperties_walk
        ReDim p.xData(1)
        ReDim p.yData(1)
        Me.get_xyvalues_fromgridcolumn(p, grd1, "Time", "SpO2")
        Me.get_xyvalues_fromgridcolumn(p, grd2, "Time", "SpO2")

        Dim ch As Chart = cWalkPlot.Create_plot(p)
        ch.Name = "WalkGraph"
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnlChart.Controls.Add(ch)

    End Sub

    Private Sub get_xyvalues_fromgridcolumn(ByRef p As walk_plot_properties, ByRef g As DataGridView, col_nameX As String, col_nameY As String)
        'Get values in Y column - skip non-numeric Y values and corresponding X values
        'Assume complete set of X vaues present

        Dim x As New List(Of Single)
        Dim y As New List(Of Single)
        For Each row As DataGridViewRow In g.Rows
            If IsNumeric(row.Cells(col_nameY).Value) Then
                x.Add(CSng(row.Cells(col_nameX).Value))
                y.Add(CSng(row.Cells(col_nameY).Value))
            End If
        Next
        Select Case g.Name.ToLower
            Case "grd1" : p.xData(0) = x : p.yData(0) = y
            Case "grd2" : p.xData(1) = x : p.yData(1) = y
        End Select

    End Sub

    Private Sub Load_SummaryResultsToGrid()

        Dim i As Integer = 0, j As Integer = 0, col As Integer = 0
        Dim w() As walk_trialdata = Me.Load_MemfromTestDataGrids()
        Dim s() As walk_SummaryResults = cWalk.Calculate_SummaryResults(w, _protocolID)

        For i = 0 To UBound(s)
            Select Case s(i).trialnumber
                Case 1 : j = 2
                Case 2 : j = 4
            End Select
            grdResults(j, 2).Value = s(i).FiO2_rest
            grdResults(j, 3).Value = s(i).SpO2_rest
            grdResults(j, 4).Value = s(i).HR_rest
            grdResults(j, 5).Value = s(i).Dyspnoea_rest
            grdResults(j, 6).Value = ""
            grdResults(j, 7).Value = ""
            grdResults(j, 8).Value = ""

            grdResults(j + 1, 2).Value = s(i).FiO2_exercise
            grdResults(j + 1, 3).Value = s(i).SpO2_exercise
            grdResults(j + 1, 4).Value = s(i).HR_exercise
            grdResults(j + 1, 5).Value = s(i).Dyspnoea_exercise
            grdResults(j + 1, 6).Value = s(i).Speed_exercise
            grdResults(j + 1, 7).Value = s(i).Gradient_exercise
            grdResults(j + 1, 8).Value = s(i).Distance_exercise
        Next

    End Sub

    Private Function Load_ScreenFieldsFromMem(isNewWalk As Boolean, walkID As Long) As Boolean
        'trialID stored in grd.tag property

        Dim i As Integer = 0, j As Integer = 0

        Try
            Select Case isNewWalk
                Case True
                    cmboReportStatus.Text = "Unreported"
                    grd1.Tag = 0    'TrialID stored here
                    grd2.Tag = 0
                    Return True

                Case False
                    Dim flds As New class_fields_WalkAndSession
                    Dim w As Dictionary(Of String, String) = cRfts.Get_walk_test_session(walkID)
                    If IsNothing(w) Then
                        'should never happen
                        Return False
                    Else

                        'Load the test data
                        txtReport.Text = w(flds.Report_text)
                        If IsDate(w(flds.Report_reporteddate)) Then txtReportedDate.Text = w(flds.Report_reporteddate)
                        cmboReportedBy.Text = w(flds.Report_reportedby)
                        cmboVerifiedBy.Text = w(flds.Report_verifiedby)
                        If IsDate(w(flds.Report_verifieddate)) Then txtVerifiedDate.Text = w(flds.Report_verifieddate)
                        cmboReportStatus.Text = w(flds.Report_status)
                        cmboScientist.Text = w(flds.Scientist)
                        If IsDate(w(flds.TestTime)) Then txtTestTime.Text = cMyRoutines.FormatDBTime(w(flds.TestTime))
                        cmboLastBD.Text = w(flds.BDStatus)
                        cmboLab.Text = w(flds.Lab)
                        txtTechnicalNote.Text = w(flds.TechnicalNotes)

                        'Load trials and levels
                        Dim ft As New class_fields_Walk_Trial
                        Dim fl As New class_fields_Walk_TrialLevel
                        Dim t() As Dictionary(Of String, String) = cRfts.Get_walk_trials(walkID)
                        Dim lvl As List(Of Dictionary(Of String, String)()) = cRfts.Get_walk_levels(walkID)

                        If Not IsNothing(t) Then
                            Dim g As New DataGridView
                            For i = 0 To t.Count - 1
                                Select Case t(i)(ft.trial_number)
                                    Case 1
                                        g = grd1
                                        txtDistance_trial_1.Text = t(i)(ft.trial_distance)
                                        If IsDate(t(i)(ft.trial_timeofday)) Then txtTime_trial_1.Text = t(i)(ft.trial_timeofday) Else txtTime_trial_1.Text = "__:__"
                                    Case 2
                                        g = grd2
                                        txtDistance_trial_2.Text = t(i)(ft.trial_distance)
                                        If IsDate(t(i)(ft.trial_timeofday)) Then txtTime_trial_2.Text = t(i)(ft.trial_timeofday) Else txtTime_trial_2.Text = "__:__"
                                End Select
                                g.Tag = t(i)(ft.trialID)

                                'Load timepoint data
                                If Not IsNothing(lvl(i)) Then
                                    For j = 0 To lvl(i).Count - 1
                                        If Not IsNothing(lvl(i)(j)(fl.levelID)) Then g(0, j).Value = lvl(i)(j)(fl.levelID) Else g(0, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_minute)) Then g(1, j).Value = lvl(i)(j)(fl.time_minute) Else g(1, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_label)) Then g(2, j).Value = lvl(i)(j)(fl.time_label) Else g(2, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_speed_kph)) Then g(3, j).Value = lvl(i)(j)(fl.time_speed_kph) Else g(3, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_gradient)) Then g(4, j).Value = lvl(i)(j)(fl.time_gradient) Else g(4, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_o2flow)) Then g(5, j).Value = lvl(i)(j)(fl.time_o2flow) Else g(5, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_spo2)) Then g(6, j).Value = lvl(i)(j)(fl.time_spo2) Else g(6, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_hr)) Then g(7, j).Value = lvl(i)(j)(fl.time_hr) Else g(7, j).Value = ""
                                        If Not IsNothing(lvl(i)(j)(fl.time_dyspnoea)) Then g(8, j).Value = lvl(i)(j)(fl.time_dyspnoea) Else g(8, j).Value = ""
                                    Next j
                                Else
                                    MsgBox("No level data to load", vbOKOnly, "ResLab")
                                End If
                            Next i
                            g = Nothing
                        Else
                            MsgBox("No trial data to load", vbOKOnly, "ResLab")
                        End If

                        Me.Load_SummaryResultsToGrid()
                        Me.plot_walkdata()

                        Return True
                    End If
                Case Else
                    Return False
            End Select

        Catch ex As Exception
            MsgBox("Error in form_walktest.Load_ScreenFieldsFromMem" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Private Sub lbltestdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        txtTestDate.Text = Format(Now, "dd/MM/yyyy")
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        If cRfts.Delete_rft(Me._RecordID, eTables.r_walktests_v1heavy) Then
            'Refresh the list of tests and pdf to reflect any changes
            If Me._callingFormName = "Form_MainNew" Then gRefreshMainForm = True
            Me.Close()
        Else
            'Keep form open if delete is cancelled
        End If

    End Sub

    Private Sub lblTestTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not IsDate(txtTestTime.Text) Then txtTestTime.Text = Format(Now, "HH:mm")
    End Sub

    Private Sub lblReportedBy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not IsDate(txtReportedDate.Text) Then
            txtReportedDate.Text = Format(Now, "dd/MM/yyyy")
        End If
    End Sub

    Private Sub lblVerifiedBy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not IsDate(txtVerifiedDate.Text) Then
            txtVerifiedDate.Text = Format(Now, "dd/MM/yyyy")
        End If
    End Sub

    Private Sub txtYearsSmoked_LostFocus(sender As Object, e As System.EventArgs)
        Me.Do_calculations()
    End Sub

    Private Sub txtCigsPerDay_LostFocus(sender As Object, e As System.EventArgs)
        Me.Do_calculations()
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

    Private Sub Do_calculations()

        'BMI
        txtBMI.Text = cMyRoutines.calc_BMI(txtHeight.Text, txtWeight.Text)

        'Age
        txtAge.Text = "(" & cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text)) & ")"

        'Pack years
        txtPackYears.Text = cMyRoutines.calc_PackYears(txtCigsPerDay.Text, txtYearsSmoked.Text)

    End Sub

    Private Sub tabTrials_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles tabTrials.SelectedIndexChanged

        If tabTrials.SelectedIndex = 0 Then
            Me.Load_SummaryResultsToGrid()
            Me.plot_walkdata()
        End If

    End Sub

    Private Sub btnCalc_Trial1_Distance_Click(sender As Object, e As System.EventArgs) Handles btnCalc_Trial1_Distance.Click

        txtDistance_trial_1.Text = cWalk.Calc_distance(grd1, _protocolID)

    End Sub

    Private Sub btnCalc_Trial2_Distance_Click(sender As Object, e As System.EventArgs) Handles btnCalc_Trial2_Distance.Click

        txtDistance_trial_2.Text = cWalk.Calc_distance(grd2, _protocolID)

    End Sub

    Private Sub grdResults_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdResults.CellContentClick

    End Sub

    Private Sub grdResults_SelectionChanged(sender As Object, e As System.EventArgs) Handles grdResults.SelectionChanged
        grdResults.ClearSelection()
    End Sub

    Private Sub btnFillDown_grd1_Click(sender As System.Object, e As System.EventArgs) Handles btnFillDown_grd1.Click

        'Get a reference to the runtime created grid
        For Each c As Control In tabTrials.TabPages("tabpage_trial_1").Controls("split_trial_1").Controls
            If TypeOf (c) Is SplitterPanel Then
                For Each cc As Control In c.Controls
                    If TypeOf (cc) Is DataGridView And CType(cc, DataGridView).SelectedCells(0).ColumnIndex > 2 Then
                        For i = CType(cc, DataGridView).SelectedCells(0).RowIndex + 1 To CType(cc, DataGridView).Rows.Count - 1
                            CType(cc, DataGridView)(CType(cc, DataGridView).SelectedCells(0).ColumnIndex, i).Value = CType(cc, DataGridView).SelectedCells(0).Value
                        Next
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Next
            End If
        Next

    End Sub

    Private Sub btnFillDown_grd2_Click(sender As Object, e As System.EventArgs) Handles btnFillDown_grd2.Click

        'Get a reference to the runtime created grid
        For Each c As Control In tabTrials.TabPages("tabpage_trial_2").Controls("split_trial_2").Controls
            If TypeOf (c) Is SplitterPanel Then
                For Each cc As Control In c.Controls
                    If TypeOf (cc) Is DataGridView And CType(cc, DataGridView).SelectedCells(0).ColumnIndex > 2 Then
                        For i = CType(cc, DataGridView).SelectedCells(0).RowIndex + 1 To CType(cc, DataGridView).Rows.Count - 1
                            CType(cc, DataGridView)(CType(cc, DataGridView).SelectedCells(0).ColumnIndex, i).Value = CType(cc, DataGridView).SelectedCells(0).Value
                        Next
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Next
            End If
        Next

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

        Dim dicR As New Dictionary(Of String, String)

        _formPhrases = New form_rft_phrases(eAutoreport_testgroups.Walk_tests, Me)
        _formPhrases.ReportTextBox = Me.txtReport
        _formPhrases.Show()

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


    Private Sub t_Click(sender As System.Object, e As System.EventArgs) Handles t.Click

    End Sub
End Class