Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_rft_cpet

    Private _callingFormName As String

    Private _cpetID As Long
    Private _patientID As Long
    Private _isNewCpet As Boolean

    Private _formPhrases As form_rft_phrases
    Private _form_session As form_rft_session


    Public Structure cpet_visitdata
        Dim testID As Long
        Dim patientID As Long
        Dim testdate As Date
        Dim testtime As Date
        Dim testtype As String
        Dim data() As cpet_loaddata
    End Structure
    Public Structure cpet_loaddata
        Dim loadID As Integer
        Dim testID As Long
        Dim loadNumber As Integer
        Dim load As String
        Dim loadtime As String
        Dim ve As String
        Dim vo2 As String
        Dim vco2 As String
        Dim hr As String
        Dim spo2 As String
        Dim peto2 As String
        Dim petco2 As String
        Dim rer As String
        Dim vt As String
        Dim bp As String
        Dim ph As String
        Dim paco2 As String
        Dim pao2 As String
        Dim sao2 As String
        Dim be As String
        Dim hco3 As String
        Dim aapo2 As String
        Dim vdvt As String
    End Structure

    Private Enum eGetValuesModes
        All_values
        End_of_minute_values
    End Enum

    Private Enum eCpetSystems
        Medisoft_ergocard
        Sensormedics_VMax2900
    End Enum

    Public Sub New(ByVal cpetID As Long, ByVal CallingForm As Form)

        InitializeComponent()
        _callingFormName = CallingForm.Name
        _cpetID = cpetID

        Select Case _cpetID
            Case 0 : _isNewCpet = True
            Case Else : _isNewCpet = False
        End Select

        'use a local patientid because can come from reporting list or logged pt
        Select Case _cpetID
            Case 0 : _patientID = cPt.PatientID
            Case Is > 0 : _patientID = cPt.Get_PatientIDFromPkID(_cpetID, eTables.r_cpet)
        End Select

    End Sub

    Private Sub form_rft_cpet_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated

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

    Private Sub form_rft_cpet_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not IsNothing(_formPhrases) Then
            If _formPhrases.Visible Then _formPhrases.Close()
        End If

    End Sub

    Private Sub form_rft_cpet_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Do session
        _form_session = New form_rft_session(_patientID, _cpetID, eTables.r_cpet, Me, _isNewCpet)
        If _isNewCpet Then
            Dim response As Integer = _form_session.ShowDialog()
            If response = DialogResult.Cancel Then
                Me.Close()
                Exit Sub
            End If
        End If
        Me.RefreshFromSession()
        Me.Text = cPt.Get_PtNameString(_patientID, eNameStringFormats.Name_UR)

        'Load combo box options
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLab, "labs", _isNewCpet)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboLastBD, "Last_BD", _isNewCpet)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboScientist, "scientist", _isNewCpet)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboReportedBy, "ReportingMO names", _isNewCpet)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboVerifiedBy, "ReportingMO names", _isNewCpet)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboTechnicalNote, "Technical comments", _isNewCpet)
        cMyRoutines.Combo_LoadItemsFromList(cmboReportStatus, eTables.List_ReportStatuses)

        'Load test 
        Me.setup_datagrid()
        Select Case _isNewCpet
            Case True
                cmboReportStatus.Text = "Unreported"
            Case False
                Me.Load_ScreenFieldsFromMem(cRfts.Get_rft_cpet_testdata(_cpetID), cRfts.Get_rft_cpet_workloaddata(_cpetID))
                Me.load_textboxes_maxvaluesfromgrid()
                Me.do_calculations()
                Me.plot_graphs()
        End Select

        cUser.set_access(Me)

    End Sub

    Private Sub do_calculations()

        Dim dPred As Dictionary(Of String, String) = Me.get_CpetPreds(class_Pred.eLoadNormalsMode.UseCurrentPrefs)
        Me.load_textboxeswithpreds(dPred)

        txtBMI.Text = cMyRoutines.calc_BMI(txtHeight.Text, txtWeight.Text)
        txtPackYears.Text = cMyRoutines.calc_PackYears(txtCigsPerDay.Text, txtYearsSmoked.Text)
        txtAge.Text = cMyRoutines.Calc_Age(txtDOB.Text, txtTestDate.Text)

        If Val(txtmax_hr.Text) > 0 And Val(txtmax_pred_hr.Tag) > 0 Then txtmax_ppred_hr.Text = Format(100 * Val(txtmax_hr.Text) / Val(txtmax_pred_hr.Tag), "(###)") Else txtmax_ppred_hr.Text = "---"
        If Val(txtmax_ve.Text) > 0 And Val(txtmax_pred_ve.Tag) > 0 Then txtmax_ppred_ve.Text = Format(100 * Val(txtmax_ve.Text) / Val(txtmax_pred_ve.Tag), "(###)") Else txtmax_ppred_ve.Text = "---"
        If Val(txtmax_vo2.Text) > 0 And Val(txtmax_pred_vo2.Tag) > 0 Then txtmax_ppred_vo2.Text = Format(100 * Val(txtmax_vo2.Text) / Val(txtmax_pred_vo2.Tag), "(###)") Else txtmax_ppred_vo2.Text = "---"
        If Val(txtmax_vco2.Text) > 0 And Val(txtmax_pred_vco2.Tag) > 0 Then txtmax_ppred_vco2.Text = Format(100 * Val(txtmax_vco2.Text) / Val(txtmax_pred_vco2.Tag), "(###)") Else txtmax_ppred_vco2.Text = "---"
        If Val(txtmax_work.Text) > 0 And Val(txtmax_pred_work.Tag) > 0 Then txtmax_ppred_work.Text = Format(100 * Val(txtmax_work.Text) / Val(txtmax_pred_work.Tag), "(###)") Else txtmax_ppred_work.Text = "---"
        If Val(txtmax_vt.Text) > 0 And Val(txtmax_pred_vt.Tag) > 0 Then txtmax_ppred_vt.Text = Format(100 * Val(txtmax_vt.Text) / Val(txtmax_pred_vt.Tag), "(###)") Else txtmax_ppred_vt.Text = "---"
        If Val(txtmax_o2pulse.Text) > 0 And Val(txtmax_pred_o2pulse.Tag) > 0 Then txtmax_ppred_o2pulse.Text = Format(100 * Val(txtmax_o2pulse.Text) / Val(txtmax_pred_o2pulse.Tag), "(###)") Else txtmax_ppred_o2pulse.Text = "---"
        If Val(txtmax_vo2kg.Text) > 0 And Val(txtmax_pred_vo2kg.Tag) > 0 Then txtmax_ppred_vo2kg.Text = Format(100 * Val(txtmax_vo2kg.Text) / Val(txtmax_pred_vo2kg.Tag), "(###)") Else txtmax_ppred_vo2kg.Text = "---"

    End Sub

    Private Sub cmboTechnicalNote_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmboTechnicalNote.DropDownClosed

        Dim kv As KeyValuePair(Of String, Integer) = cmboTechnicalNote.SelectedItem
        If Len(txtTechnicalNote.Text) = 0 Then txtTechnicalNote.Text = kv.Key Else txtTechnicalNote.Text = txtTechnicalNote.Text & " " & kv.Key
        cmboTechnicalNote.SelectedIndex = -1
        txtTechnicalNote.Focus()
        txtTechnicalNote.SelectionStart = txtTechnicalNote.Text.Length
        txtTechnicalNote.SelectionLength = 0

    End Sub

    Private Function get_demo_fromtextboxes() As Pred_demo

        Dim demo As Pred_demo = Nothing
        If IsDate(txtDOB.Text) And IsDate(txtTestDate.Text) Then demo.Age = cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text))
        demo.Htcm = Val(txtHeight.Text)
        demo.Wtkg = Val(txtWeight.Text)
        demo.GenderID = cMyRoutines.Lookup_list_ByDescription(txtGender.Text, eTables.Pred_ref_genders)
        demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(txtEthnicity.Text, eTables.Pred_ref_ethnicities)
        demo.Ethnicity = txtEthnicity.Text
        If IsDate(txtTestDate.Text) Then demo.TestDate = txtTestDate.Text Else demo.TestDate = Nothing
        demo.SourcesString = txtPredsRaw.Text
        demo.FEV1 = Val(txtSpiro_fev1.Text)
        demo.FVC = Val(txtSpiro_fvc.Text)
        Return demo

    End Function

    Private Function get_CpetPreds(ByVal Method As class_Pred.eLoadNormalsMode) As Dictionary(Of String, String)

        If Not IsDate(txtTestDate.Text) Then
            Return Nothing
        Else
            Dim dicPreds As Dictionary(Of String, String) = cPred.Get_Pred_cpet_values(Me.get_demo_fromtextboxes, Method)   'ParameterID|StatTypeID, result
            Return dicPreds
        End If

    End Function

    Private Sub load_textboxeswithpreds(d As Dictionary(Of String, String))
        'ParameterID|StatTypeID, result

        If Not (d Is Nothing) Then
            cPred.LoadLbl(txtmax_pred_hr, "hrmax|lln", d, 0)
            txtmax_pred_hr.Tag = d("hrmax|mpv")

            cPred.LoadLbl(txtmax_pred_ve, "vemax|lln", d, 1)
            txtmax_pred_ve.Tag = d("vemax|mpv")

            cPred.LoadLbl(txtmax_pred_vo2, "vo2max|lln", d, 2)
            txtmax_pred_vo2.Tag = d("vo2max|mpv")

            cPred.LoadLbl(txtmax_pred_vo2kg, "vo2kg_max|lln", d, 2)
            txtmax_pred_vo2kg.Tag = d("vo2kg_max|mpv")

            cPred.LoadLbl(txtmax_pred_o2pulse, "o2pulsemax|lln", d, 2)
            txtmax_pred_o2pulse.Tag = d("o2pulsemax|mpv")

            cPred.LoadLbl(txtmax_pred_vco2, "vco2max|lln", d, 2)
            txtmax_pred_vco2.Tag = d("vco2max|mpv")

            cPred.LoadLbl(txtmax_pred_work, "wmax|lln", d, 2)
            txtmax_pred_work.Tag = d("wmax|mpv")

            cPred.LoadLbl(txtmax_pred_vt, "vtmax|lln", d, 2)
            txtmax_pred_vt.Tag = d("vtmax|mpv")

            'txtPredsRaw.Text = cPred.Get_PrefSources_AsCodedString(Method, demo)
            'txtPreds.Text = cPred.Get_PrefSources_AsFormattedString(CDate(txtTestDate.Text), demo)

        End If

    End Sub

    Private Function get_vo2_rer_min() As Single

        Dim vo2 = Me.get_values_fromgridcolumn(grdData, "vo2", eGetValuesModes.All_values)
        Dim min = 0
        Dim index As Integer = 0
        For i = 0 To vo2.Count - 1
            If vo2(i) < min Then
                min = vo2(i)
                index = i
            End If
        Next
        Return index

    End Function

    Private Function get_vco2_rer_min() As Single

        Dim vco2 = Me.get_values_fromgridcolumn(grdData, "vco2", eGetValuesModes.All_values)
        Dim min = 0
        Dim index As Integer = 0
        For i = 0 To vco2.Count - 1
            If vco2(i) < min Then
                min = vco2(i)
                index = i
            End If
        Next
        Return index

    End Function

    Private Sub plot_graphs()

        'CHeck if any data in the grid
        If grdData.Rows.Count = 0 Then
            Exit Sub
        End If


        Dim c As New class_plot_cpet
        Dim ch As Chart
        Dim p As class_plot_cpet.cpet_plot_properties
        Dim pred_values As New Dictionary(Of String, String)
        Dim i As Integer = 0

        Select Case _isNewCpet
            Case True : pred_values = Me.get_CpetPreds(class_Pred.eLoadNormalsMode.UseCurrentPrefs)
            Case False : pred_values = Me.get_CpetPreds(class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)
        End Select

        Dim pred_demo As Pred_demo = Me.get_demo_fromtextboxes

        'Clear any existing graphs
        Me.pnlVeVO2.Controls.Clear()
        Me.pnlHR.Controls.Clear()

        'Ve/VO2
        p = cCpet.Get_plotproperties_VeVO2(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "vo2", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "ve", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnlVeVO2.Controls.Add(ch)

        'HR_SpO2/VO2
        p = cCpet.Get_plotproperties_HrSPo2_VO2(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "vo2", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "hr", eGetValuesModes.All_values)
        p.y2Data = Me.get_values_fromgridcolumn(grdData, "spo2", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnlHR.Controls.Add(ch)

        'Ve_time
        p = cCpet.Get_plotproperties_ve_time(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "time_min", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "ve", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_ve_time.Controls.Add(ch)

        'Ve_VCO2
        p = cCpet.Get_plotproperties_VeVCO2(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "vco2", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "ve", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_ve_vco2.Controls.Add(ch)

        'Vt_Ve
        p = cCpet.Get_plotproperties_VtVe(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "ve", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "vt", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_vt_ve.Controls.Add(ch)

        'HR_O2Pulse vs time
        p = cCpet.Get_plotproperties_Hr_O2Pulse_time(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "time_min", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "hr", eGetValuesModes.All_values)
        p.y2Data = Me.get_values_fromgridcolumn(grdData, "o2pulse", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_hr_o2pulse_time.Controls.Add(ch)

        'HR_VCO2 vs VO2
        p = cCpet.Get_plotproperties_Hr_VCO2_VO2(pred_demo, pred_values, Me.get_vo2_rer_min, Me.get_vco2_rer_min)
        p.xData = Me.get_values_fromgridcolumn(grdData, "vo2", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "hr", eGetValuesModes.All_values)
        p.y2Data = Me.get_values_fromgridcolumn(grdData, "vco2", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_hr_vco2_vo2.Controls.Add(ch)

        'RER vs time
        p = cCpet.Get_plotproperties_RER_time(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "time_min", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "rer", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_rq_time.Controls.Add(ch)

        'VO2_VCO2_Work vs time
        p = cCpet.Get_plotproperties_vo2_vco2_work_time(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "time_min", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "vo2", eGetValuesModes.All_values)
        p.y1_2_Data = Me.get_values_fromgridcolumn(grdData, "vco2", eGetValuesModes.All_values)
        p.y2Data = Me.get_values_fromgridcolumn(grdData, "load", eGetValuesModes.All_values)
        'Set the load data as a reference line, Filter out zero loads and duplicates
        Dim wrk = p.y2Data
        Dim v As Single = 0
        For i = 0 To wrk.Count - 2
            v = wrk(i)
            If v = 0 Then
                wrk(i) = Single.NaN
            Else
                Do While v = wrk(i + 1)
                    wrk(i + 1) = Single.NaN
                    i = i + 1
                    If i = wrk.Count - 1 Then Exit Do
                Loop
            End If
        Next
        ' Get first and last loads
        Dim first As Integer = 0, last As Integer = 0
        For i = 0 To wrk.Count - 1
            If Not Single.IsNaN(wrk(i)) Then
                first = i
                Exit For
            End If
        Next
        For i = wrk.Count - 1 To 0 Step -1
            If Not Single.IsNaN(wrk(i)) Then
                last = i
                Exit For
            End If
        Next
        ReDim Preserve p.y2Axis_plot_ref_lines(0)
        p.y2Axis_plot_ref_lines(0).xData = {p.xData(first), p.xData(last)}
        p.y2Axis_plot_ref_lines(0).yData = {wrk(first), wrk(last)}
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_vo2_vco2_work_time.Controls.Add(ch)

        'VEVO2_VECO2 vs time
        p = cCpet.Get_plotproperties_vevo2_vevco2_time(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "time_min", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "vevo2", eGetValuesModes.All_values)
        p.y2Data = Me.get_values_fromgridcolumn(grdData, "vevco2", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_vevo2_vevco2_time.Controls.Add(ch)


        'peto2_petco2_SpO2 vs time
        p = cCpet.Get_plotproperties_peto2_petco2_spo2_time(pred_demo, pred_values)
        p.xData = Me.get_values_fromgridcolumn(grdData, "time_min", eGetValuesModes.All_values)
        p.y1_1_Data = Me.get_values_fromgridcolumn(grdData, "peto2", eGetValuesModes.All_values)
        p.y1_2_Data = Me.get_values_fromgridcolumn(grdData, "petco2", eGetValuesModes.All_values)
        p.y2Data = Me.get_values_fromgridcolumn(grdData, "spo2", eGetValuesModes.All_values)
        ch = cCpet.Create_plot(pred_demo, p)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        pnl_eto2_etco2_sao2_time.Controls.Add(ch)


    End Sub

    Private Sub setup_datagrid()

        Dim i As Integer = 0
        Dim c As DataGridViewColumn
        Dim Align_mc As DataGridViewContentAlignment = DataGridViewContentAlignment.MiddleCenter

        Dim header_text() As String = {"levelID", "#", "Load (watts)", "Time (min)", "VE (L/min)", "VO2 (L/min)", "VCO2 (L/min)", "HR (bpm)", "SpO2 (%)", "VT (L)", "PetO2 (mmHg)", "PetCO2 (mmHg)", "RER", "VE/VO2", "VE/VCO2", "O2 Pulse (ml/beat)", "BP (mmHg)", "pH", "PaCO2 (mmHg)", "PaO2 (mmHg)", "SaO2 (%)", "HCO3 (mmol/L)", "BE (mmol/L)", "A-aPO2 (mmHg)", "Vd/Vt"}
        Dim col_name() As String = {"levelid", "#", "load", "time_min", "ve", "vo2", "vco2", "hr", "spo2", "vt", "peto2", "petco2", "rer", "vevo2", "vevco2", "o2pulse", "bp", "ph", "paco2", "pao2", "sao2", "hco3", "be", "aapo2", "vdvt"}
        Dim col_width() As Integer = {20, 30, 45, 40, 40, 40, 40, 40, 40, 40, 45, 45, 40, 45, 55, 40, 55, 45, 45, 45, 45, 45, 45, 45, 45}
        Dim col_visible() As Integer = {False, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True, True}
        Dim col_textalign() As DataGridViewContentAlignment = {Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc, Align_mc}

        Dim cell As DataGridViewCell = New DataGridViewTextBoxCell

        grdData.Columns.Clear()
        For i = 0 To UBound(header_text)
            c = New DataGridViewColumn
            c.HeaderText = header_text(i)
            c.Name = col_name(i)
            c.Width = col_width(i)
            c.Visible = col_visible(i)
            c.DefaultCellStyle.Alignment = col_textalign(i)
            c.CellTemplate = cell
            Me.grdData.Columns.Add(c)
        Next


        grdData.RowsDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8)
        grdData.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8)
        grdData.ColumnHeadersDefaultCellStyle.Alignment = Align_mc
        grdData.ColumnHeadersDefaultCellStyle.ForeColor = Color.SlateBlue

        grdData.RowHeadersVisible = True
        grdData.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        grdData.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        grdData.RowHeadersDefaultCellStyle.ForeColor = Color.White
        grdData.RowHeadersDefaultCellStyle.BackColor = Color.Black

    End Sub

    Private Sub btnImport_Click(sender As System.Object, e As System.EventArgs) Handles btnImport.Click

        If grdData.Rows.GetRowCount(DataGridViewElementStates.Visible) > 0 Then
            If MsgBox("Current data in the grid will be lost. Continue?", vbOKCancel, "CPET") = vbCancel Then Exit Sub
        End If

        'Need to preserve the levelIDs for deleting records
        For i = 0 To grdData.Rows.Count - 1
            If Val(grdData(0, i).Value) > 0 Then 'Don't undelete already deleted items
                grdData(0, i).Value = Val(grdData(0, i).Value) * -1
                grdData.Rows(i).Visible = False
            End If
        Next

        Dim w() As cpet_loaddata = Me.Import_cpetdata_ErgoCPXData
        Me.load_datagrid(w)


    End Sub

    Private Function Import_cpetdata_ErgoCPXData() As cpet_loaddata()

        Dim path As String = "C:\Users\Public\Vb.net applications\ResLab V1 (Heavy).Net\"
        Dim reader As StreamReader = New StreamReader(path & "Example CPX raw data from medisoft.txt")
        Dim cpetSystem As eCpetSystems
        Dim c() As cpet_loaddata = Nothing
        Dim lineofdata As String = ""

        lineofdata = reader.ReadLine
        Select Case Trim(lineofdata)
            Case "???"
                cpetSystem = eCpetSystems.Medisoft_ergocard
                c = Me.Decode_ErgoCPXData(reader)
            Case Else

        End Select

        Return c

    End Function

    Private Sub load_datagrid(c() As cpet_loaddata)
        'Load the grid

        Dim vevo2 As String = "", vevco2 As String = "", o2pulse As String = "", loadlabel As String = "", timelabel As String = "", MeasPerMin As Integer = 0
        Dim s As String()
        Dim Row As Integer = 1

        If Not IsNothing(c) Then
            MeasPerMin = 60 / Val(txtInterval.Text)
            Row = 1
            Dim r As DataGridViewRow
            For Each d As cpet_loaddata In c
                r = New DataGridViewRow
                r = grdData.RowTemplate
                'r.DefaultCellStyle = grdData.RowTemplate
                If Val(d.vo2) > 0 And Val(d.ve) > 0 Then vevo2 = Format(d.ve / d.vo2, "0.#") Else vevo2 = ""
                If Val(d.vco2) > 0 And Val(d.ve) > 0 Then vevco2 = Format(d.ve / d.vco2, "0.#") Else vevco2 = ""
                If Val(d.vo2) > 0 And Val(d.hr) > 0 Then o2pulse = Format(1000 * d.vo2 / d.hr, "0.#") Else o2pulse = ""

                Select Case Row
                    Case Is <= Val(txtRests.Text) : loadlabel = "Rest " & Row
                    Case Is <= Val(txtRests.Text) + Val(txtZeros.Text) : loadlabel = "Zero " & Row - Val(txtRests.Text)
                    Case Else : loadlabel = Val(txtStart.Text) + Math.Truncate((Row - 1 - Val(txtRests.Text) - Val(txtZeros.Text)) / MeasPerMin) * Val(txtIncrements.Text)
                End Select
                timelabel = Format(Row * CSng(txtInterval.Text) / 60.0, "0.0")
                s = New String() {d.loadID, Row, loadlabel, timelabel, d.ve, d.vo2, d.vco2, d.hr, d.spo2, d.vt, d.peto2, d.petco2, d.rer, vevo2, vevco2, o2pulse, d.bp, d.ph, d.paco2, d.pao2, d.sao2, d.hco3, d.be, d.aapo2, d.vdvt}
                r.SetValues(s.ToArray)
                grdData.Rows.Add(r)
                'grdData.Rows(grdData.Rows.Count - 1).HeaderCell.Value = Row
                Row = Row + 1
            Next
        End If

    End Sub

    Private Sub load_gridpattern()

        Dim loadlabel As String = ""
        Dim timelabel As String = ""
        Dim rowCounter As Integer = 1
        Dim R As Integer = Val(txtRests.Text)
        Dim Z As Integer = Val(txtZeros.Text)
        Dim L As Integer = (Val(txtFinishLoad.Text) - Val(txtStart.Text)) / Val(txtIncrements.Text) + 1
        Dim nAllLoads As Integer = L + Z + R
        Dim nMeasPerLoad As Integer = Val(txtLoadIncInterval.Text) / Val(txtInterval.Text)
        Dim Inc As Integer = Val(txtIncrements.Text)
        Dim Interval As Integer = Val(txtInterval.Text)
        Dim s() As String

        For i = 0 To nAllLoads - 1
            Select Case i
                Case Is <= R : If R > 0 Then loadlabel = "Rest " & i
                Case Is <= R + Z : If Z > 0 Then loadlabel = "Zero " & i - R
                Case Else : If L > 0 Then loadlabel = Val(txtStart.Text) + Math.Truncate(i - R - Z) * Inc
            End Select
            If loadlabel <> "" Then
                For j = 1 To nMeasPerLoad
                    timelabel = Format(rowCounter * Interval / 60.0, "0.0#")
                    s = New String() {"0", rowCounter.ToString, loadlabel, timelabel}
                    grdData.Rows.Add(s)
                    rowCounter = rowCounter + 1
                Next
            End If
        Next

    End Sub

    Private Function Decode_ErgoCPXData(str As StreamReader) As cpet_loaddata()

        Dim c() As cpet_loaddata = Nothing
        Dim s As String = ""
        Dim vals() As String
        Dim noDataFound As Boolean = True
        Dim i As Integer = 0

        'Skip header to first line of data
        Do While Not str.EndOfStream And Strings.InStr(s, "Every ") = 0
            s = str.ReadLine
        Loop

        'Decode data strings
        i = 0
        s = str.ReadLine
        Do While Not str.EndOfStream
            s = Replace(s, " ", "")   'remove any spaces
            If Not IsNothing(s) Then
                If s.Length > 0 Then
                    ReDim Preserve c(i)
                    s = Mid(s, 2, Len(s) - 2)   'remove first and last separator chars
                    vals = Split(s, "�")
                    c(i).load = vals(0)
                    c(i).ve = vals(1)
                    c(i).vo2 = vals(2)
                    c(i).vco2 = vals(3)
                    c(i).hr = vals(4)
                    c(i).spo2 = vals(5)
                    c(i).vt = vals(6)
                    c(i).peto2 = vals(7)
                    c(i).petco2 = vals(8)
                    c(i).rer = vals(9)
                    i = i + 1
                    noDataFound = False
                End If
            End If
            s = str.ReadLine
        Loop

        'Handle no data condition
        If noDataFound Then
            MsgBox("No CPET data found in the file", vbOKOnly, "CPET data import")
        End If

        Return c

    End Function

    Private Sub load_textboxes_maxvaluesfromgrid()

        txtmax_ve.Text = Me.get_maxvalue_fromgridcolumn(grdData, 4, 1)
        txtmax_vo2.Text = Me.get_maxvalue_fromgridcolumn(grdData, 5, 2)
        txtmax_vco2.Text = Me.get_maxvalue_fromgridcolumn(grdData, 6, 2)
        txtmax_hr.Text = Me.get_maxvalue_fromgridcolumn(grdData, 7, 0)
        If Val(Val(txtWeight.Text)) > 0 Then txtmax_vo2kg.Text = Format(1000 * Val(txtmax_vo2.Text) / Val(txtWeight.Text), "0.0") Else txtmax_vo2kg.Text = ""
        txtmax_work.Text = Me.get_maxvalue_fromgridcolumn(grdData, 2, 0)
        txtmax_o2pulse.Text = Me.get_maxvalue_fromgridcolumn(grdData, 15, 1)
        txtmax_vt.Text = Me.get_maxvalue_fromgridcolumn(grdData, 9, 2)

    End Sub

    Private Function get_maxvalue_fromgridcolumn(ByRef g As DataGridView, col As Integer, decplaces As Integer) As String
        'Don't consider empty cells or zero values

        Dim a() As Single = Nothing
        Dim i As Integer = 0
        Dim formatmask As String = "0." & Strings.StrDup(decplaces, "0")

        For Each row As DataGridViewRow In g.Rows
            If row.Visible Then
                If IsNumeric(row.Cells(col).Value) And Val(row.Cells(col).Value) <> 0 Then
                    ReDim Preserve a(i)
                    a(i) = CSng(row.Cells(col).Value)
                    i = i + 1
                End If
            End If
        Next

        If Not IsNothing(a) Then
            Array.Sort(a)
            Return Format(a(UBound(a)), formatmask)
        Else
            Return ""
        End If

    End Function

    Private Function get_minvalue_fromgridcolumn(ByRef g As DataGridView, col As Integer) As String
        'Don't consider empty cells or zero values or invisible rows

        Dim a() As Single = Nothing
        Dim i As Integer = 0

        For Each row As DataGridViewRow In g.Rows
            If row.Visible Then
                If IsNumeric(row.Cells(col).Value) And Val(row.Cells(col).Value) <> 0 Then
                    ReDim Preserve a(i)
                    a(i) = CSng(row.Cells(col).Value)
                    i = i + 1
                End If
            End If
        Next

        If Not IsNothing(a) Then
            Array.Sort(a)
            Return a(LBound(a)).ToString
        Else
            Return ""
        End If

    End Function

    Private Function get_values_fromgridcolumn(ByRef g As DataGridView, col_name As String, mode As eGetValuesModes) As Single()
        'Get all values in column (only visible rows) - non-numeric values are set to nothing

        Dim a() As Single = Nothing
        Dim i As Integer = 0

        Select Case mode
            Case eGetValuesModes.All_values
                ReDim a(g.Rows.GetRowCount(DataGridViewElementStates.Visible) - 1)
                For Each row As DataGridViewRow In g.Rows
                    If row.Visible Then
                        If IsNumeric(row.Cells(col_name).Value) Then
                            a(i) = CSng(row.Cells(col_name).Value)
                        Else
                            a(i) = Nothing
                        End If
                        i = i + 1
                    End If
                Next
            Case eGetValuesModes.End_of_minute_values



        End Select

        Return a

    End Function

    Private Sub tabResults_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles tabResults.SelectedIndexChanged

        If tabResults.SelectedTab.Name.ToLower = "tabpage_main" Then
            Select Case _isNewCpet
                Case True : Me.load_textboxeswithpreds(Me.get_CpetPreds(class_Pred.eLoadNormalsMode.UseCurrentPrefs))
                Case False : Me.load_textboxeswithpreds(Me.get_CpetPreds(class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate))
            End Select

            Me.load_textboxes_maxvaluesfromgrid()
            Me.do_calculations()
            Me.plot_graphs()
        End If

    End Sub

    Private Sub btnSession_Click(sender As System.Object, e As System.EventArgs) Handles btnSession.Click
        _form_session._cancelButton_enable = False
        _form_session.Visible = True
        _form_session.BringToFront()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As System.EventArgs) Handles btnClose.Click
        _form_session.Dispose()
        Me.Close()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click

        'Save session
        Dim sessionID As Long = _form_session._sessionID
        Select Case sessionID
            Case 0 : sessionID = cRfts.Insert_rft_session(Me.Load_MemFromScreenFields_session(True))
            Case Else : cRfts.Update_rft_session(Me.Load_MemFromScreenFields_session(True))
        End Select

        'Save cpet test and level data
        Select Case _isNewCpet
            Case True : Me._cpetID = cRfts.Insert_rft_cpet(sessionID, Me.Load_MemFromScreenFields_cpet(True), Me.Load_MemFromScreenFields_cpet_levels(grdData, True))
            Case False : cRfts.Update_rft_cpet(sessionID, Load_MemFromScreenFields_cpet(True), Me.Load_MemFromScreenFields_cpet_levels(grdData, True))
        End Select

        'Refresh the list of tests and pdf to reflect any changes
        If Me._callingFormName = "Form_MainNew" Then gRefreshMainForm = True

        Me.Close()

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
        d(R.Req_address) = q & Trim(_form_session.cmboReqMO_address.Text) & q
        d(R.Req_providernumber) = q & Trim(_form_session.txtReqMO_pn.Text) & q
        d(R.Req_phone) = q & Trim(_form_session.txtReqMO_phone.Text) & q
        d(R.Req_fax) = q & Trim(_form_session.txtReqMO_fax.Text) & q
        d(R.Req_email) = q & Trim(_form_session.txtReqMO_email.Text) & q
        d(R.Report_copyto) = q & Trim(_form_session.cmboReportCopyTo.Text) & q
        d(R.Req_clinicalnotes) = q & Trim(_form_session.txtClinicalNote.Text) & q
        d(R.Req_healthservice) = q & Trim(_form_session.cmboReqMO_HealthService.Text) & q
        d(R.AdmissionStatus) = q & Trim(_form_session.cmboAdmission.Text) & q
        d(R.Billing_billedto) = q & Trim(_form_session.cmboBilledTo.Text) & q
        d(R.Billing_billingMO) = q & Trim(_form_session.cmboBillingMO_name.Text) & q
        d(R.Billing_billingMOproviderno) = q & Trim(_form_session.cmboBillingMO_pn.Text) & q
        d(R.LastUpdated_session) = cMyRoutines.FormatDBDateTime(Date.Now)
        d(R.LastUpdatedBy_session) = q & My.User.Name & q
        Return d

    End Function

    Private Function Load_MemFromScreenFields_cpet(AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicCPET
        Dim R As New class_fields_CPETAndSessionFields

        d(R.patientID) = Me._patientID
        d(R.sessionID) = Nothing        'set in save
        d(R.cpetID) = Me._cpetID
        d(R.TestTime) = cMyRoutines.FormatDBTime(txtTestTime.Text)
        d(R.Lab) = q & cmboLab.Text & q
        d(R.TestType) = "'CPET'"
        d(R.Scientist) = q & cmboScientist.Text & q
        d(R.BDStatus) = q & cmboLastBD.Text & q
        d(R.TechnicalNotes) = q & txtTechnicalNote.Text & q
        d(R.Report_text) = q & txtReport.Text & q
        d(R.Report_reportedby) = q & cmboReportedBy.Text & q
        d(R.Report_reporteddate) = cMyRoutines.FormatDBDate(txtReportedDate.Text)
        d(R.Report_verifiedby) = q & cmboVerifiedBy.Text & q
        d(R.Report_verifieddate) = cMyRoutines.FormatDBDate(txtVerifiedDate.Text)
        d(R.Report_status) = q & cmboReportStatus.Text & q

        If _callingFormName = "form_Reporting" And cmboReportStatus.SelectedIndex < cmboReportStatus.Items.Count - 1 Then
            Dim f As New form_UpdateReportStatus(cmboReportStatus.Text)
            Dim Response As String = f.ShowDialog(Me)
            If Response = Windows.Forms.DialogResult.OK Then
                d(R.Report_status) = q & f.cmboUpdate.Text & q
            Else
                d(R.Report_status) = q & cmboReportStatus.Text & q
            End If
            f = Nothing
        Else
            d(R.Report_status) = q & cmboReportStatus.Text & q
        End If

        d(R.r_bp_pre) = q & txtbp_pre.Text & q
        d(R.r_bp_post) = q & txtbp_post.Text & q
        d(R.r_spiro_pre_fev1) = q & txtSpiro_fev1.Text & q
        d(R.r_spiro_pre_fvc) = q & txtSpiro_fvc.Text & q
        d(R.r_symptoms_dyspnoea_borg) = q & txtBorg_dyspnoea.Text & q
        d(R.r_symptoms_legs_borg) = q & txtBorg_legs.Text & q
        d(R.r_symptoms_other_borg) = q & txtBorg_other.Text & q
        d(R.r_symptoms_other_text) = q & txtBorg_othersymptomtext.Text & q
        d(R.r_abg_post_be) = q & txtabg_be.Text & q
        d(R.r_abg_post_fio2) = q & "" & q
        d(R.r_abg_post_hco3) = q & txtabg_hco3.Text & q
        d(R.r_abg_post_paco2) = q & txtabg_co2.Text & q
        d(R.r_abg_post_pao2) = q & txtabg_o2.Text & q
        d(R.r_abg_post_sao2) = q & txtabg_sao2.Text & q
        d(R.r_abg_post_ph) = q & txtabg_ph.Text & q
        d(R.r_max_hr) = q & txtmax_hr.Text & q
        d(R.r_max_hr_lln) = q & txtmax_pred_hr.Text & q
        d(R.r_max_hr_mpv) = q & txtmax_pred_hr.Tag & q
        d(R.r_max_o2pulse) = q & txtmax_o2pulse.Text & q
        d(R.r_max_o2pulse_lln) = q & txtmax_pred_o2pulse.Text & q
        d(R.r_max_o2pulse_mpv) = q & txtmax_o2pulse.Tag & q
        d(R.r_max_vco2) = q & txtmax_vco2.Text & q
        d(R.r_max_vco2_lln) = q & txtmax_pred_vco2.Text & q
        d(R.r_max_vco2_mpv) = q & txtmax_vco2.Tag & q
        d(R.r_max_vo2) = q & txtmax_vo2.Text & q
        d(R.r_max_vo2_lln) = q & txtmax_pred_vo2.Text & q
        d(R.r_max_vo2_mpv) = q & txtmax_vo2.Tag & q
        d(R.r_max_ve) = q & txtmax_ve.Text & q
        d(R.r_max_ve_lln) = q & txtmax_pred_ve.Text & q
        d(R.r_max_ve_mpv) = q & txtmax_ve.Tag & q
        d(R.r_max_vt) = q & txtmax_vt.Text & q
        d(R.r_max_vt_lln) = q & txtmax_pred_vt.Text & q
        d(R.r_max_vt_mpv) = q & txtmax_vt.Tag & q
        d(R.r_max_vo2kg) = q & txtmax_vo2kg.Text & q
        d(R.r_max_vo2kg_lln) = q & txtmax_pred_vo2kg.Text & q
        d(R.r_max_vo2kg_mpv) = q & txtmax_vo2kg.Tag & q
        d(R.r_max_workload) = q & txtmax_work.Text & q
        d(R.r_max_workload_lln) = q & txtmax_pred_work.Text & q
        d(R.r_max_workload_mpv) = q & txtmax_work.Tag & q
        d(R.LastUpdated_cpet) = cMyRoutines.FormatDBDateTime(Date.Now)
        d(R.LastUpdatedBy_cpet) = q & My.User.Name & q

        Return d

    End Function

    Private Function Load_MemFromScreenFields_cpet_levels(g As DataGridView, AddQuotesAroundStrings As Boolean) As Dictionary(Of String, String)()

        Dim q As String = ""
        If AddQuotesAroundStrings Then q = "'"

        Dim i As Integer = 0, j As Integer = 0
        Dim row As DataRow = Nothing
        Dim d() As Dictionary(Of String, String) = Nothing
        Dim R As New class_fields_Cpet_Levels

        For j = 0 To g.Rows.Count - 1
            If g.Rows(j).Visible = False And g(0, j).Value = "0" Then
                'Skip new but deleted rows
            Else
                ReDim Preserve d(i)
                d(i) = cMyRoutines.MakeEmpty_dicCPET_levels
                d(i)(R.levelID) = CLng(g("levelid", i).Value)
                d(i)(R.cpetID) = Me._cpetID
                d(i)(R.level_number) = q & g("#", i).Value & q
                d(i)(R.level_time) = q & g("time_min", i).Value & q
                d(i)(R.level_workload) = q & g("load", i).Value & q
                d(i)(R.level_vo2) = q & g("vo2", i).Value & q
                d(i)(R.level_vco2) = q & g("vco2", i).Value & q
                d(i)(R.level_ve) = q & g("ve", i).Value & q
                d(i)(R.level_vt) = q & g("vt", i).Value & q
                d(i)(R.level_spo2) = q & g("spo2", i).Value & q
                d(i)(R.level_hr) = q & g("hr", i).Value & q
                d(i)(R.level_peto2) = q & g("peto2", i).Value & q
                d(i)(R.level_petco2) = q & g("petco2", i).Value & q
                d(i)(R.level_rer) = q & g("rer", i).Value & q
                d(i)(R.level_vevo2) = q & g("vevo2", i).Value & q
                d(i)(R.level_vevco2) = q & g("vevco2", i).Value & q
                d(i)(R.level_o2pulse) = q & g("o2pulse", i).Value & q

                d(i)(R.level_bp) = q & g("bp", i).Value & q
                d(i)(R.level_ph) = q & g("ph", i).Value & q
                d(i)(R.level_paco2) = q & g("paco2", i).Value & q
                d(i)(R.level_pao2) = q & g("pao2", i).Value & q
                d(i)(R.level_sao2) = q & g("sao2", i).Value & q
                d(i)(R.level_hco3) = q & g("hco3", i).Value & q
                d(i)(R.level_be) = q & g("be", i).Value & q
                d(i)(R.level_aapo2) = q & g("aapo2", i).Value & q
                d(i)(R.level_vdvt) = q & g("vdvt", i).Value & q

                i = i + 1
            End If
        Next

        Return d

    End Function

    Private Sub Load_ScreenFieldsFromMem(c As Dictionary(Of String, String), w() As Dictionary(Of String, String))

        Select Case _isNewCpet
            Case False
                'Load the test data
                Dim f As New class_fields_CPETAndSessionFields
                txtReport.Text = c(f.Report_text)
                If IsDate(c(f.Report_reporteddate)) Then txtReportedDate.Text = c(f.Report_reporteddate)
                cmboReportedBy.Text = c(f.Report_reportedby)
                cmboVerifiedBy.Text = c(f.Report_verifiedby)
                If IsDate(c(f.Report_verifieddate)) Then txtVerifiedDate.Text = c(f.Report_verifieddate)
                cmboReportStatus.Text = c(f.Report_status)
                cmboScientist.Text = c(f.Scientist)
                If IsDate(c(f.TestTime).ToString) Then txtTestTime.Text = c(f.TestTime)
                cmboLastBD.Text = c(f.BDStatus)
                cmboLab.Text = c(f.Lab)
                txtTechnicalNote.Text = c(f.TechnicalNotes)

                txtbp_pre.Text = c(f.r_bp_pre)
                txtbp_post.Text = c(f.r_bp_post)
                txtSpiro_fev1.Text = c(f.r_spiro_pre_fev1)
                txtSpiro_fvc.Text = c(f.r_spiro_pre_fvc)
                txtBorg_dyspnoea.Text = c(f.r_symptoms_dyspnoea_borg)
                txtBorg_legs.Text = c(f.r_symptoms_legs_borg)
                txtBorg_other.Text = c(f.r_symptoms_other_borg)
                txtBorg_othersymptomtext.Text = c(f.r_symptoms_other_text)
                'txtabg_be.Text = c(f.r_abg_post_be)
                'c(f.r_abg_post_fio2) = ""
                'txtabg_hco3.Text = c(f.r_abg_post_hco3)
                'txtabg_co2.Text = c(f.r_abg_post_paco2)
                'txtabg_o2.Text = c(f.r_abg_post_pao2)
                'txtabg_sao2.Text = c(f.r_abg_post_sao2)
                'txtabg_ph.Text = c(f.r_abg_post_ph)

                'Load the workload data into data structure to pass to grid load routine 
                Dim i As Integer = 0
                Dim ff As New class_fields_Cpet_Levels
                Dim d() As cpet_loaddata = Nothing
                If Not IsNothing(w) Then
                    For Each wl In w
                        ReDim Preserve d(i)
                        d(i).hr = wl(ff.level_hr)
                        d(i).load = wl(ff.level_workload)
                        d(i).loadID = wl(ff.levelID)
                        d(i).loadNumber = wl(ff.level_number)
                        d(i).loadtime = wl(ff.level_time)
                        d(i).petco2 = wl(ff.level_petco2)
                        d(i).peto2 = wl(ff.level_peto2)
                        d(i).rer = wl(ff.level_rer)
                        d(i).spo2 = wl(ff.level_spo2)
                        d(i).testID = wl(ff.cpetID)
                        d(i).vco2 = wl(ff.level_vco2)
                        d(i).ve = wl(ff.level_ve)
                        d(i).vo2 = wl(ff.level_vo2)
                        d(i).vt = wl(ff.level_vt)
                        d(i).bp = wl(ff.level_bp)
                        d(i).ph = wl(ff.level_ph)
                        d(i).paco2 = wl(ff.level_paco2)
                        d(i).pao2 = wl(ff.level_pao2)
                        d(i).sao2 = wl(ff.level_sao2)
                        d(i).hco3 = wl(ff.level_hco3)
                        d(i).be = wl(ff.level_be)
                        d(i).aapo2 = wl(ff.level_aapo2)
                        d(i).vdvt = wl(ff.level_vdvt)

                        i = i + 1
                    Next
                End If
                Me.load_datagrid(d)

            Case True

        End Select

    End Sub



    Private Sub Label85_DoubleClick(sender As Object, e As System.EventArgs) Handles Label85.DoubleClick

        If Not IsDate(txtReportedDate.Text) Then txtReportedDate.Text = Now

    End Sub


    Private Sub Label61_DoubleClick(sender As Object, e As System.EventArgs) Handles Label61.DoubleClick

        If Not IsDate(txtVerifiedDate.Text) Then txtVerifiedDate.Text = Now

    End Sub

    Private Sub btnManualEntry_Click(sender As Object, e As EventArgs) Handles btnManualEntry.Click

        gbIncrementPattern.Visible = True
        gbDataEntry.Enabled = False
        gbRowEditing.Enabled = False

    End Sub

    Private Sub btnAddIncrementPattern_Click(sender As Object, e As EventArgs) Handles btnAddIncrementPattern.Click

        If grdData.Rows.GetRowCount(DataGridViewElementStates.Visible) > 0 Then
            If MsgBox("Current data in the grid will be lost. Continue?", vbOKCancel, "CPET") = vbCancel Then Exit Sub
        End If

        'Need to preserve the levelIDs for deleting records
        For i = 0 To grdData.Rows.Count - 1
            grdData(0, i).Value = Val(grdData(0, i).Value) * -1
            grdData.Rows(i).Visible = False
        Next

        load_gridpattern()
        gbRowEditing.Visible = True
        gbRowEditing.Enabled = True
        gbIncrementPattern.Visible = False
        gbDataEntry.Enabled = True

    End Sub

    Private Sub btnRow_clearall_Click(sender As Object, e As EventArgs) Handles btnRow_clearall.Click

        If grdData.Rows.GetRowCount(DataGridViewElementStates.Visible) > 0 Then
            If MsgBox("Current data in the grid will be lost. Continue?", vbOKCancel, "CPET") = vbCancel Then Exit Sub
        End If
        Select Case _isNewCpet
            Case True
                grdData.Rows.Clear()
            Case False
                'Need to preserve the levelIDs for deleting records
                For i = 0 To grdData.Rows.Count - 1
                    If Val(grdData(0, i).Value) > 0 Then    'Dont undelete already deleted records
                        grdData(0, i).Value = (Val(grdData(0, i).Value) * -1).ToString
                        grdData.Rows(i).Visible = False
                    End If
                Next
        End Select

    End Sub

    Private Sub btnRow_add_Click(sender As Object, e As EventArgs) Handles btnRow_add.Click

        Dim rowcounter As Integer = 0
        Dim LastVisibleRow As Integer = grdData.Rows.GetLastRow(DataGridViewElementStates.Visible)
        If LastVisibleRow > -1 Then
            rowcounter = grdData(1, LastVisibleRow).Value + 1
        Else
            rowcounter = 1
        End If

        grdData.Rows.Add(0, rowcounter)
        LastVisibleRow = grdData.Rows.GetLastRow(DataGridViewElementStates.Visible)
        grdData.CurrentCell = grdData(2, LastVisibleRow)

    End Sub

    Private Sub btnRow_delete_Click(sender As Object, e As EventArgs) Handles btnRow_delete.Click

        If IsNothing(grdData.CurrentRow) Then
            MsgBox("Please select a row to delete.", vbOKOnly, "CPET")
            Exit Sub
        End If

        If MsgBox("About to delete row " & grdData(1, grdData.CurrentRow.Index).Value & ". Continue?", vbOKCancel, "CPET") = vbCancel Then
            Exit Sub
        Else
            grdData(0, grdData.CurrentRow.Index).Value = (Val(grdData(0, grdData.CurrentRow.Index).Value) * -1).ToString
            grdData.CurrentRow.Visible = False
            'Re-number visible rows
            Dim j As Integer = 1
            For Each r As DataGridViewRow In grdData.Rows
                If r.Visible Then
                    r.Cells(1).Value = j.ToString
                    j = j + 1
                End If
            Next
        End If

    End Sub

    Private Sub btnCancelPattern_Click(sender As Object, e As EventArgs) Handles btnCancelPattern.Click

        gbDataEntry.Enabled = True
        gbIncrementPattern.Visible = False
        gbRowEditing.Enabled = True

    End Sub

    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click

        If cRfts.Delete_rft(Me._cpetID, eTables.r_cpet) Then
            'Refresh the list of tests and pdf to reflect any changes
            If Me._callingFormName = "Form_MainNew" Then gRefreshMainForm = True
            Me.Close()
        Else
            'If delete cancelled keep form open
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

        _formPhrases = New form_rft_phrases(eAutoreport_testgroups.Cpx, Me)
        _formPhrases.ReportTextBox = Me.txtReport
        _formPhrases.Show()

    End Sub


End Class