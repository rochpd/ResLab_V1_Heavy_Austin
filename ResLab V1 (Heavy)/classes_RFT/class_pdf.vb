
Imports Gnostice.PDFOne
Imports System.Drawing.Printing
Imports System.Drawing.Imaging
Imports Microsoft.VisualBasic.FileIO
Imports System.Text
Imports ResLab_V1_Heavy.class_walktest
Imports System.Windows.Forms.DataVisualization.Charting
Imports ResLab_V1_Heavy.cDatabaseInfo


Public Class class_pdf

    Public Enum eReportTypes
        RoutineRft = 1
        Challenge = 2
        GenericProv = 6
        HAST = 3
        SPT = 4
        CPET = 5
        WalkTest = 6
    End Enum

    Private Structure myMargins
        Dim left As Single
        Dim right As Single
        Dim top As Single
        Dim bottom As Single
        Dim xMin As Single
        Dim xMax As Single
        Dim yMin As Single
        Dim yMax As Single
        Dim xWidth As Single
        Dim yHeight As Single
    End Structure

    Private m As myMargins
    Dim pdf As PDFDocument
    Private _License As String = "3617-45ED-E413-ABEA-54FC-A2FC-9BA0-6686"
    Dim fCourier As PDFFont
    Dim fHel7 As PDFFont
    Dim fHel8 As PDFFont
    Dim fHel8b As PDFFont
    Dim fHel10 As PDFFont
    Dim fHel10b As PDFFont
    Dim fHel10i As PDFFont
    Dim fHel11 As PDFFont
    Dim fHel11b As PDFFont
    Dim fHel12 As PDFFont
    Dim fHel12b As PDFFont
    Dim fHel14 As PDFFont
    Dim fHel14b As PDFFont
    Dim fHel16 As PDFFont
    Dim fTimes As PDFFont

    Const ctp As Single = 72 / 2.54


    Public Sub New()

        'Note - PDFOne.net has a bug. Cursor position get and set only works in points (ie 1 inch=72 points)

        'Setup the blank pdf doc
        pdf = New PDFDocument(Me.License)
        pdf.MeasurementUnit = PDFMeasurementUnit.Centimeters

        'Set margin variables in cm
        m.left = 1  'actually 2cm from left edge of paper
        m.right = m.left + 17
        m.top = 1.5   'actually 0.6cm from top edge of paper
        m.bottom = m.top + 28

        Dim p As New PDFPage(PaperKind.A4, 2, 0.5, 0, 0, PDFMeasurementUnit.Centimeters)
        pdf.AddPage(p)

        m.xMin = m.left
        m.xMax = m.right    'p.GetWidth(PDFMeasurementUnit.Centimeters) - m.right
        m.yMin = m.top
        m.yMax = m.bottom   'p.GetHeight(PDFMeasurementUnit.Centimeters) - m.bottom
        m.xWidth = m.xMax - m.xMin
        m.yHeight = m.yMax - m.yMin

        fCourier = New PDFFont(StdType1Font.Courier, PDFFontStyle.Fill, 14)
        fHel7 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 7)
        fHel8 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 8)
        fHel8 = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 8)
        fHel10 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 9)
        fHel10b = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 9)
        fHel10i = New PDFFont(StdType1Font.Helvetica_Oblique, PDFFontStyle.Fill, 9)
        fHel11 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 11)
        fHel11b = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 11)
        fHel12 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 12)
        fHel12b = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 12)
        fHel14 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 14)
        fHel14b = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 16)
        fHel16 = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 16)
        fTimes = New PDFFont(StdType1Font.Times_Roman, PDFFontStyle.Fill, 14)

    End Sub

    Public Function cp(ByVal cm As Single) As Single
        'Converts cm to points
        Return 72 * cm / 2.54

    End Function

    Public ReadOnly Property License() As String
        Get
            Return _License
        End Get
    End Property

    Public Sub Draw_Pdf_F1_rft_main_routine(ByVal Results As Dictionary(Of String, String), ByVal Demographics As Dictionary(Of String, String))

        Dim yPos As Single = 0
        Dim ReturnVar As Boolean = False

        Try
            yPos = Me.Draw_Pdf_F1_rft_header(Results, Demographics)
            yPos = Me.Draw_Pdf_F1_rft_results_routine(Results, Demographics, yPos)
            ReturnVar = Me.Draw_Pdf_F1_report(Results, Demographics("race_forrfts"), yPos)
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F1_rft_main_routine'")
        End Try

    End Sub

    Public Sub Draw_Pdf_F1_rft_main_cpet(cpetID As Long, ByVal Demographics As Dictionary(Of String, String))

        Dim yPos As Single = 0
        Dim ReturnVar As Single = 0
        Dim Results As Dictionary(Of String, String) = cRfts.Get_rft_cpet_testdata(cpetID)

        Try
            yPos = Draw_Pdf_F1_rft_header(Results, Demographics)
            yPos = Me.Draw_Pdf_F1_rft_results_cpet(Results, Demographics, yPos)
            ReturnVar = Me.Draw_Pdf_F1_report(Results, Demographics("race_forrfts"), yPos)
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F1_rft_main_cpet'")
        End Try

    End Sub

    Public Sub Draw_Pdf_F1_rft_main_spt(sptID As Long, ByVal Demographics As Dictionary(Of String, String))

        Dim yPos As Single = 0
        Dim ReturnVar As Single = 0
        Dim Results_Test As Dictionary(Of String, String) = cRfts.Get_rft_spt_test_session(sptID)
        Dim Results_Allergens() As Dictionary(Of String, String) = cRfts.Get_rft_spt_allergenresults(sptID)

        Try
            yPos = Draw_Pdf_F1_rft_header(Results_Test, Demographics)
            yPos = Me.Draw_Pdf_F1_rft_results_spt(Results_Test, Results_Allergens, yPos)
            ReturnVar = Me.Draw_Pdf_F1_report(Results_Test, Demographics("race_forrfts"), yPos)
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F1_rft_main_spt'")
        End Try

    End Sub

    Public Sub Draw_Pdf_F1_rft_main_hast(hastID As Long, ByVal Demographics As Dictionary(Of String, String))

        Dim yPos As Single = 0
        Dim ReturnVar As Single = 0
        Dim Results As Dictionary(Of String, String) = cRfts.Get_rft_hast_test_session(hastID)

        Try
            yPos = Draw_Pdf_F1_rft_header(Results, Demographics)
            yPos = Me.Draw_Pdf_F1_rft_results_hast(Results, Demographics, yPos)
            ReturnVar = Me.Draw_Pdf_F1_report(Results, Demographics("race_forrfts"), yPos)
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F1_rft_main_hast'")
        End Try

    End Sub

    Public Sub Draw_Pdf_F1_rft_main_walk(walkID As Long, ByVal Demographics As Dictionary(Of String, String))

        Dim yPos As Single = 0
        Dim ReturnVar As Single = 0

        Try

            'Get walk info from results record into a dic object
            Dim WalkInfo As Dictionary(Of String, String) = cRfts.Get_walk_test_session(walkID)

            'Get walk trials from results record into a dic object
            Dim Trials() As Dictionary(Of String, String) = cRfts.Get_walk_trials(walkID)

            'Transfer trial fields to data structure
            Dim i As Integer = 0
            Dim T() As class_walktest_plot.walk_trialdata = Nothing

            Dim f1 As New class_fields_Walk_Trial
            For Each trial In Trials
                ReDim Preserve T(i)
                T(i).trial_distance = trial(f1.trial_distance)
                T(i).trial_label = trial(f1.trial_label)
                T(i).trial_number = trial(f1.trial_number)
                T(i).trial_timeofday = trial(f1.trial_timeofday)
                T(i).trialID = trial(f1.trialID)
                T(i).walkID = walkID
                i += 1
            Next

            'Get stages from results record into a dic object
            Dim Stages As List(Of Dictionary(Of String, String)()) = cRfts.Get_walk_levels(walkID)

            'Transfer stage fields to data structure
            Dim f2 As New class_fields_Walk_TrialLevel
            Dim j As Integer = 0
            i = 0
            For Each Trial As Dictionary(Of String, String)() In Stages
                ReDim Preserve T(i).timepoints(Trial.Count - 1)
                j = 0
                For Each stage As Dictionary(Of String, String) In Trial
                    T(i).timepoints(j).time_dyspnoea = stage(f2.time_dyspnoea)
                    T(i).timepoints(j).time_gradient = stage(f2.time_gradient)
                    T(i).timepoints(j).time_hr = stage(f2.time_hr)
                    T(i).timepoints(j).time_label = stage(f2.time_label)
                    T(i).timepoints(j).time_minute = stage(f2.time_minute)
                    T(i).timepoints(j).time_o2flow = stage(f2.time_o2flow)
                    T(i).timepoints(j).time_speed_kph = stage(f2.time_speed_kph)
                    T(i).timepoints(j).time_spo2 = stage(f2.time_spo2)
                    T(i).timepoints(j).timepointID = stage(f2.levelID)
                    j += 1
                Next
                i += 1
            Next





            yPos = Me.Draw_Pdf_F1_rft_header(WalkInfo, Demographics)
            yPos = Me.Draw_Pdf_F1_rft_results_walk(WalkInfo("testtype"), WalkInfo("protocolid"), T, yPos)
            yPos = 20
            ReturnVar = Me.Draw_Pdf_F1_report(WalkInfo, Demographics("race_forrfts"), yPos)
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F1_rft_main_walk'")
        End Try

    End Sub

    'Public Sub Draw_Pdf_F3_RoutineRft(ByVal Results As Dictionary(Of String, String), ByVal Demographics As Dictionary(Of String, String))

    '    Dim yPos As Single = 0
    '    Dim ReturnVar As Boolean = False

    '    Try
    '        yPos = Draw_Pdf_F3_Reportform_HeaderSection(Results, Demographics)
    '        yPos = Draw_Pdf_F2_Reportform_ResultsSection_RoutineRft(Results, Demographics, yPos)
    '        ReturnVar = Draw_Pdf_F2_Reportform_ReportSection(Results, Demographics("race_forrfts"), yPos)
    '        Me.Draw_FlowVolImage(Results("rftid"), eTables.Rft_Routine, m.xMin, m.yMax - 6.5, 5, 5)
    '    Catch
    '        MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F2_RoutineRft'")
    '    End Try

    'End Sub

    'Public Sub Draw_Pdf_F4_RoutineRft(ByVal Results As Dictionary(Of String, String), ByVal Demographics As Dictionary(Of String, String))

    '    Dim yPos As Single = 0
    '    Dim ReturnVar As Single = 0

    '    Try
    '        yPos = Draw_Pdf_F4_Reportform_HeaderSection(Results, Demographics)
    '        yPos = Draw_Pdf_F2_Reportform_ResultsSection_RoutineRft(Results, Demographics, yPos)
    '        ReturnVar = Draw_Pdf_F2_Reportform_ReportSection(Results, Demographics("race_forrfts"), yPos)
    '        Me.Draw_FlowVolImage(Results("rftid"), eTables.Rft_Routine, m.xMin, m.yMax - 7.5, 6, 6)

    '    Catch
    '        MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F4_RoutineRft'")
    '    End Try

    'End Sub

    Public Function Get_Pref_ReportStyle(ByVal ReportType As eReportTypes) As String

        Dim sql As String = "SELECT style_code FROM prefs_reports_styles WHERE report_typeID=" & ReportType

        Try

            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0).Item("style_code")
            Else
                Return Nothing
            End If

        Catch
            MsgBox("Error generating pdf in 'class_pdf.Get_Pref_ReportStyle'")
            Return Nothing
        End Try

    End Function

    Public Function Get_Pref_ReportStrings(Optional ByVal ReportType As eReportTypes = 0) As Dictionary(Of String, Dictionary(Of String, String))()

        Dim sql As String = "SELECT * FROM prefs_reportstrings WHERE typeID IS NULL"
        If ReportType <> 0 Then sql = sql & " OR typeID=" & ReportType

        Dim D() As Dictionary(Of String, Dictionary(Of String, String)) = Nothing
        Dim D2 As Dictionary(Of String, String) = Nothing

        Try
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            With Ds.Tables(0)
                For Each row As DataRow In .Rows
                    ReDim Preserve D(UBound(D) + 1)
                    D(UBound(D)) = New Dictionary(Of String, Dictionary(Of String, String))
                    D2.Clear()
                    For i As Integer = 0 To .Columns.Count - 1
                        D2.Add(.Columns(i).ColumnName.ToLower, row.Item(i).value)
                    Next
                    D(UBound(D)).Add(row.Item("Name").ToLower, D2)
                Next
            End With
            Return D
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Get_Pref_ReportStrings'")
            Return Nothing
        End Try

    End Function

    Public Function Draw_Pdf(ByVal PatientID As Long, ByVal RecordID As Long, ByVal eTbl As eTables) As PDFDocument

        Dim Results As Dictionary(Of String, String)
        Dim Demographics As Dictionary(Of String, String) = cPt.get_demographics(PatientID)
        Dim style As String = ""

        Select Case eTbl
            Case eTables.rft_routine
                Results = cRfts.Get_rft_byRftID(RecordID)
                style = Me.Get_Pref_ReportStyle(eReportTypes.RoutineRft)
                Select Case style
                    Case "F1" : Draw_Pdf_F1_rft_main_routine(Results, Demographics)
                        'Case "F2" : Draw_Pdf_F2_RoutineRft(Results, Demographics)
                        'Case "F3" : Draw_Pdf_F3_RoutineRft(Results, Demographics)
                        'Case "F4" : Draw_Pdf_F4_RoutineRft(Results, Demographics)
                    Case Else : MsgBox("A report style has not been set for this report")
                End Select
            Case eTables.Prov_test
                style = Me.Get_Pref_ReportStyle(eReportTypes.GenericProv)
                Select Case style
                    Case "F1" : Draw_Pdf_F1_rft_main_prov(RecordID, Demographics)
                        ' Case "F2" : Draw_Pdf_F2_GenericProv(RecordID, Demographics)
                        'Case "F3" : Draw_Pdf_F3_GenericProv(RecordID, Demographics)
                        'Case "F4" : Draw_Pdf_F4_GenericProv(RecordID, Demographics)
                    Case Else : MsgBox("A report style has not been set for this report")
                End Select
            Case eTables.r_walktests_v1heavy
                style = Me.Get_Pref_ReportStyle(eReportTypes.WalkTest)
                Select Case style
                    Case "F1" : Me.Draw_Pdf_F1_rft_main_walk(RecordID, Demographics)
                        'Case "F2" : Draw_Pdf_F2_WalkTest(RecordID, Demographics)
                        'Case "F3" : Draw_Pdf_F3_WalkTest(RecordID, Demographics)
                        'Case "F4" : Draw_Pdf_F4_WalkTest(RecordID, Demographics)
                    Case Else : MsgBox("A report style has not been set for this report")
                End Select
            Case eTables.r_cpet
                style = Me.Get_Pref_ReportStyle(eReportTypes.CPET)
                Select Case style
                    Case "F1" : Draw_Pdf_F1_rft_main_cpet(RecordID, Demographics)
                        'Case "F2"
                        'Case "F3"
                        'Case "F4" 
                    Case Else : MsgBox("A report style has not been set for this report")
                End Select
            Case eTables.r_spt
                style = Me.Get_Pref_ReportStyle(eReportTypes.SPT)
                Select Case style
                    Case "F1" : Draw_Pdf_F1_rft_main_spt(RecordID, Demographics)
                        'Case "F2"
                        'Case "F3"
                        'Case "F4" 
                    Case Else : MsgBox("A report style has not been set for this report")
                End Select
            Case eTables.r_hast
                style = Me.Get_Pref_ReportStyle(eReportTypes.HAST)
                Select Case style
                    Case "F1" : Draw_Pdf_F1_rft_main_hast(RecordID, Demographics)
                        'Case "F2"
                        'Case "F3"
                        'Case "F4" 
                    Case Else : MsgBox("A report style has not been set for this report")
                End Select
        End Select


        Return pdf

        pdf.Dispose()

    End Function

    Public Function Draw_Pdf_trend(ByVal PatientID As Long, g As DataGridView, ch As Chart) As PDFDocument

        Dim ypos As Single = 0
        Dim style As String = "F2"  'Me.Get_Pref_ReportStyle(eReportTypes.RoutineRft)

        Select Case style
            Case "F1" : Me.Draw_Pdf_F1_rft_trend(cPt.get_demographics(PatientID), g, ch)
            Case "F2"
            Case "F3"
            Case "F4"
        End Select

        Return pdf

        pdf.Dispose()

    End Function


    'Private Sub Draw_Pdf_F2_WalkTest(ByVal walkID As Long, ByVal Demographics As Dictionary(Of String, String))

    '    Dim yPos As Single = 0
    '    Dim returnVar As Boolean = False

    '    Dim i As Integer = 0

    '    Try
    '        Dim walkdata As Dictionary(Of String, String) = cPt.Get_walk_test_session(walkID)
    '        Dim trialsData As Dictionary(Of String, String)() = cPt.Get_walk_trials(walkID)
    '        Dim levelsData As List(Of Dictionary(Of String, String)()) = cPt.Get_walk_levels(walkID)

    '        yPos = Draw_Pdf_F1_rft_header(walkdata, Demographics)
    '        'yPos = Draw_Pdf_F2_Reportform_ResultsSection_WalkTest(walkdata, Demographics, yPos)
    '        'returnVar = Draw_Pdf_F2_Reportform_ReportSection(visitData, Demographics("race_forrfts"), yPos)
    '    Catch
    '        MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F2_WalkTest'" & vbCrLf & Err.Description)
    '    End Try


    'End Sub

    Private Sub Draw_Pdf_F1_rft_main_prov(ByVal provid As Long, ByVal Demographics As Dictionary(Of String, String))

        Dim yPos As Single = 0
        Dim ReturnVar As Boolean = False
        Dim P As class_challenge.ProtocolData = Nothing
        Dim f1 As New class_fields_ProvAndSession
        Dim f2 As New class_ProvTestDataFields
        Dim i As Integer = 0

        Try

            'Get prov info from results record into a dic object
            Dim ProvInfo As Dictionary(Of String, String) = cRfts.Get_prov_test_session(provid)
            'Get prov test results from results record into a dic object
            Dim ProvResults() As Dictionary(Of String, String) = cRfts.Get_Prov_TestDataByProvID(provid)

            'Transfer needed fields to structure
            P.provid = provid
            P.pd_thresh = ProvInfo(f1.Protocol_threshold)
            P.pd_decimalplaces = ProvInfo(f1.Protocol_pd_decimalplaces)
            P.agent_units = ProvInfo(f1.Protocol_doseunits)
            P.agent = ProvInfo(f1.Protocol_drug)
            P.pd_method = ProvInfo(f1.Protocol_method)
            P.pd_method_reference = ProvInfo(f1.Protocol_method_reference)
            P.parameter = ProvInfo(f1.Protocol_parameter)
            P.pd_dose_effect = ProvInfo(f1.Protocol_dose_effect)
            P.pd = ProvInfo(f1.Pd)
            P.plot_ymin = ProvInfo(f1.plot_ymin)
            P.plot_ymax = ProvInfo(f1.plot_ymax)
            P.plot_ystep = ProvInfo(f1.plot_ystep)
            P.plot_xtitle = ProvInfo(f1.plot_xtitle)
            P.plot_xscaling_type = ProvInfo(f1.plot_xscaling_type)
            P.plot_title = ProvInfo(f1.plot_title)
            i = 0
            For Each dic As Dictionary(Of String, String) In ProvResults
                ReDim Preserve P.doses(i)
                P.doses(i).xaxislabel = dic(f2.xaxis_label)
                P.doses(i).dose_cumulative = dic(f2.dose_cumulative)
                P.doses(i).dose_discrete = dic(f2.dose_discrete)
                P.doses(i).dosenumber = dic(f2.dose_number)
                P.doses(i).doseID = dic(f2.doseid)
                P.doses(i).response = dic(f2.response)
                P.doses(i).result = dic(f2.result)
                P.doses(i).time_min = dic(f2.dose_time_min)
                P.doses(i).xaxislabel = dic(f2.xaxis_label)
                i = i + 1
            Next

            yPos = Draw_Pdf_F1_rft_header(ProvInfo, Demographics)
            yPos = Draw_Pdf_F1_rft_results_prov(ProvResults, ProvInfo, Demographics, yPos)
            ReturnVar = Me.Draw_Pdf_F1_report(ProvInfo, Demographics("race_forrfts"), yPos)
        Catch
            MsgBox("Error generating pdf in 'class_pdf.Draw_Pdf_F1_rft_main_prov'" & vbCrLf & Err.Description)
        End Try


    End Sub

    Private Function Draw_Pdf_F1_rft_results_prov(ByVal R() As Dictionary(Of String, String), ByVal T As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String), ByVal yPos As Single) As Single

        Try

            Dim p As PDFPage = pdf.GetPage(1)
            Dim Ystart As Single = 0
            Dim Y As Single = m.yMin + 5
            Dim incResultsHeadings As Single = 0.25
            Dim incResultsLines As Single = 0.075
            Dim c1 As Single = m.xMin * ctp                     'Test header
            Dim c2 As Single = c1 + 1.5 * ctp                   'Units
            Dim c3 As Single = c2 + 3 * ctp                     'Normal range
            Dim c4 As Single = c3 + 3 * ctp                     'Baseline
            Dim c5 As Single = c4 + 1.6 * ctp                   'PPN
            Dim c6 As Single = c5 + 2.2 * ctp                   'Post mannitol
            Dim c7 As Single = c6 + 3.5 * ctp                   'Change
            Dim ppn As String = ""
            Dim fer_pre As String = ""
            Dim PD As String = ""
            Dim fld As New class_fields_ProvAndSession

            'Transfer raw data from dictionaries to data structure for processing
            Dim c As class_challenge.ProtocolData = cChall.load_provdataStructure_from_dictionaries(T, R)

            'Get normals
            Dim demo As New Pred_demo
            demo.Age = cMyRoutines.Calc_Age(CDate(D("dob")), CDate(T(fld.TestDate)))
            demo.Htcm = Val(T(fld.Height))
            demo.Wtkg = Val(T(fld.Weight))
            demo.GenderID = cMyRoutines.Lookup_list_ByDescription(D("gender"), eTables.Pred_ref_genders)
            demo.Gender = D("gender")
            demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(D("race_forrfts"), eTables.Pred_ref_ethnicities)
            demo.TestDate = T(fld.TestDate)
            demo.SourcesString = T(fld.Pred_SourceIDs)
            Dim dPreds As Dictionary(Of String, String) = cPred.Get_PredValues(demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)    'ParameterID|StatTypeID, result

            'Results header
            Ystart = yPos + incResultsHeadings * ctp
            p.CursorPosY = Ystart
            p.Font = fHel10
            p.CursorPosX = c1 : p.WriteText("BRONCHIAL CHALLENGE TEST: " & c.agent & vbCrLf & vbCrLf, fHel10b)

            p.CursorPosX = c1 : p.WriteText("SPIROMETRY", fHel10b)
            p.CursorPosX = c3 : p.WriteText("Normal range", fHel10b)
            p.CursorPosX = c4 : p.WriteText("Baseline", fHel10b)
            p.CursorPosX = c5 : p.WriteText("(%mpv)" & vbCrLf, fHel10b)

            'Spiro
            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("FEV1")
            p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
            p.CursorPosX = c4 : p.WriteText(cMyRoutines.fmt(T(fld.R_bl_Fev1), 2))
            If dPreds.ContainsKey("FEV1|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FEV1|LLN", dPreds, 1), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(T(fld.R_bl_Fev1)) / Val(dPreds("FEV1|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("FVC")
            p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
            p.CursorPosX = c4 : p.WriteText(cMyRoutines.fmt(T(fld.R_bl_Fvc), 2))
            If dPreds.ContainsKey("FVC|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FVC|LLN", dPreds, 1), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(T(fld.R_bl_Fvc)) / Val(dPreds("FVC|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("VC")
            p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
            p.CursorPosX = c4 : p.WriteText(cMyRoutines.fmt(T(fld.R_bl_Vc), 2))
            If dPreds.ContainsKey("VC|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("VC|LLN", dPreds, 1), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(T(fld.R_bl_Vc)) / Val(dPreds("VC|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("FER")
            p.CursorPosX = c2 : p.WriteText("(%)")
            If T(fld.R_Bl_Fer) = Nothing Then
                fer_pre = cMyRoutines.Calc_Fer(Val(T(fld.R_bl_Fev1)), Val(T(fld.R_bl_Fvc)), Val(T(fld.R_bl_Vc)))
            Else
                fer_pre = T(fld.R_Bl_Fer)
            End If
            p.CursorPosX = c4 : p.WriteText(fer_pre)
            If dPreds.ContainsKey("FER|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FER|LLN", dPreds, 0), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(fer_pre) / Val(dPreds("FER|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + 0.5 * ctp
            p.CursorPosX = c1 : p.WriteText("DOSE RESPONSE" & vbCrLf, fHel10b)
            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("PD" & T(fld.Protocol_threshold))
            p.CursorPosX = c2 : p.WriteText("(" & T(fld.Protocol_doseunits) & ")")
            p.CursorPosX = c3 : p.WriteText("", fHel10i)

            'Draw graph           
            PD = cChall.Calculate_PDx(c)
            p.CursorPosX = c4 : p.WriteText(PD, fHel10)

            Y = p.CursorPosY / ctp + 1
            Dim bmp As Bitmap = cChall.Draw_ProvocationGraph(c, 500, 900, New Font("Helvetica", 16))
            Dim rect As New Rectangle(c1 / ctp, Y, 10.0, 6.0)
            p.DrawRectangle(rect, False, True)
            p.DrawImage(bmp, rect, True, True)

            Me.Draw_FlowVolImage(T("provid"), eTables.Prov_test, c1 / ctp + 12, Y + 0.4, 6, 6)

            'Reference Ypos for report section to come
            Return Y + 6.5

            p = Nothing
            demo = Nothing


        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_prov" & vbCrLf & Err.Description)
            Return 20
        End Try

    End Function


    Private Function Draw_Pdf_F2_rft_results_prov(ByVal Prov As class_challenge.ProtocolData, ByVal V As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String), ByVal yPos As Single) As Single
        'Prov is data structure - contains key test data
        'V is dic object - contains all test visit data
        'D contains the demographic related data
        'T contains the test related data - loaded below and transferred to structure 'Doses()'

        Try

            Dim p As PDFPage = pdf.GetPage(1)
            Dim fld As New class_fields_ProvAndSession
            Dim incResultsLines As Single = 0.075
            Dim incResultsHeadings As Single = 0.4
            Dim Y As Single = m.yMin + 5
            Dim Ystart As Single = 0
            Dim ppn As String = ""
            Dim fer_pre As String = ""
            Dim PD As String = ""
            Dim td As TestsDone = cMyRoutines.Get_TestsDone(V)
            Dim c1 As Single = m.xMin * ctp             'Test header
            Dim c2 As Single = c1 + 1.5 * ctp                   'Units
            Dim c3 As Single = c2 + 3 * ctp                     'Normal range
            Dim c4 As Single = c3 + 3 * ctp                     'Baseline
            Dim c5 As Single = c4 + 1.6 * ctp                   'PPN
            Dim c6 As Single = c5 + 2.2 * ctp                   'Post mannitol
            Dim c7 As Single = c6 + 3.5 * ctp                   'Change


            'Get normals
            Dim demo As New Pred_demo
            demo.Age = cMyRoutines.Calc_Age(CDate(D("dob")), CDate(V("testdate")))
            demo.Htcm = Val(V("height"))
            demo.Wtkg = Val(V("weight"))
            demo.GenderID = cMyRoutines.Lookup_list_ByDescription(D("gender"), eTables.Pred_ref_genders)
            demo.Gender = D("gender")
            demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(D("ethnicity"), eTables.Pred_ref_ethnicities)
            demo.TestDate = V(fld.TestDate)
            demo.SourcesString = V(fld.Pred_SourceIDs)

            Dim dPreds As Dictionary(Of String, String) = cPred.Get_PredValues(demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)    'ParameterID|StatTypeID, result

            'Results header
            Ystart = yPos + 0.2 * ctp
            p.Font = fHel10

            p.CursorPosY = Ystart : p.CursorPosX = c1 : p.WriteText(V("testtype").ToUpper & vbCrLf & vbCrLf, fHel10b)

            p.CursorPosX = c3 : p.WriteText("Normal range", fHel10b)
            p.CursorPosX = c4 : p.WriteText("Baseline", fHel10b)
            p.CursorPosX = c5 : p.WriteText("(%mpv)" & vbCrLf, fHel10b)

            'Spiro
            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("FEV1")
            p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
            p.CursorPosX = c4 : p.WriteText(cMyRoutines.fmt(V(fld.R_bl_Fev1), 2))
            If dPreds.ContainsKey("FEV1|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FEV1|LLN", dPreds, 1), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(V(fld.R_bl_Fev1)) / Val(dPreds("FEV1|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("FVC")
            p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
            p.CursorPosX = c4 : p.WriteText(cMyRoutines.fmt(V(fld.R_bl_Fvc), 2))
            If dPreds.ContainsKey("FVC|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FVC|LLN", dPreds, 1), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(V(fld.R_bl_Fvc)) / Val(dPreds("FVC|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("VC")
            p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
            p.CursorPosX = c4 : p.WriteText(cMyRoutines.fmt(V(fld.R_bl_Vc), 2))
            If dPreds.ContainsKey("VC|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("VC|LLN", dPreds, 1), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(V(fld.R_bl_Vc)) / Val(dPreds("VC|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + incResultsLines * ctp
            p.CursorPosX = c1 : p.WriteText("FER")
            p.CursorPosX = c2 : p.WriteText("(%)")
            If V(fld.R_Bl_Fer) = Nothing Then
                fer_pre = cMyRoutines.Calc_Fer(Val(V(fld.R_bl_Fev1)), Val(V(fld.R_bl_Fvc)), Val(V(fld.R_bl_Vc)))
            Else
                fer_pre = V(fld.R_Bl_Fer)
            End If
            p.CursorPosX = c4 : p.WriteText(fer_pre)
            If dPreds.ContainsKey("FER|MPV") Then
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FER|LLN", dPreds, 0), fHel10i)
                ppn = cPred.Format_Pred(100 * Val(fer_pre) / Val(dPreds("FER|MPV")), 0)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            Else
                p.CursorPosX = c3 : p.WriteText("---", fHel10)
                p.WriteText(vbCrLf)
            End If

            p.CursorPosY = p.CursorPosY + 0.3 * ctp
            p.CursorPosX = c1 : p.WriteText("PD" & V(fld.Protocol_threshold))
            p.CursorPosX = c2 : p.WriteText("(" & V(fld.Protocol_doseunits) & ")")
            p.CursorPosX = c3 : p.WriteText("", fHel10i)

            'Draw graph           
            PD = cChall.Calculate_PDx(Prov)
            p.CursorPosX = c4 : p.WriteText(PD, fHel10)

            Y = p.CursorPosY / ctp + 1
            Dim bmp As Bitmap = cChall.Draw_ProvocationGraph(Prov, 500, 900, New Font("Helvetica", 13))
            Dim rect As New Rectangle(c1 / ctp, Y, 10.0, 6.0)
            p.DrawRectangle(rect, False, True)
            p.DrawImage(bmp, rect, True, True)

            Me.Draw_FlowVolImage(V("provid"), eTables.Prov_test, c1 / ctp + 12, Y + 0.4, 6, 6)

            'Reference Ypos for report section to come
            Return Y + 6.5

            p = Nothing
            demo = Nothing


        Catch ex As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F2_Reportform_ResultsSection_GenericProv" & vbCrLf & Err.Description)
            Return Nothing
        End Try

    End Function

    Private Function Draw_Pdf_F2_Reportform_ResultsSection_WalkTest(w As class_walktest_plot.walk_testdata, ByVal D As Dictionary(Of String, String), ByVal yPos As Single) As Single
        'D contains the demographic related data

        Try

            Dim p As PDFPage = pdf.GetPage(1)
            Dim incResultsLines As Single = 0.075
            Dim incResultsHeadings As Single = 0.4
            Dim Y As Single = m.yMin + 5
            Dim Ystart As Single = 0
            Dim ppn As String = ""
            Dim x1 As Single = 0, x2 As Single = 0, y1 As Single = 0, y2 As Single = 0
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = c1 + 1.5 * ctp
            Dim c3 As Single = c2 + 1.5 * ctp
            Dim c4 As Single = c3 + 1.5 * ctp
            Dim c5 As Single = c4 + 1.5 * ctp
            Dim c6 As Single = c5 + 1.5 * ctp
            Dim c7 As Single = c6 + 1.5 * ctp
            Dim c8 As Single = c7 + 1.5 * ctp
            Dim r1 As Single = 0
            Dim r2 As Single = 0
            Dim r3 As Single = 0
            Dim r4 As Single = 0
            Dim r5 As Single = 0
            Dim r6 As Single = 0

            'Results header
            Ystart = yPos + 0.2 * ctp
            p.Font = fHel10
            p.CursorPosY = Ystart : p.CursorPosX = c1 : p.WriteText(w.TestType & vbCrLf & vbCrLf, fHel10b)

            'Table
            r1 = p.CursorPosY + 0.5 * ctp
            r2 = r1 + 0.3 * ctp
            r3 = r2 + 0.5 * ctp
            r4 = r3 + 0.5 * ctp
            r5 = r4 + 0.5 * ctp
            r6 = r5 + 0.5 * ctp
            p.CursorPosY = r1
            p.CursorPosX = c4 : p.WriteText("Suppl O2", fHel8)
            p.CursorPosX = c5 : p.WriteText("SpO2", fHel8)
            p.CursorPosX = c6 : p.WriteText("HR", fHel8)
            p.CursorPosX = c7 : p.WriteText("Dyspnoea", fHel8)
            p.CursorPosX = c8 : p.WriteText("Distance", fHel8)
            p.CursorPosY = r2
            p.CursorPosX = c4 : p.WriteText("(L/min)", fHel8)
            p.CursorPosX = c5 : p.WriteText("(%)", fHel8)
            p.CursorPosX = c6 : p.WriteText("(bpm)", fHel8)
            p.CursorPosX = c7 : p.WriteText("(Borg)", fHel8)
            p.CursorPosX = c8 : p.WriteText("(m)", fHel8)
            p.CursorPosY = r3
            p.CursorPosX = c2 : p.WriteText("Trial 1", fHel8)
            p.CursorPosX = c3 : p.WriteText("Rest", fHel8)
            p.CursorPosY = r4
            p.CursorPosX = c3 : p.WriteText("Walk", fHel8)
            p.CursorPosY = r5
            p.CursorPosX = c2 : p.WriteText("Trial 2", fHel8)
            p.CursorPosX = c3 : p.WriteText("Rest", fHel8)
            p.CursorPosY = r6
            p.CursorPosX = c3 : p.WriteText("Walk", fHel8)

            Dim s() As class_walktest_plot.walk_SummaryResults = cWalk.Calculate_SummaryResults(w.trials, w.ProtocolID)
            'Trial 1
            p.CursorPosY = r3
            p.CursorPosX = c4 : p.WriteText(s(0).FiO2_rest, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(0).SpO2_rest, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(0).HR_rest, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(0).Dyspnoea_rest, fHel8)
            p.CursorPosX = c8 : p.WriteText("---", fHel8)
            p.CursorPosY = r4
            p.CursorPosX = c4 : p.WriteText(s(0).FiO2_exercise, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(0).SpO2_exercise, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(0).HR_exercise, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(0).Dyspnoea_exercise, fHel8)
            p.CursorPosX = c8 : p.WriteText(s(0).Distance_exercise, fHel8)
            'Trial 2
            p.CursorPosY = r5
            p.CursorPosX = c4 : p.WriteText(s(1).FiO2_rest, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(1).SpO2_rest, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(1).HR_rest, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(1).Dyspnoea_rest, fHel8)
            p.CursorPosX = c8 : p.WriteText("---", fHel8)
            p.CursorPosY = r6
            p.CursorPosX = c4 : p.WriteText(s(1).FiO2_exercise, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(1).SpO2_exercise, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(1).HR_exercise, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(1).Dyspnoea_exercise, fHel8)
            p.CursorPosX = c8 : p.WriteText(s(1).Distance_exercise, fHel8)

            Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Black, 0.1)
            x1 = c2 / ctp - 0.2 : x2 = x1 + 11 : y1 = r1 / ctp - 0.1
            p.DrawLine(myPen, x1, y1, x2, y1)
            y1 = r3 / ctp - 0.1
            p.DrawLine(myPen, x1, y1, x2, y1)
            y1 = r5 / ctp - 0.1
            p.DrawLine(myPen, x1, y1, x2, y1)
            y1 = r6 / ctp + 0.4
            p.DrawLine(myPen, x1, y1, x2, y1)

            'Draw graph                    
            'Dim pr = cWalkPlot.Get_plotproperties_walk
            'Dim a() As List(Of Single) = cWalkPlot.Convert_XYdatatoArray(w)

            'ReDim pr.xData(0 To 1)
            'ReDim pr.yData(0 To 1)
            'pr.xData(0) = a(0)
            'pr.yData(0) = a(1)
            'pr.xData(1) = a(2)
            'pr.yData(1) = a(3)

            'Dim ch As Chart = cWalkPlot.Create_plot(pr)
            'ch.SaveImage("c:\temp\pdr_chart.jpeg", ImageFormat.Jpeg)

            'Y = p.CursorPosY / ctp + 1
            'Dim rect As New Rectangle(x1 + 1, Y, 10.0, 6.0)
            'p.DrawRectangle(rect, False, True)
            'p.DrawImage("C:\temp\pdr_chart.jpeg", rect, True, True)

            Return m.yMax - 10      'Reference Ypos for report section to come

        Catch ex As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F2_Reportform_ResultsSection_WalkTest" & vbCrLf & Err.Description)
            Return Nothing
        End Try

    End Function

    Private Function Draw_Pdf_F1_rft_header(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String)) As Single

        'Note - PDFOne.net has a bug. Cursor position get and set only works in points (ie 1 inch=72 points)

        Try
            Dim p As PDFPage = pdf.GetPage(1)
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = 0
            Dim c3 As Single = 0
            Dim c4 As Single = 0
            Dim c5 As Single = 0
            Dim c6 As Single = 0
            Dim Y As Single = m.yMin * ctp
            Dim Y1 As Single = 0
            Dim yLittleBit As Single = 0.05
            Dim demo As New class_DemographicFields
            Dim rft As New class_Rft_RoutineAndSessionFields
            Dim pen1 As Pen = New Pen(Color.Black, 0.5)

            p.MeasurementUnit = PDFMeasurementUnit.Centimeters

            Dim s As New class_reportstrings
            Dim ReportTitle As String = ""
            Dim ReportTitleFont As PDFFont = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 10)
            Select Case Left(R("testtype"), 4).ToLower
                Case "rfts" : ReportTitle = s.rft_reporttitle_routine.text
                Case "skin" : ReportTitle = s.rft_reporttitle_spt.text
                Case "alti" : ReportTitle = s.rft_reporttitle_hast.text
                Case "cpet" : ReportTitle = s.rft_reporttitle_cpet.text
                Case "walk" : ReportTitle = s.rft_reporttitle_walk.text
                Case "mann", "hist", "meth" : ReportTitle = s.rft_reporttitle_bhr.text
            End Select

            'Service details
            p.CursorPosX = m.xMin * ctp : p.CursorPosY = Y
            p.WriteText(ReportTitle & vbCrLf, ReportTitleFont)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)
            Y1 = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y1
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_1.text & vbCrLf, s.rft_serviceline_1.font) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_2.text & vbCrLf, s.rft_serviceline_2.font) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_3.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_4.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_5.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_6.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp

            p.Font = fHel10

            'Patient details. If very long name then print surname on another line
            c2 = (m.xMin + 10) * ctp
            c3 = c2 + (1.8 * ctp)
            p.CursorPosY = Y1
            p.CursorPosX = c2 : p.WriteText("Name:", fHel10b)
            Select Case Len(D(demo.Firstname) & " " & D(demo.Surname))
                Case Is < 29
                    p.CursorPosX = c3 : p.WriteText(D(demo.Firstname) & " " & D(demo.Surname) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
                Case Else
                    p.CursorPosX = c3 : p.WriteText(D(demo.Firstname) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
                    p.CursorPosX = c3 : p.WriteText(D(demo.Surname) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            End Select

            Dim UR As String = cPt.get_ur_for_hsid(D(demo.PatientID), cHs.get_healthservice_HSID_fromName(R(rft.Req_healthservice)), class_pas.eURformats.URandHS)
            p.CursorPosX = c2 : p.WriteText(gURlabel & ":", fHel10b)
            p.CursorPosX = c3 : p.WriteText(UR & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("Gender:", fHel10b)
            p.CursorPosX = c3 : p.WriteText(D(demo.Gender) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("DOB:", fHel10b)
            p.CursorPosX = c3 : If IsDate(D(demo.DOB)) Then p.WriteText(Format(CDate(D(demo.DOB)), "dd/MM/yyyy") & "  (" & cMyRoutines.Calc_Age(D(demo.DOB), R("testdate")) & " yrs)" & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            'p.CursorPosX = c3 : If IsDate(D(demo.DOB)) Then p.WriteText(Format(CDate(D(demo.DOB)), "dd/MM/yyyy") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("Address:", fHel10b)
            p.CursorPosX = c3 : p.WriteText(D(demo.Address_1) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c3 : p.WriteText(D(demo.Suburb) & " " & D(demo.PostCode) & vbCrLf, fHel10)

            Y = 0.2 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Ref details
            Y = 0.25 * ctp + p.CursorPosY
            p.CursorPosY = Y
            c4 = c1 + (0.7 * ctp)
            c5 = c1 + (12 * ctp)
            c6 = c5 + (2 * ctp)
            p.CursorPosX = c1 : p.WriteText("To:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c1 : p.WriteText("Cc:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosY = Y
            p.CursorPosX = c2 : p.WriteText("Test date:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("Time:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosY = Y
            p.CursorPosX = c4 : p.WriteText(R(rft.Req_name) & ", " & R("req_address") & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c4 : p.WriteText(R(rft.Report_copyto) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosY = Y
            p.CursorPosX = c3 : If IsDate(R(rft.TestDate)) Then p.WriteText(Format(CDate(R(rft.TestDate)), "dd/MM/yyyy") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c3 : If IsDate(R(rft.TestTime)) Then p.WriteText(Format(CDate(R(rft.TestTime)), "HH:mm") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Height etc details       
            c4 = c1 + (2.3 * ctp)
            c5 = c2 + (2.3 * ctp)
            Y = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y
            p.CursorPosX = c2 : p.WriteText("Height (cm):" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("Weight (kg)" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("BMI (kg/m2):" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosY = Y
            p.CursorPosX = c5 : p.WriteText(cMyRoutines.fmt(R(rft.Height), 1) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c5 : p.WriteText(cMyRoutines.fmt(R(rft.Weight), 1) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c5 : p.WriteText(cMyRoutines.calc_BMI(R(rft.Height), R(rft.Weight)) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosY = Y
            p.CursorPosX = c1 : p.WriteText("Smoking hx:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c1 : p.WriteText("Pack years:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c1 : p.WriteText("Last BD:" & vbCrLf, fHel10b) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosY = Y
            p.CursorPosX = c4 : p.WriteText(R(rft.Smoke_hx) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c4 : p.WriteText(R(rft.Smoke_packyears) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c4 : p.WriteText(R(rft.BDStatus) & vbCrLf, fHel10)
            p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Clin notes 
            c4 = c1 + (2.5 * ctp)
            Y = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y : p.CursorPosX = c1 : p.WriteText("Clinical notes:  ", fHel10b)
            p.CursorPosY = Y : p.CursorPosX = c4 : p.WriteText(Left(R(rft.Req_clinicalnotes), 90) & vbCrLf)

            Y = 0.2 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            Return Y * ctp

        Catch ex As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_header" & vbCrLf & Err.Description)
            Return Nothing
        End Try

    End Function

    Private Sub Draw_Pdf_F1_rft_trend(ByVal D As Dictionary(Of String, String), g As DataGridView, ch As Chart)

        'Note - PDFOne.net has a bug. Cursor position get and set only works in points (ie 1 inch=72 points)

        Try
            Dim p As PDFPage = pdf.GetPage(1)
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = 0
            Dim c3 As Single = 0
            Dim c4 As Single = 0
            Dim c5 As Single = 0
            Dim c6 As Single = 0
            Dim Y As Single = m.yMin * ctp
            Dim Y1 As Single = 0
            Dim yLittleBit As Single = 0.05
            Dim demo As New class_DemographicFields
            Dim rft As New class_Rft_RoutineAndSessionFields
            Dim pen1 As Pen = New Pen(Color.Black, 0.5)
            Dim i As Integer = 0

            p.MeasurementUnit = PDFMeasurementUnit.Centimeters

            Dim s As New class_reportstrings

            'Service details
            p.CursorPosX = m.xMin * ctp : p.CursorPosY = Y
            p.WriteText(s.rft_reporttitle_routine.text & vbCrLf, s.rft_reporttitle_routine.font)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)
            Y1 = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y1
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_1.text & vbCrLf, s.rft_serviceline_1.font) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_2.text & vbCrLf, s.rft_serviceline_2.font) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_3.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_4.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_5.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_6.text & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp

            p.Font = fHel10

            'Patient details. If very long name then print surname on another line
            c2 = (m.xMin + 10) * ctp
            c3 = c2 + (1.8 * ctp)
            p.CursorPosY = Y1
            p.CursorPosX = c2 : p.WriteText("Name:", fHel10b)
            Select Case Len(D(demo.Firstname) & " " & D(demo.Surname))
                Case Is < 29
                    p.CursorPosX = c3 : p.WriteText(D(demo.Firstname) & " " & D(demo.Surname) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
                Case Else
                    p.CursorPosX = c3 : p.WriteText(D(demo.Firstname) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
                    p.CursorPosX = c3 : p.WriteText(D(demo.Surname) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            End Select

            p.CursorPosX = c2 : p.WriteText("UR:", fHel10b)
            p.CursorPosX = c3 : p.WriteText(D(demo.UR) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("Gender:", fHel10b)
            p.CursorPosX = c3 : p.WriteText(D(demo.Gender) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("DOB:", fHel10b)
            p.CursorPosX = c3 : If IsDate(D(demo.DOB)) Then p.WriteText(Format(CDate(D(demo.DOB)), "dd/MM/yyyy") & vbCrLf, fHel10)
            p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c2 : p.WriteText("Address:", fHel10b)
            p.CursorPosX = c3 : p.WriteText(D(demo.Address_1) & vbCrLf, fHel10) : p.CursorPosY = p.CursorPosY + yLittleBit * ctp
            p.CursorPosX = c3 : p.WriteText(D(demo.Suburb) & " " & D(demo.PostCode) & vbCrLf, fHel10)

            Y = 0.2 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Report title
            p.CursorPosY = (0.4 + Y) * ctp
            p.CursorPosX = c1 : p.WriteText("Lung Function Trend Report:", fHel10b)

            'Do table
            Dim yStart As Single = 2 + Y
            c2 = c1 + (2.3 * ctp)
            'Dim rect As New Rectangle(m.xMin, yStart, 17.5, 10.0)
            Dim param As String = vbEmpty, units As String = vbEmpty
            p.CursorPosY = yStart * ctp
            For Each r As DataGridViewRow In g.Rows
                If r.Index < g.Rows.Count - 2 Then
                    'Do row headers
                    If InStr(r.HeaderCell.Value, "(") > 0 And InStr(r.HeaderCell.Value, ")") > 0 Then
                        units = Mid(r.HeaderCell.Value, InStr(r.HeaderCell.Value, "("), InStr(r.HeaderCell.Value, ")") - InStr(r.HeaderCell.Value, "(") + 1)
                        param = Left(r.HeaderCell.Value, InStr(r.HeaderCell.Value, "(") - 1)
                    Else
                        units = ""
                        param = r.HeaderCell.Value
                    End If
                    p.CursorPosX = c1 : p.WriteText(param, fHel8)
                    p.CursorPosX = c2 : p.WriteText(units & vbCrLf, fHel8)
                    p.CursorPosY = p.CursorPosY + 0.1 * ctp
                End If
            Next

            'Do test columns
            Dim xStart As Single = c2 + 2.4 * ctp
            Dim xC As Single = 0
            Dim colWidth As Single = 1.3 * ctp
            Dim col As Integer = 0
            Dim colMax As Integer = 0
            Dim TrimCols As Boolean = g.Columns.Count > 10

            If TrimCols Then colMax = 9 Else colMax = g.Columns.Count - 1

            Select Case TrimCols
                Case True
                    'First 8
                    For col = 0 To 7
                        xC = xStart + col * colWidth
                        p.CursorPosY = yStart * ctp
                        For i = 0 To g.Rows.Count - 3
                            p.CursorPosX = xC
                            If i = 2 Then p.WriteText(Left(g(col, i).Value, 5) & vbCrLf, fHel8) Else p.WriteText(Left(g(col, i).Value, 8) & vbCrLf, fHel8)
                            p.CursorPosY = p.CursorPosY + 0.1 * ctp
                        Next
                    Next
                    'Ninth is a gap indicator
                    xC = xStart + 8 * colWidth
                    p.CursorPosX = xC
                    p.CursorPosY = yStart * ctp
                    p.WriteText("< x" & g.Columns.Count - 9 & " >", fHel8)

                    'Display 10th
                    xC = xStart + 9 * colWidth
                    p.CursorPosY = yStart * ctp
                    For i = 0 To g.Rows.Count - 3
                        p.CursorPosX = xC
                        If i = 2 Then p.WriteText(Left(g(g.Columns.Count - 1, i).Value, 5) & vbCrLf, fHel8) Else p.WriteText(Left(g(g.Columns.Count - 1, i).Value, 8) & vbCrLf, fHel8)
                        p.CursorPosY = p.CursorPosY + 0.1 * ctp
                    Next
                Case False
                    For col = 0 To g.Columns.Count - 1
                        xC = xStart + col * colWidth
                        p.CursorPosY = yStart * ctp
                        For i = 0 To g.Rows.Count - 3
                            p.CursorPosX = xC
                            If i = 2 Then p.WriteText(Left(g(col, i).Value, 5) & vbCrLf, fHel8) Else p.WriteText(Left(g(col, i).Value, 8) & vbCrLf, fHel8)
                            p.CursorPosY = p.CursorPosY + 0.1 * ctp
                        Next
                    Next
            End Select

            'Do test separator lines
            Y = yStart + 3 * 0.378
            p.DrawLine(pen1, m.xMin, Y, 18.5, Y)
            Y = yStart + 5 * 0.378
            p.DrawLine(pen1, m.xMin, Y, 18.5, Y)
            Y = yStart + 13 * 0.378
            p.DrawLine(pen1, m.xMin, Y, 18.5, Y)
            Y = yStart + 18 * 0.38
            p.DrawLine(pen1, m.xMin, Y, 18.5, Y)
            Y = yStart + 22 * 0.38
            p.DrawLine(pen1, m.xMin, Y, 18.5, Y)
            Y = yStart + 24 * 0.38
            p.DrawLine(pen1, m.xMin, Y, 18.5, Y)

            'Draw chart
            Dim tempfile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\trend.jpg"
            If FileSystem.FileExists(tempfile) Then FileSystem.DeleteFile(tempfile)
            Y = Y + 3
            Dim rect As New Rectangle(c1 / ctp, Y, 17, 7.0)
            p.DrawRectangle(rect, False, True)
            ch.SaveImage(tempfile, ImageFormat.Jpeg)
            p.DrawImage(tempfile, rect, True, True)


            'Print timedate
            Dim yoffset As Single = 1.2
            p.DrawLine(pen1, m.xMin, m.yMax - yoffset, m.xMax, m.yMax - yoffset)
            p.Font = fHel7
            p.CursorPosX = c1 : p.CursorPosY = (m.yMax - yoffset + 0.1) * ctp
            p.WriteText(" PRINTED: " & Format(Now, "dd/MM/yyyy  HH:mm"))


        Catch ex As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_trend" & vbCrLf & Err.Description)
        End Try

    End Sub

    Private Function Draw_Pdf_F3_Reportform_HeaderSection(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String)) As Single

        'Note - PDFOne.net has a bug. Cursor position get and set only works in points (ie 1 inch=72 points)

        Try
            Dim p As PDFPage = pdf.GetPage(1)
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = 0
            Dim c3 As Single = 0
            Dim c4 As Single = 0
            Dim c5 As Single = 0
            Dim c6 As Single = 0
            Dim c7 As Single = 0
            Dim c8 As Single = 0
            Dim Y As Single = m.yMin * ctp
            Dim Y1 As Single = 0

            Dim demo As New class_DemographicFields
            Dim rft As New class_Rft_RoutineAndSessionFields
            Dim pen1 As Pen = New Pen(Color.Black, 0.5)

            p.MeasurementUnit = PDFMeasurementUnit.Centimeters

            Dim s As New class_reportstrings

            'Draw BSV logo
            c2 = (m.xMin) * ctp
            Dim rect As New RectangleF(m.xMin - 0.3, 1, m.xMax, 3)
            p.DrawImage(My.Resources.bsv_reslab_header, rect, False, False)
            p.CursorPosX = c2 : p.CursorPosY = 4 * ctp : p.WriteText("LUNG FUNCTION REPORT", fHel14)
            p.DrawLine(pen1, m.xMin, 4.6, m.xMax, 4.6)
            Y1 = 4.8 * ctp
            p.CursorPosY = Y1
            p.Font = fHel10

            'Patient details
            c3 = c2 + (2 * ctp)
            p.CursorPosY = Y1
            p.CursorPosX = c2 : p.WriteText("Name:" & vbCrLf, fHel10b)
            'p.CursorPosX = c2 : p.WriteText("" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("Gender:" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("DOB:" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("Address:" & vbCrLf, fHel10b)
            p.CursorPosY = Y1
            p.CursorPosX = c3 : p.WriteText(D(demo.Firstname) & " " & D(demo.Surname) & vbCrLf, fHel10)
            'p.CursorPosX = c3 : p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosX = c3 : p.WriteText(D(demo.Gender) & vbCrLf, fHel10)
            p.CursorPosX = c3 : If IsDate(D(demo.DOB)) Then p.WriteText(Format(CDate(D(demo.DOB)), "dd/MM/yyyy") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosX = c3 : p.WriteText(D(demo.Address_1) & ", " & D(demo.Suburb) & " " & D(demo.PostCode) & vbCrLf, fHel10)
            Y = 0.25 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Ref details
            Y = 0.3 * ctp + p.CursorPosY
            p.CursorPosY = Y
            c4 = c1 + (1 * ctp)
            c5 = c1 + (13 * ctp)
            c6 = c5 + (2.5 * ctp)
            p.CursorPosX = c1 : p.WriteText("To:" & vbCrLf, fHel10b)
            p.CursorPosX = c1 : p.WriteText("Cc:" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c5 : p.WriteText("Test date:" & vbCrLf, fHel10b)
            p.CursorPosX = c5 : p.WriteText("Time:" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c4 : p.WriteText(R(rft.Req_name) & ", " & R("req_address") & vbCrLf, fHel10)
            p.CursorPosX = c4 : p.WriteText(R(rft.Report_copyto) & vbCrLf, fHel10)
            p.CursorPosY = Y
            p.CursorPosX = c6 : If IsDate(R(rft.TestDate)) Then p.WriteText(Format(CDate(R(rft.TestDate)), "dd/MM/yyyy") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosX = c6 : If IsDate(R(rft.TestTime)) Then p.WriteText(Format(CDate(R(rft.TestTime)), "HH:mm") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Height etc details       
            c7 = c1 + (2.5 * ctp)
            c8 = c2 + (3 * ctp)
            Y = 0.25 * ctp + p.CursorPosY
            p.CursorPosY = Y
            p.CursorPosX = c5 : p.WriteText("Height (cm):" & vbCrLf, fHel10b)
            p.CursorPosX = c5 : p.WriteText("Weight (kg):" & vbCrLf, fHel10b)
            p.CursorPosX = c5 : p.WriteText("BMI (kg/m2):" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c6 : p.WriteText(cMyRoutines.fmt(R(rft.Height), 1) & vbCrLf, fHel10)
            p.CursorPosX = c6 : p.WriteText(cMyRoutines.fmt(R(rft.Weight), 1) & vbCrLf, fHel10)
            p.CursorPosX = c6 : p.WriteText(cMyRoutines.calc_BMI(R(rft.Height), R(rft.Weight)) & vbCrLf, fHel10)

            p.CursorPosY = Y
            p.CursorPosX = c1 : p.WriteText("Smoking hx:" & vbCrLf, fHel10b)
            p.CursorPosX = c1 : p.WriteText("Pack years:" & vbCrLf, fHel10b)
            p.CursorPosX = c1 : p.WriteText("Last BD:" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c7 : p.WriteText(R(rft.Smoke_hx) & vbCrLf, fHel10)
            p.CursorPosX = c7 : p.WriteText(R(rft.Smoke_packyears) & vbCrLf, fHel10)
            p.CursorPosX = c7 : p.WriteText(R(rft.BDStatus) & vbCrLf, fHel10)
            Y = 0.15 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Clin notes 
            c4 = c1 + (2.5 * ctp)
            Y = 0.25 * ctp + p.CursorPosY
            p.CursorPosY = Y : p.CursorPosX = c1 : p.WriteText("Clinical notes:  ", fHel10b)
            p.CursorPosY = Y : p.CursorPosX = c4 : p.WriteText(Left(R(rft.Req_clinicalnotes), 90) & vbCrLf)

            Y = 0.25 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            Return Y * ctp

        Catch ex As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F2_Reportform_HeaderSection" & vbCrLf & Err.Description)
            Return Nothing
        End Try

    End Function

    Private Function Draw_Pdf_F4_Reportform_HeaderSection(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String)) As Single

        'Note - PDFOne.net has a bug. Cursor position get and set only works in points (ie 1 inch=72 points)

        Try
            Dim p As PDFPage = pdf.GetPage(1)
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = 0
            Dim c3 As Single = 0
            Dim c4 As Single = 0
            Dim c5 As Single = 0
            Dim c6 As Single = 0
            Dim Y As Single = m.yMin * ctp
            Dim Y1 As Single = 0

            Dim demo As New class_DemographicFields
            Dim rft As New class_Rft_RoutineAndSessionFields
            Dim pen1 As Pen = New Pen(Color.Black, 0.5)

            p.MeasurementUnit = PDFMeasurementUnit.Centimeters

            Dim s As New class_reportstrings

            'Service details
            p.CursorPosX = m.xMin * ctp : p.CursorPosY = Y
            p.WriteText(s.rft_reporttitle_routine.text & vbCrLf, s.rft_reporttitle_routine.font)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)
            Y1 = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y1
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_1.text & vbCrLf, s.rft_serviceline_1.font)
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_2.text & vbCrLf, s.rft_serviceline_2.font)
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_3.text & vbCrLf, s.rft_serviceline_3.font)
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_4.text & vbCrLf, s.rft_serviceline_4.font)
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_5.text & vbCrLf, s.rft_serviceline_5.font)
            p.CursorPosX = m.xMin * ctp : p.WriteText(s.rft_serviceline_6.text & vbCrLf, s.rft_serviceline_6.font)

            p.Font = fHel10

            'Patient details
            c2 = (m.xMin + 10) * ctp
            c3 = c2 + (2 * ctp)
            p.CursorPosY = Y1
            p.CursorPosX = c2 : p.WriteText("Name:" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("Gender:" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("DOB:" & vbCrLf, fHel10b)
            p.CursorPosX = c2 : p.WriteText("Address:" & vbCrLf, fHel10b)
            p.CursorPosY = Y1
            p.CursorPosX = c3 : p.WriteText(D(demo.Firstname) & " " & D(demo.Surname) & vbCrLf, fHel10b)
            p.CursorPosX = c3 : p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosX = c3 : p.WriteText(D(demo.Gender) & vbCrLf, fHel10)
            p.CursorPosX = c3 : If IsDate(D(demo.DOB)) Then p.WriteText(Format(CDate(D(demo.DOB)), "dd/MM/yyyy") & vbCrLf, fHel10b) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosX = c3 : p.WriteText(D(demo.Address_1) & vbCrLf, fHel10)
            p.CursorPosX = c3 : p.WriteText(D(demo.Suburb) & " " & D(demo.PostCode) & vbCrLf, fHel10)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Ref details
            Y = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y
            c4 = c1 + (1 * ctp)
            c5 = c1 + (12 * ctp)
            c6 = c5 + (2.5 * ctp)
            p.CursorPosX = c1 : p.WriteText("To:" & vbCrLf, fHel10b)
            p.CursorPosX = c1 : p.WriteText("Cc:" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c5 : p.WriteText("Test date:" & vbCrLf, fHel10b)
            p.CursorPosX = c5 : p.WriteText("Time:" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c4 : p.WriteText(R(rft.Req_name) & ", " & R("req_address") & vbCrLf, fHel10)
            p.CursorPosX = c4 : p.WriteText(R(rft.Report_copyto) & vbCrLf, fHel10)
            p.CursorPosY = Y
            p.CursorPosX = c6 : If IsDate(R(rft.TestDate)) Then p.WriteText(Format(CDate(R(rft.TestDate)), "dd/MM/yyyy") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            p.CursorPosX = c6 : If IsDate(R(rft.TestTime)) Then p.WriteText(Format(CDate(R(rft.TestTime)), "HH:mm") & vbCrLf, fHel10) Else p.WriteText("" & vbCrLf, fHel10)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Height etc details       
            c4 = c1 + (3 * ctp)
            'c5 = c2 + (3 * ctp)
            Y = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y
            p.CursorPosX = c5 : p.WriteText("Height (cm):" & vbCrLf, fHel10b)
            p.CursorPosX = c5 : p.WriteText("Weight (kg)" & vbCrLf, fHel10b)
            p.CursorPosX = c5 : p.WriteText("BMI (kg/m2):" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c6 : p.WriteText(cMyRoutines.fmt(R(rft.Height), 1) & vbCrLf, fHel10)
            p.CursorPosX = c6 : p.WriteText(cMyRoutines.fmt(R(rft.Weight), 1) & vbCrLf, fHel10)
            p.CursorPosX = c6 : p.WriteText(cMyRoutines.calc_BMI(R(rft.Height), R(rft.Weight)) & vbCrLf, fHel10)
            p.CursorPosY = Y
            p.CursorPosX = c1 : p.WriteText("Smoking hx:" & vbCrLf, fHel10b)
            p.CursorPosX = c1 : p.WriteText("Pack years" & vbCrLf, fHel10b)
            p.CursorPosX = c1 : p.WriteText("Last BD:" & vbCrLf, fHel10b)
            p.CursorPosY = Y
            p.CursorPosX = c4 : p.WriteText(R(rft.Smoke_hx) & vbCrLf, fHel10)
            p.CursorPosX = c4 : p.WriteText(R(rft.Smoke_packyears) & vbCrLf, fHel10)
            p.CursorPosX = c4 : p.WriteText(R(rft.BDStatus) & vbCrLf, fHel10)
            Y = 0.1 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Clin notes 
            c4 = c1 + (2.5 * ctp)
            Y = 0.2 * ctp + p.CursorPosY
            p.CursorPosY = Y : p.CursorPosX = c1 : p.WriteText("Clinical notes:  ", fHel10b)
            p.CursorPosY = Y : p.CursorPosX = c4 : p.WriteText(Left(R(rft.Req_clinicalnotes), 90) & vbCrLf)

            Y = 0.2 + p.CursorPosY / ctp
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            Return Y * ctp

        Catch ex As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F4_Reportform_HeaderSection" & vbCrLf & Err.Description)
            Return Nothing
        End Try

    End Function

    Private Function Draw_Pdf_F1_report(ByVal R As Dictionary(Of String, String), Ethnicity As String, ByVal yStartPos_cm As Single) As Single

        Dim p As PDFPage = pdf.GetPage(1)
        Dim Y As Single = 0
        Dim c1 As Single = m.xMin * ctp
        Dim c2 As Single = c1 + (3.1 * ctp)
        Dim rft As New class_Rft_RoutineAndSessionFields
        Dim pen1 As Pen = New Pen(Color.Black, 0.5)
        Dim FlowVolTop As Single = 0

        Try

            'Tech notes 
            p.DrawLine(pen1, m.xMin, yStartPos_cm, m.xMax, yStartPos_cm)
            Y = (yStartPos_cm + 0.15)
            p.CursorPosX = c1 : p.CursorPosY = Y * ctp : p.WriteText("Technical notes:  ", fHel10b)

            Dim s As String = "                                Ethnicity=" & Ethnicity & ". " & R("technicalnotes") & " (Tested by: " & R("scientist") & ")"
            Dim rec As New RectangleF(c1 / ctp, p.CursorPosY / ctp, m.xWidth, 0.8)
            p.WriteText(s, rec, PDFAlignment.Left)
            Y = Y + 0.9
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Report header
            Y = (Y + 0.2) * ctp
            p.CursorPosY = Y : p.CursorPosX = c1 : p.WriteText("REPORT:", fHel10b)

            'Report text
            p.Font = fHel10
            Y = Y / ctp + 0.5
            If IsDate(R("report_reporteddate")) Then
                s = R("report_text") & vbCrLf & vbCrLf & R("report_reportedby") & "          " & Format(CDate(R("report_reporteddate")), "dd/MM/yyyy")
            Else
                s = R("report_text") & vbCrLf & vbCrLf & R("report_reportedby")
            End If
            p.WriteText(s, New RectangleF(c1 / ctp, Y, m.xWidth, m.yHeight - Y), PDFAlignment.Left)

            'Fax privacy statement
            Dim PrintFaxStatement As Boolean = cMyRoutines.get_prefs_showfaxstatementonreports
            If PrintFaxStatement Then FlowVolTop = m.yMax - 7.2 Else FlowVolTop = m.yMax - 6.5
            Dim Msg As New StringBuilder
            Msg.Append("Attention.  This facsimile, including any attachments, is confidential and for the sole use of the intended recipient(s).")
            Msg.Append("This confidentiality is not waived or lost if you receive it and you are not the intended recipient(s).")
            Msg.Append("Any unauthorised use, alteration, disclosure or review of this facsimile is prohibited. If you are not the intended recipient(s),")
            Msg.Append("you are asked to immediately notify the sender by telephone or return fax and destroy any copies produced.")
            p.Font = fHel7
            If PrintFaxStatement Then p.WriteText(Msg.ToString, New RectangleF(c1 / ctp, m.yMax - 2, m.xWidth, 1), PDFAlignment.Left)


            'Report status and print timedate
            Dim yoffset As Single = 1.2
            If R.ContainsKey("pred_sourceids") Then
                s = cPred.Decode_SourcesStringForDisplay(R("pred_sourceids"))
                s = Trim(Replace(s, vbCrLf, ", "))
                s = Replace(s, "Lung volumes", "Lung vols")
                s = Replace(s, "CO Transfer", "TLCO")
                s = Replace(s, "Spirometry", "Spiro")
                s = Left(s, Len(s))
            Else
                s = "n/a"
            End If
            p.DrawLine(pen1, m.xMin, m.yMax - yoffset, m.xMax, m.yMax - yoffset)
            p.Font = fHel7
            p.CursorPosX = c1 : p.CursorPosY = (m.yMax - yoffset + 0.1) * ctp
            p.WriteText("NORMAL VALUES SOURCES:  " & s)
            p.CursorPosX = c1 : p.CursorPosY = (m.yMax - yoffset + 0.4) * ctp
            p.WriteText("REPORT STATUS:  " & R("report_status") & "         PRINTED: " & Format(Now, "dd/MM/yyyy  HH:mm"))

            Return FlowVolTop
        Catch
            MsgBox("Error in class_pdf.Draw_Pdf_F1_report" & vbCrLf & Err.Description)
            Return m.yMax - 6
        End Try

    End Function

    Private Function Draw_Pdf_F2_Reportform_ReportSection(ByVal R As Dictionary(Of String, String), Ethnicity As String, ByVal yStartPos_cm As Single) As Single

        Dim p As PDFPage = pdf.GetPage(1)
        Dim Y As Single = 0
        Dim c1 As Single = m.xMin * ctp
        Dim c2 As Single = c1 + (3.1 * ctp)
        Dim rft As New class_Rft_RoutineAndSessionFields
        Dim pen1 As Pen = New Pen(Color.Black, 0.5)
        Dim FlowVolTop As Single = 0

        Try

            'Tech notes 
            p.DrawLine(pen1, m.xMin, yStartPos_cm, m.xMax, yStartPos_cm)
            Y = (yStartPos_cm + 0.15)
            p.CursorPosX = c1 : p.CursorPosY = Y * ctp : p.WriteText("Technical notes:  ", fHel10b)

            Dim s As String = "                                Ethnicity=" & Ethnicity & ". " & R("technicalnotes") & " (" & R("scientist") & ")"
            Dim rec As New RectangleF(c1 / ctp, p.CursorPosY / ctp, m.xWidth, 0.8)
            p.WriteText(s, rec, PDFAlignment.Left)
            Y = Y + 0.9
            p.DrawLine(pen1, m.xMin, Y, m.xMax, Y)

            'Report header
            Y = (Y + 0.2) * ctp
            p.CursorPosY = Y : p.CursorPosX = c1 : p.WriteText("REPORT:", fHel10b)

            'Report text
            p.Font = fHel10
            Y = Y / ctp + 0.5
            If IsDate(R("report_reporteddate")) Then
                s = R("report_text") & vbCrLf & vbCrLf & R("report_reportedby") & "          " & Format(CDate(R("report_reporteddate")), "dd/MM/yyyy")
            Else
                s = R("report_text") & vbCrLf & vbCrLf & R("report_reportedby")
            End If
            p.WriteText(s, New RectangleF(c1 / ctp, Y, m.xWidth, m.yHeight - Y), PDFAlignment.Left)

            'Fax privacy statement
            Dim PrintFaxStatement As Boolean = cMyRoutines.get_prefs_showfaxstatementonreports
            If PrintFaxStatement Then FlowVolTop = m.yMax - 7.2 Else FlowVolTop = m.yMax - 6.5
            Dim Msg As New StringBuilder
            Msg.Append("Attention.  This facsimile, including any attachments, is confidential and for the sole use of the intended recipient(s).")
            Msg.Append("This confidentiality is not waived or lost if you receive it and you are not the intended recipient(s).")
            Msg.Append("Any unauthorised use, alteration, disclosure or review of this facsimile is prohibited. If you are not the intended recipient(s),")
            Msg.Append("you are asked to immediately notify the sender by telephone or return fax and destroy any copies produced.")
            p.Font = fHel7
            If PrintFaxStatement Then p.WriteText(Msg.ToString, New RectangleF(c1 / ctp, m.yMax - 2, m.xWidth, 1), PDFAlignment.Left)


            'Report status and print timedate
            Dim yoffset As Single = 1.2
            If R.ContainsKey("pred_sourceids") Then
                s = cPred.Decode_SourcesStringForDisplay(R("pred_sourceids"))
                s = Trim(Replace(s, vbCrLf, ", "))
                s = Replace(s, "Lung volumes", "Lung vols")
                s = Replace(s, "CO Transfer", "TLCO")
                s = Replace(s, "Spirometry", "Spiro")
                s = Left(s, Len(s))
            Else
                s = "n/a"
            End If
            p.DrawLine(pen1, m.xMin, m.yMax - yoffset, m.xMax, m.yMax - yoffset)
            p.Font = fHel7
            p.CursorPosX = c1 : p.CursorPosY = (m.yMax - yoffset + 0.1) * ctp
            p.WriteText("NORMAL VALUES SOURCES:  " & s)
            p.CursorPosX = c1 : p.CursorPosY = (m.yMax - yoffset + 0.4) * ctp
            p.WriteText("REPORT STATUS:  " & R("report_status") & "         PRINTED: " & Format(Now, "dd/MM/yyyy  HH:mm"))

            Return FlowVolTop
        Catch
            MsgBox("Error in class_pdf.Draw_Pdf_F2_Reportform_ReportSection" & vbCrLf & Err.Description)
            Return m.yMax - 6
        End Try

    End Function

    Private Function pcChange(ByVal pre As Single, ByVal post As Single) As String

        If pre > 0 And post > 0 Then
            Dim ch As Single = (post - pre) / pre
            Dim plus As String = ""
            If ch >= 0 Then plus = "+" Else plus = ""
            Return Format(ch, plus & "0%")
        Else
            Return ""
        End If

    End Function

    Private Function Draw_Pdf_F1a_rft_results_routine(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String)) As Single

        Dim p As PDFPage = pdf.GetPage(1)
        Dim fld As New class_Rft_RoutineAndSessionFields
        Dim InitialLineFeed As Boolean = False
        Dim Y As Single = m.yMin + 5
        Dim ppn As String = ""
        Dim fer_pre As String = ""
        Dim fer_post As String = ""
        Dim rvtlc As String = ""
        Dim td As TestsDone = cMyRoutines.Get_TestsDone(R)
        Dim c1 As Single = (m.xMin + 0.2) * ctp             'Test header
        Dim c2 As Single = c1 + 0.3 * ctp                   'Parameter label
        Dim c3 As Single = c2 + 2 * ctp                     'Units label
        Dim c4 As Single = c3 + 3 * ctp                     'Normal range
        Dim c5 As Single = c4 + 2.2 * ctp                    'Baseline
        Dim c6 As Single = c5 + 2.8 * ctp                   'Post
        Dim c7 As Single = c6 + 2.5 * ctp                 'Post % change



        Try
            'Get normals
            Dim demo As Pred_demo = Nothing
            demo.Age = cMyRoutines.Calc_Age(CDate(D("dob")), CDate(R("testdate")))
            demo.Htcm = Val(R("height"))
            demo.Wtkg = Val(R("weight"))
            demo.GenderID = cMyRoutines.Lookup_list_ByDescription(D("gender"), eTables.Pred_ref_genders)
            demo.Gender = D("gender")
            demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(D("ethnicity"), eTables.Pred_ref_ethnicities)
            demo.Ethnicity = D("ethnicity")
            demo.TestDate = R(fld.TestDate)
            demo.SourcesString = R(fld.Pred_SourceIDs)

            Dim dPreds As Dictionary(Of String, String) = cPred.Get_PredValues(demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)    'ParameterID|StatTypeID, result

            'Clinical notes area
            Dim customColor As Color = Color.FromArgb(50, Color.Gray)
            Dim brsh As SolidBrush = New SolidBrush(customColor)
            p.DrawRectangle(New RectangleF(c1 / ctp, Y, m.xWidth, 0.6), False, False)
            p.Font = fHel10b
            p.WriteText("CLINICAL NOTES/REASON FOR TEST:", New RectangleF(c1 / ctp, Y + 0.2, m.xWidth, 0.6), PDFAlignment.Left)
            p.Font = fHel10
            p.WriteText(R("req_clinicalnotes"), New RectangleF((c1 / ctp) + 0.3, Y + 0.6, m.xWidth - 0.2, 1.2), PDFAlignment.Left)

            'Results header
            Y = Y + 1.7
            p.DrawRectangle(New RectangleF(c1 / ctp, Y, m.xWidth, 0.6), False, False)
            p.Font = fHel10b
            p.WriteText("RESULTS:", New RectangleF(c1 / ctp, Y + 0.2, m.xWidth, 0.6), PDFAlignment.Left)
            Y = Y + 0.5
            p.WriteText("Units", fHel10, c3 / ctp, Y)
            p.WriteText("Normal", fHel10, c4 / ctp, Y)
            p.WriteText("Baseline", fHel10, c5 / ctp, Y)
            p.WriteText("Post", fHel10, c6 / ctp, Y)
            p.WriteText("% Change" & vbCrLf, fHel10, c7 / ctp, Y)
            Y = Y + 0.4
            p.WriteText("Range", fHel10, c4 / ctp, Y)
            p.WriteText("(%mpv)", fHel10, c5 / ctp, Y)
            p.WriteText("(%mpv)", fHel10, c6 / ctp, Y)
            p.Font = fHel10
            Y = Y + 0.2

            'Flow vol image
            Dim flowvol_yBottom As Single = m.yMin + 17
            Dim f As String = cMyRoutines.Get_TempDirectory & "\reslab_temp.tmp"
            If My.Computer.FileSystem.FileExists(f) Then My.Computer.FileSystem.DeleteFile(f)
            If cDAL.Get_imageAsFile(R("rftid"), eTables.rft_routine, "flowvolloop", f) = True Then
                Dim rect As New RectangleF(m.xMin + 10.5, m.yMin + 10.7, 6, 6)
                p.DrawRectangle(rect, False, True)
                p.DrawImage(f, rect, True, True)
            End If

            'Spiro
            If td.anyspir_done Then
                InitialLineFeed = True
                p.CursorPosX = c1 : p.CursorPosY = Y * ctp
                p.WriteText("Spirometry" & vbCrLf, fHel10b)
                p.CursorPosX = c2 : p.WriteText("FEV1")
                p.CursorPosX = c3 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("FEV1|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("FEV1|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_bl_Fev1)) / Val(dPreds("FEV1|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_bl_Fev1) & ppn)
                If dPreds.ContainsKey("FEV1|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_post_Fev1)) / Val(dPreds("FEV1|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(fld.R_post_Fev1) & ppn)
                p.CursorPosX = c7 : p.WriteText(pcChange(Val(R(fld.R_bl_Fev1)), Val(R(fld.R_post_Fev1))) & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("FVC")
                p.CursorPosX = c3 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("FVC|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("FVC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_bl_Fvc)) / Val(dPreds("FVC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_bl_Fvc) & ppn)
                If dPreds.ContainsKey("FVC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_post_Fvc)) / Val(dPreds("FVC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(fld.R_post_Fvc) & ppn)
                p.CursorPosX = c7 : p.WriteText(pcChange(Val(R(fld.R_bl_Fvc)), Val(R(fld.R_post_Fvc))) & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("VC")
                p.CursorPosX = c3 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("VC|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("VC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_bl_Vc)) / Val(dPreds("VC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_bl_Vc) & ppn)
                If dPreds.ContainsKey("VC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_post_Vc)) / Val(dPreds("VC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(fld.R_post_Vc) & ppn)
                p.CursorPosX = c7 : p.WriteText(pcChange(Val(R(fld.R_bl_Vc)), Val(R(fld.R_post_Vc))) & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("FER")
                p.CursorPosX = c3 : p.WriteText("(%)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("FER|LLN", dPreds, 0), fHel10i)
                If R(fld.R_Bl_Fer) = Nothing Then
                    fer_pre = cMyRoutines.Calc_Fer(Val(R(fld.R_bl_Fev1)), Val(R(fld.R_bl_Fvc)), Val(R(fld.R_bl_Vc)))
                Else
                    fer_pre = R(fld.R_Bl_Fer)
                End If
                If dPreds.ContainsKey("FER|MPV") Then ppn = cPred.Format_Pred(100 * Val(fer_pre) / Val(dPreds("FER|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(fer_pre & ppn)
                If R(fld.R_Post_Fer) = Nothing Then
                    fer_post = cMyRoutines.Calc_Fer(Val(R(fld.R_post_Fev1)), Val(R(fld.R_post_Fvc)), Val(R(fld.R_post_Vc)))
                Else
                    fer_post = R(fld.R_Post_Fer)
                End If
                If dPreds.ContainsKey("FER|MPV") Then ppn = cPred.Format_Pred(100 * Val(fer_post) / Val(dPreds("FER|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(fer_post & ppn)
                p.CursorPosX = c7 : p.WriteText(pcChange(Val(fer_pre), Val(fer_post)) & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("FEF25-75")
                p.CursorPosX = c3 : p.WriteText("(L/min BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("FEF2575|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("FEF2575|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_bl_Fef2575)) / Val(dPreds("FEF2575|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_bl_Fef2575) & ppn)
                If dPreds.ContainsKey("FEF2575|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_post_Fef2575)) / Val(dPreds("FEF2575|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(fld.R_post_Fef2575) & ppn)
                p.CursorPosX = c7 : p.WriteText(pcChange(Val(R(fld.R_bl_Fef2575)), Val(R(fld.R_post_Fef2575))) & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("PEF")
                p.CursorPosX = c3 : p.WriteText("(L/min BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("PEF|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("PEF|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_bl_Pef)) / Val(dPreds("PEF|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_bl_Pef) & ppn)
                If dPreds.ContainsKey("PEF|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_post_Pef)) / Val(dPreds("PEF|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(fld.R_post_Pef) & ppn)
                p.CursorPosX = c7 : p.WriteText(pcChange(Val(R(fld.R_bl_Pef)), Val(R(fld.R_post_Pef))) & vbCrLf)
            End If

            'TLCO
            If td.tlco_done Then
                If Not InitialLineFeed Then p.WriteText(vbCrLf)
                InitialLineFeed = True

                p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + 0.1 * ctp
                p.WriteText("CO Transfer Factor" & vbCrLf, fHel10b)
                p.CursorPosX = c2 : p.WriteText("TLCO")
                p.CursorPosX = c3 : p.WriteText("(ml/min/mmHg)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("TLCO|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("TLCO|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Tlco)) / Val(dPreds("TLCO|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Tlco) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("TLCO Hb")
                p.CursorPosX = c3 : p.WriteText("(ml/min/mmHg)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("TLCO|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("TLCO|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Tlco)) / Val(dPreds("TLCO|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Tlco) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("KCO")
                p.CursorPosX = c3 : p.WriteText("(ml/min/mmHg/L)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("KCO|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("KCO|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Kco)) / Val(dPreds("KCO|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Kco) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("KCO Hb")
                p.CursorPosX = c3 : p.WriteText("(ml/min/mmHg/L)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("KCO|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("KCO|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Kco)) / Val(dPreds("KCO|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Kco) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("VA")
                p.CursorPosX = c3 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("VA|LLN", dPreds, 1), fHel10i)
                If dPreds.ContainsKey("VA|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Va)) / Val(dPreds("VA|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Va) & ppn & vbCrLf)
            End If

            'LVs
            If td.vols_done Then
                If Not InitialLineFeed Then p.WriteText(vbCrLf)
                InitialLineFeed = True

                p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + 0.1 * ctp
                p.WriteText("Static Lung Volumes" & vbCrLf, fHel10b)
                p.CursorPosX = c2 : p.WriteText("TLC")
                p.CursorPosX = c3 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("TLC|LLN", dPreds, 2), fHel10i)
                If dPreds.ContainsKey("TLC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Tlc)) / Val(dPreds("TLC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Tlc) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("FRC")
                p.CursorPosX = c3 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("FRC|LLN", dPreds, 2), fHel10i)
                If dPreds.ContainsKey("FRC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Frc)) / Val(dPreds("FRC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Frc) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("RV")
                p.CursorPosX = c3 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("RV|LLN", dPreds, 2), fHel10i)
                If dPreds.ContainsKey("RV|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Rv)) / Val(dPreds("RV|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Rv) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("RV/TLC")
                p.CursorPosX = c3 : p.WriteText("(%)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("RV/TLC|ULN", dPreds, 0), fHel10i)
                If R(fld.R_Bl_RvTlc) = Nothing Then
                    rvtlc = Format(100 * CSng(R(fld.R_Bl_Rv)) / CSng(R(fld.R_Bl_Tlc)), "0")
                Else
                    rvtlc = R(fld.R_Bl_RvTlc)
                End If
                If dPreds.ContainsKey("RV/TLC|MPV") Then ppn = cPred.Format_Pred(100 * rvtlc / Val(dPreds("RV/TLC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(rvtlc & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("VC")
                p.CursorPosX = c3 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("VC|LLN", dPreds, 2), fHel10i)
                If dPreds.ContainsKey("VC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_LvVc)) / Val(dPreds("VC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_LvVc) & ppn & vbCrLf)
            End If

            'MRPs
            If td.mrps_done Then
                If Not InitialLineFeed Then p.WriteText(vbCrLf)
                InitialLineFeed = True

                p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + 0.1 * ctp
                p.WriteText("Maximum Respiratory Pressures" & vbCrLf, fHel10b)
                p.CursorPosX = c2 : p.WriteText("MIP")
                p.CursorPosX = c3 : p.WriteText("(cmH2O)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("MIP|LLN", dPreds, 0), fHel10i)
                If dPreds.ContainsKey("MIP|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Mip)) / Val(dPreds("MIP|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Mip) & ppn & vbCrLf)

                p.CursorPosX = c2 : p.WriteText("MEP")
                p.CursorPosX = c3 : p.WriteText("(cmH2O)")
                p.CursorPosX = c4 : p.WriteText(cPred.Get_Pred("MEP|LLN", dPreds, 0), fHel10i)
                If dPreds.ContainsKey("MEP|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(fld.R_Bl_Mep)) / Val(dPreds("MEP|MPV")), 0) Else ppn = ""
                p.CursorPosX = c5 : p.WriteText(R(fld.R_Bl_Mep) & ppn & vbCrLf)
            End If

            'SpO2
            If td.oximetry_done Then
                If Not InitialLineFeed Then p.WriteText(vbCrLf)
                InitialLineFeed = True

                p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + 0.1 * ctp
                p.WriteText("Pulse Oximetry" & vbCrLf, fHel10b)
                p.CursorPosX = c2 : p.WriteText("SpO2")
                p.CursorPosX = c3 : p.WriteText("(%)")
                p.CursorPosX = c4 : p.WriteText("")
                p.CursorPosX = c5 : p.WriteText(R(fld.R_SpO2_1) & vbCrLf)
            End If

            'Reference Ypos for report section to come
            If p.CursorPosY / ctp > flowvol_yBottom Then
                Return p.CursorPosY / ctp
            Else
                Return flowvol_yBottom
            End If

        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_routine" & vbCrLf & Err.Description)
            Return 0
        End Try

    End Function

    Private Function Draw_Pdf_F1_rft_results_cpet(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String), ByVal yPos As Single) As Single

        Try

            Dim p As PDFPage = pdf.GetPage(1)
            Dim Ystart As Single = 0
            Dim incResultsHeadings As Single = 0.25
            Dim c1 As Single = m.xMin * ctp             'Test header

            'Results header
            Ystart = yPos + incResultsHeadings * ctp
            p.CursorPosY = Ystart
            p.Font = fHel10
            p.CursorPosX = c1 : p.WriteText("CARDIO-PULMONARY EXERCISE TEST", fHel10b)





            Return 20
        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_cpet" & vbCrLf & Err.Description)
            Return 20
        End Try

    End Function

    Private Function Draw_Pdf_F1_rft_results_spt(ByVal Test As Dictionary(Of String, String), Allergens() As Dictionary(Of String, String), ByVal yPos As Single) As Single


        Try
            Dim i As Integer = 0
            Dim p As PDFPage = pdf.GetPage(1)
            Dim Ystart As Single = 0
            Dim Yfinish As Single = 0
            Dim incResultsHeadings As Single = 0.25
            Dim incResultsLines As Single = 0.075
            Dim c1 As Single = m.xMin * ctp             'Test header
            Dim c2 As Single = c1 + 0.5 * ctp
            Dim c3 As Single = c2 + 1 * ctp
            Dim c4 As Single = c3 + 3 * ctp
            Dim c5 As Single = c4 + 3 * ctp
            Dim c6 As Single = c5 + 5 * ctp

            'Results header
            Ystart = yPos + incResultsHeadings * ctp
            p.CursorPosY = Ystart
            p.Font = fHel10
            p.CursorPosX = c1 : p.WriteText("SKIN PRICK ALLERGEN TEST" & vbCrLf & vbCrLf, fHel10b)

            p.CursorPosX = c2 : p.WriteText("#", fHel10b)
            p.CursorPosX = c3 : p.WriteText("Allergen Group", fHel10b)
            p.CursorPosX = c4 : p.WriteText("Allergen", fHel10b)
            p.CursorPosX = c5 : p.WriteText("Wheal (mm)" & vbCrLf, fHel10b)

            Dim f As New class_fields_Spt_Allergens
            i = 1
            For Each dic As Dictionary(Of String, String) In Allergens
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c2 : p.WriteText(i)
                p.CursorPosX = c3 : p.WriteText(dic(f.allergen_category_name))
                p.CursorPosX = c4 : p.WriteText(dic(f.allergen_name))
                p.CursorPosX = c5 : p.WriteText(dic(f.wheal_mm) & vbCrLf)
                i += 1
            Next

            p.WriteText(vbCrLf & vbCrLf)
            p.CursorPosX = c1 : p.WriteText("Site: ", fHel10b)
            p.CursorPosX = c1 + 0.8 * ctp : p.WriteText(Test("site"), fHel10)

            p.CursorPosX = c4 : p.WriteText("Medications last 48hrs:", fHel10b)
            Dim meds() As String = Split(Test("medications"), vbCrLf)
            Dim med As String = ""
            For i = 0 To meds.Count - 1
                If meds(i) <> "" Then med = med & meds(i) & ", "
            Next
            If med.Length > 0 Then med = Left(med, med.Length - 2)
            p.CursorPosX = c4 + 3.7 * ctp : p.WriteText(med, fHel10)

            p.WriteText(vbCrLf & vbCrLf & vbCrLf)
            Return p.CursorPosY / ctp

        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_spt" & vbCrLf & Err.Description)
            Return 20
        End Try

    End Function

    Private Function Draw_Pdf_F1_rft_results_walk(testtype As String, protocolID As Integer, trials() As class_walktest_plot.walk_trialdata, ByVal yPos As Single) As Single

        Try

            Dim p As PDFPage = pdf.GetPage(1)
            Dim incResultsLines As Single = 0.075
            Dim incResultsHeadings As Single = 0.4
            Dim Y As Single = m.yMin + 5
            Dim Ystart As Single = 0
            Dim ppn As String = ""
            Dim x1 As Single = 0, x2 As Single = 0, y1 As Single = 0, y2 As Single = 0
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = c1 + 1.5 * ctp
            Dim c3 As Single = c2 + 1.5 * ctp
            Dim c4 As Single = c3 + 1.5 * ctp
            Dim c5 As Single = c4 + 1.5 * ctp
            Dim c6 As Single = c5 + 1.5 * ctp
            Dim c7 As Single = c6 + 1.5 * ctp
            Dim c8 As Single = c7 + 1.5 * ctp
            Dim r1 As Single = 0
            Dim r2 As Single = 0
            Dim r3 As Single = 0
            Dim r4 As Single = 0
            Dim r5 As Single = 0
            Dim r6 As Single = 0

            'Results header
            Ystart = yPos + 0.2 * ctp
            p.Font = fHel10
            p.CursorPosY = Ystart : p.CursorPosX = c1 : p.WriteText(testtype & vbCrLf & vbCrLf, fHel10b)

            'Table
            r1 = p.CursorPosY + 0.5 * ctp
            r2 = r1 + 0.3 * ctp
            r3 = r2 + 0.5 * ctp
            r4 = r3 + 0.5 * ctp
            r5 = r4 + 0.5 * ctp
            r6 = r5 + 0.5 * ctp
            p.CursorPosY = r1
            p.CursorPosX = c4 : p.WriteText("Suppl O2", fHel8)
            p.CursorPosX = c5 : p.WriteText("SpO2", fHel8)
            p.CursorPosX = c6 : p.WriteText("HR", fHel8)
            p.CursorPosX = c7 : p.WriteText("Dyspnoea", fHel8)
            p.CursorPosX = c8 : p.WriteText("Distance", fHel8)
            p.CursorPosY = r2
            p.CursorPosX = c4 : p.WriteText("(L/min)", fHel8)
            p.CursorPosX = c5 : p.WriteText("(%)", fHel8)
            p.CursorPosX = c6 : p.WriteText("(bpm)", fHel8)
            p.CursorPosX = c7 : p.WriteText("(Borg)", fHel8)
            p.CursorPosX = c8 : p.WriteText("(m)", fHel8)
            p.CursorPosY = r3
            p.CursorPosX = c2 : p.WriteText("Trial 1", fHel8)
            p.CursorPosX = c3 : p.WriteText("Rest", fHel8)
            p.CursorPosY = r4
            p.CursorPosX = c3 : p.WriteText("Walk", fHel8)
            p.CursorPosY = r5
            p.CursorPosX = c2 : p.WriteText("Trial 2", fHel8)
            p.CursorPosX = c3 : p.WriteText("Rest", fHel8)
            p.CursorPosY = r6
            p.CursorPosX = c3 : p.WriteText("Walk", fHel8)


            Dim s() As class_walktest_plot.walk_SummaryResults = cWalk.Calculate_SummaryResults(trials, protocolID)
            'Trial 1
            p.CursorPosY = r3
            p.CursorPosX = c4 : p.WriteText(s(0).FiO2_rest, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(0).SpO2_rest, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(0).HR_rest, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(0).Dyspnoea_rest, fHel8)
            p.CursorPosX = c8 : p.WriteText("---", fHel8)
            p.CursorPosY = r4
            p.CursorPosX = c4 : p.WriteText(s(0).FiO2_exercise, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(0).SpO2_exercise, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(0).HR_exercise, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(0).Dyspnoea_exercise, fHel8)
            p.CursorPosX = c8 : p.WriteText(s(0).Distance_exercise, fHel8)
            'Trial 2
            p.CursorPosY = r5
            p.CursorPosX = c4 : p.WriteText(s(1).FiO2_rest, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(1).SpO2_rest, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(1).HR_rest, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(1).Dyspnoea_rest, fHel8)
            p.CursorPosX = c8 : p.WriteText("---", fHel8)
            p.CursorPosY = r6
            p.CursorPosX = c4 : p.WriteText(s(1).FiO2_exercise, fHel8)
            p.CursorPosX = c5 : p.WriteText(s(1).SpO2_exercise, fHel8)
            p.CursorPosX = c6 : p.WriteText(s(1).HR_exercise, fHel8)
            p.CursorPosX = c7 : p.WriteText(s(1).Dyspnoea_exercise, fHel8)
            p.CursorPosX = c8 : p.WriteText(s(1).Distance_exercise, fHel8)

            Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Black, 0.1)
            x1 = c2 / ctp - 0.2 : x2 = x1 + 11 : y1 = r1 / ctp - 0.1
            p.DrawLine(myPen, x1, y1, x2, y1)
            y1 = r3 / ctp - 0.1
            p.DrawLine(myPen, x1, y1, x2, y1)
            y1 = r5 / ctp - 0.1
            p.DrawLine(myPen, x1, y1, x2, y1)
            y1 = r6 / ctp + 0.4
            p.DrawLine(myPen, x1, y1, x2, y1)

            'Draw graph                    
            Dim pr = cWalkPlot.Get_plotproperties_walk
            Dim a() As List(Of Single) = cWalkPlot.Convert_XYdatatoArray(trials)

            ReDim pr.xData(0 To 1)
            ReDim pr.yData(0 To 1)
            pr.xData(0) = a(0)
            pr.yData(0) = a(1)
            pr.xData(1) = a(2)
            pr.yData(1) = a(3)

            Dim ch As Chart = cWalkPlot.Create_plot(pr)
            ch.SaveImage("c:\temp\pdr_chart.jpeg", ImageFormat.Jpeg)

            Y = p.CursorPosY / ctp + 1
            Dim rect As New Rectangle(x1 + 1, Y, 10.0, 6.0)
            p.DrawRectangle(rect, False, True)
            p.DrawImage("C:\temp\pdr_chart.jpeg", rect, True, True)

            Return m.yMax - 10      'Reference Ypos for report section to come






        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_walk" & vbCrLf & Err.Description)
            Return 20
        End Try

    End Function

    Private Function Draw_Pdf_F1_rft_results_hast(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String), ByVal yPos As Single) As Single

        Try


            Dim p As PDFPage = pdf.GetPage(1)
            Dim incResultsLines As Single = 0.075
            Dim incResultsHeadings As Single = 0.4
            Dim Y As Single = m.yMin + 5
            Dim Ystart As Single = 0
            Dim ppn As String = ""
            Dim x1 As Single = 0, x2 As Single = 0, y1 As Single = 0, y2 As Single = 0
            Dim c1 As Single = m.xMin * ctp
            Dim c2 As Single = c1 + 1.5 * ctp
            Dim c3 As Single = c2 + 1.5 * ctp
            Dim c4 As Single = c3 + 1.5 * ctp
            Dim c5 As Single = c4 + 1.5 * ctp
            Dim c6 As Single = c5 + 1.5 * ctp
            Dim c7 As Single = c6 + 1.5 * ctp
            Dim c8 As Single = c7 + 1.5 * ctp
            Dim r1 As Single = 0
            Dim r2 As Single = 0
            Dim r3 As Single = 0
            Dim r4 As Single = 0
            Dim r5 As Single = 0
            Dim r6 As Single = 0

            'Results header
            Ystart = yPos + incResultsHeadings * ctp
            p.CursorPosY = Ystart
            p.Font = fHel10
            p.CursorPosX = c1 : p.WriteText("ALTITUDE SIMULATION TEST", fHel10b)


            'Table
            r1 = p.CursorPosY + 0.5 * ctp
            r2 = r1 + 0.3 * ctp
            r3 = r2 + 0.5 * ctp
            r4 = r3 + 0.5 * ctp
            r5 = r4 + 0.5 * ctp
            r6 = r5 + 0.5 * ctp
            p.CursorPosY = r1
            p.CursorPosX = c6 : p.WriteText("Altitude", fHel8)
            p.CursorPosX = c4 : p.WriteText("Suppl O2", fHel8)
            p.CursorPosX = c5 : p.WriteText("SpO2", fHel8)

            p.CursorPosY = r2
            p.CursorPosX = c4 : p.WriteText("(feet)", fHel8)
            p.CursorPosX = c5 : p.WriteText("(L/min)", fHel8)
            p.CursorPosX = c6 : p.WriteText("(%)", fHel8)




            Return 20
        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_hast" & vbCrLf & Err.Description)
            Return 20
        End Try

    End Function


    Private Function Draw_Pdf_F1_rft_results_routine(ByVal R As Dictionary(Of String, String), ByVal D As Dictionary(Of String, String), ByVal yPos As Single) As Single

        Dim p As PDFPage = pdf.GetPage(1)
        Dim rf As New class_Rft_RoutineAndSessionFields
        Dim df As New class_DemographicFields
        Dim FirstTestDone As Boolean = False
        Dim incResultsLines As Single = 0.075
        Dim incResultsHeadings As Single = 0.25
        Dim Y As Single = m.yMin + 5
        Dim Ystart As Single = 0
        Dim ppn As String = ""
        Dim fer_pre As String = ""
        Dim fer_post As String = ""
        Dim rvtlc As String = ""
        Dim td As TestsDone = cMyRoutines.Get_TestsDone(R)
        Dim c1 As Single = m.xMin * ctp             'Test header
        Dim c2 As Single = c1 + 1.5 * ctp                   'Units
        Dim c3 As Single = c2 + 3 * ctp                     'Normal range
        Dim c4 As Single = c3 + 3 * ctp                     'Baseline
        Dim c5 As Single = c4 + 1.6 * ctp                   'PPN
        Dim c6 As Single = c5 + 2.8 * ctp                   'Post
        Dim c7 As Single = c6 + 1.6 * ctp                   'PPN
        Dim c8 As Single = c7 + 2 * ctp                   'Post % change

        Try
            'Get normals
            Dim demo As Pred_demo = Nothing
            demo.Age = cMyRoutines.Calc_Age(CDate(D(df.DOB)), CDate(R(rf.TestDate)))
            If IsDate(D(df.DOB)) Then demo.DOB = CDate(D(df.DOB)) Else D(df.DOB) = Nothing
            demo.Htcm = Val(R(rf.Height))
            demo.Wtkg = Val(R(rf.Weight))
            demo.GenderID = cMyRoutines.Lookup_list_ByDescription(D(df.Gender), eTables.Pred_ref_genders)
            demo.Gender = D(df.Gender)
            demo.EthnicityID = cMyRoutines.Lookup_list_ByDescription(D(df.Race_forRfts), eTables.Pred_ref_ethnicities)
            demo.Ethnicity = D(df.Race_forRfts)
            demo.TestDate = R(rf.TestDate)
            demo.SourcesString = R(rf.Pred_SourceIDs)

            Dim dPreds As Dictionary(Of String, String) = cPred.Get_PredValues(demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)    'ParameterID|StatTypeID, result

            'Results header
            Ystart = yPos + incResultsHeadings * ctp
            p.CursorPosY = Ystart
            p.Font = fHel10

            If td.anyspir_done Then
                FirstTestDone = True

                p.CursorPosX = c1 : p.WriteText("SPIROMETRY", fHel10b)
                p.CursorPosX = c3 : p.WriteText("Normal range", fHel10b)
                p.CursorPosX = c4 : p.WriteText("Baseline", fHel10b)
                p.CursorPosX = c5 : p.WriteText("(%mpv)", fHel10b)
                p.CursorPosX = c6 : p.WriteText(Left(R(rf.R_post_Condition), 9), fHel10b)
                p.CursorPosX = c7 : p.WriteText("(%mpv)", fHel10b)
                p.CursorPosX = c8 : p.WriteText("% Change" & vbCrLf, fHel10b)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("FEV1")
                p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FEV1|LLN", dPreds, 1))
                If dPreds.ContainsKey("FEV1|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_bl_Fev1)) / Val(dPreds("FEV1|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_bl_Fev1))
                p.CursorPosX = c5 : p.WriteText(ppn)
                If dPreds.ContainsKey("FEV1|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_post_Fev1)) / Val(dPreds("FEV1|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(rf.R_post_Fev1))
                p.CursorPosX = c7 : p.WriteText(ppn)
                p.CursorPosX = c8 : p.WriteText(pcChange(Val(R(rf.R_bl_Fev1)), Val(R(rf.R_post_Fev1))) & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("FVC")
                p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FVC|LLN", dPreds, 1))
                If dPreds.ContainsKey("FVC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_bl_Fvc)) / Val(dPreds("FVC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_bl_Fvc))
                p.CursorPosX = c5 : p.WriteText(ppn)
                If dPreds.ContainsKey("FVC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_post_Fvc)) / Val(dPreds("FVC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(rf.R_post_Fvc))
                p.CursorPosX = c7 : p.WriteText(ppn)
                p.CursorPosX = c8 : p.WriteText(pcChange(Val(R(rf.R_bl_Fvc)), Val(R(rf.R_post_Fvc))) & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("VC")
                p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("VC|LLN", dPreds, 1))
                If dPreds.ContainsKey("VC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_bl_Vc)) / Val(dPreds("VC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_bl_Vc))
                p.CursorPosX = c5 : p.WriteText(ppn)
                If dPreds.ContainsKey("VC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_post_Vc)) / Val(dPreds("VC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(R(rf.R_post_Vc))
                p.CursorPosX = c7 : p.WriteText(ppn)
                p.CursorPosX = c8 : p.WriteText(pcChange(Val(R(rf.R_bl_Vc)), Val(R(rf.R_post_Vc))) & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("FER")
                p.CursorPosX = c2 : p.WriteText("(%)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FER|LLN", dPreds, 0))
                If R(rf.R_Bl_Fer) = Nothing Then
                    fer_pre = cMyRoutines.Calc_Fer(Val(R(rf.R_bl_Fev1)), Val(R(rf.R_bl_Fvc)), Val(R(rf.R_bl_Vc)))
                Else
                    fer_pre = R(rf.R_Bl_Fer)
                End If
                If dPreds.ContainsKey("FER|MPV") Then ppn = cPred.Format_Pred(100 * Val(fer_pre) / Val(dPreds("FER|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(fer_pre)
                p.CursorPosX = c5 : p.WriteText(ppn)
                If R(rf.R_Post_Fer) = Nothing Then
                    fer_post = cMyRoutines.Calc_Fer(Val(R(rf.R_post_Fev1)), Val(R(rf.R_post_Fvc)), Val(R(rf.R_post_Vc)))
                Else
                    fer_post = R(rf.R_Post_Fer)
                End If
                If dPreds.ContainsKey("FER|MPV") Then ppn = cPred.Format_Pred(100 * Val(fer_post) / Val(dPreds("FER|MPV")), 0) Else ppn = ""
                p.CursorPosX = c6 : p.WriteText(fer_post)
                p.CursorPosX = c7 : p.WriteText(ppn)
                p.CursorPosX = c8 : p.WriteText(pcChange(Val(fer_pre), Val(fer_post)) & vbCrLf)

            Else
                FirstTestDone = False

                p.CursorPosX = c1 : p.WriteText("", fHel10b)
                p.CursorPosX = c3 - 0.5 / ctp : p.WriteText("Normal range", fHel10b)
                p.CursorPosX = c4 : p.WriteText("Baseline", fHel10b)
                p.CursorPosX = c5 : p.WriteText("(%mpv)", fHel10b)
            End If

            If td.tlco_done Then

                Dim HbFactor As Single = cMyRoutines.calc_HbFac(R(rf.R_Bl_Hb), demo.DOB, D(df.Gender), demo.TestDate)
                Dim Hb_Done As Boolean = CBool(HbFactor)

                If FirstTestDone = False Then
                    p.CursorPosX = c1 : p.CursorPosY = Ystart : p.WriteText("CO TRANSFER FACTOR" & vbCrLf, fHel10b)
                Else
                    p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + incResultsHeadings * ctp
                    p.WriteText("CO TRANSFER FACTOR" & vbCrLf, fHel10b)
                End If
                FirstTestDone = True

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("TLCO")
                p.CursorPosX = c2 : p.WriteText("(ml/min/mmHg)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("TLCO|LLN", dPreds, 1))
                If dPreds.ContainsKey("TLCO|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Tlco)) / Val(dPreds("TLCO|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Tlco))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                If Hb_Done Then
                    p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                    p.CursorPosX = c1 : p.WriteText("TLCO Hb")
                    p.CursorPosX = c2 : p.WriteText("(ml/min/mmHg)")
                    p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("TLCO|LLN", dPreds, 1))
                    If dPreds.ContainsKey("TLCO|MPV") Then ppn = cPred.Format_Pred(100 * HbFactor * Val(R(rf.R_Bl_Tlco)) / Val(dPreds("TLCO|MPV")), 0) Else ppn = ""
                    p.CursorPosX = c4 : p.WriteText(Format(HbFactor * Val(R(rf.R_Bl_Tlco)), "##.0"))
                    p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
                End If

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("KCO")
                p.CursorPosX = c2 : p.WriteText("(ml/min/mmHg/L)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("KCO|Range", dPreds, 1))
                If dPreds.ContainsKey("KCO|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Kco)) / Val(dPreds("KCO|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Kco))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                If Hb_Done Then
                    p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                    p.CursorPosX = c1 : p.WriteText("KCO Hb")
                    p.CursorPosX = c2 : p.WriteText("(ml/min/mmHg/L)")
                    p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("KCO|Range", dPreds, 1))
                    If dPreds.ContainsKey("KCO|MPV") Then ppn = cPred.Format_Pred(100 * HbFactor * Val(R(rf.R_Bl_Kco)) / Val(dPreds("KCO|MPV")), 0) Else ppn = ""
                    p.CursorPosX = c4 : p.WriteText(Format(HbFactor * Val(R(rf.R_Bl_Kco)), "##.00"))
                    p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                    p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                    p.CursorPosX = c1 : p.WriteText("Hb")
                    p.CursorPosX = c2 : p.WriteText("(gm/dL)")
                    p.CursorPosX = c4 : p.WriteText(Format(Val(R(rf.R_Bl_Hb)), "##.0") & vbCrLf)
                End If

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("VA")
                p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("VA|LLN", dPreds, 1))
                If dPreds.ContainsKey("VA|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Va)) / Val(dPreds("VA|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Va))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("Vin")
                p.CursorPosX = c2 : p.WriteText("(L,BTPS)")
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Ivc) & vbCrLf)

            End If

            If td.vols_done Then
                If FirstTestDone = False Then
                    p.CursorPosX = c1 : p.CursorPosY = Ystart : p.WriteText("LUNG VOLUMES (" & R(rf.LungVolumes_method) & ")" & vbCrLf, fHel10b)
                Else
                    p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + incResultsHeadings * ctp
                    p.WriteText("LUNG VOLUMES (" & R(rf.LungVolumes_method) & ")" & vbCrLf, fHel10b)
                End If
                FirstTestDone = True

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("VC")
                p.CursorPosX = c2 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("VC|LLN", dPreds, 2))
                If dPreds.ContainsKey("VC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_LvVc)) / Val(dPreds("FVC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_LvVc))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("TLC")
                p.CursorPosX = c2 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("TLC|Range", dPreds, 2))
                If dPreds.ContainsKey("TLC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Tlc)) / Val(dPreds("TLC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Tlc))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("FRC")
                p.CursorPosX = c2 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FRC|Range", dPreds, 2))
                If dPreds.ContainsKey("FRC|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Frc)) / Val(dPreds("FRC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Frc))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("RV")
                p.CursorPosX = c2 : p.WriteText("(L/BTPS)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("RV|Range", dPreds, 2))
                If dPreds.ContainsKey("RV|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Rv)) / Val(dPreds("RV|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Rv))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("RV/TLC")
                p.CursorPosX = c2 : p.WriteText("(%)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("RV/TLC|ULN", dPreds, 0))
                If R(rf.R_Bl_RvTlc) = Nothing Then
                    rvtlc = Format(100 * CSng(R(rf.R_Bl_Rv)) / CSng(R(rf.R_Bl_Tlc)), "0")
                Else
                    rvtlc = R(rf.R_Bl_RvTlc)
                End If
                If dPreds.ContainsKey("RV/TLC|MPV") Then ppn = cPred.Format_Pred(100 * rvtlc / Val(dPreds("RV/TLC|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(rvtlc)
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            End If

            'MRPs
            If td.mrps_done Then
                If FirstTestDone = False Then
                    p.CursorPosX = c1 : p.CursorPosY = Ystart : p.WriteText("MAX RESP PRESSURES" & vbCrLf, fHel10b)
                Else
                    p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + incResultsHeadings * ctp
                    p.WriteText("MAX RESP PRESSURES" & vbCrLf, fHel10b)
                End If
                FirstTestDone = True

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("MIP")
                p.CursorPosX = c2 : p.WriteText("(cmH2O)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("MIP|LLN", dPreds, 0))
                If dPreds.ContainsKey("MIP|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Mip)) / Val(dPreds("MIP|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Mip))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("MEP")
                p.CursorPosX = c2 : p.WriteText("(cmH2O)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("MEP|LLN", dPreds, 0))
                If dPreds.ContainsKey("MEP|MPV") Then ppn = cPred.Format_Pred(100 * Val(R(rf.R_Bl_Mep)) / Val(dPreds("MEP|MPV")), 0) Else ppn = ""
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_Mep))
                p.CursorPosX = c5 : p.WriteText(ppn & vbCrLf)
            End If


            'FeNO
            If td.feno_done Then
                If FirstTestDone = False Then
                    p.CursorPosX = c1 : p.CursorPosY = Ystart : p.WriteText("EXHALED NITRIC OXIDE" & vbCrLf, fHel10b)
                Else
                    p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + incResultsHeadings * ctp
                    p.WriteText("EXHALED NITRIC OXIDE" & vbCrLf, fHel10b)
                End If
                FirstTestDone = True

                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("FeNO")
                p.CursorPosX = c2 : p.WriteText("(ppb)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("FeNO|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_Bl_FeNO) & vbCrLf)
            End If

            'ABGs
            If td.anyabgs_done Then
                If FirstTestDone = False Then
                    p.CursorPosX = c1 : p.CursorPosY = Ystart : p.WriteText("ARTERIAL BLODD GASES" & vbCrLf, fHel10b)
                Else
                    p.CursorPosX = c1 : p.CursorPosY = p.CursorPosY + incResultsHeadings * ctp
                    p.WriteText("ARTERIAL BLOOD GASES" & vbCrLf, fHel10b)
                End If
                FirstTestDone = True

                Dim cAbg2 = c5 + 1.5 * ctp
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("FiO2")
                p.CursorPosX = c2 : p.WriteText("")
                p.CursorPosX = c3 : p.WriteText("")
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_fio2))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_fio2) & vbCrLf)
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("pH")
                p.CursorPosX = c2 : p.WriteText("")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("ph|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_ph))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_ph) & vbCrLf)
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("PaCO2")
                p.CursorPosX = c2 : p.WriteText("(mmHg)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("paco2|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_paco2))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_paco2) & vbCrLf)
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("PaO2")
                p.CursorPosX = c2 : p.WriteText("(mmHg)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("pao2|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_pao2))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_pao2) & vbCrLf)
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("HCO3")
                p.CursorPosX = c2 : p.WriteText("(mmol/L)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("hco3|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_hco3))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_hco3) & vbCrLf)
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("BE")
                p.CursorPosX = c2 : p.WriteText("(mmol/L)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("be|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_be))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_be) & vbCrLf)
                p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                p.CursorPosX = c1 : p.WriteText("SaO2")
                p.CursorPosX = c2 : p.WriteText("(%)")
                p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("be|Range", dPreds, 0))
                p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_sao2))
                p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_sao2) & vbCrLf)
                If td.shunt_done Then
                    p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                    p.CursorPosX = c1 : p.WriteText("Shunt (Anatomic)   (%)")
                    p.CursorPosX = c2 : p.WriteText("")
                    p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("shunt|Range", dPreds, 0))
                    p.CursorPosX = c4 : p.WriteText(R(rf.R_abg1_shunt))
                    p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_abg2_shunt) & vbCrLf)
                End If
                If td.oximetry_done Then
                    p.CursorPosY = p.CursorPosY + incResultsLines * ctp
                    p.CursorPosX = c1 : p.WriteText("SpO2")
                    p.CursorPosX = c2 : p.WriteText("(%)")
                    p.CursorPosX = c3 : p.WriteText(cPred.Get_Pred("SpO2|Range", dPreds, 0))
                    p.CursorPosX = c4 : p.WriteText(R(rf.R_SpO2_1))
                    p.CursorPosX = cAbg2 : p.WriteText(R(rf.R_SpO2_2) & vbCrLf)
                End If

            End If

            Y = 0.3 + p.CursorPosY / ctp
            Return Y

        Catch e As Exception
            MsgBox("Error in class_pdf.Draw_Pdf_F1_rft_results_routine" & vbCrLf & Err.Description)
            Return 20
        End Try

    End Function

    Private Function Draw_FlowVolImage(rftID As Long, tbl As eTables, xPos As Single, yPos As Single, xSize As Single, ySize As Single)
        'Dimensions in cm

        Dim p As PDFPage = pdf.GetPage(1)
        Dim f As String = cMyRoutines.Get_TempDirectory & "\reslab_temp.tmp"
        If My.Computer.FileSystem.FileExists(f) Then My.Computer.FileSystem.DeleteFile(f)
        If cDAL.Get_imageAsFile(rftID, tbl, "flowvolloop", f) = True Then
            Dim rect As New RectangleF(xPos, yPos, xSize, ySize)
            p.DrawRectangle(rect, False, True)
            p.DrawImage(f, rect, True, True)
        End If

        Return yPos + ySize

    End Function








End Class




