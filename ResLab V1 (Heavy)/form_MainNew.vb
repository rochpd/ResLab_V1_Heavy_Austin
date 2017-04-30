
Imports Gnostice.PDFOne
Imports System.Drawing
Imports System.Drawing.Printing
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports Gnostice.PDFOne.Windows.PDFViewer
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class Form_MainNew

    Const SPLITTERCLOSEDDISTANCE As Integer = 25
    Public flagReturnKeyPressed As Boolean = False
    Dim formLoading As Boolean = False
    Public WithEvents grdTrend As DataGridView = Nothing
    Public WithEvents chrt As Chart = Nothing

    Dim _pt_fore_colour_focus = Color.Black
    Dim _pt_back_colour_focus = Color.White
    Dim _pt_fore_colour_nofocus = Color.White
    Dim _pt_back_colour_nofocus = Color.SteelBlue

    Private Sub Form_MainNew_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If Not Me.Enabled Then Me.Enabled = True

        If gRefreshMainForm Then
            Me.ResetForm(Get_CurrentUnit())
            Me.DisplayRecallItems(cPt.PatientID)
            gRefreshMainForm = False
            cUser.set_access(Me)
            ts_lblAccess.Text = "Access level: " & cUser.AccessLevel
        End If

        lvRecall.Focus()

    End Sub

    Private Sub frmMainNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        formLoading = True
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        'Check that can sucessfully connect to the database server
        Dim CnCheck As String = cDAL.CanConnectToDb
        Select Case CnCheck
            Case "Success"

                'Ping the reslab central server
                Dim p As New class_pingversion
                'p.ping_version()
                p = Nothing

                module_Initialise.InitialiseApp()

                'Check app version is latest
                Dim VersionMessage As String = ""
                ts_lblVersionMsg.ForeColor = Color.Red
                Select Case module_Initialise.CheckIfCurrentversion()
                    Case -1 : VersionMessage = "Can't connect to check that installed version is latest"
                    Case 0 : VersionMessage = "Latest version/build of ResLab (Heavy) installed" : ts_lblVersionMsg.ForeColor = Color.Green
                    Case 1 : VersionMessage = "This is an out of date build of ResLab (Heavy)" & " (Latest: " & module_Initialise.Get_CurrentBuild & ")"
                    Case 2 : VersionMessage = "This is an out of date version of ResLab (Heavy)" & " (Latest: " & module_Initialise.Get_Currentversion & ")"
                End Select
                'ts_lblVersionMsg.Text = VersionMessage
                ts_lblVersion.Text = "Version: " & My.Settings.Version
                ts_lblBuild.Text = "Build: " & My.Settings.BuildDate
                ts_lblLoginID.Text = "LoginID: " & Environment.UserName
                ts_lblServer.Text = "Server: " & My.Settings.Item(My.Settings.Selected_DB_type)
                ts_lblAccess.Text = "Access level: " & cUser.AccessLevel
                ts_lblHealthService.Text = "Location: " & cHs.Current_UserLocation_HSID_Name

                'Initial setup of splitter windows
                Me.Height = Me.Height * 1.01
                splitterUnits.SplitterDistance = SPLITTERCLOSEDDISTANCE

                splitterPts.SplitterDistance = splitterPts.Height / 2
                lvPatients.Height = splitterPts.Panel2.Height - panelDates.Height - tsPtList.Height

                lvUnits.Height = 1
                lvRecall.Height = splitterPts.Panel1.Height - tsTests.Height

                Me.Add_ProvProtocolsToMainMenu(tsMenuItem_BronchialChallengeTests)
                Me.Add_WalkProtocolsToMainMenu(tsMenuItem_WalkTests)

                'Trend plot initial settings
                chkShowValues.Checked = True
                optPlot_absolute.Tag = class_plot_trend.eYValueType.Absolute
                optPlot_pcChange.Tag = class_plot_trend.eYValueType.PercentChange
                optPlot_ppn.Tag = class_plot_trend.eYValueType.PercentPredicted
                optPlot_earliest.Checked = True

                'cPt.PatientID = 1
                gRefreshMainForm = True



            Case Else
                Dim Msg As String = "Can't connect to the backend database - a network problem has occurred." & vbCrLf & vbCrLf
                Msg = Msg & CnCheck & vbCrLf & vbCrLf
                Msg = Msg & "Connection string: " & cDAL.Get_ConnString & vbCrLf
                Msg = Msg & "Please contact your database administrator."
                MsgBox(Msg)
        End Select



        formLoading = False

    End Sub

    Public Function get_trendplotproperties(grd As DataGridView) As class_plot_trend.plotproperties

        Dim p As class_plot_trend.plotproperties = Nothing

        If Not IsNothing(grd) And grd.SelectedRows.Count > 0 Then
            p.Seriesdata.x = Me.Get_xDataToPlot(grd)
            p.Seriesdata.y = Me.Get_yDataToPlot(grd.Rows(0), grd.SelectedRows)
            p.Seriesdata.yLabels = Me.Get_yLabelsToPlot(grd.SelectedRows)
        End If

        If optPlot_absolute.Checked Then
            p.PlotYValueAs = class_plot_trend.eYValueType.Absolute
            p.yRef_Type = Nothing
            p.GridCanMultiselect = False
        ElseIf optPlot_pcChange.Checked Then
            p.PlotYValueAs = class_plot_trend.eYValueType.PercentChange
            p.GridCanMultiselect = True
            If optPlot_earliest.Checked Then
                p.yRef_Type = class_plot_trend.eReferenceType.EarliestTest
                p.yRef_Index = grdTrend.Columns.Count - 1
            Else
                p.yRef_Type = class_plot_trend.eReferenceType.SelectedTest
                p.yRef_Index = cmboRefTestDate.SelectedIndex
                p.yRef_Date = cmboRefTestDate.Text
            End If
        ElseIf optPlot_ppn.Checked Then
            p.PlotYValueAs = class_plot_trend.eYValueType.PercentPredicted
            p.yRef_Type = Nothing
            p.GridCanMultiselect = True
            p.yRef_Index = Nothing
        End If

        p.Autoscale = chkAutoscale.Checked
        If Not chkAutoscale.Checked Then
            p.Scale_Ymin = txtyaxis_min.Text
            p.Scale_Ymax = txtyaxis_max.Text
        End If

        p.ShowDataLabels = chkShowValues.Checked

        Return p

    End Function

    Public Sub EventHandler_ToolBar_rft(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'event handlers for toolbar items set in class_buildtoolstripmenu

        If cPt.PatientID = 0 Then
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        Else
            Select Case sender.name.ToString.ToLower
                Case "ts_contact_btn_rft"
                    Dim f As New form_Rft(0, Me)
                    f.Show()
                Case "ts_contact_btn_provs_histamine provocation", "ts_contact_btn_provs_methacholine provocation", "ts_contact_btn_provs_mannitol provocation"
                    Dim f As New form_rft_challenge(0, Me, CLng(sender.tag))
                    f.Show()
                Case "ts_contact_btn_spt"
                    Dim f As New form_spt(0, Me)
                    f.Show()
                Case "ts_contact_btn_cpx"
                    Dim f As New form_rft_cpet(0, Me)
                    f.Show()
                Case "ts_contact_btn_6mwd"
                    MsgBox("Not enabled", vbOKOnly, "ResLab")
                Case "ts_contact_btn_hast"
                    Dim f As New form_hast(0, Me)
                    f.Show()
                Case "ts_contact_btn_walk_treadmill test (nrg protocol)", "ts_contact_btn_walk_6mwd test (ats protocol)"
                    Dim f As New form_walktest(0, Me, CLng(sender.tag))
                    f.Show()
                Case "ts_contact_btn_edit"
                    Me.Edit_SelectedTest()
                Case "ts_contact_btn_trend"
                    tabReports.SelectTab(1)
            End Select
        End If

    End Sub

    Public Sub EventHandler_ToolBar_sleep(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'event handlers for toolbar items set in class_buildtoolstripmenu

        If cPt.PatientID = 0 Then
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        Else
            Select Case sender.ToString.ToUpper
                Case ""

            End Select
        End If

    End Sub

    Public Sub EventHandler_ToolBar_cpap(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'event handlers for toolbar items set in class_buildtoolstripmenu

        If cPt.PatientID = 0 Then
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        Else
            Select Case sender.ToString.ToUpper
                Case ""

            End Select
        End If

    End Sub

    Public Sub EventHandler_ToolBar_O2clinic(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'event handlers for toolbar items set in class_buildtoolstripmenu

        If cPt.PatientID = 0 Then
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        Else
            Select Case sender.ToString.ToUpper
                Case ""

            End Select
        End If

    End Sub

    Public Sub EventHandler_ToolBar_vrss(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'event handlers for toolbar items set in class_buildtoolstripmenu

        If cPt.PatientID = 0 Then
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        Else
            Select Case sender.ToString.ToUpper
                Case ""

            End Select
        End If

    End Sub

    Public Sub EventHandler_ToolBar_generic(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'event handlers for toolbar items set in class_buildtoolstripmenu

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        If cPt.PatientID = 0 Then
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        Else
            Select Case sender.name.ToString.ToLower
                Case "ts_contact_btn_print_ts_contact_report"
                    If PdfViewer1.PDFLoaded Then
                        Dim p = New PrintDialog
                        Dim returnVar As Boolean = True
                        gRefreshMainForm = False
                        If p.ShowDialog() = DialogResult.OK Then
                            Dim cPrt As New class_Printing
                            cPrt.PrintMe(PdfViewer1.Document, p.PrinterSettings.Copies)
                            cPrt = Nothing
                        End If
                        p.Dispose()
                    End If
                Case "ts_contact_btn_print_ts_contact_trend"
                    If Not IsNothing(chrt) And Not IsNothing(grdTrend) Then
                        Dim p = New PrintDialog
                        Dim returnVar As Boolean = True
                        gRefreshMainForm = False
                        If p.ShowDialog() = DialogResult.OK Then
                            Dim tempfile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\trend.pdf"
                            If FileSystem.FileExists(tempfile) Then FileSystem.DeleteFile(tempfile)
                            Dim cPrt As New class_Printing
                            Dim cPdf As New class_pdf
                            Dim pdf As PDFDocument = cPdf.Draw_Pdf_trend(cPt.PatientID, grdTrend, chrt)
                            pdf.Save(tempfile)
                            Dim v As New PDFViewer
                            v.LoadDocument(tempfile)
                            cPrt.PrintMe(v.Document, p.PrinterSettings.Copies)
                            cPrt = Nothing
                        End If
                        p.Dispose()
                    Else
                        MsgBox("Trend chart is not displayed - printing cancelled", vbOKOnly, "Print trend page")
                    End If
                Case "ts_contact_btn_zoom_ts_contact_fitwidth"
                    PdfViewer1.StdZoomType = Windows.PDFViewer.StandardZoomType.FitWidth
                    tsButton_View.Tag = Windows.PDFViewer.StandardZoomType.FitWidth
                Case "ts_contact_btn_zoom_ts_contact_fitpage"
                    PdfViewer1.StdZoomType = Windows.PDFViewer.StandardZoomType.FitPage
                    tsButton_View.Tag = Windows.PDFViewer.StandardZoomType.FitPage
                Case "ts_contact_btn_zoom_ts_contact_clearpage"
                    If PdfViewer1.PDFLoaded Then
                        PdfViewer1.CloseDocument()
                        PdfViewer1.Visible = False
                    End If
                Case "ts_contact_btn_savetofile"
                    If PdfViewer1.PDFLoaded Then
                        Dim defaultpath As String = cMyRoutines.Get_AppString("pdf_filepath")
                        Dim p = New SaveFileDialog
                        If defaultpath = "" Then
                            p.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Temp
                        Else
                            p.InitialDirectory = defaultpath
                        End If
                        p.Filter = "pdf document |*.pdf"
                        p.FilterIndex = 1
                        p.FileName = PdfViewer1.Tag
                        Dim returnVar As Boolean = True
                        gRefreshMainForm = False
                        If p.ShowDialog() = DialogResult.OK Then
                            PdfViewer1.Document.Save(p.FileName)
                        End If
                        cMyRoutines.Update_AppString("pdf_filepath", IO.Path.GetDirectoryName(p.FileName).ToString)
                        p.Dispose()
                    End If
            End Select
        End If

        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Public Sub Setup_lvUnits()

        lvUnits.Clear()
        lvUnits.View = View.Details
        lvUnits.FullRowSelect = True
        lvUnits.GridLines = True

        ' Create columns for the items and subitems
        lvUnits.Columns.Add("", 0, HorizontalAlignment.Left)
        lvUnits.Columns.Add("unitID", 0, HorizontalAlignment.Left)
        lvUnits.Columns.Add("Unit", splitter.SplitterDistance - 10, HorizontalAlignment.Left)
        lvUnits.HeaderStyle = ColumnHeaderStyle.None

    End Sub

    Private Sub Setup_lvtests()

        lvRecall.Clear()
        lvRecall.View = View.Details
        lvRecall.FullRowSelect = True
        lvRecall.GridLines = True

        ' Create columns for the items and subitems
        lvRecall.Columns.Add("", 0, HorizontalAlignment.Left)
        lvRecall.Columns.Add("PkID", 0, HorizontalAlignment.Left)
        lvRecall.Columns.Add("Table", 0, HorizontalAlignment.Left)
        lvRecall.Columns.Add("#", 30, HorizontalAlignment.Left)
        lvRecall.Columns.Add("Test date", 90, HorizontalAlignment.Left)
        lvRecall.Columns.Add("Time", 60, HorizontalAlignment.Left)
        lvRecall.Columns.Add("Test type", 180, HorizontalAlignment.Left)
        lvRecall.Columns.Add("Report status", 130, HorizontalAlignment.Left)

    End Sub

    Public Sub ResetForm(Unit As eUnits)

        'Remove current toolbar
        If splitter.Panel2.Controls.ContainsKey("ts_contact") Then
            splitter.Panel2.Controls.RemoveByKey("ts_contact")
        End If

        'Clear the pdf viewer
        Me.PdfViewer1.CloseDocument()
        Me.PdfViewer1.Visible = False

        'Clear the trend viewer
        Me.splitTrend.Panel1.Controls.Clear()
        Me.splitTrend.Panel2.Controls.Clear()

        Select Case Unit
            Case eUnits.Respiratory_Laboratory

                'Get rft menu
                splitter.Panel2.Controls.Add(cBuildTsMenu.RftMenu())

                'Show reports tab
                tabReports.Visible = True
                tabReports.SelectedTab = Me.tabReports.TabPages(0)

            Case eUnits.Sleep_Laboratory
                splitter.Panel2.Controls.Add(cBuildTsMenu.SleepMenu())
                tabReports.Visible = False

            Case eUnits.O2_Therapy_Clinic

                tabReports.Visible = False
            Case eUnits.CPAP_Clinic
                tabReports.Visible = False
            Case eUnits.Victorian_Respiratory_Support_Service
                tabReports.Visible = False
        End Select

        lvRecall.Focus()

    End Sub

    Private Sub SetupLv_patients()

        lvPatients.Clear()
        lvPatients.View = View.Details
        lvPatients.FullRowSelect = True
        lvPatients.GridLines = True

        ' Create columns for the items and subitems
        lvPatients.Columns.Add("", 0, HorizontalAlignment.Left)
        lvPatients.Columns.Add("#", 30, HorizontalAlignment.Left)
        lvPatients.Columns.Add("PkID", 0, HorizontalAlignment.Left)
        lvPatients.Columns.Add("Surname", 100, HorizontalAlignment.Left)
        lvPatients.Columns.Add("Firstname", 60, HorizontalAlignment.Left)
        lvPatients.Columns.Add("UR", 80, HorizontalAlignment.Left)
        lvPatients.Columns.Add("Gender", 50, HorizontalAlignment.Left)
        lvPatients.Columns.Add("DOB", 80, HorizontalAlignment.Left)

    End Sub

    Public Sub Get_Rfts(PatientID As Long)

        'Populate the list view with current patient's tests  
        Dim testtime As String = ""
        Dim testdate As String = ""
        Dim lvi As ListViewItem

        Me.Setup_lvtests()

        'Dim dicR() As Dictionary(Of String, String) = cPt.Get_AllRftsByPatientID(PatientID)
        Dim dicR() As Dictionary(Of String, String) = cRfts.Get_rfts_all_byPatientID(PatientID)
        If Not (dicR Is Nothing) Then
            For i As Integer = 0 To UBound(dicR)
                lvi = New ListViewItem
                testtime = dicR(i)("TestTime")
                If testtime = Nothing Then testtime = "" Else testtime = Format(CDate(testtime), "HH:mm")
                testdate = dicR(i)("TestDate")
                If testdate = Nothing Then testdate = "" Else testdate = Format(CDate(testdate), "dd/MM/yyyy")
                lvi.SubItems.Add(dicR(i)("ID"))
                lvi.SubItems.Add(dicR(i)("tbl"))
                lvi.SubItems.Add(i + 1)
                lvi.SubItems.Add(testdate)
                lvi.SubItems.Add(testtime)
                lvi.SubItems.Add(dicR(i)("TestType"))
                lvi.SubItems.Add(dicR(i)("Report_status"))
                lvRecall.Items.Add(lvi)
            Next
        End If

    End Sub

    Public Sub Get_Patients()
        'Populate the list view with patients  

        Dim dob
        Dim lvi As ListViewItem
        Dim flds As New class_DemographicFields
        Dim startdate As Date, enddate As Date

        'Check dates are valid
        If (IsDate(txtStart.Text) And IsDate(txtEnd.Text)) Then
            If CDate(txtEnd.Text) < CDate(txtStart.Text) Then
                MsgBox("Start date cannot be ealier than end date", vbOKOnly, "Error")
                Exit Sub
            Else
                startdate = CDate(txtStart.Text)
                enddate = CDate(txtEnd.Text)
            End If
        Else
            If Not (IsDate(txtStart.Text) Or IsDate(txtEnd.Text)) Then
                lvPatients.Clear()
                Exit Sub
            Else
                If Not IsDate(txtStart.Text) Then startdate = CDate(txtEnd.Text) : enddate = CDate(txtEnd.Text)
                If Not IsDate(txtEnd.Text) Then startdate = CDate(txtStart.Text) : enddate = CDate(txtStart.Text)
            End If
        End If

        'Get patients
        Dim dicR() As Dictionary(Of String, String) = cPt.Get_PatientsByDate(startdate, enddate, CInt(lvUnits.SelectedItems(0).SubItems(1).Text))
        If dicR Is Nothing Then
            lvPatients.Clear()
        Else
            Me.SetupLv_patients()
            For i As Integer = 0 To UBound(dicR)
                dob = dicR(i)(flds.DOB)
                If dob = Nothing Then dob = "" Else dob = Format(CDate(dicR(i)(flds.DOB)), "dd/MM/yyyy")
                lvi = New ListViewItem
                lvi.SubItems.Add(i + 1)
                lvi.SubItems.Add(dicR(i)(flds.PatientID))
                lvi.SubItems.Add(dicR(i)(flds.Surname))
                lvi.SubItems.Add(dicR(i)(flds.Firstname))
                lvi.SubItems.Add(dicR(i)(flds.UR))
                lvi.SubItems.Add(dicR(i)(flds.Gender))
                lvi.SubItems.Add(dob)
                lvPatients.Items.Add(lvi)
            Next
        End If
        lvPatients.Focus()

    End Sub

    Private Sub FindPatientToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindPatientToolStripMenuItem.Click

        Dim f As New form_Find("", "", "")
        f.Show()

    End Sub

    Private Sub toolbtnFindPt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolbtnFindPt.Click

        Dim f As New form_Find("", "", "")
        f.Show()

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsButton_exit.Click
        End
    End Sub

    Private Sub NewPatientToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewPatientToolStripMenuItem.Click

        Dim f As New form_Demographics(Me)
        f.Tag = 0
        f.Show()

    End Sub

    Private Sub toolbtnNewPt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolbtnNewPt.Click

        Dim f As New form_Demographics(Me)
        f.Tag = 0
        f.Show()

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click

        If cPt.PatientID > 0 Then
            Dim f As New form_Demographics(Me)
            f.Tag = cPt.PatientID
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Private Sub ListsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListsToolStripMenuItem.Click

        form_prefs_list.Show()

    End Sub

    Private Sub NormalValuesToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NormalValuesToolStripMenuItem.Click

        form_prefs_normals.Show()

    End Sub

    Private Sub PredictedEquationsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsMenuItem_PredsManager.Click

        form_Pred.Show()

    End Sub

    Private Sub ReportingToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsButton_Reporting.Click

        form_Reporting.Show()

    End Sub

    Private Sub Do_pdf_report(ByVal PtName As String, ByVal TestDateTime As String, ByVal RecordID As Integer, ByVal eTbl As eTables)

        Dim filenamepath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & PtName & "_" & TestDateTime & ".pdf"
        Dim filename As String = PtName & "_" & TestDateTime & ".pdf"

        If PdfViewer1.PDFLoaded Then PdfViewer1.CloseDocument()
        Try
            If FileSystem.FileExists(filenamepath) Then FileSystem.DeleteFile(filenamepath)
            Dim cPdf As New class_pdf
            Dim pdf As PDFDocument = cPdf.Draw_Pdf(cPt.PatientID, RecordID, eTbl)
            pdf.Save(filenamepath)
            PdfViewer1.Visible = True
            PdfViewer1.LoadDocument(filenamepath)
            PdfViewer1.StdZoomType = tsButton_View.Tag
            PdfViewer1.Tag = filename  'used for default filename when user saves pdf to disk
        Catch e As Exception
            If Err.Number = 57 Then
                MsgBox("An error has occurred." & vbCr & "The PDF file '" & filename & "' in use already. Please close and try again.")
            Else
                MsgBox("Error in routine 'Form_MainNew.do_pdf' trying to generate PDF." & vbCrLf & Err.Description)
            End If
        End Try

    End Sub

    Private Function do_pdf_trend(patientID As Long) As PDFDocument

        Dim filename As String = cPt.Get_PtNameString(patientID, eNameStringFormats.NameForPdfFilename) & "_trend" & ".pdf"
        Dim filenamepath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & filename

        Try
            If FileSystem.FileExists(filenamepath) Then FileSystem.DeleteFile(filenamepath)
            Dim cPdf As New class_pdf
            Dim pdf As PDFDocument = cPdf.Draw_Pdf_trend(cPt.PatientID, grdTrend, chrt)
            Dim pdf_viewer As New PDFViewer
            pdf.Save(filenamepath)
            pdf_viewer.Visible = False
            pdf_viewer.LoadDocument(filenamepath)
            Return pdf
        Catch e As Exception
            If Err.Number = 57 Then
                MsgBox("An error has occurred." & vbCr & "The PDF file '" & filename & "' in use already. Please close and try again.")
            Else
                MsgBox("Error in routine 'Form_MainNew.do_pdf_trend' trying to generate PDF." & vbCrLf & Err.Description)
            End If
            Return Nothing
        End Try

    End Function

    Private Sub lvRecall_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRecall.Click

        If lvRecall.SelectedItems.Count = 0 Then
            MsgBox("No tests to display ...")
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim RecordID As Long = 0
            Dim Tbl As String = ""
            Dim testtime
            Dim testdate

            RecordID = CLng(lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(1).Text)
            Tbl = lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(2).Text
            Dim eTbl As eTables = [Enum].Parse(GetType(eTables), Tbl, True)
            testdate = lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(4).Text
            If IsDate(testdate) Then testdate = Replace(testdate, "/", "_") Else testdate = ""
            testtime = lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(5).Text
            If IsDate(testtime) Then testtime = Replace(testtime, ":", "") Else testtime = ""
            Me.Do_pdf_report(cPt.Get_PtNameString(cPt.PatientID, eNameStringFormats.NameForPdfFilename), testdate & "_" & testtime, RecordID, eTbl)

            lvRecall.Focus()
            tabReports.SelectTab(0)

            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub lvRecall_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRecall.DoubleClick

        Dim RecordID As Long = CLng(lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(1).Text)
        Dim tbl As String = lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(2).Text
        Me.Edit_test(RecordID, tbl)

    End Sub

    Private Sub Edit_test(ByVal RecordID As Long, ByVal tbl As String, Optional ByVal ChallengeType As String = "")

        If RecordID > 0 Then
            Dim f As Form = Nothing
            Select Case tbl.ToLower
                Case "rft_routine" : f = New form_Rft(RecordID, Me)
                Case "prov_test" : f = New form_rft_challenge(RecordID, Me)
                Case "r_walktests_v1heavy" : f = New form_walktest(RecordID, Me)
                Case "r_cpet" : f = New form_rft_cpet(RecordID, Me)
                Case "r_spt" : f = New form_spt(RecordID, Me)
                Case "r_hast" : f = New form_hast(RecordID, Me)
            End Select

            f.ShowDialog()
        End If

    End Sub

    Private Sub Edit_SelectedTest()

        If lvRecall.SelectedItems.Count = 1 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim RecordID As Long = CLng(lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(1).Text)
            Dim tbl As String = lvRecall.Items(lvRecall.SelectedIndices(0)).SubItems(2).Text
            Me.Edit_test(RecordID, tbl)
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub Form_MainNew_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        splitter.Height = Me.Height - tsMainMenu.Height - ToolStrip2.Height - sstrip.Height - 46
        lvRecall.Height = splitterPts.Panel1.Height - tsTests.Height - 6
        tabReports.Height = splitter.Panel1.Height - 20 - 6

    End Sub

    Private Sub ReportStylesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        form_customize.Show()

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsMenuItem_about.Click
        form_AboutBox.Show()
    End Sub

    Private Sub dtp_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Me.Get_Patients()
        lvRecall.Focus()

    End Sub

    Public Sub Add_ProvProtocolsToMainMenu(ByVal ts_menu As ToolStripMenuItem)

        Dim p() As class_challenge.ProtocolMenuData = cChall.Get_Protocols(True)

        If Not IsNothing(p) Then
            For i As Integer = 0 To UBound(p)
                Dim m As New ToolStripMenuItem
                m.Name = "tsMenuItem_" & p(i).Label
                m.Text = p(i).Label
                m.Tag = p(i).protocolID.ToString
                AddHandler m.Click, AddressOf FireUpChallForm
                ts_menu.DropDownItems.Add(m)
                m = Nothing
            Next
        End If

    End Sub

    Public Sub Add_WalkProtocolsToMainMenu(ByVal ts_menu As ToolStripMenuItem)

        Dim p() As class_walktest.ProtocolMenuData = cWalk.Get_ProtocolList(True)

        If Not IsNothing(p) Then
            For i As Integer = 0 To UBound(p)
                Dim m As New ToolStripMenuItem
                m.Name = "tsMenuItem_" & p(i).Label
                m.Text = p(i).Label
                m.Tag = p(i).protocolID.ToString
                m.Enabled = True
                AddHandler m.Click, AddressOf Me.FireUpWalkTestForm
                ts_menu.DropDownItems.Add(m)
                m = Nothing
            Next
        End If

    End Sub

    Public Sub FireUpChallForm(ByVal sender As Object, ByVal e As EventArgs)

        If cPt.PatientID > 0 Then
            Dim f As New form_rft_challenge(0, Me, CLng(sender.tag))
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Public Sub FireUpWalkTestForm(ByVal sender As Object, ByVal e As EventArgs)

        If cPt.PatientID > 0 Then
            Dim f As New form_walktest(0, Me, CLng(sender.tag))
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Private Sub toolbtnEditPt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolbtnEditPt.Click

        If cPt.PatientID > 0 Then
            Dim f As New form_Demographics(Me)
            f.Tag = cPt.PatientID
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Private Sub splitterPts_SplitterMoved(sender As System.Object, e As System.Windows.Forms.SplitterEventArgs) Handles splitterPts.SplitterMoved

        If Not formLoading Then
            lvPatients.Height = splitterPts.Panel2.Height - panelDates.Height - tsPtList.Height
            lvRecall.Height = splitterPts.Panel1.Height - tsTests.Height
        End If

    End Sub

    Private Sub splitterUnits_SplitterMoved(sender As Object, e As System.Windows.Forms.SplitterEventArgs) Handles splitterUnits.SplitterMoved

        If Not formLoading Then
            lvUnits.Height = splitterUnits.SplitterDistance - tsUnits.Height
        End If

    End Sub

    Private Sub dtp_CloseUp(sender As Object, e As System.EventArgs) Handles dtp.CloseUp

        dtp.Tag.text = Format(dtp.Value, "dd/MM/yyyy")
        dtp.Visible = False
        Me.Get_Patients()

    End Sub

    Private Sub ContextMenuStrip_date_ItemClicked(sender As Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip_date.ItemClicked

        Dim myCI As New CultureInfo("en-US")
        Dim Cal As Calendar = myCI.Calendar
        Dim daysFromMonday As Integer = Date.Today.DayOfWeek - DayOfWeek.Monday
        Dim MondayDate As Date = Date.Today.AddDays(-daysFromMonday)
        Dim SundayDate As Date = MondayDate.AddDays(6)
        Dim week As Integer = Cal.GetWeekOfYear(DateTime.Today, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)
        Dim thisMonthStart As Date = DateSerial(Today.Year, Today.Month, 1)
        Dim thisMonthEnd As Date = DateSerial(Today.Year, Today.Month + 1, 0)
        DateSerial(Today.Year, Today.Month, 0)
        Select Case UCase(e.ClickedItem.Text)
            Case "CLEAR DATE ENTRIES"
                txtStart.Text = ""
                txtEnd.Text = ""
                lvPatients.Clear()
            Case "TODAY"
                txtStart.Text = Date.Today.ToShortDateString
                txtEnd.Text = Date.Today.ToShortDateString
            Case "YESTERDAY"
                txtStart.Text = DateAdd(DateInterval.Day, -1, Date.Today).ToShortDateString
                txtEnd.Text = DateAdd(DateInterval.Day, -1, Date.Today).ToShortDateString
            Case "THIS WEEK"
                txtStart.Text = MondayDate.ToShortDateString
                txtEnd.Text = SundayDate.ToShortDateString
            Case "LAST WEEK"
                txtStart.Text = MondayDate.AddDays(-7).ToShortDateString
                txtEnd.Text = SundayDate.AddDays(-7).ToShortDateString
            Case "THIS MONTH"
                txtStart.Text = thisMonthStart.ToShortDateString
                txtEnd.Text = thisMonthEnd.ToShortDateString
            Case "LAST MONTH"
                txtStart.Text = thisMonthStart.AddMonths(-1).ToShortDateString
                txtEnd.Text = thisMonthEnd.AddMonths(-1).ToShortDateString
        End Select
        Me.Get_Patients()

    End Sub

    Private Sub txtStart_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtStart.MouseDown

        Select Case e.Button
            Case System.Windows.Forms.MouseButtons.Left
                dtp.Visible = True
                dtp.Tag = txtStart
                If IsDate(txtStart.Text) Then dtp.Value = txtStart.Text
                dtp.Focus()
                SendKeys.Send("{F4}")
            Case System.Windows.Forms.MouseButtons.Right

        End Select

    End Sub

    Private Sub txtEnd_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtEnd.MouseDown

        Select Case e.Button
            Case System.Windows.Forms.MouseButtons.Left
                dtp.Visible = True
                dtp.Tag = txtEnd
                If IsDate(txtEnd.Text) Then dtp.Value = txtEnd.Text
                dtp.Focus()
                SendKeys.Send("{F4}")
            Case System.Windows.Forms.MouseButtons.Right

        End Select

    End Sub

    Private Sub lvUnits_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lvUnits.SelectedIndexChanged

        If lvUnits.SelectedItems.Count = 1 Then
            tsLblUnits.Text = lvUnits.SelectedItems(0).SubItems(2).Text
            Me.ResetForm(Me.Get_CurrentUnit())
            lvRecall.Clear()
            lvPatients.Clear()
            Me.Get_Patients()
            Me.DisplayRecallItems(cPt.PatientID)
        End If

    End Sub

    Private Sub lvPatients_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles lvPatients.SelectedIndexChanged

        If lvPatients.SelectedItems.Count = 1 Then
            Dim PatientID As Long = Val(lvPatients.SelectedItems(0).SubItems(2).Text)
            If PatientID > 0 Then
                cPt.PatientID = PatientID
                DisplayRecallItems(PatientID)
                gRefreshMainForm = True
            End If
        End If

    End Sub

    Private Function Get_CurrentUnit() As eUnits

        If tsLblUnits.Text = "" Then
            Return Nothing
        Else
            Dim CurrentUnit As String = Replace(tsLblUnits.Text, " ", "_")
            Dim eCurrentUnit As eUnits = [Enum].Parse(GetType(eUnits), CurrentUnit)
            Return eCurrentUnit
        End If

    End Function

    Private Sub DisplayRecallItems(PatientID As Long)
        'also reset the report page view

        lvRecall.Clear()
        Select Case Me.Get_CurrentUnit()
            Case eUnits.Respiratory_Laboratory : Me.Get_Rfts(PatientID)
            Case eUnits.Sleep_Laboratory
            Case eUnits.O2_Therapy_Clinic
            Case eUnits.CPAP_Clinic
            Case eUnits.Victorian_Respiratory_Support_Service
        End Select

    End Sub

    Private Sub tstxt_UR_GotFocus(sender As Object, e As System.EventArgs) Handles tstxt_UR.GotFocus

        tstxt_UR.ForeColor = _pt_fore_colour_focus
        tstxt_UR.BackColor = _pt_back_colour_focus
        tstxt_UR.Tag = tstxt_UR.Text    're-apply if esc

    End Sub

    Private Sub tstxt_UR_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles tstxt_UR.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape : tstxt_UR.Text = tstxt_UR.Tag
            Case Keys.Return
                flagReturnKeyPressed = True
                Dim f As New form_Find("", "", tstxt_UR.Text)
        End Select

    End Sub

    Private Sub tstxt_UR_LostFocus(sender As Object, e As System.EventArgs) Handles tstxt_UR.LostFocus

        'Refresh the original
        tstxt_UR.ForeColor = _pt_fore_colour_nofocus
        tstxt_UR.BackColor = _pt_back_colour_nofocus
        tstxt_UR.Text = tstxt_UR.Tag

    End Sub


    Private Sub tstxt_UR_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles tstxt_UR.MouseDown

        tstxt_UR.SelectAll()

    End Sub

    Private Sub tstxt_PatientName_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tstxt_PatientName.Click

        tstxt_PatientName.SelectAll()

    End Sub

    Private Sub tstxt_PatientName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tstxt_PatientName.GotFocus

        tstxt_PatientName.ForeColor = _pt_fore_colour_focus
        tstxt_PatientName.BackColor = _pt_back_colour_focus
        tstxt_PatientName.Tag = tstxt_PatientName.Text
        tstxt_PatientName.SelectAll()

    End Sub

    Private Sub tstxt_PatientName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tstxt_PatientName.KeyDown

        Select Case e.KeyCode
            Case Keys.Return
                Dim Surname As String = ""
                Dim Firstname As String = ""
                Dim UR As String = ""
                flagReturnKeyPressed = True
                If InStr(tstxt_PatientName.Text, ",") = 0 Then
                    Surname = tstxt_PatientName.Text
                    Firstname = ""
                Else
                    Surname = Trim(Mid(tstxt_PatientName.Text, 1, InStr(tstxt_PatientName.Text, ",") - 1))
                    Firstname = Trim(Mid(tstxt_PatientName.Text, InStr(tstxt_PatientName.Text, ",") + 1))
                End If
                Dim f As New form_Find(Surname, Firstname, "")
            Case Else
                'Allow typing of text
        End Select

    End Sub

    Private Sub tstxt_PatientName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tstxt_PatientName.LostFocus

        'Refresh the original
        tstxt_PatientName.ForeColor = _pt_fore_colour_nofocus
        tstxt_PatientName.BackColor = _pt_back_colour_nofocus
        tstxt_PatientName.Text = tstxt_PatientName.Tag

    End Sub

    Private Sub tabReports_Deselecting(sender As Object, e As System.Windows.Forms.TabControlCancelEventArgs) Handles tabReports.Deselecting

        'Used to prevent rebuilding the trend grid when returning from options page
        tabReports.Tag = tabReports.TabPages(tabReports.SelectedIndex).Name

        'Check that manual scale options are valid
        Dim Msg As String = ""
        If chkAutoscale.Checked = False Then
            If Not IsNumeric(txtyaxis_min.Text) Then
                MsgBox("Please enter a valid value for Ymin", vbOKOnly, "Trend plot")
                txtyaxis_min.Focus()
                e.Cancel = True
            ElseIf Not IsNumeric(txtyaxis_max.Text) Then
                MsgBox("Please enter a valid value for Ymax", vbOKOnly, "Trend plot")
                txtyaxis_max.Focus()
                e.Cancel = True
            Else
                If Val(txtyaxis_max.Text) <= Val(txtyaxis_min.Text) Then
                    MsgBox("Ymax must be > Ymin", vbOKOnly, "Trend plot")
                    txtyaxis_max.Focus()
                    e.Cancel = True
                End If
            End If
        End If

    End Sub


    Private Sub tabReports_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles tabReports.SelectedIndexChanged

        If cPt.PatientID > 0 And lvRecall.Items.Count > 0 Then

            Dim p As class_plot_trend.plotproperties = Nothing

            'Build table if not done yet
            If splitTrend.Panel1.Controls.Count = 0 Then
                splitTrend.Panel1.Controls.Clear()
                grdTrend = cTrend.Create_trend_table(cPt.PatientID, Me.optPlot_pcChange.Checked)
                splitTrend.Panel1.Controls.Add(grdTrend)

                'Add dates to combo
                Dim dates(grdTrend.Rows(0).Cells.Count - 1) As String
                For i = 0 To grdTrend.Rows(0).Cells.Count - 1
                    dates(i) = grdTrend(i, 0).Value & ""
                Next
                cmboRefTestDate.Items.Clear()
                cmboRefTestDate.Items.AddRange(dates)
            End If

            Select Case tabReports.SelectedTab.Name.ToLower
                Case "tabpagetrend"
                    p = Me.get_trendplotproperties(grdTrend)
                    If tabReports.Tag.ToString.ToLower = "tabpagetrendoptions" Then
                        grdTrend.MultiSelect = p.GridCanMultiselect
                    End If
                    'Build chart
                    splitTrend.Panel2.Controls.Clear()
                    chrt = cTrend.Create_trend_plot(p)
                    splitTrend.Panel2.Controls.Add(chrt)
                Case "tabpagetrendoptions"
                    'MsgBox("Under construction")
            End Select
        Else
            'Clear
            splitTrend.Panel1.Controls.Clear()
            splitTrend.Panel2.Controls.Clear()
        End If

    End Sub

    Private Function Get_xDataToPlot(g As DataGridView) As Nullable(Of Date)()
        'Assumes the x data is in row 0 of the grid 

        If IsNothing(g) Then Return Nothing

        Dim x(g.Columns.Count - 1) As Nullable(Of Date)

        For i As Integer = 0 To g.ColumnCount - 1
            If IsDate(g(i, 0).Value) Then
                x(i) = CDate(g(i, 0).Value)
            Else
                MsgBox("Invalid date for test #" & i + 1 & ". Unable to plot this test", vbOKOnly, "Error")
                x(i) = Nothing
            End If
        Next

        Return x

    End Function

    Private Function Get_yDataToPlot(xRow As DataGridViewRow, ByVal GridRows As DataGridViewSelectedRowCollection) As Nullable(Of Double)(,)
        'Assumes the y data is in selected gridrow(s)

        If GridRows.Count = 0 Then
            Return Nothing
        Else
            If GridRows(0).Cells.Count = 0 Then
                Return Nothing
            End If
        End If

        Dim y(GridRows.Count - 1, GridRows(0).Cells.Count - 1) As Nullable(Of Double)
        For r = 0 To GridRows.Count - 1
            For i As Integer = 0 To xRow.Cells.Count - 1
                If IsDate(xRow.Cells(i).Value) And IsNumeric(GridRows(r).Cells(i).Value) Then
                    y(r, i) = CDbl(GridRows(r).Cells(i).Value)
                Else
                    y(r, i) = Nothing
                End If
            Next
        Next

        Return y

    End Function

    Private Function Get_yLabelsToPlot(ByVal GridRows As DataGridViewSelectedRowCollection) As String()
        'Assumes the y data is in selected gridrow(s)

        If GridRows.Count = 0 Then
            Return Nothing
        Else
            If GridRows(0).Cells.Count = 0 Then
                Return Nothing
            End If
        End If

        Dim lbls(GridRows.Count - 1) As String
        For i = 0 To GridRows.Count - 1
            lbls(i) = GridRows(i).HeaderCell.Value
        Next

        Return lbls

    End Function

    Private Sub tsUnits_Click(sender As Object, e As System.EventArgs) Handles tsUnits.Click

        If splitterUnits.SplitterDistance = SPLITTERCLOSEDDISTANCE Then
            splitterUnits.SplitterDistance = lvUnits.Items.Count * 31

        Else
            splitterUnits.SplitterDistance = SPLITTERCLOSEDDISTANCE
        End If

    End Sub

    Private Sub tsPtList_Click(sender As Object, e As System.EventArgs) Handles tsPtList.Click

        If splitterPts.SplitterDistance > splitterPts.Height * 0.9 Then
            splitterPts.SplitterDistance = splitterPts.Height / 2
        Else
            splitterPts.SplitterDistance = splitterPts.Height - tsPtList.Height
        End If
        lvPatients.Height = splitterPts.Panel2.Height - panelDates.Height - tsPtList.Height

    End Sub

    Private Sub tsMenuItem_EditTest_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_EditTest.Click

        Me.Edit_SelectedTest()

    End Sub

    Private Sub tsMenuItem_RoutineRft_Click(sender As Object, e As System.EventArgs) Handles tsMenuItem_RoutineRft.Click

        If cPt.PatientID > 0 Then
            Dim f As New form_Rft(0, Me)
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Private Sub tsMenuItem_cpx_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_cpx.Click

        If cPt.PatientID > 0 Then
            Dim f As New form_rft_cpet(0, Me)
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Private Sub tsMenuItem_spt_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_spt.Click

        If cPt.PatientID > 0 Then
            Dim f As New form_spt(0, Me)
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If

    End Sub

    Private Sub tsMenuItem_AltitudeSimulation_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_AltitudeSimulation.Click
        If cPt.PatientID > 0 Then
            Dim f As New form_hast(0, Me)
            f.Show()
        Else
            MsgBox("Please select a patient", vbOKOnly, "Problem detected")
        End If
    End Sub

    Private Sub toolbtnReporting_Click(sender As System.Object, e As System.EventArgs) Handles toolbtnReporting.Click

        form_Reporting.Show()

    End Sub

    Private Sub ListsToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ListsToolStripMenuItem1.Click

        form_prefs_list.Show()

    End Sub

    Private Sub PredictedValuesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PredictedValuesToolStripMenuItem.Click

        form_prefs_normals.Show()

    End Sub

    Private Sub PredictedValuesEditorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PredictedValuesEditorToolStripMenuItem.Click

        form_Pred.Show()

    End Sub

    Private Sub toolbtnBookings_Click(sender As System.Object, e As System.EventArgs) Handles toolbtnBookings.Click
        MsgBox("Not implemented", vbOKOnly, "ResLab (Lite)")

    End Sub

    Private Sub MigrateJjpToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MigrateJjpToolStripMenuItem.Click

        If form_password.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            form_migrate_jjp.Show()
        End If

    End Sub

    Private Sub tsMenuItem_SptPanelBuilder_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_SptPanelBuilder.Click

        Dim f As New form_spt_panelbuilder(False)
        f.ShowDialog()

    End Sub

    Private Sub tsMenuItem_ActivityReport_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_ActivityReport.Click

        Dim f As New form_reportsdb(tsMenuItem_ActivityReport.Text)
        f.Show()

    End Sub

    Private Sub tsMenuItem_peoplemanager_Click(sender As System.Object, e As System.EventArgs) Handles tsMenuItem_peoplemanager.Click

        Dim f As New form_user(0, Me)
        f.ShowDialog()

    End Sub

    Private Sub optPlot_pcChange_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optPlot_pcChange.CheckedChanged

        Select Case optPlot_pcChange.Checked
            Case True : panelChangeRef.Visible = True
            Case False : panelChangeRef.Visible = False
        End Select

    End Sub

    Private Sub chkAutoscale_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAutoscale.CheckedChanged

        panelScale.Visible = Not chkAutoscale.Checked

    End Sub

    Private Sub grdTrend_Click(sender As Object, e As System.EventArgs) Handles grdTrend.Click

        Dim r As Integer = 0
        Dim c As Integer = 0

        'Display report text
        r = grdTrend.CurrentCell.RowIndex
        c = grdTrend.CurrentCell.ColumnIndex
        If r > grdTrend.RowCount - 3 And c > 0 Then
            Select Case grdTrend.Rows(r).HeaderCell.Value.ToString
                Case "Technical note", "Report"
                    Dim Msg As String = "TECHNICAL NOTE:" & vbCrLf & grdTrend(c, grdTrend.Rows.Count - 2).Value
                    Msg = Msg & vbCrLf & vbCrLf & "REPORT:" & vbCrLf & grdTrend(c, grdTrend.Rows.Count - 1).Value
                    MsgBox(Msg, vbOKOnly, "Report: " & grdTrend(c, 0).Value)

            End Select
        End If

        gRefreshMainForm = False

    End Sub

    Private Sub grdTrend_RowHeaderMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles grdTrend.RowHeaderMouseClick

        If grdTrend.SelectedRows.Count = 0 Or grdTrend.Columns.Count = 0 Then
            Exit Sub
        End If

        If optPlot_pcChange.Checked = True Then
            If optPlot_selected.Checked = True Then
                If cmboRefTestDate.SelectedIndex = -1 Then
                    grdTrend.CurrentRow.Selected = False
                    Exit Sub
                Else
                    If Not IsNumeric(grdTrend(cmboRefTestDate.SelectedIndex, grdTrend.CurrentRow.Index).Value) Then
                        MsgBox("No valid reference value for " & grdTrend.CurrentRow.HeaderCell.Value & vbCrLf & "Can't be plotted as % change")
                        grdTrend.CurrentRow.Selected = False
                        Exit Sub
                    End If
                End If

            ElseIf optPlot_earliest.Checked = True Then
                If Not IsNumeric(grdTrend(grdTrend.Columns.Count - 1, grdTrend.CurrentRow.Index).Value) Then
                    MsgBox("No valid reference value for " & grdTrend.CurrentRow.HeaderCell.Value & vbCrLf & "Can't be plotted as % change")
                    grdTrend.CurrentRow.Selected = False
                    Exit Sub
                End If
            End If
        End If


        Dim r As Integer = grdTrend.SelectedRows(0).Index
        Select Case r
            Case Is < 3 : grdTrend.SelectedRows(0).Selected = False
            Case Is > grdTrend.RowCount - 3 : grdTrend.SelectedRows(0).Selected = False
            Case Else
                'Kill off any existing graph
                splitTrend.Panel2.Controls.Clear()
                'Generate and display graph
                chrt = cTrend.Create_trend_plot(Me.get_trendplotproperties(grdTrend))
                splitTrend.Panel2.Controls.Add(chrt)
        End Select


    End Sub


    Private Sub optPlot_selected_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optPlot_selected.CheckedChanged

        cmboRefTestDate.Visible = optPlot_selected.Checked

    End Sub


    Private Sub tsMainMenu_ItemClicked(sender As System.Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles tsMainMenu.ItemClicked

    End Sub

    Private Sub txtEnd_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtEnd.TextChanged

    End Sub
End Class



