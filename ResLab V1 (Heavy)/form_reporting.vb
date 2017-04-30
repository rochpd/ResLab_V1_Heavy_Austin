Imports System.Text
Imports Gnostice.PDFOne
Imports Gnostice.PDFOne.Windows.PDFViewer
Imports Microsoft.VisualBasic.FileIO
Imports System.Security.Permissions
Imports ResLab_V1_Heavy.cDatabaseInfo



Public Class form_Reporting

    Private _form_loading As Boolean
    Private _refreshReportingList As Boolean

    Private Sub Load_listview(ByRef lv As ListView, ByVal filter_reportingMO As String, ByVal filter_reportstatus As String)

        Dim eTbls() As eTables = {eTables.rft_routine, eTables.r_cpet, eTables.Prov_test, eTables.r_walktests_v1heavy, eTables.r_hast, eTables.r_spt}
        Dim i As Integer

        lv.Items.Clear()

        Dim s As String = ""
        Dim s1 As String = ""
        Dim sqlWhere As New StringBuilder

        If filter_reportingMO <> "" Then s = "req_name='" & Replace(filter_reportingMO, "'", "''") & "'"
        If filter_reportstatus = "" Then
            s1 = "report_status<>'Completed'"
        Else
            s1 = "report_status='" & filter_reportstatus & "'"
        End If


        If Len(s) > 0 And Len(s1) > 0 Then
            sqlWhere.Append(" WHERE " & s & " AND " & s1)
        ElseIf Len(s) > 0 Then
            sqlWhere.Append(" WHERE " & s)
        ElseIf Len(s1) > 0 Then
            sqlWhere.Append(" WHERE " & s1)
        End If

        Dim sql As New StringBuilder
        sql.Clear()
        For i = 0 To UBound(eTbls)
            sql.Append("SELECT patientdetails.PatientID, patientdetails.Surname, patientdetails.Firstname, patientdetails.UR, ")
            sql.Append(cDbInfo.table_name(eTbls(i)) & ".TestType, " & cDbInfo.table_name(eTbls(i)) & ".Report_Status AS Stat, " & cDbInfo.table_name(eTbls(i)) & ".TestTime, ")
            sql.Append(cDbInfo.table_name(eTbls(i)) & "." & cDbInfo.primarykey(eTbls(i)) & " AS RecordID, '" & cDbInfo.table_name(eTbls(i)) & "' AS tbl, ")
            sql.Append("r_sessions.testdate, r_sessions.req_name ")
            sql.Append("FROM patientdetails INNER JOIN " & cDbInfo.table_name(eTbls(i)) & " ON patientdetails.PatientID = " & cDbInfo.table_name(eTbls(i)) & ".PatientID ")
            sql.Append("INNER JOIN r_sessions ON " & cDbInfo.table_name(eTbls(i)) & ".SessionID = r_sessions.sessionID " & sqlWhere.ToString)
            If i < UBound(eTbls) Then sql.Append(" UNION ALL ")
        Next
        sql.Append(" ORDER BY TestDate")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        i = 1
        Dim item As ListViewItem
        If Ds IsNot Nothing Then
            With Ds.Tables(0)

                For Each record As DataRow In .Rows
                    item = New ListViewItem
                    item.SubItems.Clear()
                    Select Case record("tbl")
                        Case "rft_routine" : item.ImageKey = 0
                        Case "prov_test" : item.ImageKey = 1
                    End Select
                    item.SubItems.Add(record("tbl"))
                    item.SubItems.Add(record("RecordID"))
                    item.SubItems.Add(i)
                    item.SubItems.Add(Format(record("TestDate"), "dd/MM/yyyy"))
                    If IsDate(record("Testtime").ToString) Then item.SubItems.Add(Format(CDate(record("Testtime").ToString), "HH:mm")) Else item.SubItems.Add("")
                    item.SubItems.Add(record("Surname"))
                    item.SubItems.Add(record("Firstname") & "")
                    item.SubItems.Add(record("UR") & "")
                    item.SubItems.Add(record("Req_name") & "")
                    item.SubItems.Add(record("TestType") & "")
                    item.SubItems.Add(record("stat") & "")
                    item.SubItems.Add(record("patientID") & "")

                    lv.Items.Add(item)
                    item = Nothing
                    i = i + 1
                Next
            End With
        End If
        ss.Items.Item(0).Text = "Reports in list: " & lv.Items.Count

    End Sub

    Private Sub lv_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles lv.ColumnClick

        _refreshReportingList = False
        If Me.lv.Sorting = SortOrder.Ascending Then Me.lv.Sorting = SortOrder.Descending Else Me.lv.Sorting = SortOrder.Ascending
        Me.lv.ListViewItemSorter = New ListViewItemComparer(e.Column, lv.Sorting)

    End Sub




    Private Sub Setup_lv()

        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        lv.MultiSelect = True
        lv.HeaderStyle = ColumnHeaderStyle.Clickable
        lv.Sorting = SortOrder.Ascending

        ' Create columns for the items and subitems.
        lv.Columns.Add("", 0, HorizontalAlignment.Left)
        lv.Columns.Add("tbl", 0, HorizontalAlignment.Left)
        lv.Columns.Add("RecordID", 0, HorizontalAlignment.Left)
        lv.Columns.Add("#", 30, HorizontalAlignment.Left)
        lv.Columns.Add("Test Date", 100, HorizontalAlignment.Left)
        lv.Columns.Add("Time", 70, HorizontalAlignment.Left)
        lv.Columns.Add("Surname", 90, HorizontalAlignment.Left)
        lv.Columns.Add("Firstname", 90, HorizontalAlignment.Left)
        lv.Columns.Add("UR", 90, HorizontalAlignment.Left)
        lv.Columns.Add("Requesting MO", 170, HorizontalAlignment.Left)
        lv.Columns.Add("Test type", 220, HorizontalAlignment.Left)
        lv.Columns.Add("Report status", 150, HorizontalAlignment.Left)
        lv.Columns.Add("PatientID", 1, HorizontalAlignment.Left)


        'Add the items to the ListView.
        'lv.Items.AddRange(New ListViewItem() {item1, item2, item3, item4})

        ' Create two ImageList objects.
        'Dim imageListSmall As New ImageList()
        'Dim imageListLarge As New ImageList()

        '' Initialize the ImageList objects with bitmaps.
        'imageListSmall.Images.Add(Bitmap.FromFile("C:\Users\peter\Me\VB.NET applications\ResLab.Net\Resources\lungs.ico"))
        'imageListSmall.Images.Add(Bitmap.FromFile("C:\Users\peter\Me\VB.NET applications\ResLab.Net\Resources\mannitol.ico"))

        ''Assign the ImageList objects to the ListView.
        'lv.LargeImageList = imageListLarge
        'lv.SmallImageList = imageListSmall

    End Sub

    Private Sub form_Reporting_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If _form_loading Or _refreshReportingList Then
            Load_listview(lv, tsCmbo_requestingmo.Text, tsCmbo_reportstatus.Text)
        End If
        _form_loading = False
        _refreshReportingList = False

    End Sub

    Private Sub form_Reporting_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        gRefreshMainForm = True
    End Sub


    Private Sub form_Reporting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        _form_loading = True

        Setup_lv()

        'Load combos
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboDefaultReporter, "ReportingMO names", True)
        cMyRoutines.Combo_ts_LoadItemsFromList(tsCmbo_reportstatus, eTables.List_ReportStatuses, True)
        tsCmbo_reportstatus.Items.RemoveAt(tsCmbo_reportstatus.FindStringExact("Completed"))
        tsCmbo_reportstatus.Items.Add("")
        tsCmbo_reportstatus.SelectedItem = ""
        cMyRoutines.Combo_ts_LoadRequestingMOPermutations(tsCmbo_requestingmo)
        If tsCmbo_requestingmo.FindString("") = -1 Then tsCmbo_requestingmo.Items.Add("")
        tsCmbo_requestingmo.SelectedItem = ""

        txtDefaultDate.Text = Format(Now, "dd/MM/yyyy")

        _form_loading = False
        _refreshReportingList = True

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        gRefreshMainForm = True
        Me.Close()

    End Sub

    Private Sub tsCmbo_reportstatus_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsCmbo_reportstatus.TextChanged

        If Not _form_loading Then
            cMyRoutines.Combo_ts_LoadRequestingMOPermutations(tsCmbo_requestingmo)
            Load_listview(lv, tsCmbo_requestingmo.Text, tsCmbo_reportstatus.Text)
        End If

    End Sub

    Private Sub tsCmbo_reportingmo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsCmbo_requestingmo.TextChanged

        If Not _form_loading Then Load_listview(lv, tsCmbo_requestingmo.Text, tsCmbo_reportstatus.Text)

    End Sub



    Private Sub SelectTest(ByVal sender As Object, ByVal e As System.EventArgs) Handles lv.DoubleClick, btnSelect.Click

        If lv.SelectedItems.Count = 1 Then
            Dim RecordID As Long = CLng(lv.Items(lv.SelectedIndices(0)).SubItems(2).Text)
            Dim tbl As String = lv.Items(lv.SelectedIndices(0)).SubItems(1).Text
            Dim ChallengeType As String = lv.Items(lv.SelectedIndices(0)).SubItems(10).Text

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
                f.Show()
                _refreshReportingList = True
            End If
        End If

    End Sub

    Private Sub tsPrint_all_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsPrint_all.Click

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim p = New PrintDialog
        Dim cPrt As New class_Printing
        Dim cPDF As New class_pdf
        Dim returnVar As Boolean = True
        Dim ResultsID As Long
        Dim eTbl As eTables
        Dim patientID As Long = 0
        Dim updateReportstatus As Boolean = False
        Dim pdf As PDFDocument
        Dim v As New PDFViewer
        Dim filename As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\reslab_pdf.tmp"

        If lv.Items.Count > 0 Then
            If p.ShowDialog() = DialogResult.OK Then
                Dim Response As Integer = MsgBox("Update status to 'Completed'?", vbYesNoCancel, "Batch print")
                Select Case Response
                    Case vbCancel : GoTo Quit
                    Case vbYes : updateReportstatus = True
                    Case vbNo : updateReportstatus = False
                End Select

                For i As Integer = 0 To lv.Items.Count - 1
                    If FileSystem.FileExists(filename) Then FileSystem.DeleteFile(filename)
                    ResultsID = CLng(lv.Items(i).SubItems(2).Text)
                    eTbl = [Enum].Parse(GetType(eTables), lv.Items(i).SubItems(1).Text, True)
                    patientID = CLng(lv.Items(i).SubItems(12).Text)

                    If updateReportstatus Then
                        cRfts.Update_ReportStatus(eReportStatus.Completed, eTbl, ResultsID)
                    End If

                    pdf = cPDF.Draw_Pdf(patientID, ResultsID, eTbl)
                    pdf.Save(filename)
                    v.LoadDocument(filename)
                    cPrt.PrintMe(v.Document, 1)
                    v.CloseDocument()
                    ss.Items.Item(0).Text = "Printing " & i + 1 & " of " & lv.Items.Count
                    ss.Refresh()
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Next

                'Refresh list
                If updateReportstatus Then Load_listview(lv, tsCmbo_requestingmo.Text, tsCmbo_reportstatus.Text)

            End If
        End If

Quit:
        cPrt = Nothing
        cPDF = Nothing
        pdf = Nothing
        v = Nothing
        p.Dispose()
        ss.Items.Item(0).Text = "Printing completed."
        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub


    Private Sub tsSave_all_Click(sender As System.Object, e As System.EventArgs) Handles tsSave_all.Click

        Dim i As Integer = 0
        Dim pdf As PDFDocument
        Dim cPDF As New class_pdf
        Dim patientID As Long = 0
        Dim recordID As Long = 0
        Dim eTbl As eTables
        Dim filenamepluspath As String, filename As String, filepath As String
        Dim PtName As String = "", TestDate As String = "", TestTime As String = "", TestType As String = ""
        Dim updateReportstatus As Boolean = False

        'Do file dialog to get save directory
        filepath = cMyRoutines.Get_AppString("pdf_filepath")
        filepath = InputBox("Save reports as pdf files to folder", "Save reports as pdf files", filepath)
        If filepath = "" Then
            Exit Sub
        Else
            If Not FileSystem.DirectoryExists(filepath) Then
                MsgBox("Folder does not exist.", vbOKOnly, "Save to folder")
                Exit Sub
            End If
        End If

        Dim Response As Integer = MsgBox("Update status to 'Completed'?", vbYesNoCancel, "Batch print")
        Select Case Response
            Case vbCancel
            Case vbYes : updateReportstatus = True
            Case vbNo : updateReportstatus = False
        End Select

        If Response <> vbCancel Then
            'Loop through all reports
            For i = 0 To lv.Items.Count - 1

                'Get info from list
                recordID = CLng(lv.Items(i).SubItems(2).Text)
                eTbl = [Enum].Parse(GetType(eTables), lv.Items(i).SubItems(1).Text, True)
                patientID = CLng(lv.Items(i).SubItems(12).Text)
                PtName = cPt.Get_PtNameString(patientID, eNameStringFormats.NameForPdfFilename)
                TestDate = Replace(lv.Items(i).SubItems(4).Text, "/", "")
                TestType = lv.Items(i).SubItems(10).Text
                filename = "Rft_" & PtName & "_" & TestDate & "_" & i & "_" & TestType & ".pdf"
                filenamepluspath = filepath & "\" & filename

                If updateReportstatus Then
                    cRfts.Update_ReportStatus(eReportStatus.Completed, eTbl, recordID)
                End If

                'Delete any existing file if already exists
                If FileSystem.FileExists(filenamepluspath) Then FileSystem.DeleteFile(filenamepluspath)

                'Save to file
                pdf = cPDF.Draw_Pdf(patientID, recordID, eTbl)
                pdf.Save(filenamepluspath)
                pdf = Nothing
                ss.Items.Item(0).Text = "Saving " & i + 1 & " of " & lv.Items.Count
                ss.Refresh()
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Next
            ss.Items.Item(0).Text = "Saved " & i & " of " & lv.Items.Count & " as pdf files to '" & filepath & "'"

            'Save folder
            cMyRoutines.Update_AppString("pdf_filepath", filepath)

            'Refresh list
            If updateReportstatus Then Load_listview(lv, tsCmbo_requestingmo.Text, tsCmbo_reportstatus.Text)
        End If

    End Sub

    Private Sub Label2_DoubleClick(sender As Object, e As System.EventArgs) Handles Label2.DoubleClick

        If Not IsDate(txtDefaultDate.Text) Then txtDefaultDate.Text = Format(Today.Date, "dd/MM/yyyy")

    End Sub

End Class



