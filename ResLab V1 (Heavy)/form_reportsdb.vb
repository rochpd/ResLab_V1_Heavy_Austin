Imports System.Text
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_reportsdb

    Private _reportName As String = ""
    Private _TestToFind As String = ""
    Dim _t As New List(Of String), _t1 As New List(Of String)

    Private Structure gridrowindexes
        Public patients_rowstart As Integer
        Public patients_rowcount As Integer
        Public tests_rowstart As Integer
        Public tests_rowcount As Integer
        Public reportscompleted_rowstart As Integer
        Public reportscompleted_rowcount As Integer
        Public requestsby_rowstart As Integer
        Public requestsby_rowcount As Integer
    End Structure

    Private Enum eListTypes
        ListByMonth
        PatientList
    End Enum

    Public Sub New(reportName As String)

        InitializeComponent()
        _reportName = reportName
        _t.AddRange({"Spirometry", "CO Transfer", "Lung Volumes", "MRPs", "ABGs", "Shunt", "Oximetry", "FeNO", "CPET", "Walk test", "Provocation test", "Altitude simulation", "Skin prick test"})
        _t1.AddRange({"sp", "tl", "lv", "mrp", "abg", "sh", "ox", "feno", "cpet", "walk", "prov", "hast", "spt"})

    End Sub

    Private Function Create_grid_activityreport(sDate As Date, fDate As Date) As DataGridView
        'Tag property contains the row numbers for sections etc

        Dim i As Integer = 0
        Dim g As New DataGridView

        'Get the months
        Dim DaysInMonth As Integer = Date.DaysInMonth(fDate.Year, fDate.Month)
        Dim sDate_new As Date = CDate("1/" & sDate.Month & "/" & sDate.Year)
        Dim fDate_new As Date = CDate(DaysInMonth & "/" & fDate.Month & "/" & fDate.Year)
        Dim nMonths As Integer = DateDiff(DateInterval.Month, sDate_new, fDate_new) + 1

        'Create columns
        Dim c As DataGridViewTextBoxColumn
        g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        g.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.ButtonFace
        g.RowHeadersVisible = False

        For i = 0 To nMonths + 1
            c = New DataGridViewTextBoxColumn
            Select Case i
                Case Is < 2
                    c.HeaderText = ""
                    c.Name = i.ToString
                    c.Width = 150
                Case Else
                    c.HeaderText = Format(DateAdd(DateInterval.Month, i - 2, sDate), "MMM yy")
                    c.Name = c.HeaderText
                    c.Width = 70
            End Select
            c.SortMode = DataGridViewColumnSortMode.NotSortable
            g.Columns.Add(c)
        Next
        g.Columns(0).DefaultCellStyle.BackColor = SystemColors.ButtonFace
        g.Columns(1).DefaultCellStyle.BackColor = SystemColors.ButtonFace

        'Create rows
        Dim nPatients, nTests, nRequesters
        Dim nRequests As Integer = 0
        Dim requests() As String = cMyRoutines.get_rfts_requested(sDate_new, fDate_new)
        Dim requesters() As String = cMyRoutines.get_RequestMOPermutations_rft(sDate_new, fDate_new)
        nPatients = 1
        nTests = _t.Count
        If IsNothing(requesters) Then nRequesters = 1 Else nRequesters = requesters.Count
        i = nPatients + nTests + nRequesters - 1
        g.Rows.Add(i)

        'Patients
        i = 0
        g(0, i).Value = "Patient test sessions"

        'Tests
        i = nPatients
        g(0, i).Value = "Tests performed"
        For Each item In _t
            g(1, i).Value = item
            i = i + 1
        Next

        'Tests requested by
        i = nPatients + nTests
        g(0, i).Value = "Tests requested by"
        If Not IsNothing(requesters) Then
            For Each re As String In requesters
                If re = "" Then re = "<No entry>"
                g(1, i).Value = re
                i = i + 1
            Next
        End If

        Dim gRows As gridrowindexes
        gRows.patients_rowstart = 0
        gRows.patients_rowcount = nPatients
        gRows.tests_rowstart = nPatients
        gRows.tests_rowcount = nTests
        gRows.requestsby_rowstart = nPatients + nTests
        gRows.requestsby_rowcount = nRequesters
        g.Tag = gRows

        Me.Fill_grid_activityreport(g, sDate, fDate)

        Return g

    End Function

    Private Function Get_gridrowindex_bycolvalue(g As DataGridView, txt As String, col As Integer) As Integer

        Dim i As Integer = 0

        For i = 0 To g.Rows.Count - 1
            If g(col, i).Value = txt Then Return i
        Next

        Return -1

    End Function

    Private Sub txtStart_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtStart.MouseDown

        Select Case e.Button
            Case System.Windows.Forms.MouseButtons.Left
                pnlDates.Tag = "start"
                dtp.Visible = True
                lblStart.Tag = txtStart.Text
                If IsDate(txtStart.Text) Then dtp.Value = txtStart.Text
                dtp.Focus()
                SendKeys.Send("{F4}")
            Case System.Windows.Forms.MouseButtons.Right

        End Select

    End Sub

    Private Sub txtEnd_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtEnd.MouseDown


        Select Case e.Button
            Case System.Windows.Forms.MouseButtons.Left
                pnlDates.Tag = "end"
                dtp.Visible = True
                lblEnd.Tag = txtEnd.Text
                If IsDate(txtEnd.Text) Then dtp.Value = txtEnd.Text
                dtp.Focus()
                SendKeys.Send("{F4}")
            Case System.Windows.Forms.MouseButtons.Right

        End Select

    End Sub

    Private Sub dtp_CloseUp(sender As Object, e As EventArgs) Handles dtp.CloseUp

        If optPatientList.Checked Then
            'dtp.Tag = Format(dtp.Value, "dd/MM/yyyy")
            Select Case pnlDates.Tag
                Case "start"
                    txtStart.Text = Format(dtp.Value, "dd/MM/yyyy")
                    lblStart.Tag = txtStart.Text
                Case "end"
                    txtEnd.Text = Format(dtp.Value, "dd/MM/yyyy")
                    lblEnd.Tag = txtEnd.Text
            End Select
        ElseIf optStatsByMonth.Checked Then
            'dtp.Tag = Format(dtp.Value, "MMM yyyy")
            Select Case pnlDates.Tag
                Case "start"
                    txtStart.Text = Format(dtp.Value, "MMM yyyy")
                    lblStart.Tag = txtStart.Text
                Case "end"
                    txtEnd.Text = Format(dtp.Value, "MMM yyyy")
                    lblEnd.Tag = txtEnd.Text
            End Select
        End If

        dtp.Visible = False

    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As System.EventArgs) Handles btnGenerate.Click


        If Not (IsDate(txtStart.Text) And IsDate(txtEnd.Text)) Then
            Exit Sub
        End If
        If CDate(txtStart.Text) > CDate(txtEnd.Text) Then
            MsgBox("Start date must be before end date")
            Exit Sub
        End If

        split.Panel2.Controls.Clear()

        Select Case tslbl_ReportTitle.Text.ToLower
            Case "activity report"
                Dim g As DataGridView = Me.Create_grid_activityreport(CDate(txtStart.Text), CDate(txtEnd.Text))
                g.Dock = DockStyle.Fill
                split.Panel2.Controls.Add(g)
        End Select

    End Sub

    Private Sub form_reportsdb_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        tslbl_ReportTitle.Text = _reportName

    End Sub

    Private Function FindTestInList(ByVal findthis As String, L As List(Of String)) As Integer

        Dim i As Integer = 0

        For i = 0 To L.Count - 1
            If L(i) = findthis Then
                Return i
            End If
        Next

        Return Nothing

    End Function


    Public Function Get_data_activityreport_tests(sDate As Date, eDate As Date, testtype As String) As DataSet

        Dim sql_test As String = ""
        Dim test As String = ""
        Dim testindex As Integer = 0
        Dim tbl As String = "", pk As String = ""
        Dim sql As New StringBuilder


        testindex = Me.FindTestInList(testtype, _t)
        test = _t1(testindex)
        Select Case testindex
            Case 0 To 7
                tbl = cDbInfo.table_name(eTables.rft_routine)
                pk = cDbInfo.primarykey(eTables.rft_routine)
            Case 8
                tbl = cDbInfo.table_name(eTables.r_cpet)
                pk = cDbInfo.primarykey(eTables.r_cpet)
            Case 9
                tbl = cDbInfo.table_name(eTables.r_walktests_v1heavy)
                pk = cDbInfo.primarykey(eTables.r_walktests_v1heavy)
            Case 10
                tbl = cDbInfo.table_name(eTables.Prov_test)
                pk = cDbInfo.primarykey(eTables.Prov_test)
            Case 11
                tbl = cDbInfo.table_name(eTables.r_hast)
                pk = cDbInfo.primarykey(eTables.r_hast)
            Case 12
                tbl = cDbInfo.table_name(eTables.r_spt)
                pk = cDbInfo.primarykey(eTables.r_spt)
        End Select

        sql.Clear()
        sql.Append("SELECT YEAR(testdate) AS yr, MONTH(testdate) AS mnth, COUNT(" & pk & ") AS nTests FROM " & tbl & " INNER JOIN r_sessions ON r_sessions.sessionID = " & tbl & ".sessionID ")
        sql.Append(" WHERE testdate >='" & Format(sDate, "yyyy-MM-dd") & "' AND testdate <='" & Format(eDate, "yyyy-MM-dd") & "' ")
        sql.Append("AND TestType LIKE '%" & test & "%'  ")
        sql.Append("GROUP BY YEAR(testdate), MONTH(testdate) ")
        sql.Append("ORDER BY yr DESC, mnth DESC")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count > 0 Then
                Return Ds
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Public Function Get_data_activityreport_requestsby(sDate As Date, eDate As Date) As DataSet

        Dim sql As New StringBuilder
        Dim eTbls() As eTables = {eTables.rft_routine, eTables.r_cpet, eTables.Prov_test, eTables.r_walktests_v1heavy, eTables.r_hast, eTables.r_spt}
        sql.Clear()
        sql.Append("SELECT t.Req_name, t.mnth, t.yr, sum(t.n) AS num FROM (")
        For i = 0 To UBound(eTbls)
            sql.Append("SELECT r_sessions.Req_name, YEAR(r_sessions.testdate) AS yr ,MONTH(r_sessions.testdate) AS mnth, COUNT(" & cDbInfo.primarykey(eTbls(i)) & ") AS n ")
            sql.Append("FROM r_sessions INNER JOIN " & cDbInfo.table_name(eTbls(i)) & " ON r_sessions.sessionid = " & cDbInfo.table_name(eTbls(i)) & ".sessionid ")
            sql.Append("WHERE (r_sessions.testdate BETWEEN '" & Format(sDate, "yyyy-MM-dd") & "' AND '" & Format(eDate, "yyyy-MM-dd") & "') ")
            sql.Append("GROUP BY Req_name, YEAR(testdate), MONTH(testdate)  ")
            If i < UBound(eTbls) Then sql.Append(" UNION ALL ")
        Next
        sql.Append(") AS t GROUP BY t.Req_name, t.yr,t.mnth  ORDER BY t.Req_name, t.yr DESC, t.mnth DESC")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count > 0 Then
                Return Ds
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Public Function Get_data_activityreport_patients(sDate As Date, eDate As Date) As DataSet

        Dim etbls() As eTables = {eTables.rft_routine, eTables.Prov_test, eTables.r_cpet, eTables.r_walktests_v1heavy, eTables.r_hast, eTables.r_spt}
        'Dim s_select As String = "SELECT patientid, testdate, YEAR(testdate) AS yr, MONTH(testdate) AS mnth FROM " & cDbInfo.Table_name(etbls(Integer)
        'Dim s_where As String = 


        Dim sql As New StringBuilder
        sql.Clear()
        sql.Append("SELECT YEAR(testdate) AS yr, MONTH(testdate) AS mnth, COUNT(patientid) AS countofpts FROM ")
        sql.Append("(SELECT DISTINCT patientid, testdate, YEAR(testdate) AS yr, MONTH(testdate) AS mnth FROM (")
        For i = 0 To UBound(etbls)
            sql.Append("SELECT r_sessions.patientid, testdate, YEAR(testdate) AS yr, MONTH(testdate) AS mnth ")
            sql.Append("FROM " & cDbInfo.table_name(etbls(i)) & " INNER JOIN r_sessions ON " & cDbInfo.table_name(etbls(i)) & ".sessionID = r_sessions.sessionID ")
            sql.Append(" WHERE testdate >='" & Format(sDate, "yyyy-MM-dd") & "' AND testdate <='" & Format(eDate, "yyyy-MM-dd") & "' ")
            If i < UBound(etbls) Then sql.Append(" UNION ALL ")
        Next
        sql.Append(") as p) as p1 ")
        sql.Append("GROUP BY YEAR(testdate), MONTH(testdate)  ")
        sql.Append("ORDER BY YEAR(testdate) DESC, MONTH(testdate) DESC ")
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count > 0 Then
                Return Ds
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Private Function Fill_grid_activityreport(g As DataGridView, sDate As Date, fDate As Date) As Boolean

        Dim colname As String = ""
        Dim rowIndex As Integer = 0
        Dim r As DataRow
        Dim rowinfo As gridrowindexes = g.Tag

        Dim Ds As DataSet
        'Do patients section
        Ds = Me.Get_data_activityreport_patients(sDate, fDate)
        If Not IsNothing(Ds) Then
            For Each r In Ds.Tables(0).Rows
                colname = Format(CDate("1/" & r("mnth") & "/" & r("yr")), "MMM yy")
                g(colname, rowinfo.patients_rowstart).Value = r("countofpts")
            Next
        End If

        'Do tests section
        For i = 0 To _t.Count - 1
            Ds = Me.Get_data_activityreport_tests(sDate, fDate, _t(i))
            If Not IsNothing(Ds) Then
                rowIndex = rowinfo.tests_rowstart + i
                For Each r In Ds.Tables(0).Rows
                    colname = Format(CDate("1/" & r("mnth") & "/" & r("yr")), "MMM yy")
                    g(colname, rowIndex).Value = r("ntests")
                Next
            End If
        Next

        'Do requests by
        Ds = Me.Get_data_activityreport_requestsby(sDate, fDate)
        If Not IsNothing(Ds) Then
            For Each r In Ds.Tables(0).Rows
                colname = Format(CDate("1/" & r("mnth") & "/" & r("yr")), "MMM yy")
                rowIndex = Get_gridrowindex_bycolvalue(g, r("req_name"), 1)
                g(colname, rowIndex).Value = r("num")
            Next
        End If

        Return True

    End Function



    Private Sub optStatsByMonth_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optStatsByMonth.CheckedChanged

        If optStatsByMonth.Checked Then
            dtp.CustomFormat = "MMM, yyyy"
            dtp.Format = DateTimePickerFormat.Custom
            Me.do_opt(eListTypes.ListByMonth)
        Else

        End If


    End Sub

    Private Sub do_opt(eType As eListTypes)


        Select Case eType
            Case eListTypes.ListByMonth
                txtStart.Mask = "??? ####"
                txtEnd.Mask = "??? ####"
                If IsDate(lblStart.Tag) Then txtStart.Text = Format(CDate(lblStart.Tag), "MMM yyyy")
                If IsDate(lblEnd.Tag) Then txtEnd.Text = Format(CDate(lblEnd.Tag), "MMM yyyy")


            Case eListTypes.PatientList
                txtStart.Mask = "0#/##/####"
                txtEnd.Mask = "0#/##/####"
                If IsDate(lblStart.Tag) Then txtStart.Text = Format(CDate(lblStart.Tag), "dd/MM/yyyy")
                If IsDate(lblEnd.Tag) Then txtEnd.Text = Format(CDate(lblEnd.Tag), "dd/MM/yyyy")

        End Select



    End Sub


    Private Sub optPatientList_CheckedChanged(sender As Object, e As System.EventArgs) Handles optPatientList.CheckedChanged

        If optPatientList.Checked Then Me.do_opt(eListTypes.PatientList)

    End Sub




End Class