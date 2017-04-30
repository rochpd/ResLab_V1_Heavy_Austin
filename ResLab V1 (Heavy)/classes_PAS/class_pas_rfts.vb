Imports System
Imports System.Data
Imports System.Text
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Configuration
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_pas_rfts

    Private mRftID As Long


    Private Function Delete_cpet(cpetID As Long) As Boolean
        'Only called from me.delete_rft

        'Get trial IDs
        Dim i As Integer = 0
        Dim Success As Boolean
        Dim levelIDs() As Long = Me.Get_rft_cpet_levelids(cpetID)

        'First delete the level data record(s)
        If Not IsNothing(levelIDs) Then
            For i = 0 To UBound(levelIDs)
                Success = cDAL.Delete_Record(levelIDs(i), eTables.r_cpet, cDbInfo.primarykey(eTables.r_cpet), , True)
                If Not Success Then
                    MsgBox("Problem deleting a CPET level record", vbOKOnly, "Reslab")
                End If
            Next
        End If

        'Now delete the test data record
        If Not IsNothing(cpetID) Then
            Success = cDAL.Delete_Record(cpetID, eTables.r_cpet, cDbInfo.primarykey(eTables.r_cpet_levels), , True)
            If Not Success Then
                MsgBox("Problem deleting CPET test record", vbOKOnly, "Reslab")
            End If
        End If

        Return True

    End Function

    Private Function Delete_spt(sptID As Long) As Boolean
        'Only called from me.delete_rft

        'Get trial IDs
        Dim i As Integer = 0
        Dim Success As Boolean
        Dim IDs() As Long = Me.Get_rft_spt_allergenids(sptID)

        'First delete the allergen data record(s)
        If Not IsNothing(IDs) Then
            For i = 0 To UBound(IDs)
                Success = cDAL.Delete_Record(IDs(i), eTables.r_spt_allergens, cDbInfo.primarykey(eTables.r_spt), sptID, True)
                If Not Success Then
                    MsgBox("Problem deleting a SPT allergen record", vbOKOnly, "Reslab")
                End If
            Next
        End If

        'Now delete the test data record
        If Not IsNothing(sptID) Then
            Success = cDAL.Delete_Record(sptID, eTables.r_spt, , , True)
            If Not Success Then
                MsgBox("Problem deleting SPT test record", vbOKOnly, "Reslab")
            End If
        End If

        Return True

    End Function

    Private Function Delete_hast(hastID As Long) As Boolean
        'Only called from me.delete_rft

        'Get level IDs
        Dim i As Integer = 0
        Dim Success As Boolean
        Dim IDs() As Long = Me.Get_rft_hast_levelids(hastID)

        'First delete the allergen data record(s)
        If Not IsNothing(IDs) Then
            For i = 0 To UBound(IDs)
                Success = cDAL.Delete_Record(IDs(i), eTables.r_hast_levels, cDbInfo.primarykey(eTables.r_hast), hastID, True)
                If Not Success Then
                    MsgBox("Problem deleting a HAST level record", vbOKOnly, "Reslab")
                End If
            Next
        End If

        'Now delete the test data record
        If Not IsNothing(hastID) Then
            Success = cDAL.Delete_Record(hastID, eTables.r_hast, , , True)
            If Not Success Then
                MsgBox("Problem deleting HAST test record", vbOKOnly, "Reslab")
            End If
        End If

        Return True

    End Function

    Private Function Delete_prov(provID As Long) As Boolean
        'Only called from me.delete_rft

        'Get trial IDs
        Dim i As Integer = 0
        Dim Success As Boolean
        Dim levelIDs() As Long = Me.Get_rft_prov_levelids(provID)

        'First delete the level data record(s)
        If Not IsNothing(levelIDs) Then
            For i = 0 To UBound(levelIDs)
                Success = cDAL.Delete_Record(levelIDs(i), eTables.Prov_testdata, cDbInfo.primarykey(eTables.Prov_test), provID, True)
                If Not Success Then
                    MsgBox("Problem deleting a provocation level record", vbOKOnly, "Reslab")
                End If
            Next
        End If

        'Now delete the test data record
        If Not IsNothing(provID) Then
            Success = cDAL.Delete_Record(provID, eTables.Prov_test, , , True)
            If Not Success Then
                MsgBox("Problem deleting provocation test record", vbOKOnly, "Reslab")
            End If
        End If

        Return True

    End Function

    Public Function Delete_walktest(walkID As Long) As Boolean

        'Get trial IDs
        Dim i As Integer = 0
        Dim Success As Boolean
        Dim trialIDs() As Long = Me.Get_rft_walk_trialIDs(walkID)

        'First delete the timepoint data records
        If Not IsNothing(trialIDs) Then
            For i = 0 To UBound(trialIDs)
                Success = cDAL.Delete_Record(trialIDs(i), eTables.r_walktests_trials_levels, cDbInfo.primarykey(eTables.r_walktests_trials), , True)
                If Not Success Then
                    MsgBox("Problem deleting walk timepoint record", vbOKOnly, "Reslab")
                End If
            Next
        End If

        'Now delete the trial data record(s)
        If Not IsNothing(trialIDs) Then
            For i = 0 To UBound(trialIDs)
                Success = cDAL.Delete_Record(trialIDs(i), eTables.r_walktests_trials, cDbInfo.primarykey(eTables.r_walktests_v1heavy), , True)
                If Not Success Then
                    MsgBox("Problem deleting a walk trial record", vbOKOnly, "Reslab")
                End If
            Next
        End If

        'Now delete the test data record
        If Not IsNothing(walkID) Then
            Success = cDAL.Delete_Record(walkID, eTables.r_walktests_v1heavy, cDbInfo.primarykey(eTables.r_walktests_v1heavy), , True)
            If Not Success Then
                MsgBox("Problem deleting walk test record", vbOKOnly, "Reslab")
            End If
        End If

        Return True

    End Function

    Public Function Delete_rft(testID As Long, eTbl As eTables) As Boolean

        If testID = 0 Then
            MsgBox("Database error - invalid record ID (ID = " & testID & ")." & vbCrLf & "Operation aborted.", vbOKOnly, "Delete test record")
            Return False
        Else
            Dim Response As Integer = 0
            Response = MsgBox("You are about to delete this record. Continue?", vbYesNo, "Delete database record")
            If Response = vbNo Then
                Return False
            Else
                'Get sessionID
                Dim sessionID As Long = cRfts.Get_RftSessionIDfromRftTestID(testID, eTbl)

                'Delete test record(s)
                Select Case eTbl
                    Case eTables.rft_routine : cDAL.Delete_Record(testID, eTbl, , , True)
                    Case eTables.Prov_test : Me.Delete_prov(testID)
                    Case eTables.r_cpet : Me.Delete_cpet(testID)
                    Case eTables.r_walktests_v1heavy : Me.Delete_walktest(testID)
                    Case eTables.r_spt : Me.Delete_spt(testID)
                    Case eTables.r_hast : Me.Delete_hast(testID)
                    Case Else : Return False
                End Select
                MsgBox("Test deleted", vbOKOnly, "Delete test")

                'If no other tests attached to session, then delete it too.
                Dim d() As Dictionary(Of String, String) = cRfts.Get_rfts_all_bySessionID(sessionID, eSortMode.Ascending)
                If IsNothing(d) Then
                    cDAL.Delete_Record(sessionID, eTables.r_sessions, , , True)
                    MsgBox("Session deleted", vbOKOnly, "Delete test")
                End If
                Return True
            End If
        End If

    End Function

    Public Function Get_rfts_all_byPatientID(ByVal PatientID As Long, Optional ByVal sortorder As eSortMode = eSortMode.Descending) As Dictionary(Of String, String)()
        'Returns a set of tests with abbreviated data for recall 

        Dim i As Integer = 0
        Dim s As New StringBuilder
        Dim sSort As String = ""
        Dim tbl() As String = {cDbInfo.table_name(eTables.rft_routine), cDbInfo.table_name(eTables.Prov_test), cDbInfo.table_name(eTables.r_walktests_v1heavy), cDbInfo.table_name(eTables.r_cpet), cDbInfo.table_name(eTables.r_spt), cDbInfo.table_name(eTables.r_hast)}
        Dim pk() As String = {cDbInfo.primarykey(eTables.rft_routine), cDbInfo.primarykey(eTables.Prov_test), cDbInfo.primarykey(eTables.r_walktests_v1heavy), cDbInfo.primarykey(eTables.r_cpet), cDbInfo.primarykey(eTables.r_spt), cDbInfo.primarykey(eTables.r_hast)}

        Select Case sortorder
            Case eSortMode.Ascending : sSort = "ASC"
            Case eSortMode.Descending : sSort = "DESC"
        End Select
        s.Clear()
        For i = 0 To UBound(tbl)
            s.Append("SELECT r_sessions.sessionID, r_sessions.testdate, ")
            s.Append("      '" & tbl(i) & "' AS tbl, " & tbl(i) & ".TestTime, " & tbl(i) & ".TestType, " & tbl(i) & ".Report_status, " & tbl(i) & "." & pk(i) & " AS ID ")
            s.Append("FROM r_sessions INNER JOIN " & tbl(i) & " ON r_sessions.sessionID = " & tbl(i) & ".SessionID ")
            s.Append("WHERE " & tbl(i) & ".PatientID=" & PatientID)
            If i < UBound(tbl) Then s.Append(" UNION ALL ")
        Next i
        s.Append(" ORDER BY TestDate " & sSort & ", testtime " & sSort)

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(s.ToString)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Dim dicR() As Dictionary(Of String, String) = Nothing
            With Ds.Tables(0)
                i = 0
                For Each record As DataRow In .Rows
                    ReDim Preserve dicR(i)
                    dicR(i) = New Dictionary(Of String, String)
                    dicR(i).Add("ID", record("ID"))
                    dicR(i).Add("TestDate", record("TestDate"))
                    If IsDBNull(record("TestTime")) Then dicR(i).Add("TestTime", "") Else dicR(i).Add("TestTime", record("TestTime").ToString)
                    dicR(i).Add("TestType", record("TestType") & "")
                    dicR(i).Add("Report_status", record("Report_status") & "")
                    dicR(i).Add("tbl", record("tbl"))
                    i = i + 1
                Next
            End With
            Return dicR
        End If

    End Function

    Public Function Get_rfts_all_bySessionID(ByVal sessionID As Long, Optional ByVal sortorder As eSortMode = eSortMode.Descending) As Dictionary(Of String, String)()
        'Returns the set of tests with abbreviated data  

        Dim s As New StringBuilder
        Dim sSort As String = ""
        Dim i As Integer = 0

        Select Case sortorder
            Case eSortMode.Ascending : sSort = "ASC"
            Case eSortMode.Descending : sSort = "DESC"
        End Select

        Dim eTbls() As eTables = {eTables.rft_routine, eTables.r_cpet, eTables.Prov_test, eTables.r_walktests_v1heavy, eTables.r_hast, eTables.r_spt}
        Dim sql As New StringBuilder
        sql.Clear()
        For i = 0 To UBound(eTbls)
            s.Append("SELECT '" & cDbInfo.table_name(eTbls(i)) & "' AS tbl, " & cDbInfo.table_name(eTbls(i)) & ".TestTime, " & cDbInfo.table_name(eTbls(i)) & ".TestType, " & cDbInfo.table_name(eTbls(i)) & ".Report_status, " & cDbInfo.table_name(eTbls(i)) & "." & cDbInfo.primarykey(eTbls(i)) & " AS ID, ")
            s.Append("r_sessions.sessionID, r_sessions.testdate ")
            s.Append("FROM r_sessions INNER JOIN " & cDbInfo.table_name(eTbls(i)) & " ON r_sessions.sessionID = " & cDbInfo.table_name(eTbls(i)) & ".SessionID")
            s.Append(" WHERE " & cDbInfo.table_name(eTbls(i)) & ".sessionID=" & sessionID)
            If i < UBound(eTbls) Then s.Append(" UNION ALL ")
        Next i
        s.Append(" ORDER BY TestDate " & sSort & ", testtime " & sSort)

        Dim dicR(0) As Dictionary(Of String, String)
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(s.ToString)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)
                For Each record As DataRow In .Rows
                    ReDim Preserve dicR(i)
                    dicR(i) = New Dictionary(Of String, String)
                    dicR(i).Add("ID", record("ID"))
                    dicR(i).Add("TestDate", record("TestDate"))
                    If IsDBNull(record("TestTime")) Then dicR(i).Add("TestTime", "") Else dicR(i).Add("TestTime", record("TestTime").ToString)
                    dicR(i).Add("TestType", record("TestType") & "")
                    dicR(i).Add("Report_status", record("Report_status") & "")
                    dicR(i).Add("tbl", record("tbl"))
                    i = i + 1
                Next
            End With
            Return dicR
        End If

    End Function

    Public Function Get_rfts_ForTrend(ByVal PatientID As Long, Optional ByVal sortorder As eSortMode = eSortMode.Descending) As Dictionary(Of String, String)()
        'Returns relevant test types with abbreviated data for trend table/plot

        Dim s As New StringBuilder
        Dim sSort As String = ""

        Select Case sortorder
            Case eSortMode.Ascending : sSort = "ASC"
            Case eSortMode.Descending : sSort = "DESC"
        End Select
        s.Clear()
        s.Append("SELECT rft_routine.rftID AS ID, r_sessions.testdate, rft_routine.testtime, rft_routine.testtype, rft_routine.report_status, 'rft_routine' as tbl, r_sessions.height, r_sessions.weight ")
        s.Append("FROM rft_routine INNER JOIN r_sessions ON rft_routine.sessionID = r_sessions.sessionID WHERE r_sessions.patientID=" & PatientID)
        s.Append(" UNION ALL ")
        s.Append("SELECT prov_test.provID AS ID, r_sessions.testdate, prov_test.testtime, prov_test.testtype, prov_test.report_status, 'prov_test' as tbl, r_sessions.height, r_sessions.weight ")
        s.Append("FROM prov_test INNER JOIN r_sessions ON prov_test.sessionID = r_sessions.sessionID WHERE r_sessions.patientID=" & PatientID)
        s.Append(" ORDER BY TestDate " & sSort & ", testtime " & sSort)

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(s.ToString)
        Dim dicR() As Dictionary(Of String, String) = Nothing
        Dim i As Integer = 0

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)
                For Each record As DataRow In .Rows
                    ReDim Preserve dicR(i)
                    dicR(i) = New Dictionary(Of String, String)
                    dicR(i).Add("ID", record("ID"))
                    dicR(i).Add("testdate", record("testdate"))
                    If IsDBNull(record("testtime")) Then dicR(i).Add("testtime", "") Else dicR(i).Add("testtime", record("testtime").ToString)
                    dicR(i).Add("testtype", record("testtype") & "")
                    dicR(i).Add("report_status", record("report_status") & "")
                    dicR(i).Add("height", record("height") & "")
                    dicR(i).Add("weight", record("weight") & "")
                    dicR(i).Add("tbl", record("tbl"))
                    i = i + 1
                Next
            End With
            Return dicR
        End If

    End Function

    Public Function Get_RftSessionIDfromRftTestID(ByVal testID As Long, eTbl As eTables) As Long

        Dim sql As String = "SELECT sessionID FROM " & eTbl.ToString & " WHERE " & cDbInfo.primarykey(eTbl) & "=" & testID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Dim sessionid As Long = Ds.Tables(0).Rows(0)("sessionid")
            Return sessionid
        End If

    End Function

    Public Function Get_RftSessionIDforDate(patientID As Long, testdate As Date) As Long

        Dim sql As String = "SELECT sessionID FROM r_sessions WHERE patientID=" & patientID & " AND TestDate='" & Format(testdate, "yyyy-MM-dd") & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Dim sessionid As Long = Ds.Tables(0).Rows(0)("sessionid")
            Return sessionid
        End If

    End Function

    Public Function get_rft_session(sessionID As Long) As Dictionary(Of String, String)

        Dim sql As String = "SELECT * FROM r_sessions WHERE sessionID=" & sessionID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_Session()
        Dim f As New class_Rft_RoutineAndSessionFields

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)
                d(f.PatientID) = .Rows(0)(f.PatientID)
                d(f.SessionID) = .Rows(0)(f.SessionID)
                If IsDate(.Rows(0)(f.TestDate)) Then d(f.TestDate) = .Rows(0)(f.TestDate) Else d(f.TestDate) = Nothing
                d(f.Height) = .Rows(0)(f.Height) & ""
                d(f.Weight) = .Rows(0)(f.Weight) & ""
                d(f.Smoke_hx) = .Rows(0)(f.Smoke_hx) & ""
                d(f.Smoke_yearssmoked) = .Rows(0)(f.Smoke_yearssmoked) & ""
                d(f.Smoke_cigsperday) = .Rows(0)(f.Smoke_cigsperday) & ""
                d(f.Smoke_packyears) = .Rows(0)(f.Smoke_packyears) & ""
                If IsDate(.Rows(0)(f.Req_date)) Then d(f.Req_date) = .Rows(0)(f.Req_date) Else d(f.Req_date) = Nothing
                d(f.Req_name) = .Rows(0)(f.Req_name) & ""
                d(f.Req_providernumber) = .Rows(0)(f.Req_providernumber) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_email) = .Rows(0)(f.Req_email) & ""
                d(f.Req_fax) = .Rows(0)(f.Req_fax) & ""
                d(f.Req_phone) = .Rows(0)(f.Req_phone) & ""
                d(f.Report_copyto) = .Rows(0)(f.Report_copyto) & ""
                d(f.Req_clinicalnotes) = .Rows(0)(f.Req_clinicalnotes) & ""
                d(f.Req_healthservice) = .Rows(0)(f.Req_healthservice) & ""
                d(f.AdmissionStatus) = .Rows(0)(f.AdmissionStatus) & ""
                d(f.Billing_billedto) = .Rows(0)(f.Billing_billedto) & ""
                d(f.Billing_billingMO) = .Rows(0)(f.Billing_billingMO) & ""
                d(f.Billing_billingMOproviderno) = .Rows(0)(f.Billing_billingMOproviderno) & ""
                If IsDate(.Rows(0)(f.LastUpdated_session)) Then d(f.LastUpdated_session) = .Rows(0)(f.LastUpdated_session) Else d(f.LastUpdated_session) = ""
                d(f.LastUpdatedBy_session) = .Rows(0)(f.LastUpdatedBy_session) & ""
            End With
        End If

        Return d

    End Function

    Public Function Get_rft_byRftID(ByVal RftID As Long) As Dictionary(Of String, String)

        Dim sql As String = "SELECT rft_routine.*, r_sessions.* FROM rft_routine INNER JOIN r_sessions ON rft_routine.sessionID=r_sessions.sessionID WHERE rft_routine.rftID =" & RftID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        Dim dicR As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicRft_RoutineAndSession()

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)
                Dim r As New class_Rft_RoutineAndSessionFields
                dicR(r.PatientID) = .Rows(0)(r.PatientID)
                dicR(r.RftID) = .Rows(0)(r.RftID)
                If IsDate(.Rows(0)(r.TestDate)) Then dicR(r.TestDate) = .Rows(0)(r.TestDate) Else dicR(r.TestDate) = ""
                If IsNumeric(.Rows(0)(r.Height) & "") Then dicR(r.Height) = Math.Round(Val(.Rows(0)(r.Height)), 1) Else dicR(r.Height) = ""
                If IsNumeric(.Rows(0)(r.Weight) & "") Then dicR(r.Weight) = Math.Round(Val(.Rows(0)(r.Weight)), 1) Else dicR(r.Weight) = ""
                dicR(r.Smoke_hx) = .Rows(0)(r.Smoke_hx) & ""
                dicR(r.Smoke_packyears) = .Rows(0)(r.Smoke_packyears) & ""
                dicR(r.Smoke_yearssmoked) = .Rows(0)(r.Smoke_yearssmoked) & ""
                dicR(r.Smoke_cigsperday) = .Rows(0)(r.Smoke_cigsperday) & ""
                If IsDate(.Rows(0)(r.Req_date)) Then dicR(r.Req_date) = .Rows(0)(r.Req_date) Else dicR(r.Req_date) = Nothing
                If IsDate(.Rows(0)(r.Req_time)) Then dicR(r.Req_time) = .Rows(0)(r.Req_time) Else dicR(r.Req_time) = Nothing
                dicR(r.Req_name) = .Rows(0)(r.Req_name) & ""
                dicR(r.Req_address) = .Rows(0)(r.Req_address) & ""
                dicR(r.Req_providernumber) = .Rows(0)(r.Req_providernumber) & ""
                dicR(r.Req_address) = .Rows(0)(r.Req_address) & ""
                dicR(r.Req_email) = .Rows(0)(r.Req_email) & ""
                dicR(r.Req_fax) = .Rows(0)(r.Req_fax) & ""
                dicR(r.Req_phone) = .Rows(0)(r.Req_phone) & ""
                dicR(r.Report_copyto) = .Rows(0)(r.Report_copyto) & ""
                dicR(r.Req_clinicalnotes) = .Rows(0)(r.Req_clinicalnotes) & ""
                dicR(r.Req_healthservice) = .Rows(0)(r.Req_healthservice) & ""
                dicR(r.AdmissionStatus) = .Rows(0)(r.AdmissionStatus) & ""
                dicR(r.BDStatus) = .Rows(0)(r.BDStatus) & ""
                dicR(r.Billing_billedto) = .Rows(0)(r.Billing_billedto) & ""
                dicR(r.Billing_billingMO) = .Rows(0)(r.Billing_billingMO) & ""
                dicR(r.Billing_billingMOproviderno) = .Rows(0)(r.Billing_billingMOproviderno) & ""

                dicR(r.TechnicalNotes) = .Rows(0)(r.TechnicalNotes) & ""
                dicR(r.Report_text) = .Rows(0)(r.Report_text) & ""
                dicR(r.Report_status) = .Rows(0)(r.Report_status) & ""
                dicR(r.Report_reportedby) = .Rows(0)(r.Report_reportedby) & ""
                If IsDate(.Rows(0)(r.Report_reporteddate)) Then dicR(r.Report_reporteddate) = .Rows(0)(r.Report_reporteddate) Else dicR(r.Report_reporteddate) = ""
                dicR(r.Report_verifiedby) = .Rows(0)(r.Report_verifiedby) & ""
                If IsDate(.Rows(0)(r.Report_verifieddate)) Then dicR(r.Report_verifieddate) = .Rows(0)(r.Report_verifieddate) Else dicR(r.Report_verifieddate) = ""
                If IsDate(.Rows(0)(r.TestTime).ToString) Then dicR(r.TestTime) = .Rows(0)(r.TestTime).ToString Else dicR(r.TestTime) = ""
                dicR(r.TestType) = .Rows(0)(r.TestType) & ""
                dicR(r.Lab) = .Rows(0)(r.Lab) & ""
                dicR(r.Scientist) = .Rows(0)(r.Scientist) & ""

                dicR(r.LungVolumes_method) = .Rows(0)(r.LungVolumes_method) & ""
                dicR(r.R_bl_Fev1) = .Rows(0)(r.R_bl_Fev1) & ""
                dicR(r.R_bl_Fvc) = .Rows(0)(r.R_bl_Fvc) & ""
                dicR(r.R_bl_Vc) = .Rows(0)(r.R_bl_Vc) & ""
                dicR(r.R_bl_Fef2575) = .Rows(0)(r.R_bl_Fef2575) & ""
                dicR(r.R_bl_Pef) = .Rows(0)(r.R_bl_Pef) & ""
                dicR(r.R_Bl_Tlco) = .Rows(0)(r.R_Bl_Tlco) & ""
                dicR(r.R_Bl_Kco) = .Rows(0)(r.R_Bl_Kco) & ""
                dicR(r.R_Bl_Va) = .Rows(0)(r.R_Bl_Va) & ""
                dicR(r.R_Bl_Hb) = .Rows(0)(r.R_Bl_Hb) & ""
                dicR(r.R_Bl_Ivc) = .Rows(0)(r.R_Bl_Ivc) & ""
                dicR(r.R_Bl_Tlc) = .Rows(0)(r.R_Bl_Tlc) & ""
                dicR(r.R_Bl_Frc) = .Rows(0)(r.R_Bl_Frc) & ""
                dicR(r.R_Bl_Rv) = .Rows(0)(r.R_Bl_Rv) & ""
                dicR(r.R_Bl_LvVc) = .Rows(0)(r.R_Bl_LvVc) & ""
                dicR(r.R_Bl_Mip) = .Rows(0)(r.R_Bl_Mip) & ""
                dicR(r.R_Bl_Mep) = .Rows(0)(r.R_Bl_Mep) & ""
                dicR(r.R_SpO2_1) = .Rows(0)(r.R_SpO2_1) & ""
                dicR(r.R_SpO2_2) = .Rows(0)(r.R_SpO2_2) & ""
                dicR(r.R_Bl_FeNO) = .Rows(0)(r.R_Bl_FeNO) & ""
                dicR(r.R_post_Fev1) = .Rows(0)(r.R_post_Fev1) & ""
                dicR(r.R_post_Fvc) = .Rows(0)(r.R_post_Fvc) & ""
                dicR(r.R_post_Vc) = .Rows(0)(r.R_post_Vc) & ""
                dicR(r.R_post_Fef2575) = .Rows(0)(r.R_post_Fef2575) & ""
                dicR(r.R_post_Pef) = .Rows(0)(r.R_post_Pef) & ""
                dicR(r.R_post_Condition) = .Rows(0)(r.R_post_Condition) & ""

                dicR(r.R_abg1_aapo2) = .Rows(0)(r.R_abg1_aapo2) & ""
                dicR(r.R_abg1_be) = .Rows(0)(r.R_abg1_be) & ""
                dicR(r.R_abg1_fio2) = .Rows(0)(r.R_abg1_fio2) & ""
                dicR(r.R_abg1_hco3) = .Rows(0)(r.R_abg1_hco3) & ""
                dicR(r.R_abg1_paco2) = .Rows(0)(r.R_abg1_paco2) & ""
                dicR(r.R_abg1_pao2) = .Rows(0)(r.R_abg1_pao2) & ""
                dicR(r.R_abg1_ph) = .Rows(0)(r.R_abg1_ph) & ""
                dicR(r.R_abg1_sao2) = .Rows(0)(r.R_abg1_sao2) & ""
                dicR(r.R_abg1_shunt) = .Rows(0)(r.R_abg1_shunt) & ""

                dicR(r.R_abg2_aapo2) = .Rows(0)(r.R_abg2_aapo2) & ""
                dicR(r.R_abg2_be) = .Rows(0)(r.R_abg2_be) & ""
                dicR(r.R_abg2_fio2) = .Rows(0)(r.R_abg2_fio2) & ""
                dicR(r.R_abg2_hco3) = .Rows(0)(r.R_abg2_hco3) & ""
                dicR(r.R_abg2_paco2) = .Rows(0)(r.R_abg2_paco2) & ""
                dicR(r.R_abg2_pao2) = .Rows(0)(r.R_abg2_pao2) & ""
                dicR(r.R_abg2_ph) = .Rows(0)(r.R_abg2_ph) & ""
                dicR(r.R_abg2_sao2) = .Rows(0)(r.R_abg2_sao2) & ""
                dicR(r.R_abg2_shunt) = .Rows(0)(r.R_abg2_shunt) & ""

                dicR(r.Pred_SourceIDs) = .Rows(0)(r.Pred_SourceIDs) & ""
            End With
        End If

        Return dicR

    End Function

    Public Function Get_LastHtEtc(PatientID As Long) As Dictionary(Of String, String)
        'Allows retrieval of Ht, smoking hx and requested by details from last test when new test created

        Dim dic() As Dictionary(Of String, String) = Me.Get_rfts_all_byPatientID(PatientID, eSortMode.Descending)
        If IsNothing(dic) Then
            Return Nothing
        Else
            Return dic(0)
        End If

    End Function

    Public Function Get_MostRecentRftSessionID(PatientID As Long) As Long
        'Allows use of Ht, smoking hx and requested by details from the latest previous session (if exists) when new test session created

        Dim sessionID As Long = 0
        Dim sql As String = "SELECT sessionID FROM r_sessions WHERE patientID=" & PatientID & " ORDER BY testdate DESC"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return 0
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return 0
        Else
            Return Ds.Tables(0).Rows(0).Item(0)
        End If

    End Function

    Public Function Get_ProtocolID_FromProvID(ByVal ProvID As Long) As Long

        Try
            Dim sql As String = "SELECT protocolid FROM prov_test WHERE provid=" & ProvID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

            If IsNothing(Ds) Then
                Return 0
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return 0
            Else
                Return Ds.Tables(0).Rows(0)("protocolid")
            End If
        Catch ex As Exception
            MsgBox("Error in class_patient.Get_Prov_VisitDataByProvID" & vbNewLine & ex.Message.ToString)
            Return -1
        End Try

    End Function

    Public Function Get_ProtocolID_FromWalkID(ByVal WalkID As Long) As Long

        Try
            Dim sql = "SELECT protocolid FROM r_walktests_v1heavy WHERE WalkID=" & WalkID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

            If IsNothing(Ds) Then
                Return 0
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return 0
            Else
                Return Ds.Tables(0).Rows(0)("protocolid")
            End If
        Catch ex As Exception
            MsgBox("Error in class_patient.Get_ProtocolID_FromWalkID" & vbNewLine & ex.Message.ToString)
            Return -1
        End Try

    End Function

    Public Function Get_walk_test_session(ByVal walkID As Long) As Dictionary(Of String, String)

        Try
            Dim sql As String = "SELECT r_walktests_v1heavy.*, r_sessions.* FROM r_walktests_v1heavy INNER JOIN r_sessions ON r_walktests_v1heavy.sessionID = r_sessions.sessionID WHERE r_walktests_v1heavy.walkID =" & walkID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            Dim j As Integer = 0
            Dim row As DataRow

            If IsNothing(Ds) Then
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicWalk_test_session
                row = Ds.Tables(0).Rows(0)
                For j = 0 To Ds.Tables(0).Columns.Count - 1
                    If d.ContainsKey(Ds.Tables(0).Columns(j).ColumnName.ToLower) Then
                        If Not IsDBNull(row(Ds.Tables(0).Columns(j).ColumnName)) Then
                            Select Case Ds.Tables(0).Columns(j).DataType.FullName
                                Case "System.TimeSpan" : d(Ds.Tables(0).Columns(j).ColumnName.ToLower) = row(Ds.Tables(0).Columns(j).ColumnName).ToString
                                Case Else : d(Ds.Tables(0).Columns(j).ColumnName.ToLower) = row(Ds.Tables(0).Columns(j).ColumnName)
                            End Select
                        Else
                            d(Ds.Tables(0).Columns(j).ColumnName.ToLower) = ""
                        End If
                    End If
                Next
                Return d
            End If

        Catch ex As Exception
            MsgBox("Error in class_patient.Get_Walk_Test_Session" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Public Function Get_walk_trials(ByVal walkID As Long) As Dictionary(Of String, String)()

        Dim d() As Dictionary(Of String, String) = Nothing
        Dim nTrial As Integer = 0

        Try
            Dim Sql As String = "SELECT * FROM r_walktests_trials WHERE WalkID=" & walkID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(Sql)

            If IsNothing(Ds) Then
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim flds As New class_fields_Walk_Trial
                nTrial = 0
                For Each t In Ds.Tables(0).Rows
                    ReDim Preserve d(nTrial)
                    d(nTrial) = New Dictionary(Of String, String)
                    d(nTrial)(flds.trialID) = t(flds.trialID)
                    d(nTrial)(flds.trial_number) = t(flds.trial_number) & ""
                    d(nTrial)(flds.trial_label) = t(flds.trial_label) & ""
                    d(nTrial)(flds.trial_distance) = t(flds.trial_distance) & ""
                    If Not IsDBNull(t(flds.trial_timeofday)) Then d(nTrial)(flds.trial_timeofday) = t(flds.trial_timeofday).ToString Else d(nTrial)(flds.trial_timeofday) = Nothing
                    nTrial = nTrial + 1
                Next t
                Return d
            End If

        Catch ex As Exception
            MsgBox("Error in class_patient.Get_walk_trials" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Public Function Get_walk_levels(ByVal walkID As Long) As List(Of Dictionary(Of String, String)())

        Dim Sql As String = ""
        Dim Ds As DataSet = Nothing


        Try
            Sql = "SELECT trialID FROM r_walktests_trials WHERE WalkID=" & walkID
            Ds = cDAL.Get_DataAsDataset(Sql)
            If IsNothing(Ds) Then
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else

                Dim nLevel As Integer = 0
                Dim lst As New List(Of Dictionary(Of String, String)())
                Dim Ds1 As DataSet = Nothing
                Dim flds As New class_fields_Walk_TrialLevel
                Dim d() As Dictionary(Of String, String) = Nothing

                For Each t In Ds.Tables(0).Rows
                    Sql = "SELECT * FROM r_walktests_trials_levels WHERE trialID=" & t("trialID")
                    Ds1 = cDAL.Get_DataAsDataset(Sql)
                    If IsNothing(Ds) Then
                        lst.Add(Nothing)
                    ElseIf Ds.Tables(0).Rows.Count = 0 Then
                        lst.Add(Nothing)
                    Else
                        d = Nothing
                        nLevel = 0
                        For Each lvl In Ds1.Tables(0).Rows
                            ReDim Preserve d(nLevel)
                            d(nLevel) = New Dictionary(Of String, String)
                            d(nLevel)(flds.levelID) = lvl(flds.levelID)
                            d(nLevel)(flds.trialID) = lvl(flds.trialID)
                            d(nLevel)(flds.time_dyspnoea) = lvl(flds.time_dyspnoea)
                            d(nLevel)(flds.time_gradient) = lvl(flds.time_gradient)
                            d(nLevel)(flds.time_hr) = lvl(flds.time_hr)
                            d(nLevel)(flds.time_label) = lvl(flds.time_label)
                            d(nLevel)(flds.time_minute) = lvl(flds.time_minute)
                            d(nLevel)(flds.time_o2flow) = lvl(flds.time_o2flow)
                            d(nLevel)(flds.time_speed_kph) = lvl(flds.time_speed_kph)
                            d(nLevel)(flds.time_spo2) = lvl(flds.time_spo2)
                            nLevel = nLevel + 1
                        Next lvl
                        lst.Add(d)
                    End If
                Next t
                Return lst
            End If

        Catch ex As Exception
            MsgBox("Error in class_patient.Get_walk_levels" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Public Function Get_prov_test_session(ByVal provID As Long) As Dictionary(Of String, String)

        Try
            Dim sql As String = "SELECT prov_test.*, r_sessions.* FROM prov_test INNER JOIN r_sessions ON prov_test.sessionID = r_sessions.sessionID WHERE prov_test.provID =" & provID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            Dim j As Integer = 0
            Dim row As DataRow

            If IsNothing(Ds) Then
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicProv_test_session
                row = Ds.Tables(0).Rows(0)
                For j = 0 To Ds.Tables(0).Columns.Count - 1
                    If d.ContainsKey(Ds.Tables(0).Columns(j).ColumnName.ToLower) Then
                        If Not IsDBNull(row(Ds.Tables(0).Columns(j).ColumnName)) Then
                            Select Case Ds.Tables(0).Columns(j).DataType.FullName
                                Case "System.TimeSpan" : d(Ds.Tables(0).Columns(j).ColumnName.ToLower) = row(Ds.Tables(0).Columns(j).ColumnName).ToString
                                Case Else : d(Ds.Tables(0).Columns(j).ColumnName.ToLower) = row(Ds.Tables(0).Columns(j).ColumnName)
                            End Select

                        Else
                            d(Ds.Tables(0).Columns(j).ColumnName.ToLower) = ""
                        End If
                    End If
                Next
                Return d
            End If

        Catch ex As Exception
            MsgBox("Error in class_patient.Get_prov_test_session" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function


    Public Function Get_Prov_VisitDataByProvID(ByVal ProvID As Long) As Dictionary(Of String, String)

        Try
            Dim sql = "SELECT * FROM prov_test WHERE provid=" & ProvID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            Dim j As Integer = 0
            Dim row As DataRow

            If IsNothing(Ds) Then
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicProv_test
                row = Ds.Tables(0).Rows(0)
                For j = 0 To Ds.Tables(0).Columns.Count - 1
                    If d.ContainsKey(Ds.Tables(0).Columns(j).ColumnName) Then d(Ds.Tables(0).Columns(j).ColumnName) = row(Ds.Tables(0).Columns(j).ColumnName).ToString
                Next
                Return d
            End If
        Catch ex As Exception
            MsgBox("Error in class_patient.Get_Prov_VisitDataByProvID" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Public Function Get_Prov_TestDataByProvID(ByVal ProvID As Long) As Dictionary(Of String, String)()

        Dim i As Integer = 0, j As Integer = 0
        Dim dicR() As Dictionary(Of String, String) = Nothing
        Dim flds = New class_ProvTestDataFields
        Dim item As DataColumn = Nothing, row As DataRow

        Dim sql = "SELECT * FROM prov_testdata WHERE provid=" & ProvID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            i = 0
            For Each row In Ds.Tables(0).Rows
                ReDim Preserve dicR(i)
                dicR(i) = cMyRoutines.MakeEmpty_dicProv_testdata()
                For j = 0 To Ds.Tables(0).Columns.Count - 1
                    If dicR(i).ContainsKey(Ds.Tables(0).Columns(j).ColumnName) Then dicR(i)(Ds.Tables(0).Columns(j).ColumnName) = row(Ds.Tables(0).Columns(j).ColumnName).ToString
                Next
                i = i + 1
            Next

            Return dicR

        End If



    End Function

    Public Function insert_rft_routine(sessionID As Long, ByVal R As Dictionary(Of String, String)) As Long

        'Build insert query
        R("sessionid") = sessionID
        Dim sql As String = cDAL.Build_InsertQuery(eTables.rft_routine, R)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new Rfts" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

    End Function

    Public Function update_rft_routine(sessionID As Long, ByVal R As Dictionary(Of String, String)) As Boolean

        R("sessionid") = sessionID
        Dim sqlString As String = cDAL.Build_UpdateQuery(eTables.rft_routine, R)

        Try
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving Rfts" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Public Function Save_rft_walk(newWalk As Boolean, sessionID As Long, session As Dictionary(Of String, String), test As Dictionary(Of String, String), trials() As Dictionary(Of String, String), levels As List(Of Dictionary(Of String, String)())) As Long
        'Data structure:
        ' 1 test record                         (table = 'r_walktests_v1heavy')
        ' n trial records                       (table = 'r_walktests_trials')
        ' n level records - set for each trial  (table = 'r_walktests_trials_levels')

        Dim i As Integer = 0, j As Integer = 0
        Dim sql As String = 0
        Dim nTrial As Integer = 0, nLevel As Integer = 0
        Dim trialID As Long
        Dim walkID As Long
        Dim levelID As Long
        Dim d() As Dictionary(Of String, String)
        Dim newTrial As Boolean
        Dim newLevel As Boolean
        Dim newSession As Boolean

        If sessionID = 0 Then newSession = True Else newSession = False

        Try

            'Do session save
            session("sessionid") = sessionID
            Select Case newSession
                Case True : sessionID = cRfts.Insert_rft_session(session)
                Case False : cRfts.Update_rft_session(session)
            End Select

            'Do test save
            test("sessionid") = sessionID
            Select Case newWalk
                Case True : walkID = cRfts.Insert_rft_walk(sessionID, test)
                Case False : walkID = test("walkid") : cRfts.update_rft_walk(sessionID, test)
            End Select

            'Do trials and levels save
            For nTrial = 0 To trials.Count - 1
                'Save trial
                trials(nTrial)("walkid") = walkID
                newTrial = Not CBool(Val(trials(nTrial)("trialid")))
                Select Case newTrial
                    Case True : trialID = cRfts.Insert_rft_walk_trial(trials(nTrial))
                    Case False : trialID = trials(nTrial)("trialid") : cRfts.Update_rft_walk_trial(trials(nTrial))
                End Select

                'Save level
                For nLevel = 0 To levels(nTrial).Count - 1
                    d = levels(nTrial)
                    d(nLevel)("trialid") = trialID
                    newLevel = Not CBool(Val(d(nLevel)("levelid")))
                    Select Case newLevel
                        Case True : levelID = cRfts.insert_rft_walk_trial_level(d(nLevel))
                        Case False : levelID = d(nLevel)("levelid") : cRfts.Update_rft_walk_trial_level(d(nLevel))
                    End Select
                Next nLevel

            Next nTrial

        Catch ex As Exception
            MsgBox("Error saving walk test record in class_patient_save_walktest" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

        Return walkID

    End Function

    Public Function Insert_rft_walk(sessionID As Long, ByVal R As Dictionary(Of String, String)) As Long

        R("sessionid") = sessionID
        Dim Sql As String = cDAL.Build_InsertQuery(eTables.r_walktests_v1heavy, R)
        Dim walkID As Long = cDAL.Insert_Record(Sql)

        Return walkID

    End Function

    Public Function Insert_rft_walk_trial(ByVal R As Dictionary(Of String, String)) As Long

        Dim Sql As String = cDAL.Build_InsertQuery(eTables.r_walktests_trials, R)
        Dim trialID As Long = cDAL.Insert_Record(Sql)

        Return trialID

    End Function

    Public Function insert_rft_walk_trial_level(ByVal R As Dictionary(Of String, String)) As Long

        Dim Sql As String = cDAL.Build_InsertQuery(eTables.r_walktests_trials_levels, R)
        Dim levelID As Long = cDAL.Insert_Record(Sql)

        Return levelID

    End Function

    Public Function update_rft_walk(sessionID As Long, ByVal R As Dictionary(Of String, String)) As Boolean

        R("sessionid") = sessionID
        Dim Sql As String = cDAL.Build_UpdateQuery(eTables.r_walktests_v1heavy, R)
        Dim success As Boolean = cDAL.Update_Record(Sql)

        Return success

    End Function

    Public Function Update_rft_walk_trial(ByVal R As Dictionary(Of String, String)) As Boolean

        Dim Sql As String = cDAL.Build_UpdateQuery(eTables.r_walktests_trials, R)
        Dim success As Boolean = cDAL.Update_Record(Sql)

        Return success

    End Function

    Public Function Update_rft_walk_trial_level(ByVal R As Dictionary(Of String, String)) As Boolean

        Dim Sql As String = cDAL.Build_UpdateQuery(eTables.r_walktests_trials_levels, R)
        Dim success As Boolean = cDAL.Update_Record(Sql)

        Return success

    End Function

    Public Function insert_rft_genericProv(sessionID As Long, ByVal dTest As Dictionary(Of String, String), ByVal dTestData() As Dictionary(Of String, String)) As Long
        'Structure = 1 x test record -> table 'prov_test' plus n x related testdata records (one for each dose) -> table 'prov_testdata'

        Dim sql As String = ""
        Dim newProvID As Long = 0
        Dim newtestdataID As Long = 0
        Dim dic As Dictionary(Of String, String) = Nothing
        Dim i As Integer = 0

        'First create the primary test record
        dTest("sessionid") = sessionID
        sql = cDAL.Build_InsertQuery(eTables.Prov_test, dTest)
        Try
            newProvID = cDAL.Insert_Record(sql)

            'Load the new foreign key into the related tables (dic)
            For i = 0 To UBound(dTestData)
                dTestData(i)(cDbInfo.primarykey(eTables.Prov_test)) = newProvID
            Next

            'Now create the secondary testdata records 
            For i = 0 To UBound(dTestData)
                sql = cDAL.Build_InsertQuery(eTables.Prov_testdata, dTestData(i))
                Try
                    newtestdataID = cDAL.Insert_Record(sql)
                Catch ex As Exception
                    MsgBox("Error creating new provocation test dose record" & vbNewLine & ex.Message.ToString)
                End Try
            Next

        Catch ex As Exception
            MsgBox("Error creating new provocation test record" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

        Return newProvID

    End Function

    Public Function insert_rft_spt(sessionID As Long, ByVal dTest As Dictionary(Of String, String), ByVal dTestData() As Dictionary(Of String, String)) As Long
        'Structure = 1 x test record -> table 'r_spt' plus n x related testdata records (one for each allergen) -> table 'r_spt_allergens'

        Dim sql As String = ""
        Dim newSptID As Long = 0
        Dim newAllergenDataID As Long = 0
        Dim dic As Dictionary(Of String, String) = Nothing
        Dim i As Integer = 0

        'First create the primary test record
        dTest("sessionid") = sessionID
        sql = cDAL.Build_InsertQuery(eTables.r_spt, dTest)
        Try
            newSptID = cDAL.Insert_Record(sql)

            'Load the new foreign key into the related tables (dic)
            If Not IsNothing(dTestData) Then
                For i = 0 To UBound(dTestData)
                    dTestData(i)(cDbInfo.primarykey(eTables.r_spt)) = newSptID
                Next

                'Now create the secondary testdata records 
                For i = 0 To UBound(dTestData)
                    sql = cDAL.Build_InsertQuery(eTables.r_spt_allergens, dTestData(i))
                    Try
                        newAllergenDataID = cDAL.Insert_Record(sql)
                    Catch ex As Exception
                        MsgBox("Error creating new spt test allergen record in class_patient.insert_rft_spt" & vbNewLine & ex.Message.ToString)
                    End Try
                Next
            End If

        Catch ex As Exception
            MsgBox("Error creating new spt test record in class_patient.insert_rft_spt" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

        Return newSptID

    End Function

    Public Function insert_rft_hast(sessionID As Long, ByVal dTest As Dictionary(Of String, String), ByVal dTestData() As Dictionary(Of String, String)) As Long
        'Structure = 1 x test record -> table 'r_hast' plus n x related testdata records (one for each level) -> table 'r_hast_levels'

        Dim sql As String = ""
        Dim newHastID As Long = 0
        Dim newLevelID As Long = 0
        Dim dic As Dictionary(Of String, String) = Nothing
        Dim i As Integer = 0

        'First create the primary test record
        dTest("sessionid") = sessionID
        sql = cDAL.Build_InsertQuery(eTables.r_hast, dTest)
        Try
            newHastID = cDAL.Insert_Record(sql)

            'Load the new foreign key into the related tables (dic)
            If Not IsNothing(dTestData) Then
                For i = 0 To UBound(dTestData)
                    dTestData(i)(cDbInfo.primarykey(eTables.r_hast)) = newHastID
                Next

                'Now create the secondary testdata records 
                For i = 0 To UBound(dTestData)
                    sql = cDAL.Build_InsertQuery(eTables.r_hast_levels, dTestData(i))
                    Try
                        newLevelID = cDAL.Insert_Record(sql)
                    Catch ex As Exception
                        MsgBox("Error creating new hast test level record in class_patient.insert_rft_hast" & vbNewLine & ex.Message.ToString)
                    End Try
                Next
            End If

        Catch ex As Exception
            MsgBox("Error creating new hast test record in class_patient.insert_rft_hast" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

        Return newHastID

    End Function

    Public Function Update_rft_genericProv(sessionID As Long, ByVal dTest As Dictionary(Of String, String), ByVal dTestData() As Dictionary(Of String, String)) As Boolean

        Dim d As Dictionary(Of String, String)
        Dim sql As String = ""
        Dim Success As Boolean
        Dim RecordID As Long = 0
        Dim nFailed As Integer = 0

        Try
            'Update once only test data first
            dTest("sessionid") = sessionID
            sql = cDAL.Build_UpdateQuery(eTables.Prov_test, dTest)
            Success = cDAL.Update_Record(sql)
            If Success Then   'Successfully saved
                'Now update dose records
                For Each d In dTestData
                    sql = cDAL.Build_UpdateQuery(eTables.Prov_testdata, d)
                    Success = cDAL.Update_Record(sql)
                    If Not Success Then
                        nFailed = nFailed + 1
                    End If
                Next
                If nFailed > 0 Then
                    MsgBox("Error saving " & nFailed & " of " & dTestData.Count & " dose data records (in class_patient.Update_GenericProv)")
                    Return False
                End If
            Else
                MsgBox("Error saving challenge - aborted (in class_patient.Update_GenericProv)")
                Return False
            End If
            Return True

        Catch ex As Exception
            MsgBox("Error saving challenge - aborted (in class_patient.Update_GenericProv)" & vbNewLine & ex.Message.ToString)
            Return False

        End Try

    End Function

    Public Function Update_rft_spt(sessionID As Long, ByVal dTest As Dictionary(Of String, String), ByVal dTestData() As Dictionary(Of String, String)) As Boolean

        Dim d As Dictionary(Of String, String)
        Dim sql As String = ""
        Dim Success As Boolean
        Dim RecordID As Long = 0
        Dim nFailed As Integer = 0

        Try
            'Update once only test data first
            dTest("sessionid") = sessionID
            sql = cDAL.Build_UpdateQuery(eTables.r_spt, dTest)
            Success = cDAL.Update_Record(sql)
            If Success Then   'Successfully saved
                'Now update dose records
                For Each d In dTestData
                    sql = cDAL.Build_UpdateQuery(eTables.r_spt_allergens, d)
                    Success = cDAL.Update_Record(sql)
                    If Not Success Then
                        nFailed = nFailed + 1
                    End If
                Next
                If nFailed > 0 Then
                    MsgBox("Error saving " & nFailed & " of " & dTestData.Count & " allergen data records (in class_patient.Update_rft_spt)")
                    Return False
                End If
            Else
                MsgBox("Error saving spt - aborted (in class_patient.Update_rft_spt)")
                Return False
            End If
            Return True

        Catch ex As Exception
            MsgBox("Error saving spt - aborted (in class_patient.Update_rft_spt)" & vbNewLine & ex.Message.ToString)
            Return False

        End Try

    End Function


    Public Function Update_rft_hast(sessionID As Long, ByVal r_test As Dictionary(Of String, String), ByVal r_levels() As Dictionary(Of String, String)) As Boolean
        'Need to handle cases where
        '  1. existing records: levelID=levelID
        '  2. deleted records:  levelID=-(levelID)
        '  3. new records:      levelID=0

        Dim d As Dictionary(Of String, String) = Nothing
        Dim sql As String = ""
        Dim Success As Boolean
        Dim RecordID As Long = 0
        Dim nFailed As Integer = 0
        Dim Msg As String = ""

        Try
            'Update test data first
            r_test("sessionid") = sessionID
            sql = cDAL.Build_UpdateQuery(eTables.r_hast, r_test)
            Success = cDAL.Update_Record(sql)
            If Success Then
                'Now handle level records
                For Each level In r_levels
                    Select Case level("levelid")
                        Case 0
                            sql = cDAL.Build_InsertQuery(eTables.r_hast_levels, level)
                            Success = cDAL.Insert_Record(sql)
                            Msg = "creating"
                        Case Is < 0
                            Success = cDAL.Delete_Record(Math.Abs(CLng(level("levelid"))), eTables.r_hast_levels, , , True)
                            Msg = "deleting"
                        Case Is > 0
                            sql = cDAL.Build_UpdateQuery(eTables.r_hast_levels, level)
                            Success = cDAL.Update_Record(sql)
                            Msg = "updating"
                    End Select
                    If Not Success Then MsgBox("Error " & Msg & " hast level #: " & level("levelID"))
                Next
            Else
                MsgBox("Error saving hast - aborted (In class_patient.Update_rft_hast)")
                Return False
            End If
            Return True

        Catch ex As Exception
            MsgBox("Error saving hast - aborted (In class_patient.Update_rft_hast)" & vbNewLine & ex.Message.ToString)
            Return False

        End Try

    End Function

    Public Function Update_ReportStatus(NewStatus As eReportStatus, eTbl As eTables, RecordID As Long) As Boolean

        Dim sql As String = ""
        Dim d As New Dictionary(Of String, String)

        Try
            d.Add("Report_Status", "'" & NewStatus.ToString & "'")
            d.Add(cDbInfo.primarykey(eTbl).ToLower, RecordID)
            sql = cDAL.Build_UpdateQuery(eTbl, d)

            Dim Success As Boolean = cDAL.Update_Record(sql)
            Return True
        Catch
            MsgBox("Error in class_Patient.Update_ReportStatus" & vbCrLf & Err.Description, vbOKOnly, "Batch print")
            Return False
        End Try

    End Function

    Public Function Get_SessionToUse_newtest(patientID As Long, thisDate As Date) As Long
        'For new tests
        '   return sessionID
        '       0:      create new session
        '       vbnull: don't wish to use existing session for selected date
        '       value:  use this session

        Dim sessionID As Nullable(Of Long)
        Dim response As Integer = 0
        Dim Msg As String = ""

        If thisDate = Date.Today Then
            Msg = "Existing testing session found for today. Add test to this session?"
        Else
            Msg = "Existing testing session found for " & thisDate.ToString & ". Add test to this session?"
        End If

        Dim ID As Long = cRfts.Exists_RftSessionForDate(patientID, thisDate)
        Select Case ID
            Case 0 : sessionID = 0
            Case Is > 0
                response = MsgBox(Msg, vbYesNo, "New test")
                If response = vbYes Then
                    sessionID = ID
                Else
                    sessionID = vbNull
                End If

            Case Is < 0
                MsgBox("Error. " & Math.Abs(ID) & " sessions found for this date. The first will be used.", vbOKOnly, "New test")
                'Should never happen - only 1 test session per day allowed
                'MsgBox("The first session for this date will be used.", vbOKOnly, "New test")
                'sessionID = Math.Abs(ID)
        End Select

        Return sessionID

    End Function

    Public Function Get_SessionToUse_existingtest(testID As Long, eTbl As eTables) As Long

        Dim sessionID As Long
        sessionID = cRfts.get_rft_sessionID(testID, eTbl)
        Return sessionID

    End Function

    Public Function Exists_RftSessionForDate(patientID As Long, thisDate As Date) As Long
        'Returns 
        '   0 if no session found
        '   sessionID if one found
        '   first sessionid as a negative number if >1 found (this is an error condition)

        If IsDate(thisDate) And patientID > 0 Then
            Dim sql As String = "SELECT sessionID FROM r_sessions WHERE patientID=" & patientID & " AND testdate='" & Format(thisDate, "yyyy-MM-dd") & "'"
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If IsNothing(Ds) Then
                Return Nothing
            ElseIf Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Select Case Ds.Tables(0).Rows.Count
                    Case 0
                        Return 0
                    Case 1
                        Return Ds.Tables(0).Rows(0).Item(0)
                    Case Is > 1
                        'Dim Msg As String = Ds.Tables(0).Rows.Count & " test session records found for this patient/date." & vbCrLf
                        'Msg = Msg & "Only one testing session per day is allowed." & vbCrLf
                        'Msg = Msg & "Contact your database administrator to sort this out."
                        'MsgBox(Msg, vbOKOnly, "Warning")
                        Return -Ds.Tables(0).Rows(0).Item(0)
                    Case Else
                        Return 0
                End Select
            End If
        Else
            Return 0
        End If

    End Function

    Public Function get_rft_sessionID(recordID As Long, eTbl As eTables) As Long
        'Returns 0 if prob or not found

        Dim sql As String = "SELECT sessionID FROM " & cDbInfo.table_name(eTbl) & " WHERE " & cDbInfo.primarykey(eTbl) & "=" & recordID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds Is Nothing Then
            Return 0
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return 0
        Else
            If IsNumeric(Ds.Tables(0).Rows(0).Item(0)) Then
                Return Ds.Tables(0).Rows(0).Item(0)
            Else
                Return 0
            End If
        End If

    End Function

    Public Function Get_rft_cpet_testdata(ByVal cpetID As Long) As Dictionary(Of String, String)

        Dim i As Integer = 0, j As Integer = 0
        Dim f As New class_fields_CPETAndSessionFields
        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicCPETandSession
        Dim item As DataColumn = Nothing, row As DataRow = Nothing

        Dim sql As String = "SELECT r_cpet.*, r_sessions.* FROM r_cpet INNER JOIN r_sessions ON r_cpet.sessionID = r_sessions.sessionID WHERE r_cpet.cpetID =" & cpetID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)

                d(f.patientID) = .Rows(0)(f.patientID)
                d(f.cpetID) = .Rows(0)(f.cpetID)
                If IsDate(.Rows(0)(f.TestDate)) Then d(f.TestDate) = .Rows(0)(f.TestDate) Else d(f.TestDate) = ""
                If IsNumeric(.Rows(0)(f.Height) & "") Then d(f.Height) = Math.Round(Val(.Rows(0)(f.Height)), 1) Else d(f.Height) = ""
                If IsNumeric(.Rows(0)(f.Weight) & "") Then d(f.Weight) = Math.Round(Val(.Rows(0)(f.Weight)), 1) Else d(f.Weight) = ""
                d(f.Smoke_hx) = .Rows(0)(f.Smoke_hx) & ""
                d(f.Smoke_packyears) = .Rows(0)(f.Smoke_packyears) & ""
                d(f.Smoke_yearssmoked) = .Rows(0)(f.Smoke_yearssmoked) & ""
                d(f.Smoke_cigsperday) = .Rows(0)(f.Smoke_cigsperday) & ""
                If IsDate(.Rows(0)(f.Req_date)) Then d(f.Req_date) = .Rows(0)(f.Req_date) Else d(f.Req_date) = Nothing
                If IsDate(.Rows(0)(f.Req_time)) Then d(f.Req_time) = .Rows(0)(f.Req_time) Else d(f.Req_time) = Nothing
                d(f.Req_name) = .Rows(0)(f.Req_name) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_providernumber) = .Rows(0)(f.Req_providernumber) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_email) = .Rows(0)(f.Req_email) & ""
                d(f.Req_fax) = .Rows(0)(f.Req_fax) & ""
                d(f.Req_phone) = .Rows(0)(f.Req_phone) & ""
                d(f.Report_copyto) = .Rows(0)(f.Report_copyto) & ""
                d(f.Req_clinicalnotes) = .Rows(0)(f.Req_clinicalnotes) & ""
                d(f.Req_healthservice) = .Rows(0)(f.Req_healthservice) & ""
                d(f.AdmissionStatus) = .Rows(0)(f.AdmissionStatus) & ""
                d(f.BDStatus) = .Rows(0)(f.BDStatus) & ""
                d(f.Billing_billedto) = .Rows(0)(f.Billing_billedto) & ""
                d(f.Billing_billingMO) = .Rows(0)(f.Billing_billingMO) & ""
                d(f.Billing_billingMOproviderno) = .Rows(0)(f.Billing_billingMOproviderno) & ""

                d(f.TechnicalNotes) = .Rows(0)(f.TechnicalNotes) & ""
                d(f.Report_text) = .Rows(0)(f.Report_text) & ""
                d(f.Report_status) = .Rows(0)(f.Report_status) & ""
                d(f.Report_reportedby) = .Rows(0)(f.Report_reportedby) & ""
                If IsDate(.Rows(0)(f.Report_reporteddate)) Then d(f.Report_reporteddate) = .Rows(0)(f.Report_reporteddate) Else d(f.Report_reporteddate) = ""
                d(f.Report_verifiedby) = .Rows(0)(f.Report_verifiedby) & ""
                If IsDate(.Rows(0)(f.Report_verifieddate)) Then d(f.Report_verifieddate) = .Rows(0)(f.Report_verifieddate) Else d(f.Report_verifieddate) = ""
                If IsDate(.Rows(0)(f.TestTime).ToString) Then d(f.TestTime) = .Rows(0)(f.TestTime).ToString Else d(f.TestTime) = ""
                d(f.TestType) = .Rows(0)(f.TestType) & ""
                d(f.Lab) = .Rows(0)(f.Lab) & ""
                d(f.Scientist) = .Rows(0)(f.Scientist) & ""

                d(f.r_bp_pre) = .Rows(0)(f.r_bp_pre) & ""
                d(f.r_bp_post) = .Rows(0)(f.r_bp_post) & ""
                d(f.r_spiro_pre_fev1) = .Rows(0)(f.r_spiro_pre_fev1) & ""
                d(f.r_spiro_pre_fvc) = .Rows(0)(f.r_spiro_pre_fvc) & ""
                d(f.r_symptoms_dyspnoea_borg) = .Rows(0)(f.r_symptoms_dyspnoea_borg) & ""
                d(f.r_symptoms_legs_borg) = .Rows(0)(f.r_symptoms_legs_borg) & ""
                d(f.r_symptoms_other_borg) = .Rows(0)(f.r_symptoms_other_borg) & ""
                d(f.r_symptoms_other_text) = .Rows(0)(f.r_symptoms_other_text) & ""
                'd(f.r_abg_post_be) = .Rows(0)(f.r_abg_post_be) & ""
                'd(f.r_abg_post_fio2) = "" & ""
                'd(f.r_abg_post_hco3) = .Rows(0)(f.r_abg_post_hco3) & ""
                'd(f.r_abg_post_paco2) = .Rows(0)(f.r_abg_post_paco2) & ""
                'd(f.r_abg_post_pao2) = .Rows(0)(f.r_abg_post_pao2) & ""
                'd(f.r_abg_post_sao2) = .Rows(0)(f.r_abg_post_sao2) & ""
                'd(f.r_abg_post_ph) = .Rows(0)(f.r_abg_post_ph) & ""
                d(f.r_max_hr) = .Rows(0)(f.r_max_hr) & ""
                d(f.r_max_hr_lln) = .Rows(0)(f.r_max_hr_lln) & ""
                d(f.r_max_hr_mpv) = .Rows(0)(f.r_max_hr_mpv) & ""
                d(f.r_max_o2pulse) = .Rows(0)(f.r_max_o2pulse) & ""
                d(f.r_max_o2pulse_lln) = .Rows(0)(f.r_max_o2pulse_lln) & ""
                d(f.r_max_o2pulse_mpv) = .Rows(0)(f.r_max_o2pulse_mpv) & ""
                d(f.r_max_vco2) = .Rows(0)(f.r_max_vco2) & ""
                d(f.r_max_vco2_lln) = .Rows(0)(f.r_max_vco2_lln) & ""
                d(f.r_max_vco2_mpv) = .Rows(0)(f.r_max_vco2_mpv) & ""
                d(f.r_max_vo2) = .Rows(0)(f.r_max_vo2) & ""
                d(f.r_max_vo2_lln) = .Rows(0)(f.r_max_vo2_lln) & ""
                d(f.r_max_vo2_mpv) = .Rows(0)(f.r_max_vo2_mpv) & ""
                d(f.r_max_ve) = .Rows(0)(f.r_max_ve) & ""
                d(f.r_max_ve_lln) = .Rows(0)(f.r_max_ve_lln) & ""
                d(f.r_max_ve_mpv) = .Rows(0)(f.r_max_ve_mpv) & ""
                d(f.r_max_vt) = .Rows(0)(f.r_max_vt) & ""
                d(f.r_max_vt_lln) = .Rows(0)(f.r_max_vt_lln) & ""
                d(f.r_max_vt_mpv) = .Rows(0)(f.r_max_vt_mpv) & ""
                d(f.r_max_vo2kg) = .Rows(0)(f.r_max_vo2kg) & ""
                d(f.r_max_vo2kg_lln) = .Rows(0)(f.r_max_vo2kg_lln) & ""
                d(f.r_max_vo2kg_mpv) = .Rows(0)(f.r_max_vo2kg_mpv) & ""
                d(f.r_max_workload) = .Rows(0)(f.r_max_workload) & ""
                d(f.r_max_workload_lln) = .Rows(0)(f.r_max_workload_lln) & ""
                d(f.r_max_workload_mpv) = .Rows(0)(f.r_max_workload_mpv) & ""
            End With
        End If

        Return d

    End Function

    Public Function Get_rft_spt_allergenresults(ByVal sptID As Long) As Dictionary(Of String, String)()

        Dim i As Integer = 0, j As Integer = 0
        Dim f As New class_fields_Spt_Allergens
        Dim d() As Dictionary(Of String, String) = Nothing
        Dim item As DataColumn = Nothing, row As DataRow

        Dim sql As String = "SELECT * FROM r_spt_allergens WHERE sptID=" & sptID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            i = 0
            For Each row In Ds.Tables(0).Rows
                ReDim Preserve d(i)
                d(i) = New Dictionary(Of String, String)
                d(i) = cMyRoutines.MakeEmpty_dicSpt_allergen
                d(i)(f.sptID) = row(f.sptID)
                d(i)(f.allergenID) = row(f.allergenID)
                d(i)(f.allergen_category_id) = row(f.allergen_category_id)
                d(i)(f.allergen_category_name) = row(f.allergen_category_name)
                d(i)(f.allergen_category_colour) = row(f.allergen_category_colour)
                d(i)(f.allergen_name) = row(f.allergen_name)
                d(i)(f.note) = row(f.note)
                d(i)(f.panelmemberID) = row(f.panelmemberID)
                d(i)(f.wheal_mm) = row(f.wheal_mm)
                i = i + 1
            Next

            Return d

        End If

    End Function


    Public Function Get_rft_spt_test_session(ByVal sptID As Long) As Dictionary(Of String, String)

        Dim i As Integer = 0, j As Integer = 0
        Dim f As New class_fields_SptAndSessionFields
        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicSPTandSession
        Dim item As DataColumn = Nothing, row As DataRow = Nothing

        Dim sql As String = "SELECT r_spt.*, r_sessions.* FROM r_spt INNER JOIN r_sessions ON r_spt.sessionID = r_sessions.sessionID WHERE r_spt.sptID =" & sptID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)

                d(f.patientID) = .Rows(0)(f.patientID)
                d(f.sptID) = .Rows(0)(f.sptID)
                If IsDate(.Rows(0)(f.TestDate)) Then d(f.TestDate) = .Rows(0)(f.TestDate) Else d(f.TestDate) = ""
                If IsNumeric(.Rows(0)(f.Height) & "") Then d(f.Height) = Math.Round(Val(.Rows(0)(f.Height)), 1) Else d(f.Height) = ""
                If IsNumeric(.Rows(0)(f.Weight) & "") Then d(f.Weight) = Math.Round(Val(.Rows(0)(f.Weight)), 1) Else d(f.Weight) = ""
                d(f.Smoke_hx) = .Rows(0)(f.Smoke_hx) & ""
                d(f.Smoke_packyears) = .Rows(0)(f.Smoke_packyears) & ""
                d(f.Smoke_yearssmoked) = .Rows(0)(f.Smoke_yearssmoked) & ""
                d(f.Smoke_cigsperday) = .Rows(0)(f.Smoke_cigsperday) & ""
                If IsDate(.Rows(0)(f.Req_date)) Then d(f.Req_date) = .Rows(0)(f.Req_date) Else d(f.Req_date) = Nothing
                If IsDate(.Rows(0)(f.Req_time)) Then d(f.Req_time) = .Rows(0)(f.Req_time) Else d(f.Req_time) = Nothing
                d(f.Req_name) = .Rows(0)(f.Req_name) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_providernumber) = .Rows(0)(f.Req_providernumber) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_email) = .Rows(0)(f.Req_email) & ""
                d(f.Req_fax) = .Rows(0)(f.Req_fax) & ""
                d(f.Req_phone) = .Rows(0)(f.Req_phone) & ""
                d(f.Report_copyto) = .Rows(0)(f.Report_copyto) & ""
                d(f.Req_clinicalnotes) = .Rows(0)(f.Req_clinicalnotes) & ""
                d(f.Req_healthservice) = .Rows(0)(f.Req_healthservice) & ""
                d(f.AdmissionStatus) = .Rows(0)(f.AdmissionStatus) & ""
                d(f.BDStatus) = .Rows(0)(f.BDStatus) & ""
                d(f.Billing_billedto) = .Rows(0)(f.Billing_billedto) & ""
                d(f.Billing_billingMO) = .Rows(0)(f.Billing_billingMO) & ""
                d(f.Billing_billingMOproviderno) = .Rows(0)(f.Billing_billingMOproviderno) & ""

                d(f.TechnicalNotes) = .Rows(0)(f.TechnicalNotes) & ""
                d(f.Report_text) = .Rows(0)(f.Report_text) & ""
                d(f.Report_status) = .Rows(0)(f.Report_status) & ""
                d(f.Report_reportedby) = .Rows(0)(f.Report_reportedby) & ""
                If IsDate(.Rows(0)(f.Report_reporteddate)) Then d(f.Report_reporteddate) = .Rows(0)(f.Report_reporteddate) Else d(f.Report_reporteddate) = ""
                d(f.Report_verifiedby) = .Rows(0)(f.Report_verifiedby) & ""
                If IsDate(.Rows(0)(f.Report_verifieddate)) Then d(f.Report_verifieddate) = .Rows(0)(f.Report_verifieddate) Else d(f.Report_verifieddate) = ""
                If IsDate(.Rows(0)(f.TestTime).ToString) Then d(f.TestTime) = .Rows(0)(f.TestTime).ToString Else d(f.TestTime) = ""
                d(f.TestType) = .Rows(0)(f.TestType) & ""
                d(f.Lab) = .Rows(0)(f.Lab) & ""
                d(f.Scientist) = .Rows(0)(f.Scientist) & ""

                d(f.panelID) = .Rows(0)(f.panelID) & ""
                d(f.panel_name) = .Rows(0)(f.panel_name) & ""
                d(f.site) = .Rows(0)(f.site) & ""
                d(f.medications) = .Rows(0)(f.medications) & ""

            End With
        End If

        Return d

    End Function

    Public Function Get_rft_hast_levelresults(ByVal hastID As Long) As Dictionary(Of String, String)()

        Dim i As Integer = 0, j As Integer = 0
        Dim f As New class_fields_Hast_Levels
        Dim d() As Dictionary(Of String, String) = Nothing
        Dim item As DataColumn = Nothing, row As DataRow

        Dim sql As String = "SELECT * FROM r_hast_levels WHERE hastID=" & hastID & " ORDER BY row_order ASC"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            i = 0
            For Each row In Ds.Tables(0).Rows
                ReDim Preserve d(i)
                d(i) = New Dictionary(Of String, String)
                d(i) = cMyRoutines.MakeEmpty_dicHAST_level
                d(i)(f.hastID) = row(f.hastID)
                d(i)(f.levelID) = row(f.levelID)
                If Not IsDBNull(row(f.row_order)) Then d(i)(f.row_order) = row(f.row_order) Else d(i)(f.row_order) = 0
                d(i)(f.altitude_fio2) = row(f.altitude_fio2)
                d(i)(f.altitude_ft) = row(f.altitude_ft)
                d(i)(f.r_be) = row(f.r_be)
                d(i)(f.r_hco3) = row(f.r_hco3)
                d(i)(f.r_note) = row(f.r_note)
                d(i)(f.r_paco2) = row(f.r_paco2)
                d(i)(f.r_pao2) = row(f.r_pao2)
                d(i)(f.r_ph) = row(f.r_ph)
                d(i)(f.r_sao2) = row(f.r_sao2)
                d(i)(f.r_spo2) = row(f.r_spo2)
                d(i)(f.suppO2_flow) = row(f.suppO2_flow)
                i = i + 1
            Next

            Return d

        End If

    End Function

    Public Function Get_rft_hast_test_session(ByVal hastID As Long) As Dictionary(Of String, String)

        Dim i As Integer = 0, j As Integer = 0
        Dim f As New class_fields_HastAndSessionFields
        Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicHASTandSession
        Dim item As DataColumn = Nothing, row As DataRow = Nothing

        Dim sql As String = "SELECT r_hast.*, r_sessions.* FROM r_hast INNER JOIN r_sessions ON r_hast.sessionID = r_sessions.sessionID WHERE r_hast.hastID =" & hastID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            With Ds.Tables(0)

                d(f.patientID) = .Rows(0)(f.patientID)
                d(f.hastID) = .Rows(0)(f.hastID)
                If IsDate(.Rows(0)(f.TestDate)) Then d(f.TestDate) = .Rows(0)(f.TestDate) Else d(f.TestDate) = ""
                If IsNumeric(.Rows(0)(f.Height) & "") Then d(f.Height) = Math.Round(Val(.Rows(0)(f.Height)), 1) Else d(f.Height) = ""
                If IsNumeric(.Rows(0)(f.Weight) & "") Then d(f.Weight) = Math.Round(Val(.Rows(0)(f.Weight)), 1) Else d(f.Weight) = ""
                d(f.Smoke_hx) = .Rows(0)(f.Smoke_hx) & ""
                d(f.Smoke_packyears) = .Rows(0)(f.Smoke_packyears) & ""
                d(f.Smoke_yearssmoked) = .Rows(0)(f.Smoke_yearssmoked) & ""
                d(f.Smoke_cigsperday) = .Rows(0)(f.Smoke_cigsperday) & ""
                If IsDate(.Rows(0)(f.Req_date)) Then d(f.Req_date) = .Rows(0)(f.Req_date) Else d(f.Req_date) = Nothing
                If IsDate(.Rows(0)(f.Req_time)) Then d(f.Req_time) = .Rows(0)(f.Req_time) Else d(f.Req_time) = Nothing
                d(f.Req_name) = .Rows(0)(f.Req_name) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_providernumber) = .Rows(0)(f.Req_providernumber) & ""
                d(f.Req_address) = .Rows(0)(f.Req_address) & ""
                d(f.Req_email) = .Rows(0)(f.Req_email) & ""
                d(f.Req_fax) = .Rows(0)(f.Req_fax) & ""
                d(f.Req_phone) = .Rows(0)(f.Req_phone) & ""
                d(f.Report_copyto) = .Rows(0)(f.Report_copyto) & ""
                d(f.Req_clinicalnotes) = .Rows(0)(f.Req_clinicalnotes) & ""
                d(f.Req_healthservice) = .Rows(0)(f.Req_healthservice) & ""
                d(f.AdmissionStatus) = .Rows(0)(f.AdmissionStatus) & ""
                d(f.BDStatus) = .Rows(0)(f.BDStatus) & ""
                d(f.Billing_billedto) = .Rows(0)(f.Billing_billedto) & ""
                d(f.Billing_billingMO) = .Rows(0)(f.Billing_billingMO) & ""
                d(f.Billing_billingMOproviderno) = .Rows(0)(f.Billing_billingMOproviderno) & ""

                d(f.TechnicalNotes) = .Rows(0)(f.TechnicalNotes) & ""
                d(f.Report_text) = .Rows(0)(f.Report_text) & ""
                d(f.Report_status) = .Rows(0)(f.Report_status) & ""
                d(f.Report_reportedby) = .Rows(0)(f.Report_reportedby) & ""
                If IsDate(.Rows(0)(f.Report_reporteddate)) Then d(f.Report_reporteddate) = .Rows(0)(f.Report_reporteddate) Else d(f.Report_reporteddate) = ""
                d(f.Report_verifiedby) = .Rows(0)(f.Report_verifiedby) & ""
                If IsDate(.Rows(0)(f.Report_verifieddate)) Then d(f.Report_verifieddate) = .Rows(0)(f.Report_verifieddate) Else d(f.Report_verifieddate) = ""
                If IsDate(.Rows(0)(f.TestTime).ToString) Then d(f.TestTime) = .Rows(0)(f.TestTime).ToString Else d(f.TestTime) = ""
                d(f.TestType) = .Rows(0)(f.TestType) & ""
                d(f.Lab) = .Rows(0)(f.Lab) & ""
                d(f.Scientist) = .Rows(0)(f.Scientist) & ""

                d(f.deliverymethod_fio2) = .Rows(0)(f.deliverymethod_fio2) & ""
                d(f.deliverymethod_suppo2) = .Rows(0)(f.deliverymethod_suppo2) & ""

            End With
        End If

        Return d

    End Function

    Public Function Get_rft_walk_trialIDs(walkID As Long) As Long()

        Dim trialID() As Long = Nothing
        Dim i As Integer = 0
        Dim sql As String = "SELECT trialID FROM r_walktests_trials WHERE walkID=" & walkID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) And Ds.Tables(0).Rows.Count > 0 Then
            For Each r As DataRow In Ds.Tables(0).Rows
                ReDim Preserve trialID(i)
                trialID(i) = r("trialID")
                i = i + 1
            Next
        End If

        Return trialID

    End Function

    Public Function Get_rft_cpet_levelids(ByVal cpetID As Long) As Long()

        Dim i As Integer = 0
        Dim IDs() As Long

        Dim sql As String = "SELECT levelID FROM r_cpet_levels WHERE cpetID=" & cpetID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            ReDim IDs(Ds.Tables(0).Rows.Count - 1)
            For i = 0 To Ds.Tables(0).Rows.Count - 1
                IDs(i) = Ds.Tables(0).Rows(i)("levelID")
            Next
            Return IDs
        End If

    End Function

    Public Function Get_rft_spt_allergenids(ByVal sptID As Long) As Long()

        Dim i As Integer = 0
        Dim IDs() As Long

        Dim sql As String = "SELECT allergenID FROM r_spt_allergens WHERE sptID=" & sptID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            ReDim IDs(Ds.Tables(0).Rows.Count - 1)
            For i = 0 To Ds.Tables(0).Rows.Count - 1
                IDs(i) = Ds.Tables(0).Rows(i)("allergenID")
            Next
            Return IDs
        End If

    End Function

    Public Function Get_rft_hast_levelids(ByVal hastID As Long) As Long()

        Dim i As Integer = 0
        Dim IDs() As Long

        Dim sql As String = "SELECT levelID FROM r_hast_levels WHERE hastID=" & hastID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            ReDim IDs(Ds.Tables(0).Rows.Count - 1)
            For i = 0 To Ds.Tables(0).Rows.Count - 1
                IDs(i) = Ds.Tables(0).Rows(i)("levelID")
            Next
            Return IDs
        End If

    End Function

    Public Function Get_rft_prov_levelids(ByVal provID As Long) As Long()

        Dim i As Integer = 0
        Dim IDs() As Long

        Dim sql As String = "SELECT testdataid FROM prov_testdata WHERE provID=" & provID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            ReDim IDs(Ds.Tables(0).Rows.Count - 1)
            For i = 0 To Ds.Tables(0).Rows.Count - 1
                IDs(i) = Ds.Tables(0).Rows(i)("testdataID")
            Next
            Return IDs
        End If

    End Function

    Public Function Get_rft_cpet_workloaddata(ByVal cpetID As Long) As Dictionary(Of String, String)()

        Dim i As Integer = 0, j As Integer = 0
        Dim f As New class_fields_Cpet_Levels
        Dim d() As Dictionary(Of String, String) = Nothing
        Dim item As DataColumn = Nothing, row As DataRow

        Dim sql As String = "SELECT * FROM r_cpet_levels WHERE cpetID=" & cpetID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            i = 0
            For Each row In Ds.Tables(0).Rows
                ReDim Preserve d(i)
                d(i) = New Dictionary(Of String, String)
                d(i) = cMyRoutines.MakeEmpty_dicCPET_levels
                d(i)(f.cpetID) = row(f.cpetID)
                d(i)(f.levelID) = row(f.levelID)
                d(i)(f.level_hr) = row(f.level_hr) & ""
                d(i)(f.level_number) = row(f.level_number) & ""
                d(i)(f.level_o2pulse) = row(f.level_o2pulse) & ""
                d(i)(f.level_petco2) = row(f.level_petco2) & ""
                d(i)(f.level_peto2) = row(f.level_peto2) & ""
                d(i)(f.level_rer) = row(f.level_rer) & ""
                d(i)(f.level_spo2) = row(f.level_spo2) & ""
                d(i)(f.level_time) = row(f.level_time) & ""
                d(i)(f.level_vco2) = row(f.level_vco2) & ""
                d(i)(f.level_ve) = row(f.level_ve) & ""
                d(i)(f.level_vevco2) = row(f.level_vevco2) & ""
                d(i)(f.level_vevo2) = row(f.level_vevo2) & ""
                d(i)(f.level_vo2) = row(f.level_vo2) & ""
                d(i)(f.level_vt) = row(f.level_vt) & ""
                d(i)(f.level_workload) = row(f.level_workload) & ""
                d(i)(f.level_bp) = row(f.level_bp) & ""
                d(i)(f.level_ph) = row(f.level_ph) & ""
                d(i)(f.level_paco2) = row(f.level_paco2) & ""
                d(i)(f.level_pao2) = row(f.level_pao2) & ""
                d(i)(f.level_sao2) = row(f.level_sao2) & ""
                d(i)(f.level_hco3) = row(f.level_hco3) & ""
                d(i)(f.level_be) = row(f.level_be) & ""
                d(i)(f.level_aapo2) = row(f.level_aapo2) & ""
                d(i)(f.level_vdvt) = row(f.level_vdvt) & ""

                i = i + 1
            Next

            Return d

        End If

    End Function

    Public Function Insert_rft_session(ByVal R As Dictionary(Of String, String)) As Long

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.r_sessions, R)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new rft session" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

    End Function

    Public Function Update_rft_session(ByVal R As Dictionary(Of String, String)) As Boolean

        Dim sqlString As String = cDAL.Build_UpdateQuery(eTables.r_sessions, R)

        Try
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving rft session" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Public Function Insert_rft_cpet(sessionID As Long, ByVal r_test As Dictionary(Of String, String), ByVal r_levels() As Dictionary(Of String, String)) As Long

        'Structure = 1 x test record -> table 'r_cpet' plus n x related testdata records (one for each level) -> table 'r_cpet_levels'

        Dim sql As String = ""
        Dim newcpetID As Long = 0
        Dim newcpetlevelID As Long = 0
        Dim dic As Dictionary(Of String, String) = Nothing
        Dim i As Integer = 0


        Try
            'First create the primary test record
            r_test("sessionid") = sessionID
            sql = cDAL.Build_InsertQuery(eTables.r_cpet, r_test)
            newcpetID = cDAL.Insert_Record(sql)

            'Now create the secondary testdata records 
            If newcpetID > 0 Then
                If Not IsNothing(r_levels) Then
                    For i = 0 To UBound(r_levels)
                        Me.Insert_rft_cpet_level(newcpetID, r_levels(i))
                    Next
                End If
            Else
                MsgBox("Error creating new cpet test record")
                Return 0
            End If

        Catch ex As Exception
            MsgBox("Error creating new cpet test record" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

        Return newcpetID

    End Function

    Public Function Insert_rft_cpet_level(cpetID As Long, level As Dictionary(Of String, String)) As Long

        Dim sql As String = ""
        Dim newLevelID As Long = 0
        Dim dic As Dictionary(Of String, String) = Nothing

        Try
            'Load the new foreign key for the test
            level("cpetid") = cpetID

            'Create the new level record
            sql = cDAL.Build_InsertQuery(eTables.r_cpet_levels, level)
            newLevelID = cDAL.Insert_Record(sql)

        Catch ex As Exception
            MsgBox("Error creating new cpet load record for load #: " & level("level_number") & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

        Return newLevelID

    End Function

    Public Function Update_rft_cpet(sessionID As Long, r_test As Dictionary(Of String, String), ByVal r_levels() As Dictionary(Of String, String)) As Boolean
        'Need to handle cases where
        '  1. existing records: levelID=levelID
        '  2. deleted records:  levelID=-(levelID)
        '  3. new records:      levelID=0

        Dim d As Dictionary(Of String, String) = Nothing
        Dim sql As String = ""
        Dim Success As Boolean
        Dim RecordID As Long = 0
        Dim nFailed As Integer = 0
        Dim Msg As String = ""

        Try
            'Update test data first
            r_test("sessionid") = sessionID
            sql = cDAL.Build_UpdateQuery(eTables.r_cpet, r_test)
            Success = cDAL.Update_Record(sql)
            If Success Then
                'Now handle level records
                For Each level In r_levels
                    Select Case level("levelid")
                        Case 0
                            sql = cDAL.Build_InsertQuery(eTables.r_cpet_levels, level)
                            Success = cDAL.Insert_Record(sql)
                            Msg = "creating"
                        Case Is < 0
                            Success = cDAL.Delete_Record(Math.Abs(CLng(level("levelid"))), eTables.r_cpet_levels, , , True)
                            Msg = "deleting"
                        Case Is > 0
                            sql = cDAL.Build_UpdateQuery(eTables.r_cpet_levels, level)
                            Success = cDAL.Update_Record(sql)
                            Msg = "updating"
                    End Select
                    If Not Success Then MsgBox("Error " & Msg & " cpet workload #: " & level("level_number"))
                Next
            Else
                MsgBox("Error saving cpet - aborted (In class_patient.Update_rft_cpet)")
                Return False
            End If
            Return True

        Catch ex As Exception
            MsgBox("Error saving cpet - aborted (In class_patient.Update_rft_cpet)" & vbNewLine & ex.Message.ToString)
            Return False

        End Try

    End Function


End Class




