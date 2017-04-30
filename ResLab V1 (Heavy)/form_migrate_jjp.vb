
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.Common
Imports System.Text

Public Class form_migrate_jjp

    Dim DBtype_source As eDBtype
    Dim DBtype_dest As eDBtype
    Dim conn_source As DbConnection
    Dim conn_dest As DbConnection

    Private Enum eDBtype
        MySQL
        SQLServer
        MyLapTop
    End Enum

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    Private Sub cmboLab_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmboLab.SelectedIndexChanged

        Select Case cmboLab.Text
            Case "MRSS"
                cmboDBType_source.Text = "MySQL"
                cmboDBType_destination.Text = "MySQL"
                DBtype_source = eDBtype.MySQL
                DBtype_dest = eDBtype.MySQL
                txtSourceConn.Text = "Driver={MySQL ODBC 5.3 ANSI Driver};server=SlpLb-PC;uid=root;pwd=Grange00;database=respdb;auto_reconnect=1"
                txtDestinationConn.Text = "Driver={MySQL ODBC 5.3 ANSI Driver};server=SlpLb-PC;uid=root;pwd=Grange00;database=reslabdb;auto_reconnect=1"
                conn_source = Create_Connection(DBtype_source, txtSourceConn.Text)
                conn_dest = Create_Connection(DBtype_dest, txtDestinationConn.Text)
            Case "NRG"
                cmboDBType_source.Text = "SQL Server"
                cmboDBType_destination.Text = "MySQL"
                
                DBtype_source = eDBtype.SQLServer
                DBtype_dest = eDBtype.MySQL
                txtSourceConn.Text = "persist security info=false;user id=sa;initial catalog=RespiratoryMedicine;data source=VIRTUALXP-48358"
                txtDestinationConn.Text = "Driver={MySQL ODBC 5.3 ANSI Driver};server=NRG;uid=respdbuser;pwd=Grange00;database=reslabdb;auto_reconnect=1"
                conn_source = Create_Connection(DBtype_source, txtSourceConn.Text)
                conn_dest = Create_Connection(DBtype_dest, txtDestinationConn.Text)
            Case "MyLapTop"
                cmboDBType_source.Text = "MySQL"
                cmboDBType_destination.Text = "MySQL"

                DBtype_source = eDBtype.MySQL
                DBtype_dest = eDBtype.MySQL
                txtSourceConn.Text = "Driver={MySQL ODBC 5.3 ANSI Driver};host=SlpLb-PC;uid=root;pwd=Grange00;database=respdb;auto_reconnect=1"
                txtDestinationConn.Text = "Driver={MySQL ODBC 5.3 ANSI Driver};host=SlpLb-PC;uid=root;pwd=Grange00;database=reslabdb;auto_reconnect=1"
                'Driver={MySQL ODBC 5.3 ANSI Driver};host=Pete;uid=root;pwd=Grange00;database=reslabdb;auto_reconnect=1
                conn_source = Create_Connection(DBtype_source, txtSourceConn.Text)
                conn_dest = Create_Connection(DBtype_dest, txtDestinationConn.Text)
        End Select

    End Sub

    Private Sub btnTestSource_Click(sender As Object, e As System.EventArgs) Handles btnTestSource.Click

        Try
            conn_source.Open()
            MsgBox("Success")
        Catch ex As Exception
            MsgBox("Failed" & vbNewLine & ex.Message)
        End Try

    End Sub

    Private Sub btnTestDestination_Click(sender As System.Object, e As System.EventArgs) Handles btnTestDestination.Click

        Try
            conn_dest.Open()
            MsgBox("Success")
        Catch ex As Exception
            MsgBox("Failed" & vbNewLine & ex.Message)
        End Try

    End Sub

    Private Function Create_Connection(dbtype As eDBtype, conn_string As String) As DbConnection

        Dim Cn As DbConnection = Nothing

        Select Case dbtype
            Case eDBtype.SQLServer : Cn = New SqlConnection(conn_string)
            Case eDBtype.MySQL : Cn = New OdbcConnection(conn_string)
        End Select

        Return Cn

    End Function

    Private Function Create_Command(dbtype As eDBtype, ByVal CommandText As String, ByVal Cn As DbConnection)

        Dim cmd As DbCommand = Nothing
        Select Case DbType
            Case eDBtype.SQLServer : cmd = New SqlCommand(CommandText, Cn)
            Case eDBtype.MySQL : cmd = New OdbcCommand(CommandText, Cn)
        End Select

        Return cmd

    End Function

    Private Sub get_dbtype()

        Select Case cmboDBType_source.Text
            Case "SQL Server" : DBtype_source = eDBtype.SQLServer
            Case "MySQL" : DBtype_source = eDBtype.MySQL
        End Select
        Select Case cmboDBType_destination.Text
            Case "SQL Server" : DBtype_dest = eDBtype.SQLServer
            Case "MySQL" : DBtype_dest = eDBtype.MySQL
        End Select

    End Sub

    Private Function Get_dataset(dbtype As eDBtype, sqlstring As String, cn As DbConnection) As DataSet

        Dim Cmd As DbCommand = Me.Create_Command(dbtype, sqlstring, cn)
        Dim Da As DbDataAdapter = Me.Create_Adapter(dbtype, Cmd)

        If sqlstring = "" Then
            Return Nothing
        End If

        Try
            Dim Ds As New DataSet("MyData")
            Da.Fill(Ds)
            Return Ds
        Catch
            MsgBox("Error in cDAL.Get_DataAsDataset" & vbNewLine & Err.Description)
            Return Nothing
        End Try

    End Function

    Private Function Create_Adapter(dbtype As eDBtype, ByVal cmd As DbCommand) As DbDataAdapter

        Dim Da As DbDataAdapter = Nothing

        Select Case dbtype
            Case eDBtype.SQLServer : Da = New SqlDataAdapter(cmd)
            Case eDBtype.MySQL : Da = New OdbcDataAdapter(cmd)
        End Select

        Return Da

    End Function


    Private Sub btnMigrate_Click(sender As Object, e As System.EventArgs) Handles btnMigrate.Click
        'SQL server method
        'SET IDENTITY_INSERT dbo.Client ON 
        'INSERT INTO dbo.CLIENT (ClientID, ClientName) VALUES (782, 'Edgewood Solutions')
        'INSERT INTO dbo.CLIENT (ClientID, ClientName) VALUES (783, 'Microsoft')
        'SET IDENTITY_INSERT dbo.Client OFF

        'In MySQL simply specify the desired ID in the INSERT statement

        Dim i As Integer = 0, j As Integer = 0
        Dim sql As New StringBuilder
        Dim r As DataRow
        Dim d As New Dictionary(Of String, String)
        Dim cmdInsert As DbCommand = Nothing
        Dim mannitoldoses As New List(Of String) From {"", "0", "5", "15", "35", "75", "155", "315", "475", "635", ""}
        Dim fev1 As String = ""
        Dim response As Single = 0

        'Get source records
        sql.Append("SELECT * FROM challengedata")
        Dim ds As DataSet = Me.Get_dataset(DBtype_source, sql.ToString, conn_source)

        'Create prov test records
        If Not IsNothing(ds) Then
            'conn_dest.Open()
            txtLog.Text = txtLog.Text & vbNewLine & "Source records from table 'Challengedata' to process: " & ds.Tables(0).Rows.Count
            For Each r In ds.Tables(0).Rows
                'Prepare the new primary test row to insert
                d.Clear()
                d.Add("provid", r("challid"))
                d.Add("patientid", r("patientid"))
                d.Add("testdate", "'" & Format(CDate(r("testdate")), "yyyy-MM-dd") & "'")
                If Not IsDBNull(r("testtime")) Then d.Add("testtime", "'" & Format(CDate(r("testtime")), "HH:mm") & "'") Else d.Add("testtime", "NULL")
                d.Add("lab", "'" & r("status") & "'")
                d.Add("testtype", "'Mannitol Bronchoprovocation'")
                d.Add("height", "'" & r("height") & "'")
                d.Add("weight", "'" & r("weight") & "'")
                d.Add("req_name", "'" & r("Physician") & "'")
                d.Add("req_address", "'" & r("ReportTo") & "'")
                d.Add("req_date", "NULL")
                d.Add("req_fax", "NULL")
                d.Add("req_email", "NULL")
                d.Add("req_clinicalnotes", "'" & r("ClinicalNote") & "'")
                d.Add("report_text", "'" & r("Report") & "'")
                d.Add("report_copyto", "'" & r("cc") & "'")
                d.Add("report_reportedby", "'" & r("reporter") & "'")
                If Not IsDBNull(r("reportdate")) Then d.Add("report_reporteddate", "'" & Format(CDate(r("reportdate")), "yyyy-MM-dd") & "'") Else d.Add("report_reporteddate", "NULL")
                d.Add("report_verifiedby", "NULL")
                If Not IsDBNull(r("VerifyDate")) Then d.Add("report_verifieddate", "'" & Format(CDate(r("VerifyDate")), "yyyy-MM-dd") & "'") Else d.Add("report_verifieddate", "NULL")
                d.Add("report_status", "'" & r("reportstatus") & "'")
                d.Add("smoke_hx", "'" & r("smokehx") & "'")
                d.Add("smoke_cigsperday", "'" & r("cigsperday") & "'")
                d.Add("smoke_yearssmoked", "'" & r("yearssmoked") & "'")
                d.Add("smoke_packyears", "'" & r("pack years") & "'")
                d.Add("diagnosticcategory", "'" & r("DiagCat") & "'")
                d.Add("protocolid", "'" & "1" & "'")
                d.Add("pd_threshold", "'" & "15" & "'")
                d.Add("pd_decimalplaces", "'" & "0" & "'")
                d.Add("pd_method", "'" & "linear interpolation" & "'")
                d.Add("p_title", "'" & "Mannitol Bronchoprovocation" & "'")
                d.Add("p_agent_units", "'" & "mg" & "'")
                d.Add("p_agent", "'" & "Mannitol" & "'")
                d.Add("p_parameter", "'" & "FEV1" & "'")
                d.Add("p_parameter_units", "'" & "L" & "'")
                d.Add("p_parameter_response", "'" & "%" & "'")
                d.Add("p_dose_effect", "'" & "Cumulative" & "'")
                d.Add("p_method_reference", "'" & "Control" & "'")
                d.Add("p_post_drug", "'" & "BD" & "'")
                d.Add("plot_xscaling_type", "'" & "log" & "'")
                d.Add("plot_xtitle", "'" & "Log cumulative dose (mg)" & "'")
                d.Add("plot_title", "'" & "FEV1 (% of control)" & "'")
                d.Add("plot_ymin", "'" & "50" & "'")
                d.Add("plot_ymax", "'" & "120" & "'")
                d.Add("plot_ystep", "'" & "10" & "'")
                d.Add("pd", "NULL")
                d.Add("r_bl_fev1", "'" & r("preFEV1") & "'")
                d.Add("r_bl_fvc", "'" & r("preFVC") & "'")
                d.Add("r_bl_vc", "NULL")
                d.Add("r_bl_fer", "'" & r("preFER") & "'")
                d.Add("r_bl_fef2575", "NULL")
                d.Add("r_bl_pef", "NULL")
                d.Add("bdstatus", "'" & r("BDstatus") & "'")
                d.Add("flowvolLoop", "NULL")
                d.Add("lastupdate", "'" & Format(Now, "yyyy-MM-dd") & "'")
                d.Add("lasteditby", "'" & "PDR" & "'")
                d.Add("technicalnote", "'" & r("Comment") & "'")
                d.Add("pred_sourceids", "NULL")
                d.Add("scientist", "'" & r("TestedBy") & "'")
                d.Add("admissionstatus", "'" & r("ward") & "'")
      
                'Insert the new row
                sql.Clear()
                sql.Append(Me.Build_InsertQuery("prov_test", d))
                cmdInsert = Me.Create_Command(DBtype_dest, sql.ToString, conn_dest)

                Try
                    cmdInsert.ExecuteNonQuery()
                    cmdInsert = Nothing
                    txtLog.AppendText(vbNewLine & i & "     ChallID: " & r("challid") & " inserted")
                    i = i + 1

                    'Create the new secondary dose result records - 11 records per mannitol
                    For j = 1 To 11
                        d.Clear()
                        d.Add("provid", r("challid"))
                        d.Add("doseid", j)
                        d.Add("dose_time_min", "'" & "0" & "'")
                        If j = 11 Then d.Add("dose_number", 100) Else d.Add("dose_number", j - 2)
                        d.Add("dose_discrete", "NULL")
                        d.Add("dose_cumulative", "'" & mannitoldoses(j - 1) & "'")
                        d.Add("dose_xaxis_label", "'" & mannitoldoses(j - 1) & "'")
                        d.Add("dose_canskip", "NULL")
                        Select Case j
                            Case 1 : fev1 = r("preFEV1") & ""
                            Case 2 : fev1 = r("1FEV1") & ""
                            Case 3 : fev1 = r("2FEV1") & ""
                            Case 4 : fev1 = r("3FEV1") & ""
                            Case 5 : fev1 = r("4FEV1") & ""
                            Case 6 : fev1 = r("5FEV1") & ""
                            Case 7 : fev1 = r("6FEV1") & ""
                            Case 8 : fev1 = r("7FEV1") & ""
                            Case 9 : fev1 = r("8FEV1") & ""
                            Case 10 : fev1 = r("9FEV1") & ""
                            Case 11 : fev1 = r("BDFEV1") & ""
                        End Select
                        d.Add("result", "'" & fev1 & "'")
                        If IsNumeric(fev1) And IsNumeric(r("1FEV1")) Then
                            response = Format(100 * fev1 / Val(r("1FEV1")), "0.0")
                            d.Add("response", "'" & response & "'")
                        Else
                            d.Add("response", "NULL")
                        End If

                        'Insert the new row
                        sql.Clear()
                        sql.Append(Me.Build_InsertQuery("prov_testdata", d))
                        Try
                            cmdInsert = Me.Create_Command(DBtype_dest, sql.ToString, conn_dest)
                            cmdInsert.ExecuteNonQuery()
                            cmdInsert = Nothing
                            txtLog.AppendText(vbNewLine & "    " & j & ":  test data record inserted")

                        Catch
                            MsgBox("Error inserting record in prov_testdata" & vbNewLine & Err.Description)
                        End Try
                    Next

                Catch
                    MsgBox("Error inserting record in prov_test. ChallID=" & r("challid") & vbNewLine & Err.Description)
                End Try
            Next
        Else
            MsgBox("Nothing returned from source table, process aborted.")
            Exit Sub
        End If

        conn_dest.Close()

    End Sub

    Private Sub cmboDBType_destination_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmboDBType_destination.SelectedValueChanged

        Select Case cmboDBType_destination.Text
            Case "MySQL" : DBtype_dest = eDBtype.MySQL
            Case "SQLServer" : DBtype_dest = eDBtype.SQLServer
        End Select

    End Sub

    Private Sub cmboDBType_source_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmboDBType_source.SelectedValueChanged

        Select Case cmboDBType_source.Text
            Case "MySQL" : DBtype_dest = eDBtype.MySQL
            Case "SQLServer" : DBtype_dest = eDBtype.SQLServer
        End Select

    End Sub

    Public Function Build_InsertQuery(Tbl As String, ByVal D As Dictionary(Of String, String)) As String

        Dim sb As New StringBuilder
        Dim val As String = ""
        Dim kv As KeyValuePair(Of String, String) = Nothing
        Dim kv1 As KeyValuePair(Of String, String) = Nothing

        sb.Clear()
        sb.Append("INSERT INTO " & Tbl & " (")
        For Each kv In D         
            sb.Append(kv.Key & ", ")
        Next
        sb.Remove(sb.Length - 2, 2)

        sb.Append(") VALUES (")
        For Each kv1 In D
            If kv1.Value = Nothing Or kv1.Value = "NULL" Then
                sb.Append("NULL,")
            Else
                'Change single apostrophes in text to '' eg 'O'Brien' to  'O''Brien'
                If Strings.Left(kv1.Value, 1) = "'" And Strings.Right(kv1.Value, 1) = "'" Then
                    val = Replace(Mid(kv1.Value, 2, Len(kv1.Value) - 2), "'", "''")
                    val = "'" & val & "'"
                Else
                    val = kv1.Value
                End If
                sb.Append(val & ",")
            End If
        Next
        sb.Remove(sb.Length - 1, 1)
        sb.Append(");")

        Return sb.ToString

    End Function

    Private Sub form_migrate_jjp_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnMigrateFlowVols_Click(sender As System.Object, e As System.EventArgs) Handles btnMigrateFlowVols.Click
       
        Dim ds1 As DataSet = Nothing
        Dim CmdPlaceholder As String = ""
        Dim cmd As DbCommand
        Dim Pr As DbParameter
        Dim i As Integer = 1
        Dim bImage As Byte() = Nothing

        'Get source records
        txtLog.Clear()
        txtLog.Text = "Getting source records" & vbNewLine
        Dim Sql As String = "SELECT rftid,FlowVolLoop FROM respdb.rftdata where length(FlowVolLoop)>0;"
        Dim ds As DataSet = Me.Get_dataset(DBtype_source, Sql, conn_source)
        If Not IsNothing(ds) Then

            For Each r In ds.Tables(0).Rows
                txtLog.Text = txtLog.Text & vbNewLine & "Source record flow vols to process: " & i & "/" & ds.Tables(0).Rows.Count

                'Get corresponding destination record
                'Sql = "SELECT FlowVolLoop FROM reslabdb_mrss_july2016migration.rft_routine where rftid=" & r("rftid")
                'ds1 = Me.Get_dataset(DBtype_dest, Sql, conn_dest)

                'Save image data to a byte array
                'If ds1.Tables(0).Rows.Count = 1 Then
                ReDim bImage(0)
                bImage = CType(ds.Tables(0).Rows(0)("FlowVolLoop"), Byte())

                Select Case DBtype_dest
                    Case eDBtype.MySQL : CmdPlaceholder = "?"
                    Case eDBtype.SQLServer : CmdPlaceholder = "@BLOBData"
                End Select
                Sql = "UPDATE rft_routine SET FlowVolLoop=" & CmdPlaceholder & " WHERE rftid=" & r("rftid")
                cmd = Me.Create_Command(DBtype_dest, Sql, conn_dest)
                Pr = cDAL.Create_Parameter
                Pr.ParameterName = "?"
                Pr.DbType = DbType.Binary
                Pr.Direction = ParameterDirection.Input
                Pr.Value = bImage
                cmd.Parameters.Add(Pr)
                cmd.ExecuteNonQuery()
                cmd = Nothing
                Pr = Nothing
                'End If
                'ds1 = Nothing
                i = i + 1
            Next

        End If

        conn_dest.Close()
        conn_source.Close()

        MsgBox("Completed")

    End Sub

End Class