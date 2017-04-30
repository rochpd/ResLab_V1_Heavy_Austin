Imports System
Imports System.Data
Imports System.Text
Imports System.IO
Imports System.Linq.Expressions
Imports ResLab_V1_Heavy.cDatabaseInfo




Public Enum eImageTransferMode
    DatabaseToPicbox
    PicboxToDatabase
End Enum

Public Structure TestsDone
    Dim TestsDoneString As String
    Dim pre_done As Boolean
    Dim post_done As Boolean
    Dim anyspir_done As Boolean
    Dim tlco_done As Boolean
    Dim vols_done As Boolean
    Dim raw_done As Boolean
    Dim abgs1_done As Boolean
    Dim abgs2_done As Boolean
    Dim anyabgs_done As Boolean
    Dim shunt_done As Boolean
    Dim mrps_done As Boolean
    Dim oximetry_done As Boolean
    Dim feno_done As Boolean
End Structure

Public Class cGeneralRoutines

    Public Function Get_TempDirectory() As String

        Dim s As String = My.Computer.FileSystem.SpecialDirectories.Temp
        Return s

    End Function

    Public Function fmt(ByVal s As String, ByVal Places As Integer) As String
        'Takes a number as string and formats it to x decimal places and returns as string

        If Val(s) > 0 Then
            Dim p As String = "0." & Strings.StrDup(Places, "0")
            Return Format(Val(s), p)
        Else
            Return ""
        End If

    End Function

    Public Function Get_Hast_Protocol(Optional protocolID As Integer = 0) As class_hast_protocoldata

        Dim sql As String = ""
        Dim h As New class_hast_protocoldata

        Select Case protocolID
            Case 0 : sql = "SELECT * FROM r_hast_protocols where protocol_enabled=1"
            Case Else : sql = "SELECT * FROM r_hast_protocols where protocolid=" & protocolID
        End Select

        Try
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If Not IsNothing(Ds) Then
                With Ds.Tables(0)
                    If .Rows.Count = 1 Then
                        h.protocolID = .Rows(0)("protocolID")
                        h.description = .Rows(0)("description") & ""
                        h.pMenuLabel = .Rows(0)("pMenuLabel") & ""
                        h.deliverymode_fio2 = .Rows(0)("deliverymode_fio2") & ""
                        h.deliverymode_suppo2 = .Rows(0)("deliverymode_suppo2") & ""
                        h.fio2_1 = .Rows(0)("fio2_1") & ""
                        h.fio2_2 = .Rows(0)("fio2_2") & ""
                        h.fio2_3 = .Rows(0)("fio2_3") & ""
                        h.fio2_4 = .Rows(0)("fio2_4") & ""
                        h.fio2_5 = .Rows(0)("fio2_5") & ""
                        h.fio2_6 = .Rows(0)("fio2_6") & ""
                        If Not IsDBNull(.Rows(0)("fio2_1_enabled")) Then h.fio2_1_enabled = .Rows(0)("fio2_1_enabled") Else h.fio2_1_enabled = False
                        If Not IsDBNull(.Rows(0)("fio2_2_enabled")) Then h.fio2_2_enabled = .Rows(0)("fio2_2_enabled") Else h.fio2_2_enabled = False
                        If Not IsDBNull(.Rows(0)("fio2_3_enabled")) Then h.fio2_3_enabled = .Rows(0)("fio2_3_enabled") Else h.fio2_3_enabled = False
                        If Not IsDBNull(.Rows(0)("fio2_4_enabled")) Then h.fio2_4_enabled = .Rows(0)("fio2_4_enabled") Else h.fio2_4_enabled = False
                        If Not IsDBNull(.Rows(0)("fio2_5_enabled")) Then h.fio2_5_enabled = .Rows(0)("fio2_5_enabled") Else h.fio2_5_enabled = False
                        If Not IsDBNull(.Rows(0)("fio2_6_enabled")) Then h.fio2_6_enabled = .Rows(0)("fio2_6_enabled") Else h.fio2_6_enabled = False
                        h.fio2_1_altitude = .Rows(0)("fio2_1_altitude") & ""
                        h.fio2_2_altitude = .Rows(0)("fio2_2_altitude") & ""
                        h.fio2_3_altitude = .Rows(0)("fio2_3_altitude") & ""
                        h.fio2_4_altitude = .Rows(0)("fio2_4_altitude") & ""
                        h.fio2_5_altitude = .Rows(0)("fio2_5_altitude") & ""
                        h.fio2_6_altitude = .Rows(0)("fio2_6_altitude") & ""
                        If Not IsDBNull(.Rows(0)("abg_enabled")) Then h.abg_enabled = .Rows(0)("abg_enabled") Else h.abg_enabled = False
                        If Not IsDBNull(.Rows(0)("abg_part_enabled")) Then h.abg_part_enabled = .Rows(0)("abg_part_enabled") Else h.abg_part_enabled = False
                        If Not IsDBNull(.Rows(0)("protocol_enabled")) Then h.protocol_enabled = .Rows(0)("protocol_enabled") Else h.protocol_enabled = False
                        h.lasteditby = .Rows(0)("lasteditby") & ""
                        If IsDate(.Rows(0)("lastedited")) Then h.lastedited = .Rows(0)("lastedited") Else h.lastedited = ""
                    Else

                    End If
                End With
                Return h
            Else
                MsgBox("Error retrieving list of units in cGeneralRoutines.Get_ListOfUnits" & vbCrLf & Err.Description, vbOKOnly, "Error")
                Return Nothing
            End If
        Catch
            MsgBox("Error retrieving list of units in cGeneralRoutines.Get_ListOfUnits" & vbCrLf & Err.Description, vbOKOnly, "Error")
            Return Nothing
        End Try



    End Function


    Public Function Get_ListOfUnits(OnlyEnabled As Boolean, longname As Boolean) As Dictionary(Of String, Integer)

        Try
            Dim name As String = ""
            Dim i As Integer
            If longname Then name = "description" Else name = "shortname"
            Dim dic As New Dictionary(Of String, Integer)

            Dim sql As String = "SELECT * FROM list_units "
            If OnlyEnabled Then sql = sql & " WHERE enabled = 1"

            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If Not IsNothing(Ds) Then
                With Ds.Tables(0)
                    For i = 0 To .Rows.Count - 1
                        dic.Add(.Rows(i).Item(name).ToString, .Rows(i).Item("code"))
                    Next
                End With
                Return dic
            Else
                MsgBox("Error retrieving list of units in cGeneralRoutines.Get_ListOfUnits" & vbCrLf & Err.Description, vbOKOnly, "Error")
                Return Nothing
            End If
        Catch
            MsgBox("Error retrieving list of units in cGeneralRoutines.Get_ListOfUnits" & vbCrLf & Err.Description, vbOKOnly, "Error")
            Return Nothing
        End Try

    End Function

    Public Function Check_ForDuplicatePatient(Surname As String, Firstname As String, DOB As Date, Gender As String) As Long()

        Try
            Dim PatientIDs() As Long = Nothing
            Surname = Replace(Surname, "'", "''")

            Dim sql As String = "SELECT patientid FROM patientdetails where Surname='" & Surname & "' and DOB='" & Format(DOB, "yyyy-MM-dd") & "' and gender='" & Gender & "'"
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If IsNothing(Ds) Or Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                ReDim PatientIDs(-1)
                For Each dr As DataRow In Ds.Tables(0).Rows
                    ReDim Preserve PatientIDs(UBound(PatientIDs) + 1)
                    PatientIDs(UBound(PatientIDs)) = dr(0)
                Next
                Return PatientIDs
            End If
        Catch
            MsgBox("Error in class_Routines.Check_ForDuplicatePatient" & vbCrLf & Err.Description, vbOKOnly, "Check for duplicate patient")
            Return Nothing
        End Try

    End Function

    Public Function WriteToFile_DataGrid(ByVal dgv As DataGridView) As String
        'Returns the file name/path

        Dim filePath As String = Application.StartupPath.ToString & "\ResLab_LastPredPreferences.log"
        Dim headers = (From header As DataGridViewColumn In dgv.Columns.Cast(Of DataGridViewColumn)() Select header.HeaderText).ToArray
        Dim rows = From row As DataGridViewRow In dgv.Rows.Cast(Of DataGridViewRow)() Where Not row.IsNewRow _
                     Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))

        Using sw As IO.StreamWriter = File.AppendText(filePath)
            sw.WriteLine()
            sw.WriteLine("Predicted value preferences just prior to the last update")
            sw.WriteLine("Updated: " & Now())
            sw.WriteLine(String.Join(",", headers))
            For Each r In rows
                sw.WriteLine(String.Join(",", r))
            Next
        End Using
        Return filePath

    End Function

    Public Sub Combo_SelectItem(ByRef cmbo As ComboBox, ByVal ItemString As String, Optional ByVal AddIfNotFound As Boolean = False)

        Dim i As Integer = -1
        i = cmbo.FindStringExact(ItemString)
        Select Case i
            Case Is >= 0
                cmbo.SelectedIndex = i
            Case Else
                If AddIfNotFound Then
                    cmbo.Items.Add(ItemString)
                    cmbo.SelectedIndex = cmbo.FindStringExact(ItemString)
                Else
                    cmbo.SelectedIndex = -1
                End If
        End Select
    End Sub

    Public Sub Combo_FillFromCollection(ByVal cmbo As ComboBox, ByVal c As Collection, Optional ByVal SortItems As Boolean = False)

        For Each Thing In c
            cmbo.Items.Add(Thing)
        Next
        cmbo.Sorted = SortItems

    End Sub

    Public Function Combo_GetKey(ByVal cmbo As ComboBox) As Integer

        Dim kv = CType(cmbo.SelectedItem, KeyValuePair(Of String, String))
        Return kv.Key

    End Function

    Public Function Combo_GetValue(ByVal cmbo As ComboBox) As String

        Dim kv = CType(cmbo.SelectedItem, KeyValuePair(Of String, String))
        Return kv.Value

    End Function

    Public Function Combo_LoadItems(ByVal cmbo As ComboBox, ByVal items As Dictionary(Of String, String)) As Boolean

        Try
            cmbo.Items.Clear()
            For Each kv As KeyValuePair(Of String, String) In items
                cmbo.Items.Add(kv)
            Next
            cmbo.DisplayMember = "value"
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function Combo_LoadItemsFromPrefs(ByVal cmbo As ComboBox, ByVal field As String, Optional ByVal SelectDefault As Boolean = False) As Boolean

        'Field can be passed as either ID or description
        Dim fieldID As String = ""
        If IsNumeric(field) Then
            fieldID = field
        Else
            fieldID = cPrefs.get_fieldID_fromFieldName(field)
            If IsNothing(fieldID) Then Return False
        End If

        Try
            cmbo.Items.Clear()
            cmbo.DisplayMember = "value"
            Dim items As Dictionary(Of String, String) = cPrefs.Get_FieldItemsForFieldID(fieldID, False)
            For Each kv As KeyValuePair(Of String, String) In items
                cmbo.Items.Add(kv)
            Next

            'Do default
            If SelectDefault Then
                Dim DefaultOptionID As Integer = cPrefs.Get_DefaultFieldOptionID(fieldID)
                If DefaultOptionID > 0 Then
                    Dim DefaultOptionTxt As String = cPrefs.Get_FieldOptionTextForID(DefaultOptionID)
                    If DefaultOptionTxt <> "" Then
                        cmbo.SelectedIndex = cmbo.FindStringExact(DefaultOptionTxt)
                    End If
                End If
            End If
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function Combo_ts_LoadItemsFromPrefs(ByVal cmbo As ToolStripComboBox, ByVal Fieldname As String) As Boolean
        'For a tool strip combo box, note that can't load the primary key along with the text in these controls

        Try
            cmbo.Items.Clear()
            Dim items As Dictionary(Of String, String) = cPrefs.Get_FieldItemsForFieldID(cPrefs.get_fieldID_fromFieldName(Fieldname))
            For Each kv As KeyValuePair(Of String, String) In items
                cmbo.Items.Add(kv.Value)
            Next
            'cmbo.DisplayMember = "value"
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function Combo_LoadItemsFromList(ByVal cmbo As ComboBox, ByVal eTbl As eTables, Optional EnabledOnly As Boolean = False) As Boolean

        Try
            cmbo.Items.Clear()
            Dim items As Dictionary(Of String, String) = Me.get_items_listTable(eTbl, EnabledOnly)
            For Each kv As KeyValuePair(Of String, String) In items
                cmbo.Items.Add(kv)
            Next
            cmbo.DisplayMember = "value"
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function Combo_ts_LoadItemsFromList(ByVal cmbo As ToolStripComboBox, ByVal eTbl As eTables, Optional EnabledOnly As Boolean = False) As Boolean
        'For a tool strip combo box, note that can't load the primary key along with the text in these controls

        Try
            cmbo.Items.Clear()
            Dim items As Dictionary(Of String, String) = Me.get_items_listTable(eTbl, EnabledOnly)
            For Each kv As KeyValuePair(Of String, String) In items
                cmbo.Items.Add(kv.Value)
            Next
            'cmbo.DisplayMember = "value"
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function Combo_ts_LoadRequestingMOPermutations(ByVal cmbo As ToolStripComboBox) As Boolean

        Dim sql As String = "SELECT DISTINCT Req_name FROM r_sessions ORDER BY Req_name"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            cmbo.Items.Clear()
            For Each r As DataRow In Ds.Tables(0).Rows
                cmbo.Items.Add(r.Item(0))
            Next
        End If

        Return True

    End Function

    Public Function get_rfts_requested(Sdate As Date, fDate As Date) As String()

        Dim s() As String = Nothing
        Dim i As Integer = 0
        Dim sql As New StringBuilder
        Dim eTbls() As eTables = {eTables.rft_routine, eTables.r_cpet, eTables.Prov_test, eTables.r_walktests_v1heavy, eTables.r_hast, eTables.r_spt}

        sql.Clear()
        For i = 0 To UBound(eTbls)
            sql.Append("SELECT r_sessions.testdate, r_sessions.req_name FROM " & cDbInfo.table_name(eTbls(i)) & " INNER JOIN r_sessions ON " & cDbInfo.table_name(eTbls(i)) & ".SessionID = r_sessions.sessionID ")
            sql.Append("WHERE (r_sessions.testdate BETWEEN '" & Format(Sdate, "yyyy-MM-dd") & "' AND '" & Format(fDate, "yyyy-MM-dd") & "')   ")
            If i < UBound(eTbls) Then sql.Append(" UNION ALL ")
        Next
        sql.Append(" ORDER BY TestDate")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count > 0 Then
                For Each r As DataRow In Ds.Tables(0).Rows
                    ReDim Preserve s(i)
                    s(i) = r.Item(0)
                    i = i + 1
                Next
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

        Return s

    End Function

    Public Function get_RequestMOPermutations_rft(Sdate As Date, fDate As Date) As String()

        Dim s() As String = Nothing
        Dim i As Integer = 0
        Dim sql As String
        sql = "SELECT DISTINCT req_name FROM r_sessions WHERE testdate >= '" & Format(Sdate, "yyyy-MM-dd") & "' AND testdate <= '" & Format(fDate, "yyyy-MM-dd") & "'"
        sql = sql & " ORDER BY req_name"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) And Ds.Tables(0).Rows.Count > 0 Then
            For Each r As DataRow In Ds.Tables(0).Rows
                ReDim Preserve s(i)
                s(i) = r.Item(0)
                i = i + 1
            Next
        Else
            Return Nothing
        End If

        Return s

    End Function

    Public Function get_prefs_showfaxstatementonreports() As Boolean

        Dim i As Integer = 0
        Dim sql As String
        sql = "SELECT fielditem FROM prefs_fielditems INNER JOIN "
        sql = sql & "prefs_fields ON prefs_fielditems.prefs_id = prefs_fields.default_fielditem_id "
        sql = sql & "WHERE prefs_fields.fieldname ='ShowFaxStatementOnReports'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count > 0 Then
                Return Ds.Tables(0).Rows(0).Item(0)
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Public Function get_items_listTable(ByVal eTbl As eTables, Optional EnabledOnly As Boolean = False) As Dictionary(Of String, String)
        'Assumes fields named 'description' and 'code' are present in the table. 
        'Returns a dictionary object with elements 'code' as key and 'description' as value
        'Checks for duplicate keys

        Dim Msg As String = ""

        Try
            'First check for duplicate keys - not allowed
            If cDAL.check_duplicatesInColumn("code", eTables.list_Nationality) = 0 Then

                Dim sql As String = "SELECT * FROM " & cDbInfo.table_name(eTbl) & " WHERE code is not null AND code<>'' "
                If EnabledOnly Then sql = sql & " AND enabled=1"
                Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

                If IsNothing(Ds) Then
                    Return Nothing
                ElseIf Ds.Tables(0).Rows.Count = 0 Then
                    Return Nothing
                Else
                    Dim d As New Dictionary(Of String, String)
                    For Each r As DataRow In Ds.Tables(0).Rows
                        d.Add(r.Item("code"), r.Item("description") & "")
                    Next
                    Return d
                End If
            Else
                Msg = "Duplicate codes found in lookup table '" & cDbInfo.table_name(eTbl) & "'." & vbCrLf
                Msg = Msg & "Please note this message and contact your database adminstrator to rectify."
                MsgBox(Msg, vbOKOnly, "ResLab")
                Return Nothing
            End If

        Catch ex As Exception
            Select Case ex.Message
                Case ""
                    Msg = ""
                Case Else
                    Msg = "Error in class_routines.Get_ListTable_Items" & vbCrLf & ex.Message
            End Select
            MsgBox(Msg, vbOKOnly, "ResLab")
            Return Nothing

        End Try

    End Function

    Public Function Listbox_LoadItems(ByVal lst As ListBox, ByVal items As Dictionary(Of String, String)) As Boolean

        Try
            lst.Items.Clear()
            For Each kv As KeyValuePair(Of String, String) In items
                lst.Items.Add(kv)
            Next
            lst.DisplayMember = "value"
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function Listbox_LoadItemsFromPrefs(ByVal lst As Object, ByVal FieldName As String, Optional ByVal SelectDefault As Boolean = False) As Boolean
        'Handles both standard and checked list boxes

        Try
            If lst.GetType = GetType(ListBox) Or lst.GetType = GetType(CheckedListBox) Then
                lst.Items.Clear()
                lst.DisplayMember = "value"
                Dim items As Dictionary(Of String, String) = cPrefs.Get_FieldItemsForFieldID(cPrefs.get_fieldID_fromFieldName(FieldName), False)
                For Each kv As KeyValuePair(Of String, String) In items
                    lst.Items.Add(kv)
                Next

                'Do default
                If SelectDefault Then
                    Dim DefaultOptionID As Integer = cPrefs.Get_DefaultFieldOptionID(FieldName)
                    If DefaultOptionID > 0 Then
                        Dim DefaultOptionTxt As String = cPrefs.Get_FieldOptionTextForID(DefaultOptionID)
                        If DefaultOptionTxt <> "" Then
                            lst.SelectedIndex = lst.FindStringExact(DefaultOptionTxt)
                        End If
                    End If
                End If
                Return True
            Else
                Return False
            End If

        Catch
            Return False
        End Try

    End Function

    Public Function Listbox_GetKey(ByVal lst As ListBox) As Integer

        Dim kv = CType(lst.SelectedItem, KeyValuePair(Of String, Integer))
        Return kv.Key

    End Function

    Public Function Listbox_GetValue(ByVal lst As ListBox) As String

        Dim kv = CType(lst.SelectedItem, KeyValuePair(Of String, Integer))
        Return kv.Value

    End Function

    Public Function Get_AppString(Name As String) As String

        Try
            Dim sql As String = "SELECT value FROM prefs_app_strings WHERE name='" & Name & "'"
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0).Item(0)
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try

    End Function

    Public Function Update_AppString(name As String, value As String) As Boolean

        Dim sqlString As String = "UPDATE prefs_app_strings SET value='" & value & "' WHERE name='" & name & "'"

        Try
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving challenge" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function


    Public Function IsRealDate(ByVal aDate) As Boolean

        If IsDate(aDate) Then
            If CDate(aDate).ToString("dd/MM/yyyy") = "01/01/0001" Then
                Return False
            Else
                Return True
            End If
        Else
            Return False
        End If

    End Function

    Public Function FormatDBDate(ByVal aDate) As String
        'format for insert sql statement

        Dim s As String = Nothing

        If IsRealDate(aDate) Then
            Return "'" & Format(CDate(aDate), "yyyy/MM/dd") & "'"
        Else
            Return "NULL"
        End If

    End Function


    Public Function FormatDBTime(ByVal aTime) As String
        'format for insert sql statement

        Dim s As String = Nothing

        If IsDate(aTime) Then
            Return "'" & Format(CDate(aTime), "HH:mm") & "'"
        Else
            Return "NULL"
        End If

    End Function

    Public Function FormatDBDateTime(ByVal aDate) As String
        'format for insert sql statement

        Dim s As String = Nothing

        If IsRealDate(aDate) Then
            Return "'" & Format(CDate(aDate), "yyyy/MM/dd HH:mm:ss") & "'"
        Else
            Return "NULL"
        End If

    End Function

    Public Function Get_TestsDone(ByVal R As Dictionary(Of String, String)) As TestsDone

        Dim flds As New class_Rft_RoutineAndSessionFields
        Dim td As TestsDone

        td.pre_done = False
        td.post_done = False
        td.anyspir_done = False
        td.tlco_done = False
        td.vols_done = False
        td.raw_done = False
        td.abgs1_done = False
        td.abgs2_done = False
        td.anyabgs_done = False
        td.shunt_done = False
        td.mrps_done = False
        td.oximetry_done = False
        td.feno_done = False

        If R.ContainsKey(flds.R_bl_Fev1) And R.ContainsKey(flds.R_bl_Fvc) And R.ContainsKey(flds.R_bl_Vc) Then
            td.pre_done = Val(Replace(R(flds.R_bl_Fev1), "'", "")) > 0 Or Val(Replace(R(flds.R_bl_Fvc), "'", "")) > 0 Or Val(Replace(R(flds.R_bl_Vc), "'", "")) > 0
        End If
        If R.ContainsKey(flds.R_post_Fev1) And R.ContainsKey(flds.R_post_Fvc) And R.ContainsKey(flds.R_post_Vc) Then
            td.post_done = Val(Replace(R(flds.R_post_Fev1), "'", "")) > 0 Or Val(Replace(R(flds.R_post_Fvc), "'", "")) > 0 Or Val(Replace(R(flds.R_post_Vc), "'", "")) > 0
        End If
        td.anyspir_done = td.pre_done Or td.post_done
        If R.ContainsKey(flds.R_Bl_Tlco) Then
            td.tlco_done = Val(Replace(R(flds.R_Bl_Tlco), "'", "")) > 0
        End If
        If R.ContainsKey(flds.R_Bl_Tlc) And R.ContainsKey(flds.R_Bl_Frc) And R.ContainsKey(flds.R_Bl_Rv) Then
            td.vols_done = Val(Replace(R(flds.R_Bl_Tlc), "'", "")) > 0 Or Val(Replace(R(flds.R_Bl_Frc), "'", "")) > 0 Or Val(Replace(R(flds.R_Bl_Rv), "'", "")) > 0
        End If
        If R.ContainsKey(flds.R_Bl_Mip) And R.ContainsKey(flds.R_Bl_Mep) Then
            td.mrps_done = Val(Replace(R(flds.R_Bl_Mip), "'", "")) > 0 Or Val(Replace(R(flds.R_Bl_Mep), "'", "")) > 0
        End If
        If R.ContainsKey(flds.R_SpO2_1) Then
            td.oximetry_done = Val(Replace(R(flds.R_SpO2_1), "'", "")) > 0
        End If
        If R.ContainsKey(flds.R_Bl_FeNO) Then
            td.feno_done = Val(Replace(R(flds.R_Bl_FeNO), "'", "")) > 0
        End If
        If R.ContainsKey(flds.R_abg1_ph) And R.ContainsKey(flds.R_abg1_paco2) And R.ContainsKey(flds.R_abg1_pao2) Then
            td.abgs1_done = Val(Replace(R(flds.R_abg1_ph), "'", "")) > 0 Or Val(Replace(R(flds.R_abg1_paco2), "'", "")) > 0 Or Val(Replace(R(flds.R_abg1_pao2), "'", "")) > 0
            td.abgs2_done = Val(Replace(R(flds.R_abg2_ph), "'", "")) > 0 Or Val(Replace(R(flds.R_abg2_paco2), "'", "")) > 0 Or Val(Replace(R(flds.R_abg2_pao2), "'", "")) > 0
            td.anyabgs_done = td.abgs1_done Or td.abgs2_done
        End If
        If R.ContainsKey(flds.R_abg1_shunt) And R.ContainsKey(flds.R_abg1_shunt) Then
            td.shunt_done = Val(Replace(R(flds.R_abg1_shunt), "'", "")) > 0 Or Val(Replace(R(flds.R_abg2_shunt), "'", "")) > 0
        End If

        'Build the test type string
        Dim sb As New StringBuilder
        sb.Clear()
        sb.Append("RFTs  (")
        If td.anyspir_done Then sb.Append("Sp")
        If td.pre_done Then sb.Append("1")
        If td.post_done Then sb.Append("2")
        If td.anyspir_done Then sb.Append(" ")
        If td.tlco_done Then sb.Append("TL ")
        If td.vols_done Then sb.Append("LV ")
        If td.mrps_done Then sb.Append("MRP ")
        If td.oximetry_done Then sb.Append("Ox ")
        If td.feno_done Then sb.Append("FeNO ")
        If td.shunt_done Then sb.Append("Sh ")
        If td.abgs1_done Or td.abgs2_done Then
            sb.Append("ABG")
            If td.abgs1_done Then sb.Append("1")
            If td.abgs2_done Then sb.Append("2")
            sb.Append(" ")
        End If

        If sb.ToString = "RFTs  (" Then
            sb.Append(")")
        Else
            sb.Remove(sb.Length - 1, 1)
            sb.Append(")")
        End If
        td.TestsDoneString = sb.ToString

        Return td

    End Function

    Public Function MakeEmpty_dicUR() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)
        Dim flds As New class_fields_ur

        D.Add(flds.UR_id, Nothing)
        D.Add(flds.patientID, Nothing)
        D.Add(flds.UR, Nothing)
        D.Add(flds.UR_mergeto, Nothing)
        D.Add(flds.UR_hsid, Nothing)
        D.Add(flds.created_inResLab, Nothing)
        D.Add(flds.create_by, Nothing)
        D.Add(flds.create_date, Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dicDemo() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)
        Dim flds As New class_DemographicFields

        D.Add(flds.PatientID, Nothing)
        D.Add(flds.Surname, Nothing)
        D.Add(flds.Firstname, Nothing)
        D.Add(flds.Title, Nothing)
        D.Add(flds.UR, Nothing)
        D.Add(flds.URandHS, Nothing)
        D.Add(flds.URandHS_all_asString, Nothing)
        D.Add(flds.UR_id, Nothing)
        D.Add(flds.UR_hsid, Nothing)
        D.Add(flds.DOB, Nothing)
        D.Add(flds.Gender, Nothing)
        D.Add(flds.Gender_code, Nothing)
        D.Add(flds.Gender_forRfts, Nothing)
        D.Add(flds.Gender_forRfts_code, Nothing)
        D.Add(flds.Email, Nothing)
        D.Add(flds.Address_1, Nothing)
        D.Add(flds.Address_2, Nothing)
        D.Add(flds.Suburb, Nothing)
        D.Add(flds.State, Nothing)
        D.Add(flds.PostCode, Nothing)
        D.Add(flds.Phone_home, Nothing)
        D.Add(flds.Phone_mobile, Nothing)
        D.Add(flds.Phone_work, Nothing)
        D.Add(flds.Medicare_No, Nothing)
        D.Add(flds.Medicare_expirydate, Nothing)
        D.Add(flds.Race, Nothing)
        D.Add(flds.Race_code, Nothing)
        D.Add(flds.Race_forRfts, Nothing)
        D.Add(flds.Race_forRfts_code, Nothing)
        D.Add(flds.PreferredLanguage, Nothing)
        D.Add(flds.PreferredLanguage_code, Nothing)
        D.Add(flds.Death_date, Nothing)
        D.Add(flds.Death_indicator, Nothing)
        D.Add(flds.CountryOfBirth, Nothing)
        D.Add(flds.CountryOfBirth_code, Nothing)
        D.Add(flds.AboriginalStatus, Nothing)
        D.Add(flds.AboriginalStatus_code, Nothing)
        D.Add(flds.gpID, Nothing)
        D.Add(flds.Lastupdated_date, Nothing)
        D.Add(flds.Lastupdated_by, Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dicUser() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)
        Dim flds As New class_Fields_User

        D.Add(flds.personID, Nothing)
        D.Add(flds.Surname, Nothing)
        D.Add(flds.Firstname, Nothing)
        D.Add(flds.Title, Nothing)
        D.Add(flds.DOB, Nothing)
        D.Add(flds.Gender, Nothing)
        D.Add(flds.Profession_category, Nothing)
        D.Add(flds.Email, Nothing)
        D.Add(flds.Address_1, Nothing)
        D.Add(flds.Address_2, Nothing)
        D.Add(flds.Suburb, Nothing)
        D.Add(flds.PostCode, Nothing)
        D.Add(flds.State, Nothing)
        D.Add(flds.Phone_home, Nothing)
        D.Add(flds.Phone_mobile, Nothing)
        D.Add(flds.Phone_work, Nothing)
        D.Add(flds.Department, Nothing)
        D.Add(flds.Institution, Nothing)
        D.Add(flds.User_name, Nothing)
        D.Add(flds.User_password, Nothing)
        D.Add(flds.Last_login, Nothing)
        D.Add(flds.Lastupdated, Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dicPrefPred_ForLoad() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_Pref_PredFields

        R.Add(flds.Age_lower, Nothing)
        R.Add(flds.Age_upper, Nothing)
        R.Add(flds.AgeGroupID, Nothing)
        R.Add(flds.AgeGroup, Nothing)
        R.Add(flds.EndDate, Nothing)
        R.Add(flds.EthnicityID, Nothing)
        R.Add(flds.Ethnicity, Nothing)
        R.Add(flds.GenderID, Nothing)
        R.Add(flds.Gender, Nothing)
        R.Add(flds.Lastedit, Nothing)
        R.Add(flds.LasteditBy, Nothing)
        R.Add(flds.ParameterID, Nothing)
        R.Add(flds.Parameter, Nothing)
        R.Add(flds.PrefID, Nothing)
        R.Add(flds.SourceID, Nothing)
        R.Add(flds.Source, Nothing)
        R.Add(flds.StartDate, Nothing)
        R.Add(flds.TestID, Nothing)
        R.Add(flds.Test, Nothing)
        R.Add(flds.EquationID_mpv, Nothing)
        R.Add(flds.EquationID_NormalRange, Nothing)

        Return R

    End Function

    Public Function Calc_Age(ByVal DOB, ByVal aDate) As String

        Dim Age As Single = 0

        If IsDate(DOB) And IsDate(aDate) Then
            Age = DateDiff(DateInterval.Day, DOB, aDate) / 365
            If Age > 0 And Age < 120 Then
                Return Format(Age, "#.0")
            Else
                Return ""
            End If
        Else
            Return ""
        End If

    End Function

    Public Function MakeEmpty_dicPrefPred_ForSave() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_Pref_PredFields

        R.Add(flds.PrefID, Nothing)
        R.Add(flds.EquationID_mpv, Nothing)
        R.Add(flds.EquationID_NormalRange, Nothing)
        R.Add(flds.Age_lower, Nothing)
        R.Add(flds.Age_upper, Nothing)
        R.Add(flds.Age_clipmethod, Nothing)
        R.Add(flds.Age_clipmethodID, Nothing)
        R.Add(flds.ht_lower, Nothing)
        R.Add(flds.ht_upper, Nothing)
        R.Add(flds.ht_clipmethod, Nothing)
        R.Add(flds.ht_clipmethodID, Nothing)
        R.Add(flds.wt_lower, Nothing)
        R.Add(flds.wt_upper, Nothing)
        R.Add(flds.wt_clipmethod, Nothing)
        R.Add(flds.wt_clipmethodID, Nothing)
        R.Add(flds.AgeGroupID, Nothing)
        R.Add(flds.EthnicityID, Nothing)
        R.Add(flds.GenderID, Nothing)
        R.Add(flds.ParameterID, Nothing)
        R.Add(flds.SourceID, Nothing)
        R.Add(flds.StartDate, Nothing)
        R.Add(flds.EndDate, Nothing)
        R.Add(flds.TestID, Nothing)
        R.Add(flds.Lastedit, Nothing)
        R.Add(flds.LasteditBy, Nothing)
        Return R

    End Function

    Public Function MakeEmpty_dicRft_Session() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_Rft_RoutineAndSessionFields

        R.Add(flds.SessionID, Nothing)
        R.Add(flds.PatientID, Nothing)
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.Pred_SourceIDs, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)
        Return R
    End Function

    Public Function MakeEmpty_dicRft_Routine() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_Rft_RoutineAndSessionFields

        R.Add(flds.PatientID, Nothing)
        R.Add(flds.RftID, Nothing)
        R.Add(flds.SessionID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)

        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.LungVolumes_method, Nothing)

        R.Add(flds.R_bl_Fev1, Nothing)
        R.Add(flds.R_bl_Fvc, Nothing)
        R.Add(flds.R_bl_Vc, Nothing)
        R.Add(flds.R_Bl_Fer, Nothing)
        R.Add(flds.R_bl_Fef2575, Nothing)
        R.Add(flds.R_bl_Pef, Nothing)
        R.Add(flds.R_Bl_Tlco, Nothing)
        R.Add(flds.R_Bl_Kco, Nothing)
        R.Add(flds.R_Bl_Va, Nothing)
        R.Add(flds.R_Bl_Hb, Nothing)
        R.Add(flds.R_Bl_Ivc, Nothing)
        R.Add(flds.R_Bl_Tlc, Nothing)
        R.Add(flds.R_Bl_Frc, Nothing)
        R.Add(flds.R_Bl_Rv, Nothing)
        R.Add(flds.R_Bl_LvVc, Nothing)
        R.Add(flds.R_Bl_RvTlc, Nothing)
        R.Add(flds.R_Bl_Mip, Nothing)
        R.Add(flds.R_Bl_Mep, Nothing)
        R.Add(flds.R_SpO2_1, Nothing)
        R.Add(flds.R_SpO2_2, Nothing)
        R.Add(flds.R_Bl_FeNO, Nothing)

        R.Add(flds.R_post_Fev1, Nothing)
        R.Add(flds.R_post_Fvc, Nothing)
        R.Add(flds.R_post_Vc, Nothing)
        R.Add(flds.R_Post_Fer, Nothing)
        R.Add(flds.R_post_Fef2575, Nothing)
        R.Add(flds.R_post_Pef, Nothing)
        R.Add(flds.R_post_Condition, Nothing)

        R.Add(flds.R_abg1_aapo2, Nothing)
        R.Add(flds.R_abg1_be, Nothing)
        R.Add(flds.R_abg1_fio2, Nothing)
        R.Add(flds.R_abg1_hco3, Nothing)
        R.Add(flds.R_abg1_paco2, Nothing)
        R.Add(flds.R_abg1_pao2, Nothing)
        R.Add(flds.R_abg1_ph, Nothing)
        R.Add(flds.R_abg1_sao2, Nothing)
        R.Add(flds.R_abg1_shunt, Nothing)
        R.Add(flds.R_abg2_aapo2, Nothing)
        R.Add(flds.R_abg2_be, Nothing)
        R.Add(flds.R_abg2_fio2, Nothing)
        R.Add(flds.R_abg2_hco3, Nothing)
        R.Add(flds.R_abg2_paco2, Nothing)
        R.Add(flds.R_abg2_pao2, Nothing)
        R.Add(flds.R_abg2_ph, Nothing)
        R.Add(flds.R_abg2_sao2, Nothing)
        R.Add(flds.R_abg2_shunt, Nothing)
        R.Add(flds.LastUpdated_rft, Nothing)
        R.Add(flds.LastUpdatedBy_rft, Nothing)
        Return R

    End Function

    Public Function MakeEmpty_dicRft_RoutineAndSession() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_Rft_RoutineAndSessionFields

        'Session
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_phone, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)

        'Routine
        R.Add(flds.PatientID, Nothing)
        R.Add(flds.RftID, Nothing)
        R.Add(flds.SessionID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)

        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.LungVolumes_method, Nothing)

        R.Add(flds.R_bl_Fev1, Nothing)
        R.Add(flds.R_bl_Fvc, Nothing)
        R.Add(flds.R_bl_Vc, Nothing)
        R.Add(flds.R_Bl_Fer, Nothing)
        R.Add(flds.R_bl_Fef2575, Nothing)
        R.Add(flds.R_bl_Pef, Nothing)
        R.Add(flds.R_Bl_Tlco, Nothing)
        R.Add(flds.R_Bl_Kco, Nothing)
        R.Add(flds.R_Bl_Va, Nothing)
        R.Add(flds.R_Bl_Hb, Nothing)
        R.Add(flds.R_Bl_Ivc, Nothing)
        R.Add(flds.R_Bl_Tlc, Nothing)
        R.Add(flds.R_Bl_Frc, Nothing)
        R.Add(flds.R_Bl_Rv, Nothing)
        R.Add(flds.R_Bl_LvVc, Nothing)
        R.Add(flds.R_Bl_RvTlc, Nothing)
        R.Add(flds.R_Bl_Mip, Nothing)
        R.Add(flds.R_Bl_Mep, Nothing)
        R.Add(flds.R_SpO2_1, Nothing)
        R.Add(flds.R_SpO2_2, Nothing)
        R.Add(flds.R_Bl_FeNO, Nothing)

        R.Add(flds.R_post_Fev1, Nothing)
        R.Add(flds.R_post_Fvc, Nothing)
        R.Add(flds.R_post_Vc, Nothing)
        R.Add(flds.R_Post_Fer, Nothing)
        R.Add(flds.R_post_Fef2575, Nothing)
        R.Add(flds.R_post_Pef, Nothing)
        R.Add(flds.R_post_Condition, Nothing)

        R.Add(flds.R_abg1_aapo2, Nothing)
        R.Add(flds.R_abg1_be, Nothing)
        R.Add(flds.R_abg1_fio2, Nothing)
        R.Add(flds.R_abg1_hco3, Nothing)
        R.Add(flds.R_abg1_paco2, Nothing)
        R.Add(flds.R_abg1_pao2, Nothing)
        R.Add(flds.R_abg1_ph, Nothing)
        R.Add(flds.R_abg1_sao2, Nothing)
        R.Add(flds.R_abg1_shunt, Nothing)
        R.Add(flds.R_abg2_aapo2, Nothing)
        R.Add(flds.R_abg2_be, Nothing)
        R.Add(flds.R_abg2_fio2, Nothing)
        R.Add(flds.R_abg2_hco3, Nothing)
        R.Add(flds.R_abg2_paco2, Nothing)
        R.Add(flds.R_abg2_pao2, Nothing)
        R.Add(flds.R_abg2_ph, Nothing)
        R.Add(flds.R_abg2_sao2, Nothing)
        R.Add(flds.R_abg2_shunt, Nothing)
        R.Add(flds.Pred_SourceIDs, Nothing)
        R.Add(flds.LastUpdated_rft, Nothing)
        R.Add(flds.LastUpdatedBy_rft, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicWalk_trial() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_Walk_Trial

        R.Add(flds.trialID, Nothing)
        R.Add(flds.walkID, Nothing)
        R.Add(flds.trial_number, Nothing)
        R.Add(flds.trial_label, Nothing)
        R.Add(flds.trial_distance, Nothing)
        R.Add(flds.trial_timeofday, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicWalk_level() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_Walk_TrialLevel

        R.Add(flds.trialID, Nothing)
        R.Add(flds.levelID, Nothing)
        R.Add(flds.time_label, Nothing)
        R.Add(flds.time_minute, Nothing)
        R.Add(flds.time_speed_kph, Nothing)
        R.Add(flds.time_gradient, Nothing)
        R.Add(flds.time_spo2, Nothing)
        R.Add(flds.time_hr, Nothing)
        R.Add(flds.time_o2flow, Nothing)
        R.Add(flds.time_dyspnoea, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicWalk_test_session() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_WalkAndSession

        'Session
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_phone, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)

        'Walk
        R.Add(flds.patientID, Nothing)
        R.Add(flds.walkID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.WalkType, Nothing)
        R.Add(flds.ProtocolID, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.LastUpdated_walk, Nothing)
        R.Add(flds.LastUpdatedBy_walk, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicWalk_test() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_WalkAndSession

        'Walk
        R.Add(flds.patientID, Nothing)
        R.Add(flds.walkID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.WalkType, Nothing)
        R.Add(flds.ProtocolID, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.LastUpdated_walk, Nothing)
        R.Add(flds.LastUpdatedBy_walk, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicProv_test_session() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_ProvAndSession

        'Session
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_phone, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)

        R.Add(flds.patientID, Nothing)
        R.Add(flds.provID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.Pred_SourceIDs, Nothing)
        R.Add(flds.R_bl_Fev1, Nothing)
        R.Add(flds.R_bl_Fvc, Nothing)
        R.Add(flds.R_bl_Vc, Nothing)
        R.Add(flds.R_Bl_Fer, Nothing)
        R.Add(flds.R_bl_Fef2575, Nothing)
        R.Add(flds.R_bl_Pef, Nothing)

        R.Add(flds.ProtocolID, Nothing)
        R.Add(flds.Protocol_title, Nothing)
        R.Add(flds.Protocol_doseunits, Nothing)
        R.Add(flds.Protocol_drug, Nothing)
        R.Add(flds.Protocol_threshold, Nothing)
        R.Add(flds.Protocol_pd_decimalplaces, Nothing)
        R.Add(flds.Protocol_dose_effect, Nothing)
        R.Add(flds.Protocol_method, Nothing)
        R.Add(flds.Protocol_method_reference, Nothing)
        R.Add(flds.Protocol_parameter, Nothing)
        R.Add(flds.Protocol_parameter_units, Nothing)
        R.Add(flds.Protocol_parameter_response, Nothing)
        R.Add(flds.Protocol_post_drug, Nothing)
        R.Add(flds.Pd, Nothing)
        R.Add(flds.plot_ymin, Nothing)
        R.Add(flds.plot_ymax, Nothing)
        R.Add(flds.plot_ystep, Nothing)
        R.Add(flds.plot_xtitle, Nothing)
        R.Add(flds.plot_xscaling_type, Nothing)
        R.Add(flds.plot_title, Nothing)
        R.Add(flds.LastUpdated_prov, Nothing)
        R.Add(flds.LastUpdatedBy_prov, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicProv_test() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_ProvAndSession

        R.Add(flds.patientID, Nothing)
        R.Add(flds.provID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)

        R.Add(flds.R_bl_Fev1, Nothing)
        R.Add(flds.R_bl_Fvc, Nothing)
        R.Add(flds.R_bl_Vc, Nothing)
        R.Add(flds.R_Bl_Fer, Nothing)
        R.Add(flds.R_bl_Fef2575, Nothing)
        R.Add(flds.R_bl_Pef, Nothing)

        R.Add(flds.ProtocolID, Nothing)
        R.Add(flds.Protocol_title, Nothing)
        R.Add(flds.Protocol_doseunits, Nothing)
        R.Add(flds.Protocol_drug, Nothing)
        R.Add(flds.Protocol_threshold, Nothing)
        R.Add(flds.Protocol_pd_decimalplaces, Nothing)
        R.Add(flds.Protocol_dose_effect, Nothing)
        R.Add(flds.Protocol_method, Nothing)
        R.Add(flds.Protocol_method_reference, Nothing)
        R.Add(flds.Protocol_parameter, Nothing)
        R.Add(flds.Protocol_parameter_units, Nothing)
        R.Add(flds.Protocol_parameter_response, Nothing)
        R.Add(flds.Protocol_post_drug, Nothing)
        R.Add(flds.Pd, Nothing)
        R.Add(flds.plot_ymin, Nothing)
        R.Add(flds.plot_ymax, Nothing)
        R.Add(flds.plot_ystep, Nothing)
        R.Add(flds.plot_xtitle, Nothing)
        R.Add(flds.plot_xscaling_type, Nothing)
        R.Add(flds.plot_title, Nothing)
        R.Add(flds.LastUpdated_prov, Nothing)
        R.Add(flds.LastUpdatedBy_prov, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicProv_testdata() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_ProvTestDataFields

        R.Add(flds.testdataid, Nothing)
        R.Add(flds.provid, Nothing)
        R.Add(flds.doseid, Nothing)
        R.Add(flds.dose_number, Nothing)
        R.Add(flds.dose_discrete, Nothing)
        R.Add(flds.dose_cumulative, Nothing)
        R.Add(flds.result, Nothing)
        R.Add(flds.response, Nothing)
        R.Add(flds.dose_time_min, Nothing)
        R.Add(flds.xaxis_label, Nothing)
        R.Add(flds.dose_canskip, Nothing)
        Return R

    End Function



    Public Function MakeEmpty_dic_Pred_source() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)
        Dim flds As New class_pred_SourceFields

        D.Add(flds.SourceID, Nothing)
        D.Add(flds.Source, Nothing)
        D.Add(flds.Pub_Reference, Nothing)
        D.Add(flds.Pub_Year, Nothing)
        D.Add(flds.Lastedit, Nothing)
        D.Add(flds.LasteditBy, Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dic_Pred_equation() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)
        Dim flds As New class_pred_EquationFields

        D.Add(flds.Test, Nothing)
        D.Add(flds.TestID, Nothing)
        D.Add(flds.Source, Nothing)
        D.Add(flds.SourceID, Nothing)
        D.Add(flds.Parameter, Nothing)
        D.Add(flds.ParameterID, Nothing)
        D.Add(flds.Gender, Nothing)
        D.Add(flds.GenderID, Nothing)
        D.Add(flds.AgeGroup, Nothing)
        D.Add(flds.AgeGroupID, Nothing)
        D.Add(flds.Age_lower, Nothing)
        D.Add(flds.Age_upper, Nothing)
        D.Add(flds.Age_clipmethod, Nothing)
        D.Add(flds.Age_clipmethodID, Nothing)
        D.Add(flds.Ht_lower, Nothing)
        D.Add(flds.Ht_upper, Nothing)
        D.Add(flds.Ht_clipmethod, Nothing)
        D.Add(flds.Ht_clipmethodID, Nothing)
        D.Add(flds.Wt_lower, Nothing)
        D.Add(flds.Wt_upper, Nothing)
        D.Add(flds.Wt_clipmethod, Nothing)
        D.Add(flds.Wt_clipmethodID, Nothing)
        D.Add(flds.Ethnicity, Nothing)
        D.Add(flds.EthnicityID, Nothing)
        D.Add(flds.Equation, Nothing)
        D.Add(flds.EquationID, Nothing)
        D.Add(flds.EquationType, Nothing)
        D.Add(flds.EquationTypeID, Nothing)
        D.Add(flds.StatType, Nothing)
        D.Add(flds.StatTypeID, Nothing)
        D.Add(flds.EthnicityCorrectionType, Nothing)
        D.Add(flds.Lastedit, Nothing)
        D.Add(flds.LasteditBy, Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dic_sptPanel() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)

        D.Add(GetPropertyName(Function() (New PanelData).panelid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelData).panelname), Nothing)
        D.Add(GetPropertyName(Function() (New PanelData).enabled), Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dic_sptPanelMember_AllData() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)

        D.Add(GetPropertyName(Function() (New PanelMember_AllData).memberid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_AllData).panelid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_AllData).allergenid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_AllData).allergenname), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_AllData).displayColour), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_AllData).allergenorder), Nothing)
        Return D

    End Function

    Public Function MakeEmpty_dic_sptPanelMember_TableDataOnly() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)

        D.Add(GetPropertyName(Function() (New PanelMember_TableDataOnly).memberid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_TableDataOnly).panelid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_TableDataOnly).allergenid), Nothing)
        D.Add(GetPropertyName(Function() (New PanelMember_TableDataOnly).allergenorder), Nothing)
        Return D

    End Function

    Public Function MakeEmpty_dic_sptCategory() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)

        D.Add(GetPropertyName(Function() (New AllergenCategoryData).categoryid), Nothing)
        D.Add(GetPropertyName(Function() (New AllergenCategoryData).categoryname), Nothing)
        D.Add(GetPropertyName(Function() (New AllergenCategoryData).displaycolour), Nothing)
        D.Add(GetPropertyName(Function() (New AllergenCategoryData).enabled), Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dic_sptAllergen() As Dictionary(Of String, String)

        Dim D As New Dictionary(Of String, String)

        D.Add(GetPropertyName(Function() (New AllergenData).allergenid), Nothing)
        D.Add(GetPropertyName(Function() (New AllergenData).allergenname), Nothing)
        D.Add(GetPropertyName(Function() (New AllergenData).enabled), Nothing)

        Return D

    End Function

    Public Function MakeEmpty_dicCPET_levels() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_Cpet_Levels

        R.Add(flds.levelID, Nothing)
        R.Add(flds.cpetID, Nothing)
        R.Add(flds.level_number, Nothing)
        R.Add(flds.level_time, Nothing)
        R.Add(flds.level_workload, Nothing)
        R.Add(flds.level_vo2, Nothing)
        R.Add(flds.level_vco2, Nothing)
        R.Add(flds.level_ve, Nothing)
        R.Add(flds.level_vt, Nothing)
        R.Add(flds.level_spo2, Nothing)
        R.Add(flds.level_hr, Nothing)
        R.Add(flds.level_peto2, Nothing)
        R.Add(flds.level_petco2, Nothing)
        R.Add(flds.level_rer, Nothing)
        R.Add(flds.level_vevo2, Nothing)
        R.Add(flds.level_vevco2, Nothing)
        R.Add(flds.level_o2pulse, Nothing)
        R.Add(flds.level_ph, Nothing)
        R.Add(flds.level_paco2, Nothing)
        R.Add(flds.level_pao2, Nothing)
        R.Add(flds.level_hco3, Nothing)
        R.Add(flds.level_be, Nothing)
        R.Add(flds.level_sao2, Nothing)
        R.Add(flds.level_aapo2, Nothing)
        R.Add(flds.level_vdvt, Nothing)
        Return R

    End Function

    Public Function MakeEmpty_dicCPETandSession() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_CPETAndSessionFields


        'Session
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_phone, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)

        R.Add(flds.patientID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.cpetID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.r_bp_pre, Nothing)
        R.Add(flds.r_bp_post, Nothing)
        R.Add(flds.r_spiro_pre_fev1, Nothing)
        R.Add(flds.r_spiro_pre_fvc, Nothing)
        R.Add(flds.r_symptoms_dyspnoea_borg, Nothing)
        R.Add(flds.r_symptoms_legs_borg, Nothing)
        R.Add(flds.r_symptoms_other_borg, Nothing)
        R.Add(flds.r_symptoms_other_text, Nothing)
        R.Add(flds.r_abg_post_fio2, Nothing)
        R.Add(flds.r_abg_post_ph, Nothing)
        R.Add(flds.r_abg_post_pao2, Nothing)
        R.Add(flds.r_abg_post_paco2, Nothing)
        R.Add(flds.r_abg_post_be, Nothing)
        R.Add(flds.r_abg_post_hco3, Nothing)
        R.Add(flds.r_abg_post_sao2, Nothing)
        R.Add(flds.r_max_workload, Nothing)
        R.Add(flds.r_max_ve, Nothing)
        R.Add(flds.r_max_vo2, Nothing)
        R.Add(flds.r_max_vo2kg_mpv, Nothing)
        R.Add(flds.r_max_vco2_mpv, Nothing)
        R.Add(flds.r_max_hr_mpv, Nothing)
        R.Add(flds.r_max_vt_mpv, Nothing)
        R.Add(flds.r_max_o2pulse_mpv, Nothing)
        R.Add(flds.r_max_workload_lln, Nothing)
        R.Add(flds.r_max_ve_lln, Nothing)
        R.Add(flds.r_max_vo2_lln, Nothing)
        R.Add(flds.r_max_vo2kg_lln, Nothing)
        R.Add(flds.r_max_vco2_lln, Nothing)
        R.Add(flds.r_max_hr_lln, Nothing)
        R.Add(flds.r_max_vt_lln, Nothing)
        R.Add(flds.r_max_o2pulse_lln, Nothing)
        R.Add(flds.LastUpdated_cpet, Nothing)
        R.Add(flds.LastUpdatedBy_cpet, Nothing)

        Return R

    End Function



    Public Function MakeEmpty_dicCPET() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_CPETAndSessionFields

        R.Add(flds.patientID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.cpetID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)
        R.Add(flds.r_bp_pre, Nothing)
        R.Add(flds.r_bp_post, Nothing)
        R.Add(flds.r_spiro_pre_fev1, Nothing)
        R.Add(flds.r_spiro_pre_fvc, Nothing)
        R.Add(flds.r_symptoms_dyspnoea_borg, Nothing)
        R.Add(flds.r_symptoms_legs_borg, Nothing)
        R.Add(flds.r_symptoms_other_borg, Nothing)
        R.Add(flds.r_symptoms_other_text, Nothing)
        R.Add(flds.r_abg_post_fio2, Nothing)
        R.Add(flds.r_abg_post_ph, Nothing)
        R.Add(flds.r_abg_post_pao2, Nothing)
        R.Add(flds.r_abg_post_paco2, Nothing)
        R.Add(flds.r_abg_post_be, Nothing)
        R.Add(flds.r_abg_post_hco3, Nothing)
        R.Add(flds.r_abg_post_sao2, Nothing)
        R.Add(flds.r_max_workload, Nothing)
        R.Add(flds.r_max_ve, Nothing)
        R.Add(flds.r_max_vo2, Nothing)
        R.Add(flds.r_max_vo2kg_mpv, Nothing)
        R.Add(flds.r_max_vco2_mpv, Nothing)
        R.Add(flds.r_max_hr_mpv, Nothing)
        R.Add(flds.r_max_vt_mpv, Nothing)
        R.Add(flds.r_max_o2pulse_mpv, Nothing)
        R.Add(flds.r_max_workload_lln, Nothing)
        R.Add(flds.r_max_ve_lln, Nothing)
        R.Add(flds.r_max_vo2_lln, Nothing)
        R.Add(flds.r_max_vo2kg_lln, Nothing)
        R.Add(flds.r_max_vco2_lln, Nothing)
        R.Add(flds.r_max_hr_lln, Nothing)
        R.Add(flds.r_max_vt_lln, Nothing)
        R.Add(flds.r_max_o2pulse_lln, Nothing)
        R.Add(flds.LastUpdated_cpet, Nothing)
        R.Add(flds.LastUpdatedBy_cpet, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicHAST() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_HastAndSessionFields

        R.Add(flds.patientID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.hastID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)

        R.Add(flds.deliverymethod_fio2, Nothing)
        R.Add(flds.deliverymethod_suppo2, Nothing)
        R.Add(flds.protocol_name, Nothing)
        R.Add(flds.protocol_id, Nothing)

        R.Add(flds.LastUpdated_hast, Nothing)
        R.Add(flds.LastUpdatedBy_hast, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicHASTandSession() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_HastAndSessionFields

        'Session
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_phone, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)

        'HAST
        R.Add(flds.patientID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.hastID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)

        R.Add(flds.deliverymethod_fio2, Nothing)
        R.Add(flds.deliverymethod_suppo2, Nothing)
        R.Add(flds.protocol_name, Nothing)
        R.Add(flds.protocol_id, Nothing)

        R.Add(flds.LastUpdated_hast, Nothing)
        R.Add(flds.LastUpdatedBy_hast, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicHAST_level() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_Hast_Levels

        R.Add(flds.levelID, Nothing)
        R.Add(flds.hastID, Nothing)
        R.Add(flds.altitude_fio2, Nothing)
        R.Add(flds.altitude_ft, Nothing)
        R.Add(flds.r_be, Nothing)
        R.Add(flds.r_hco3, Nothing)
        R.Add(flds.r_note, Nothing)
        R.Add(flds.r_paco2, Nothing)
        R.Add(flds.r_ph, Nothing)
        R.Add(flds.r_sao2, Nothing)
        R.Add(flds.r_spo2, Nothing)
        R.Add(flds.suppO2_flow, Nothing)
        R.Add(flds.row_order, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicSPT() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_SptAndSessionFields

        R.Add(flds.patientID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.sptID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)

        R.Add(flds.site, Nothing)
        R.Add(flds.medications, Nothing)
        R.Add(flds.panelID, Nothing)
        R.Add(flds.panel_name, Nothing)

        R.Add(flds.LastUpdated_spt, Nothing)
        R.Add(flds.LastUpdatedBy_spt, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicSPTandSession() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_SptAndSessionFields

        'Session
        R.Add(flds.TestDate, Nothing)
        R.Add(flds.Req_date, Nothing)
        R.Add(flds.Req_time, Nothing)
        R.Add(flds.AdmissionStatus, Nothing)
        R.Add(flds.Height, Nothing)
        R.Add(flds.Weight, Nothing)
        R.Add(flds.Smoke_hx, Nothing)
        R.Add(flds.Smoke_yearssmoked, Nothing)
        R.Add(flds.Smoke_cigsperday, Nothing)
        R.Add(flds.Smoke_packyears, Nothing)
        R.Add(flds.Req_name, Nothing)
        R.Add(flds.Req_address, Nothing)
        R.Add(flds.Req_fax, Nothing)
        R.Add(flds.Req_phone, Nothing)
        R.Add(flds.Req_email, Nothing)
        R.Add(flds.Req_healthservice, Nothing)
        R.Add(flds.Report_copyto, Nothing)
        R.Add(flds.Req_clinicalnotes, Nothing)
        R.Add(flds.LastUpdated_session, Nothing)
        R.Add(flds.LastUpdatedBy_session, Nothing)

        'SPT
        R.Add(flds.patientID, Nothing)
        R.Add(flds.sessionID, Nothing)
        R.Add(flds.sptID, Nothing)
        R.Add(flds.TestTime, Nothing)
        R.Add(flds.TestType, Nothing)
        R.Add(flds.Lab, Nothing)
        R.Add(flds.BDStatus, Nothing)
        R.Add(flds.Scientist, Nothing)
        R.Add(flds.TechnicalNotes, Nothing)
        R.Add(flds.Report_text, Nothing)
        R.Add(flds.Report_reportedby, Nothing)
        R.Add(flds.Report_reporteddate, Nothing)
        R.Add(flds.Report_verifiedby, Nothing)
        R.Add(flds.Report_status, Nothing)
        R.Add(flds.Report_verifieddate, Nothing)

        R.Add(flds.site, Nothing)
        R.Add(flds.medications, Nothing)
        R.Add(flds.panelID, Nothing)
        R.Add(flds.panel_name, Nothing)

        R.Add(flds.LastUpdated_spt, Nothing)
        R.Add(flds.LastUpdatedBy_spt, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicSpt_allergen() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_fields_Spt_Allergens

        R.Add(flds.allergen_category_id, Nothing)
        R.Add(flds.allergen_category_name, Nothing)
        R.Add(flds.allergen_name, Nothing)
        R.Add(flds.allergenID, Nothing)
        R.Add(flds.note, Nothing)
        R.Add(flds.sptID, Nothing)
        R.Add(flds.wheal_mm, Nothing)
        R.Add(flds.panelmemberID, Nothing)
        R.Add(flds.allergen_category_colour, Nothing)

        Return R

    End Function

    Public Function MakeEmpty_dicApp_config() As Dictionary(Of String, String)

        Dim R As New Dictionary(Of String, String)
        Dim flds As New class_Fields_app_config

        R.Add(flds.configID, Nothing)
        R.Add(flds.siteID, Nothing)
        R.Add(flds.site_name, Nothing)
        R.Add(flds.site_state, Nothing)
        R.Add(flds.site_institution, Nothing)
        R.Add(flds.db_type, Nothing)
        R.Add(flds.db_name, Nothing)
        R.Add(flds.db_servername, Nothing)
        R.Add(flds.db_connectstring, Nothing)
        R.Add(flds.pas_mode_local, Nothing)

        Return R

    End Function

    Public Function GetPropertyName(Of t)(ByVal PropertyExp As Expression(Of Func(Of t))) As String
        Return TryCast(PropertyExp.Body, MemberExpression).Member.Name
    End Function


    Public Function Calc_Fer(ByVal Fev1, ByVal fvc, ByVal vc) As String

        Dim s As String = ""

        If Val(fvc) = 0 And Val(vc) = 0 Or Val(Fev1) = 0 Then
            Return ""
        Else
            'Use the larger of VC and FVC
            If Fev1 > 0 Then
                If fvc > vc Then
                    If fvc > 0 Then
                        s = Format(100 * Fev1 / fvc, "0")
                    Else
                        s = ""
                    End If
                Else
                    If vc > 0 Then
                        s = Format(100 * Fev1 / vc, "0")
                    Else
                        s = ""
                    End If
                End If
            End If
            Return s
        End If

    End Function

    Public Function calc_HbFac(ByVal Hb As String, ByVal DOB As String, ByVal Sex As String, ByVal TestDate As String) As Single
        'Haemoglobin correction using ATS 1995 recommended equations
        'ATS assumes Dm/Vc = 0.7 and recommends correcting to 14.6 for males and to 13.4 for females and boys < 16.
        'Returns 0 if a problem

        If Val(Hb) = 0 Or Not IsDate(DOB) Or Not IsDate(TestDate) Or Not (Sex = "Male" Or Sex = "Female") Then
            Return 0
        End If

        Dim Age As Single = cMyRoutines.Calc_Age(CDate(DOB), CDate(TestDate))

        Select Case Sex
            Case "Male"
                Select Case Age
                    Case 16 To 120 : Return (10.22 + Val(Hb)) / (1.7 * Val(Hb))
                    Case 5 To 15.999 : Return (9.38 + Val(Hb)) / (1.7 * Val(Hb))
                    Case Else : Return 0
                End Select
            Case "Female"
                Select Case Age
                    Case 5 To 120 : Return (9.38 + Val(Hb)) / (1.7 * Val(Hb))
                    Case Else : Return 0
                End Select
            Case Else : Return 0
        End Select

    End Function

    Public Function calc_BMI(ByVal Ht As String, ByVal Wt As String) As String

        If Val(Ht) > 0 And Val(Wt) > 0 Then
            Dim bmi As Single = 10000 * Val(Wt) / Val(Ht) ^ 2
            Return Format(bmi, "##.0")
        Else
            Return ""
        End If

    End Function

    Public Function calc_PackYears(ByVal CigsPerDay As String, ByVal ByValYrsSmoked As String) As String

        If Val(CigsPerDay) > 0 And Val(ByValYrsSmoked) > 0 Then
            Dim PackYrs As Single = Val(ByValYrsSmoked) * Val(CigsPerDay) / 20
            Return Format(PackYrs, "##0.0")
        Else
            Return ""
        End If

    End Function

    Public Function Lookup_list_ByCode(ByVal Code As String, ByVal eRefTbl As eTables) As String
        'Assumes lookup table has fields named 'code' and 'description'

        Dim sql As String = "SELECT description FROM " & cDbInfo.table_name(eRefTbl) & " WHERE code = '" & Code & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds Is Nothing Or Ds.Tables(0).Rows.Count = 0 Then
            Return ""
        Else
            Return Ds.Tables(0).Rows(0).Item(0)
        End If

    End Function

    Public Function Lookup_list_ByDescription(ByVal Description As String, ByVal eRefTbl As eTables) As String
        'Assumes lookup table has fields named 'code' and 'description'

        Dim sql As String = "SELECT code FROM " & cDbInfo.table_name(eRefTbl) & " WHERE description = '" & Description & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds Is Nothing Or Ds.Tables(0).Rows.Count = 0 Then
            Return 0
        Else
            Return Ds.Tables(0).Rows(0).Item(0)
        End If

    End Function

    Function Calculate_AaPO2(ByVal PaO2 As Single, ByVal PaCO2 As Single) As Single
        'Assume FiO2=air

        Try
            Dim A As Single = 0.2095 * 713 - PaCO2 / 0.8
            Dim Aa As Single = A - PaO2
            Return Aa
        Catch
            Return 0
        End Try

    End Function

    Function Calculate_Shunt(ByVal PaO2, ByVal PaCO2, fio2) As String

        Dim NoResult As String = "---"

        If InStr(fio2, "100") = 0 Then
            Dim response = MsgBox("FiO2 doesn't indicate 100% O2. Continue?", vbOKCancel, "Warning")
            If response = vbCancel Then Return ""
        End If

        Try
            If IsNumeric(PaO2) And IsNumeric(PaCO2) Then
                If Val(PaO2) > 0 And Val(PaCO2) > 0 Then
                    Dim Shunt As Single
                    Shunt = Val(PaO2) + Val(PaCO2)
                    Shunt = 100 * (713 - Shunt) / (2305 - Shunt)
                    Return Format(Shunt, "#0.0")
                Else
                    Return NoResult
                End If
            Else
                Return NoResult
            End If

        Catch
            MsgBox("Error calculating shunt")
            Return NoResult
        End Try

    End Function

    Public Function Validate_MedicareNumber(ByVal MedicareNum As String, Optional ByRef ReturnErrorVar As String = "") As Boolean
        'Australian Medicare card numbers consist of 11 digits structured as follows:
        'Identifier 8-digits    First digit should be in the range 2-6
        'Checksum 1-digit, digits are weighted (1,3,7,9,1,3,7,9)
        'Issue Number 1-digit Indicates how many times the card has been issued
        'Individual Reference Number (IRN)1-digit The IRN appears on the left of the cardholder's name
        '     on the medicare card and distinguishes the individuals named on the card.

        Dim s As String

        'strip non numeric chars
        s = Replace(MedicareNum, "-", "")
        s = Replace(s, " ", "")

        'Check overall length

        If s.ToString = "" Then
            ReturnErrorVar = "Valid - no entry"
            Return True
        ElseIf Len(s) <> 11 Then
            ReturnErrorVar = "Error - incorrect length medicare number (must be 11 numeric characters)"
            Return False
        ElseIf InStr("23456", Mid(s, 1, 1)) = 0 Then
            'Check first digit - should be in the range 2-6
            ReturnErrorVar = "Error - incorrect first digit in medicare number (must be in the range 2-6)"
            Return False
        Else
            ReturnErrorVar = "Valid - entry"
            Return True
        End If

    End Function

    Public Function FindControl(ByVal ctrlParent As Control, controlName As String) As Control

        Dim found As Boolean = False
        Dim ctrl As Control = Nothing
        Dim c As Control = Nothing

        For Each ctrl In ctrlParent.Controls
            If ctrl.Name = controlName Then
                c = ctrl
            End If
            ' If the control has children, recursively call this function
            If ctrl.HasChildren Then
                FindControl(ctrl, controlName)
            End If
        Next

        Return c

    End Function



End Class
