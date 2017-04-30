Imports System
Imports System.Data
Imports System.Text
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Configuration
Imports ResLab_V1_Heavy.class_walktest
Imports ResLab_V1_Heavy.class_healthservices
Imports ResLab_V1_Heavy.cDatabaseInfo



Public Class class_pas

    Public Enum eMapMode
        HL7Code_RMCode
        HL7Code_RMText
        HL7Code_HL7Text
        RMCode_HL7Code
        RMCode_HL7Text
        RMCode_RMText
    End Enum
    Public Structure DeathInfo
        Dim DeathDate As String
        Dim DeathStatusString As String
    End Structure
    Public Enum eNameURFormats
        Lastname
        Firstname
        FullName
        FullName_PrimaryUR
        FullName_DisplayUR
        PrimaryUR
        DisplayUR
        URforBarcode
    End Enum
    Public Enum eURformats
        URonly
        URandHS
    End Enum

    Private _PatientID As Long
    'Private _TextBoxForDisplay_Name As New TextBox
    

    Public Property PatientID As Long

        Get
            PatientID = _PatientID
        End Get

        Set(ByVal value As Long)
            _PatientID = value

            'Log patient and display the pt's details
            Dim dicD As Dictionary(Of String, String) = Me.get_demographics(_PatientID)
            If Not IsNothing(dicD) Then
                Dim d As New class_DemographicFields
                Form_MainNew.tstxt_PatientName.Text = dicD(d.Surname) & ", " & dicD(d.Firstname)
                Form_MainNew.tstxt_UR.Text = dicD(d.URandHS)
            End If
            'Clear the pdf viewer
            Form_MainNew.PdfViewer1.CloseDocument()
            Form_MainNew.PdfViewer1.Visible = False

            ''Clear the trend viewer
            Form_MainNew.splitTrend.Panel1.Controls.Clear()
            Form_MainNew.splitTrend.Panel2.Controls.Clear()
            If Form_MainNew.splitTrend.Panel1.Controls.Find("grdTrend", True).Length > 0 Then
                Form_MainNew.splitTrend.Panel1.Controls("grdTrend").Dispose()
            End If
            If Form_MainNew.splitTrend.Panel2.Controls.Find("chrt", True).Length > 0 Then
                Form_MainNew.splitTrend.Panel2.Controls("chrt").Dispose()
            End If
            Dim tp As TabPage = Form_MainNew.tabReports.TabPages(0)
            Form_MainNew.tabReports.SelectedTab = tp

            'Clear the list of tests 
            Form_MainNew.lvRecall.Clear()

        End Set

    End Property

    Public Function Update_Demographics(ByVal D As Dictionary(Of String, String)) As Boolean

        Dim sqlString As String = cDAL.Build_UpdateQuery(eTables.PatientDetails, D)

        Try
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving demographics" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function


    Public Function Find_PtByNameEtc(ByVal Surname As String, ByVal SurnameBy As String, Optional ByVal Firstname As String = "", _
                              Optional ByVal FirstNameBy As String = "", Optional ByVal DOB As String = "", _
                              Optional ByVal Gender As String = "") As Long()

        Dim PatientID(0) As Long

        If Surname = "" Then
            Return PatientID
        End If

        Surname = Replace(Surname, "'", "''")

        'Build the query
        Dim sql As String = "SELECT PatientID FROM PatientDetails WHERE (Surname  "

        Select Case SurnameBy
            Case eFindBy.Exact.ToString : sql = sql & "='" & Surname & "') "
            Case eFindBy.First_part.ToString : sql = sql & "LIKE N'" & Surname & "%') "
            Case eFindBy.Any_part.ToString : sql = sql & "LIKE N'%" & Surname & "%') "
        End Select

        If Firstname <> "" Then
            sql = sql & " AND (Firstname LIKE "
            Select Case FirstNameBy
                Case eFindBy.Exact.ToString : sql = sql & " '" & Firstname & "') "
                Case eFindBy.First_part.ToString : sql = sql & "N'" & Firstname & "%') "
                Case eFindBy.Any_part.ToString : sql = sql & "N'%" & Firstname & "%') "
            End Select
        End If

        If Gender <> "" Then
            sql = sql & " AND (Gender = '" & Gender & "') "
        End If

        If IsDate(DOB) Then
            'sql = sql & " AND (DOB = " & DOB & ") "
        End If

        sql = sql & " ORDER BY Surname, FirstName"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In Ds.Tables(0).Rows
                ReDim Preserve PatientID(UBound(PatientID) + 1)
                PatientID(UBound(PatientID)) = row("PatientID")
            Next
        End If

        Return PatientID

    End Function


    Public Function Find_PtByUR(ByVal UR As String) As Long()

        Dim PatientID() As Long

        If UR = "" Then
            ReDim PatientID(0)
        Else
            PatientID = Me.get_patientIDFromUR(UR)
        End If
        Return PatientID

    End Function

    Public Function get_patientIDFromUR(ByVal UR As String) As Long()

        Dim PatientIDs(0) As Long
        Dim sql As String = "SELECT PatientID FROM pas_pt_ur_numbers WHERE UR='" & UR & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        Dim flds As New class_DemographicFields
        For Each row As DataRow In Ds.Tables(0).Rows
            ReDim Preserve PatientIDs(UBound(PatientIDs) + 1)
            PatientIDs(UBound(PatientIDs)) = row(flds.PatientID)
        Next
        Return PatientIDs

    End Function

    Public Function Get_PatientIDFromPkID(ByVal PkID As Long, ByVal eTbl As eTables) As Long

        Dim sql As String = ""
        Dim Tbl As String = cDbInfo.table_name(eTbl)
        Dim PK As String = cDbInfo.primarykey(eTbl)

        sql = "SELECT PatientID FROM " & Tbl & " WHERE " & PK & " =" & PkID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Ds.Tables(0).Rows.Count = 1 Then
            Return Ds.Tables(0).Rows(0)("PatientID")
        Else
            Return Nothing
        End If

    End Function

    Public Function Get_PatientsByDate_buildquery(startdate As Date, enddate As Date, eunit As eUnits) As String

        Dim eTbls() As eTables = {eTables.rft_routine, eTables.r_cpet, eTables.Prov_test, eTables.r_walktests_v1heavy, eTables.r_hast, eTables.r_spt}
        Dim sql As New StringBuilder
        Dim i As Integer = 0

        Select Case eunit
            Case eUnits.Respiratory_Laboratory
                sql.Append("SELECT DISTINCT * FROM (")
                For i = 0 To UBound(eTbls)
                    sql.Append("SELECT patientdetails.Surname, patientdetails.Firstname, patientdetails.UR, patientdetails.DOB, patientdetails.Gender, patientdetails.PatientID ")
                    sql.Append("FROM patientdetails INNER JOIN r_sessions ON patientdetails.PatientID = r_sessions.patientid INNER JOIN ")
                    sql.Append(cDbInfo.table_name(eTbls(i)) & " ON r_sessions.sessionID = " & cDbInfo.table_name(eTbls(i)) & ".SessionID ")
                    sql.Append("WHERE (dbo.r_sessions.testdate BETWEEN " & cMyRoutines.FormatDBDate(startdate) & " AND " & cMyRoutines.FormatDBDate(enddate) & ") ")
                    If i < UBound(eTbls) Then sql.Append(" UNION ALL ")
                Next
                sql.Append(") t ORDER BY t.Surname ASC, t.Firstname ASC;")
                Return sql.ToString
            Case Else
                Return ""
        End Select

    End Function

    Public Function Get_PatientsByDate(startdate As Date, enddate As Date, eunit As eUnits) As Dictionary(Of String, String)()

        'Check dates are valid, at least one must be, and endate<startdate not allowed
        If Not (IsDate(startdate) Or IsDate(enddate)) Then
            Return Nothing
        Else
            If Not IsDate(startdate) Then startdate = enddate
            If Not IsDate(enddate) Then enddate = startdate
            If enddate < startdate Then
                Return Nothing
            End If
        End If

        Dim sql As String = ""
        Select Case eunit
            Case eUnits.Respiratory_Laboratory : sql = Me.Get_PatientsByDate_buildquery(startdate, enddate, eunit)
            Case eUnits.Sleep_Laboratory
            Case eUnits.Victorian_Respiratory_Support_Service
            Case eUnits.CPAP_Clinic
            Case eUnits.O2_Therapy_Clinic
        End Select

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Dim flds As New class_DemographicFields
            Dim dicD() As Dictionary(Of String, String) = Nothing
            Dim i As Integer = 0
            With Ds.Tables(0)
                For Each record As DataRow In .Rows
                    ReDim Preserve dicD(i)
                    dicD(i) = cMyRoutines.MakeEmpty_dicDemo()

                    dicD(i)(flds.PatientID) = record(flds.PatientID)
                    dicD(i)(flds.Surname) = record(flds.Surname) & ""
                    dicD(i)(flds.Firstname) = record(flds.Firstname) & ""
                    dicD(i)(flds.UR) = record(flds.UR) & ""
                    If Not IsDBNull(record(flds.DOB)) Then dicD(i)(flds.DOB) = record(flds.DOB) Else dicD(i)(flds.DOB) = ""
                    dicD(i)(flds.Gender) = record(flds.Gender) & ""
                    i = i + 1
                Next
            End With
            Return dicD
        End If

    End Function

    Public Function Get_PtNameString(ByVal PatientID As Long, ByVal format As eNameStringFormats) As String

        If PatientID > 0 Then
            Dim dicD As Dictionary(Of String, String) = Me.get_demographics(PatientID)
            Dim d As New class_DemographicFields

            Select Case format
                Case eNameStringFormats.FirstnameSpaceSurname : Return dicD(d.Firstname) & " " & dicD(d.Surname)
                Case eNameStringFormats.Name_UR : Return dicD(d.Surname) & ", " & dicD(d.Firstname) & "    UR: " & dicD(d.UR)
                Case eNameStringFormats.SurnameCommaFirstname : Return dicD(d.Surname) & ", " & dicD(d.Firstname)
                Case eNameStringFormats.NameForPdfFilename : Return dicD(d.Firstname) & "_" & dicD(d.Surname)
                Case Else : Return ""
            End Select
        Else
            Return ""
        End If

    End Function

    Public Function get_demographics(patientID As Long) As Dictionary(Of String, String)

        Dim sql As String
        Dim msg As String = ""
        Dim r As DataRow
        Dim Ds As DataSet

        sql = "SELECT pas_pt.*, pas_pt_address.*, pas_pt_names.* "
        sql = sql & "FROM pas_pt LEFT OUTER JOIN pas_pt_address ON pas_pt.patientID = pas_pt_address.patientID LEFT OUTER JOIN "
        sql = sql & "pas_pt_names ON pas_pt.patientID = pas_pt_names.patientID "
        sql = sql & "WHERE (pas_pt_names.name_type = 'legal') AND (pas_pt.patientID = '" & patientID & "') "

        Ds = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Select Case Ds.Tables(0).Rows.Count
                Case 0
                    'MsgBox "Should never happen error - Cant find demographics records for this Pt"
                    Return Nothing
                Case Else
                    'If > 1 record returned then >1 home address and/or >1 legal name for this patient - shouldn't really ever happen
                    If Ds.Tables(0).Rows.Count > 1 Then
                        msg = ""
                        For Each r In Ds.Tables(0).Rows
                            msg = msg & r("Surname") & ", " & r("FirstName") & "   " & r("Address_1") & ", " & r("Suburb") & vbCrLf & vbCrLf
                        Next
                        msg = "More than 1 legal name and/or home address for this patient -" & vbCrLf & vbCrLf & msg & vbCrLf & "The first record will be selected."
                        msg = msg & vbCrLf & "Please rectify the problem."
                        MsgBox(msg, vbOKOnly, "ResLab")
                    End If

                    Dim f As New class_DemographicFields
                    Dim p As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicDemo
                    r = Ds.Tables(0).Select.First()
                    p(f.PatientID) = r(f.PatientID)
                    p(f.Title) = StrConv(r(f.Title), VbStrConv.ProperCase)
                    p(f.Firstname) = StrConv(r(f.Firstname), VbStrConv.ProperCase)
                    p(f.Surname) = r(f.Surname).ToString.ToUpper
                    p(f.DOB) = Format(r(f.DOB), "dd/MM/yyyy")
                    p(f.Gender) = cMyRoutines.Lookup_list_ByCode(r(f.Gender_code), eTables.list_Gender)
                    p(f.Gender_code) = Me.map_refTbl(r(f.Gender_code), eTables.list_Gender, eMapMode.HL7Code_RMCode)
                    p(f.Race_forRfts) = cMyRoutines.Lookup_list_ByCode(r(f.Race_forRfts_code), eTables.List_RacesForRFTs)
                    p(f.Race_forRfts_code) = r(f.Race_forRfts_code)
                    p(f.Address_1) = StrConv(r(f.Address_1) & "", VbStrConv.ProperCase)
                    p(f.Address_2) = StrConv(r(f.Address_2) & "", VbStrConv.ProperCase)
                    p(f.Suburb) = StrConv(r(f.Suburb) & "", VbStrConv.ProperCase)
                    p(f.State) = r(f.State) & ""
                    p(f.PostCode) = r(f.PostCode) & ""
                    p(f.Email) = r(f.Email) & ""
                    p(f.Phone_home) = r(f.Phone_home) & ""
                    p(f.Phone_work) = r(f.Phone_work) & ""
                    p(f.Phone_mobile) = r(f.Phone_mobile) & ""
                    p(f.Death_date) = r(f.Death_date) & ""
                    p(f.Death_indicator) = r(f.Death_indicator) & ""
                    p(f.Medicare_No) = r(f.Medicare_No) & ""
                    p(f.Medicare_expirydate) = r(f.Medicare_expirydate) & ""
                    p(f.CountryOfBirth_code) = r(f.CountryOfBirth_code) & ""
                    p(f.CountryOfBirth) = cMyRoutines.Lookup_list_ByCode(r(f.CountryOfBirth_code), eTables.list_Nationality)
                    p(f.PreferredLanguage_code) = r(f.PreferredLanguage_code) & ""
                    p(f.PreferredLanguage) = cMyRoutines.Lookup_list_ByCode(r(f.PreferredLanguage_code), eTables.list_Language)
                    p(f.AboriginalStatus_code) = r(f.AboriginalStatus_code) & ""
                    p(f.AboriginalStatus) = cMyRoutines.Lookup_list_ByCode(r(f.AboriginalStatus_code), eTables.list_AboriginalStatus)

                    p(f.UR) = Me.get_ur_for_hsid(patientID, cHs.Current_UserLocation_HSID_code, eURformats.URonly)
                    p(f.URandHS) = Me.get_ur_for_hsid(patientID, cHs.Current_UserLocation_HSID_code, eURformats.URandHS)
                    p(f.URandHS_all_asString) = Me.get_ur_all_string(patientID, eURformats.URandHS)
                    p(f.UR_hsid) = cHs.Current_UserLocation_HSID_code
                    p(f.UR_id) = Nothing

                    p(f.gpID) = r(f.gpID)

                    If IsDate(r(f.Lastupdated_date)) Then p(f.Lastupdated_date) = r(f.Lastupdated_date) Else p(f.Lastupdated_date) = Nothing
                    p(f.Lastupdated_by) = r(f.Lastupdated_by) & ""

                    Return p
            End Select

        End If

    End Function

    Public Function get_deathInfo(patientID As Long) As DeathInfo
        'Trak [Date of Death] = valid date PLUS Trak [DeathIndicator] = Y   --> Accurate deathdate
        'Trak [Date of Death] = empty      PLUS Trak [DeathIndicator] = Y   --> estimated death date (text not available to ResMed)

        Dim sql As String = "SELECT DateOfDeath, DeathIndicator FROM tbl_Patient WHERE ID =" & patientID
        Dim Info As DeathInfo = Nothing

        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Info.DeathDate = ""
            Info.DeathStatusString = ""
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Info.DeathDate = ""
            Info.DeathStatusString = ""
        Else
            Select Case cMyRoutines.IsRealDate(ds.Tables(0).Rows(0)("DateOfDeath"))
                Case True
                    Info.DeathDate = ds.Tables(0).Rows(0)("DateOfDeath")
                    Info.DeathStatusString = Me.map_refTbl(ds.Tables(0).Rows(0)("DeathIndicator"), eTables.list_DeathIndicator, eMapMode.HL7Code_RMText)
                Case False
                    If ds.Tables(0).Rows(0)("DeathIndicator") = "Y" Then
                        Info.DeathStatusString = Me.map_refTbl(ds.Tables(0).Rows(0)("DeathIndicator"), eTables.list_DeathIndicator, eMapMode.HL7Code_RMText)
                        Info.DeathDate = "Estimated death date"
                    Else
                        Info.DeathStatusString = Me.map_refTbl(ds.Tables(0).Rows(0)("DeathIndicator"), eTables.list_DeathIndicator, eMapMode.HL7Code_RMText)
                        Info.DeathDate = ""
                    End If
            End Select
        End If

        Return Info
        ds = Nothing

    End Function

    Public Function get_patient_namestring(patientID_primary As Long, HealthServiceID As eHealthServiceIDs, ReturnItem As eNameURFormats) As String

        Dim PrimaryUR As String = ""
        Dim DisplayUR As String = ""
        Dim Surname As String = ""
        Dim Firstname As String = ""
        Dim text As String = ""
        Dim HealthService As String = ""

        If patientID_primary = 0 Then Return ""

        'Get patients name
        Dim sql As String = "SELECT  tbl_patient.*, tbl_Name.*, tbl_PatientIdentifier.* "
        sql = sql & "FROM tbl_patient INNER JOIN tbl_PatientIdentifier "
        sql = sql & "ON tbl_patient.ID = tbl_PatientIdentifier.PatientID AND tbl_PatientIdentifier.idTypeCode = 'MR' "
        sql = sql & "INNER JOIN tbl_Name ON tbl_patient.ID = tbl_Name.TypeID AND tbl_Name.Type = 'P' "
        sql = sql & "WHERE tbl_patient.ID = '" & patientID_primary & "'"
        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            'Only occurs on DB error - no related record in tbl_Name for this PrimaryPtID
            MsgBox("Database error - no related record in tbl_Name for this PrimaryPtID. Aborted.", vbOKOnly, "ResLab")
            Return ""
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            'Only occurs on DB error - no related record in tbl_Name for this PrimaryPtID
            MsgBox("Database error - no related record in tbl_Name for this PrimaryPtID. Aborted.", vbOKOnly, "ResLab")
            Return ""
        Else
            Surname = ds.Tables(0).Rows(0)("Surname")
            Firstname = ds.Tables(0).Rows(0)("FirstName")

            'Get health service string
            If Left(ds.Tables(0).Rows(0)("URNumber"), 1) = "R" Then
                HealthService = "  (Internal DRSM)"
            Else
                HealthService = "  (Austin Health)"
            End If
            PrimaryUR = ds.Tables(0).Rows(0)("URNumber") & HealthService

            'Get site specific UR for display



            Select Case ReturnItem
                Case eNameURFormats.FullName : text = Surname & ", " & Firstname
                Case eNameURFormats.Lastname : text = Surname
                Case eNameURFormats.Firstname : text = Firstname
                Case eNameURFormats.FullName_DisplayUR : text = Surname & ", " & Firstname & "    UR: " & DisplayUR
                    'Case eNameURFormats.FullName_PrimaryUR : text = Surname & ", " & Firstname & "    UR: " & PrimaryUR
                Case eNameURFormats.PrimaryUR : text = PrimaryUR
                Case eNameURFormats.DisplayUR : text = DisplayUR
                    'Case eNameURFormats.URforBarcode : text = BarcodeUR
                Case Else : text = ""
            End Select

            Return text
        End If

    End Function

    Public Function get_ur_all_string(patientID As Integer, eFormat As eURformats, Optional eHealthServiceID As eHealthServiceIDs = eHealthServiceIDs.Unknown) As String

        Dim d() As Dictionary(Of String, String) = Me.get_ur_all(patientID, eHealthServiceID)

        If IsNothing(d) Then
            Return ""
        Else
            Dim urs As String = ""
            Dim ur As String = ""
            Dim f As New class_fields_ur
            For Each ur_dic As Dictionary(Of String, String) In d
                ur = ur_dic(f.UR)
                If eFormat = eURformats.URandHS Then ur = ur & " (" & cHs.get_healthservice_name_fromHSID(ur_dic(f.UR_hsid)) & ")"
                urs = urs & ur & ", "
            Next
            If Right(urs, 2) = ", " Then urs = Left(urs, Len(urs) - 2)
            Return urs
        End If

    End Function

    Public Function get_ur_all(patientID As Integer, Optional HealthServiceID As eHealthServiceIDs = eHealthServiceIDs.Unknown) As Dictionary(Of String, String)()

        Dim sql As String = "SELECT * FROM pas_pt_ur_numbers WHERE patientID=" & patientID
        If HealthServiceID <> eHealthServiceIDs.Unknown Then sql = sql & " AND UR_hsid='" & HealthServiceID & "'"

        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return Nothing  'should never happen
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Return Nothing  'should never happen
        Else
            Dim i As Integer = 0
            Dim d() As Dictionary(Of String, String) = Nothing
            Dim f As New class_fields_ur
            For Each r As DataRow In ds.Tables(0).Rows
                ReDim Preserve d(i)
                d(i) = cMyRoutines.MakeEmpty_dicUR
                d(i)(f.UR_id) = r(f.UR_id)
                d(i)(f.patientID) = r(f.patientID)
                d(i)(f.UR) = Trim(r(f.UR))
                d(i)(f.UR_mergeto) = r(f.UR_mergeto) & ""
                d(i)(f.UR_hsid) = r(f.UR_hsid)
                d(i)(f.created_inResLab) = r(f.created_inResLab)
                d(i)(f.create_by) = r(f.create_by) & ""
                If cMyRoutines.IsRealDate(r(f.create_date)) Then d(i)(f.create_date) = r(f.create_date) Else d(i)(f.create_date) = Nothing
                i = i + 1
            Next
            Return d
        End If

    End Function

    Public Function get_ur_for_hsid(patientID As Integer, eHealthServiceID As eHealthServiceIDs, eFormat As eURformats) As String

        Dim sql As String = "SELECT UR FROM pas_pt_ur_numbers WHERE patientID=" & patientID & " AND UR_hsid='" & eHealthServiceID & "'"

        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return ""  'should never happen
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Return ""  'should never happen
        Else
            Dim ur As String = Trim(ds.Tables(0).Rows(0)("UR"))
            If eFormat = eURformats.URandHS Then ur = ur & " (" & cHs.get_healthservice_name_fromHSID(eHealthServiceID) & ")"
            Return ur
        End If

    End Function

    Public Function insert_Demographics(ByVal D As Dictionary(Of String, String)) As Long

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.PatientDetails, D)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new demographics" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

    End Function

    Public Function map_refTbl(Item As String, eTbl As eTables, mode As eMapMode) As Object

        Dim sql As String = ""
        Dim Tbl As String = ""
        Const HL7Code As String = "HL7_CODE"
        Const HL7Desc As String = "HL7_DESCRIPTION"
        Const RESLABCode As String = "code"
        Const RESLABDesc As String = "description"

        Select Case mode
            Case eMapMode.HL7Code_RMText : sql = "SELECT " & RESLABDesc & " AS Value FROM " & cDbInfo.table_name(eTbl) & " WHERE " & HL7Code & "='" & Item & "';"
            Case eMapMode.HL7Code_RMCode : sql = "SELECT " & RESLABCode & " AS Value  FROM " & cDbInfo.table_name(eTbl) & " WHERE " & HL7Code & "='" & Item & "';"
            Case eMapMode.HL7Code_HL7Text : sql = "SELECT " & HL7Desc & " AS Value  FROM " & cDbInfo.table_name(eTbl) & " WHERE " & HL7Code & "='" & Item & "';"
            Case eMapMode.RMCode_HL7Code : sql = "SELECT " & HL7Code & " AS Value  FROM " & cDbInfo.table_name(eTbl) & " WHERE " & RESLABCode & "='" & Item & "';"
            Case eMapMode.RMCode_HL7Text : sql = "SELECT " & HL7Desc & " AS Value  FROM " & cDbInfo.table_name(eTbl) & " WHERE " & RESLABCode & "='" & Item & "';"
            Case eMapMode.RMCode_RMText : sql = "SELECT " & RESLABDesc & " AS Value  FROM " & cDbInfo.table_name(eTbl) & " WHERE " & RESLABCode & "='" & Item & "';"
        End Select

        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return Nothing
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Return ds.Tables(0).Rows(0)("Value") & ""
        End If

        ds = Nothing

    End Function


    Private Function find_patientID_ByNameEtc_buildSql(ByVal Surname As String, _
                                               ByVal SurnameBy As String, _
                                               Optional ByVal Firstname As String = "", _
                                               Optional ByVal FirstNameBy As String = "", _
                                               Optional ByVal DOB As String = "", _
                                               Optional ByVal Gender As String = "") As String
        Throw New NotImplementedException
    End Function

    Public Function find_patientID_ByNameEtc(sql As String) As Long()

        Dim patientID(0) As Long
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In Ds.Tables(0).Rows
                ReDim Preserve patientID(UBound(patientID) + 1)
                patientID(UBound(patientID)) = row("PatientID")
            Next
        End If

        Return patientID

    End Function




    Public Function insert_patient(data As Dictionary(Of String, String)) As Long
        Throw New NotImplementedException
    End Function

    Public Function update_patient(data As Dictionary(Of String, String)) As Boolean
        Throw New NotImplementedException
    End Function

    Public Function delete_patient(patientID As Long) As Boolean
        Throw New NotImplementedException
    End Function



End Class




