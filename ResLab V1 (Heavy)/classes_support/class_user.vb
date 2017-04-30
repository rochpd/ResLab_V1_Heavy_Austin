Imports System.Text
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_user

    Private _currentUser_name As String
    Private _currentUser_login_healthservice_code_default As String
    Private _Rft_can_report As Boolean
    Private _Rft_can_verifyreports As Boolean
    Private _Rft_can_request As Boolean
    Private _Rft_can_autoreport As Boolean
    Private _AccessLevel As String

    Public Property CurrentUser_name As String
        Get
            Return _currentUser_name
        End Get
        Set(CurrentUser_name As String)
            _currentUser_name = CurrentUser_name
        End Set
    End Property
    Public ReadOnly Property CurrentUser_login_healthservice_code_default As String
        Get
            Return _currentUser_login_healthservice_code_default
        End Get
    End Property
    Public Property AccessLevel As String
        Get
            Return _AccessLevel
        End Get
        Set(AccessLevel As String)
            _AccessLevel = AccessLevel
        End Set
    End Property
    Public Property Rft_can_autoreport As Boolean
        Get
            Return _Rft_can_autoreport
        End Get
        Set(Rft_can_autoreport As Boolean)
            _Rft_can_autoreport = Rft_can_autoreport
        End Set
    End Property
    Public Property Rft_can_report As Boolean
        Get
            Return _Rft_can_report
        End Get
        Set(Rft_can_report As Boolean)
            _Rft_can_report = Rft_can_report
        End Set
    End Property
    Public Property Rft_can_verifyreports As Boolean
        Get
            Return _Rft_can_verifyreports
        End Get
        Set(Rft_can_verifyreports As Boolean)
            _Rft_can_verifyreports = Rft_can_verifyreports
        End Set
    End Property
    Public Property Rft_can_request As Boolean
        Get
            Return _Rft_can_request
        End Get
        Set(Rft_can_request As Boolean)
            _Rft_can_request = Rft_can_request
        End Set
    End Property

    Public Enum ePermissions    'field names in table, must be lower case
        access
        rft_can_report
        rft_can_verifyreports
        rft_can_request
        rft_can_autoreport
    End Enum

    Public Enum ePermission_fields
        person_permissionid
        permissionid
        personid
        value
        lastupdated
    End Enum

    Public Structure permission_type
        Dim permissionid As Integer
        Dim description As String
        Dim displaytext As String
        Dim datatype As String
        Dim tablename_list As String
        Dim defaultvalue As String
        Dim enabled As Boolean
    End Structure

    Public Sub New(currentusername As String)

        Try

            _currentUser_name = currentusername
            _currentUser_login_healthservice_code_default = Me.get_currentUser_login_hs_code_default()

            Dim personID As Integer = Me.get_personID_from_username(_currentUser_name)

            'Get persmissions
            Dim p As Dictionary(Of String, String) = Me.get_user_permissions_dic(personID, True)
            If Not IsNothing(p) Then
                _Rft_can_report = p(ePermissions.rft_can_report.ToString)
                _Rft_can_verifyreports = p(ePermissions.rft_can_verifyreports.ToString)
                _Rft_can_request = p(ePermissions.rft_can_request.ToString)
                _Rft_can_autoreport = p(ePermissions.rft_can_autoreport.ToString)
                _AccessLevel = p(ePermissions.access.ToString)
            Else
                _Rft_can_report = False
                _Rft_can_verifyreports = False
                _Rft_can_request = False
                _Rft_can_autoreport = False
                _AccessLevel = "read_only"
            End If

        Catch ex As Exception
            MsgBox("Error in class_user.new" & vbCrLf & Err.Description)

        End Try

    End Sub

    Public Function get_personID_from_username(username As String) As Long

        Dim sql As String = "SELECT personID FROM persons WHERE user_name='" & username & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)("personID")
            Else
                Return 0
            End If
        Else
            Return 0
        End If

    End Function

    Public Function get_currentUser_login_hs_code_default() As String

        Dim sql As String = "SELECT login_healthservice_code_default FROM persons WHERE user_name='" & Me._currentUser_name & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)("login_healthservice_code_default")
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Public Function get_username_from_userID(personID As Long) As String

        Dim sql As String = "SELECT user_name FROM persons WHERE personID=" & personID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)("user_name")
            Else
                Return 0
            End If
        Else
            Return 0
        End If

    End Function

    Public Function get_permissions_listofvalues(tbl As String, fld As String, enabledonly As Boolean) As String()

        Dim s As New List(Of String)
        Dim sql As String = "SELECT " & fld & " FROM " & tbl
        If enabledonly Then sql = sql & " WHERE enabled=1"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Not IsNothing(Ds) Then
            For Each v In Ds.Tables(0).Rows
                s.Add(v(fld))
            Next
            Return s.ToArray
        Else
            Return Nothing
        End If

    End Function

    Public Function get_permissions_types(enabledonly As Boolean) As List(Of permission_type)

        Try
            Dim p As New List(Of permission_type)
            Dim sql As New StringBuilder
            sql.Clear()
            sql.Append("SELECT * FROM list_permissiontypes ")
            If enabledonly Then sql.Append(" WHERE list_permissiontypes.enabled=1")

            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
            Dim perm As permission_type = Nothing
            If Not IsNothing(Ds) Then
                For Each row In Ds.Tables(0).Rows
                    perm = New permission_type
                    perm.permissionid = row("code")
                    perm.tablename_list = row("tablename_list") & ""
                    perm.enabled = row("enabled")
                    perm.description = row("description") & ""
                    perm.displaytext = row("displaytext") & ""
                    perm.defaultvalue = row("defaultvalue") & ""
                    perm.datatype = row("datatype") & ""
                    p.Add(perm)
                Next
                Return p
            Else
                Return Nothing
            End If
        Catch ex As Exception
            MsgBox("Error in class_user.get_permissions_types" & vbCrLf & Err.Description)
            Return Nothing
        End Try

    End Function

    Public Function get_user_permissions_dics(userid As Long, enabledonly As Boolean) As Dictionary(Of String, String)()
        'Returns a dictionary for each permission

        Dim sql As New StringBuilder
        sql.Clear()
        sql.Append("SELECT person_permissions.person_permissionid, person_permissions.value, list_permissiontypes.description, list_permissiontypes.datatype, list_permissiontypes.code AS type_code from person_permissions INNER JOIN list_permissiontypes ")
        sql.Append("ON person_permissions.permissionID = list_permissiontypes.code ")
        sql.Append("WHERE personID=" & userid)
        If enabledonly Then
            sql.Append(" AND list_permissiontypes.enabled=1")
        End If

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            Dim p() As Dictionary(Of String, String) = Nothing
            Dim i As Integer = 0
            For Each row In Ds.Tables(0).Rows
                ReDim Preserve p(i)
                p(i) = New Dictionary(Of String, String)
                p(i).Add("description", row("description"))
                p(i).Add("permissionid", row("type_code"))
                p(i).Add("value", row("value"))
                p(i).Add("datatype", row("datatype"))
                p(i).Add("person_permissionid", row("person_permissionid"))
                i = i + 1
            Next
            Return p
        Else
            Return Nothing
        End If

    End Function

    Public Function get_user_permissions_dic(userid As Long, enabledonly As Boolean) As Dictionary(Of String, String)
        'Returns a single dictionary of key/value pairs ie no other fields other than description and value

        Try

            Dim i As Integer = 0
            Dim sql As New StringBuilder
            sql.Clear()
            sql.Append("SELECT person_permissions.*, list_permissiontypes.* from person_permissions INNER JOIN list_permissiontypes ")
            sql.Append("ON person_permissions.permissionID = list_permissiontypes.code ")
            sql.Append("WHERE personID=" & userid)
            If enabledonly Then
                sql.Append(" AND list_permissiontypes.enabled=1")
            End If

            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
            If Not IsNothing(Ds) Then
                Dim p As New Dictionary(Of String, String)
                For i = 0 To Ds.Tables(0).Rows.Count - 1
                    p.Add(Ds.Tables(0).Rows(i)("description"), Ds.Tables(0).Rows(i)("value"))
                Next
                Return p
            Else
                Return Nothing
            End If

        Catch ex As Exception
            MsgBox("Error in class_user.new" & vbNewLine & ex.Message.ToString)
            Return Nothing
        End Try

    End Function

    Public Function update_user_permissions(p As Dictionary(Of String, String)()) As Boolean

        Dim d As Dictionary(Of String, String) = Nothing
        Dim sql As String = ""
        Dim Success As Boolean
        Dim RecordID As Long = 0
        Dim nFailed As Integer = 0

        Try

            For Each d In p
                If d("person_permissionid") = 0 Then
                    sql = cDAL.Build_InsertQuery(eTables.person_permissions, d)
                    Success = cDAL.Insert_Record(sql)
                Else
                    sql = cDAL.Build_UpdateQuery(eTables.person_permissions, d)
                    Success = cDAL.Update_Record(sql)
                End If
                If Not Success Then
                    nFailed = nFailed + 1
                End If
            Next
            If nFailed > 0 Then
                MsgBox("Error saving " & nFailed & " of " & d.Count & " permission records (in class_user.update_user_permissions)")
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            MsgBox("Error saving permission - aborted (in class_user.update_user_permissions)" & vbNewLine & ex.Message.ToString)
            Return False

        End Try

    End Function

    Public Function delete_user_permissions(personid As Long) As Boolean



        Return False

    End Function

    Public Function insert_user_permissions(personID As Long, p() As Dictionary(Of String, String)) As Boolean

        Dim newRecordID As Long = 0
        Dim sql As String = ""
        Dim Ds As DataSet = Nothing
        Dim Success As Boolean
        Dim nFailed As Integer = 0

        For Each d As Dictionary(Of String, String) In p
            d("personid") = personID
            sql = cDAL.Build_InsertQuery(eTables.person_permissions, d)
            Success = cDAL.Insert_Record(sql)
        Next

        If nFailed > 0 Then
            MsgBox("Error saving " & nFailed & " of " & p.Count & " permission records (in class_user.insert_user_permissions)")
            Return False
        Else
            Return True
        End If

    End Function

    Public Function Get_all_users(enabled_only As Boolean) As Dictionary(Of Long, String)

        Dim sql As String = "SELECT personID, user_name FROM persons "
        If enabled_only Then sql = sql & "WHERE enabled=1"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count > 0 Then
                Dim users As New Dictionary(Of Long, String)
                For Each row In Ds.Tables(0).Rows
                    users.Add(row("personid"), row("user_name"))
                Next
                Return users
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If

    End Function

    Public Function Get_user_details(personID As Long) As Dictionary(Of String, String)

        Dim sql As String = "SELECT * FROM persons WHERE personid=" & personID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            Select Case Ds.Tables(0).Rows.Count
                Case 0 : Return Nothing
                Case 1
                    Dim flds As New class_Fields_User
                    Dim u As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicUser
                    u(flds.Title) = Ds.Tables(0).Rows(0)(flds.Title) & ""
                    u(flds.Surname) = Ds.Tables(0).Rows(0)(flds.Surname) & ""
                    u(flds.Firstname) = Ds.Tables(0).Rows(0)(flds.Firstname) & ""
                    If IsDate(Ds.Tables(0).Rows(0)(flds.DOB)) Then u(flds.DOB) = Ds.Tables(0).Rows(0)(flds.DOB) Else u(flds.DOB) = Nothing
                    u(flds.Gender) = Ds.Tables(0).Rows(0)(flds.Gender) & ""
                    u(flds.Address_1) = Ds.Tables(0).Rows(0)(flds.Address_1) & ""
                    u(flds.Address_2) = Ds.Tables(0).Rows(0)(flds.Address_2) & ""
                    u(flds.Suburb) = Ds.Tables(0).Rows(0)(flds.Suburb) & ""
                    u(flds.State) = Ds.Tables(0).Rows(0)(flds.State) & ""
                    u(flds.Profession_category) = Ds.Tables(0).Rows(0)(flds.Profession_category) & ""
                    u(flds.Department) = Ds.Tables(0).Rows(0)(flds.Department) & ""
                    u(flds.Institution) = Ds.Tables(0).Rows(0)(flds.Institution) & ""
                    u(flds.PostCode) = Ds.Tables(0).Rows(0)(flds.PostCode) & ""
                    u(flds.Phone_home) = Ds.Tables(0).Rows(0)(flds.Phone_home) & ""
                    u(flds.Phone_mobile) = Ds.Tables(0).Rows(0)(flds.Phone_mobile) & ""
                    u(flds.Phone_work) = Ds.Tables(0).Rows(0)(flds.Phone_work) & ""
                    u(flds.Email) = Ds.Tables(0).Rows(0)(flds.Email) & ""
                    u(flds.User_name) = Ds.Tables(0).Rows(0)(flds.User_name) & ""
                    u(flds.User_password) = Ds.Tables(0).Rows(0)(flds.User_password) & ""
                    u(flds.Last_login) = Ds.Tables(0).Rows(0)(flds.Last_login) & ""
                    Return u
                Case Else : Return Nothing
            End Select
        Else
            Return Nothing
        End If

    End Function

    Public Function Insert_user_details(user As Dictionary(Of String, String)) As Long

        Dim sql As String = cDAL.Build_InsertQuery(eTables.persons, user)
        Dim id As Long = cDAL.Insert_Record(sql)

        Return id

    End Function

    Public Function Update_user_details(user As Dictionary(Of String, String)) As Boolean

        Dim sql As String = cDAL.Build_UpdateQuery(eTables.persons, user)
        Dim success As Boolean = cDAL.Update_Record(sql)

        Return success

    End Function

    Public Function Find_user(username As String) As Long
        'returns 0 if not found, personID if found, negative count if more than one match found (should never happen)

        Dim sql As String = "SELECT * FROM persons WHERE user_name='" & username & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            Select Case Ds.Tables(0).Rows.Count
                Case 0 : Return 0
                Case 1 : Return Ds.Tables(0).Rows(0)("personID")
                Case Else : Return Ds.Tables(0).Rows.Count * -1
            End Select
        Else
            Return 0
        End If

    End Function

    Public Sub set_access(f As Form)

        Dim IsReadWrite As Boolean = False
        Select Case Me._AccessLevel
            Case "read_only" : IsReadWrite = False
            Case "read_write", "administrator" : IsReadWrite = True
        End Select

        Dim c As Control = Nothing

        Select Case f.Name.ToLower
            Case "form_mainnew"
                'c = f.Controls.Item(f.Controls.IndexOfKey("tsMainMenu"))
                'If c.Controls.ContainsKey("tsMenuItem_EditTest") Then c.Controls.Item(c.Controls.IndexOfKey("tsMenuItem_EditTest")).Enabled = IsReadWrite

            Case "form_demographics"
                f.Controls.Item(f.Controls.IndexOfKey("btnSave")).Enabled = IsReadWrite

            Case "form_rft", "form_walktest", "form_rft_cpet", "form_rft_challenge", "form_hast", "form_spt"
                f.Controls.Item(f.Controls.IndexOfKey("btnSave")).Enabled = IsReadWrite
                f.Controls.Item(f.Controls.IndexOfKey("btnDelete")).Enabled = IsReadWrite
            Case "form_prefs_list"
                c = cMyRoutines.FindControl(f, "tsbFieldOptions_new")
        End Select



    End Sub
End Class
