Public Class class_app_config

    Private _site_id As String
    Private _site_name As String
    Private _site_institution As String
    Private _site_state As String
    Private _db_type As String
    Private _db_name As String
    Private _db_servername As String
    Private _db_connectstring As String
    Private _pas_mode_local As Boolean

    Public ReadOnly Property site_id As String
        Get
            site_id = _site_id
        End Get
    End Property
    Public ReadOnly Property site_name As String
        Get
            site_name = _site_name
        End Get
    End Property
    Public ReadOnly Property site_institution As String
        Get
            site_institution = _site_institution
        End Get
    End Property
    Public ReadOnly Property site_state As String
        Get
            site_state = _site_state
        End Get
    End Property
    Public ReadOnly Property db_type As String
        Get
            db_type = _db_type
        End Get
    End Property
    Public ReadOnly Property db_name As String
        Get
            db_name = _db_name
        End Get
    End Property
    Public ReadOnly Property db_servername As String
        Get
            db_servername = _db_servername
        End Get
    End Property
    Public ReadOnly Property db_connectstring As String
        Get
            db_connectstring = _db_connectstring
        End Get
    End Property
    Public ReadOnly Property pas_mode_local As Boolean
        Get
            pas_mode_local = _pas_mode_local
        End Get
    End Property

    Private Function get_configdata_fromdb() As Dictionary(Of String, String)


        Dim sql As String = "SELECT * FROM site_config"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 1 Then
            Dim f As New class_Fields_app_config
            Dim d As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dicApp_config
            d(f.configID) = Ds.Tables(0).Rows(0)(f.configID)
            d(f.siteID) = Ds.Tables(0).Rows(0)(f.siteID)
            d(f.site_name) = Ds.Tables(0).Rows(0)(f.site_name)
            d(f.site_institution) = Ds.Tables(0).Rows(0)(f.site_institution)
            d(f.site_state) = Ds.Tables(0).Rows(0)(f.site_state)
            d(f.db_type) = Ds.Tables(0).Rows(0)(f.db_type)
            d(f.db_name) = Ds.Tables(0).Rows(0)(f.db_name)
            d(f.db_servername) = Ds.Tables(0).Rows(0)(f.db_servername)
            d(f.db_connectstring) = Ds.Tables(0).Rows(0)(f.db_connectstring)
            d(f.pas_mode_local) = Ds.Tables(0).Rows(0)(f.pas_mode_local)
            Return d
        Else
            Return Nothing
        End If

    End Function

    Public Sub intialise_app_config()

        Dim d As Dictionary(Of String, String) = Me.get_configdata_fromdb
        Dim f As New class_Fields_app_config
        If Not IsNothing(d) Then
            Me._pas_mode_local = d(f.pas_mode_local)
            Me._site_id = d(f.siteID)
            Me._site_name = d(f.site_name)
            Me._site_institution = d(f.site_institution)
            Me._site_state = d(f.site_state)
            Me._db_type = d(f.db_type)
            Me._db_name = d(f.db_name)
            Me._db_servername = d(f.db_servername)
            Me._db_connectstring = d(f.db_connectstring)
        End If

    End Sub

End Class
