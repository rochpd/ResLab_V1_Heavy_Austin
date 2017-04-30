


Public Class class_healthservices

    Public Enum eHealthServiceIDs
        nonhospital_service = 0
        austin_internal_drsm = 1
        Unknown = -1

        alfred_hospital = 1010
        austin_health = 1031
        bendigo_hospital = 1021
        bethlehem_hospital = 3050
        boxhill_hospital = 1050
        frankston_hospital = 2220
        'john_hunter_hospital=
        mercy_hospital_for_women = 1160
        monash_medical_centre = 1170
        northern_hospital = 1280
        western_hospital = 1180
    End Enum

    Private _Current_UserLocation_HSID_code As Integer
    Private _Current_UserLocation_HSID_Name As String

    Public Property Current_UserLocation_HSID_code() As Integer
        'Used to determine which UR to use 
        Get
            Return _Current_UserLocation_HSID_code
        End Get
        Set(ByVal value As Integer)
            _Current_UserLocation_HSID_code = value
            _Current_UserLocation_HSID_Name = Me.get_healthservice_name_fromHSID(_Current_UserLocation_HSID_code)
        End Set
    End Property

    Public ReadOnly Property Current_UserLocation_HSID_Name() As String
        Get
            Return _Current_UserLocation_HSID_Name
        End Get
    End Property

    Public Function get_healthservice_name_fromHSID(HSID As Integer) As String

        Dim sql As String = "SELECT description FROM List_HealthServices WHERE code=" & HSID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Return Ds.Tables(0).Rows(0)("description")
        End If

        Ds = Nothing

    End Function

    Public Function get_healthservice_HSID_fromName(hs_name As String) As Integer

        Dim sql As String = "SELECT code FROM List_HealthServices WHERE description='" & hs_name & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Return Ds.Tables(0).Rows(0)("code")
        End If

        Ds = Nothing

    End Function

    Public Function get_list_Healthservices(Optional state As String = "", Optional only_VRSS As Boolean = False, Optional only_enabled As Boolean = True) As List(Of String)

        Dim sqlState As String = "", sqlOnlyVrss As String = "", sqlEnabled As String = ""

        Dim sql As String = "SELECT * FROM List_HealthServices "
        If state <> "" Then sqlState = " state = '" & state & "' AND "
        If only_VRSS = True Then sqlOnlyVrss = " austin_vrss_satellite = 1 AND "
        If only_enabled = True Then sqlEnabled = " enabled_global = 1 AND "

        If sqlState <> "" Or sqlOnlyVrss = "" Or sqlEnabled = "" Then
            sql = sql & " WHERE "
            If sqlState <> "" Then sql = sql & sqlState
            If sqlOnlyVrss <> "" Then sql = sql & sqlOnlyVrss
            If sqlEnabled <> "" Then sql = sql & sqlEnabled
            sql = Left(sql, Len(sql) - 4)
        End If
        sql = sql & " ORDER BY description;"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        ElseIf Ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            Dim lst As New List(Of String)
            For Each r As DataRow In Ds.Tables(0).Rows
                lst.Add(r("description"))
            Next
            Return lst
        End If

        Ds = Nothing

    End Function

End Class
