Module module_intialise_globals

    Public colFindBys As New Collection

    Public cRfts As New class_pas_rfts
    Public cPt As New class_pas

    Public cHs As New class_healthservices
    Public cWalkPlot As New class_walktest_plot
    Public cWalk As New class_walktest
    Public cMyRoutines As New cGeneralRoutines
    Public cDAL As New class_DataAccessLayer
    Public cDbInfo As New cDatabaseInfo
    Public cPred As New class_Pred
    Public cPrefs As New class_Preferences
    Public cChall As New class_challenge
    Public cBuildTsMenu As New class_buildtoolstripmenu
    Public cPhrases As New class_reportphrases
    Public cSpt As New class_spt
    Public cTrend As New class_plot_trend
    Public cCpet As New class_plot_cpet
    Public cUser As class_user  'instance created in form_login
    Public gRefreshMainForm As Boolean
    Public cAppConfig As New class_app_config

    Public gUserName As String
    Public gSiteID As String

    Public gURlabel As String = get_URlabel()

    Private Function get_URlabel() As String

        Dim defaultLabel As String = "UR"

        Dim sql As String = "SELECT prefs_fielditems.fielditem FROM prefs_fields INNER JOIN prefs_fielditems ON prefs_fields.default_fielditem_id = prefs_fielditems.prefs_id "
        sql = sql & "WHERE prefs_fields.fieldname = 'MRN style'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)(0)
            Else
                Return defaultLabel
            End If
        Else
            Return defaultLabel
        End If

    End Function

End Module
