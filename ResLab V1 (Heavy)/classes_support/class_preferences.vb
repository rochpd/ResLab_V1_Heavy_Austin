Imports System.Text
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_Preferences

#Region "Application options"
    Public app_enabled_SptTest As Boolean
    Public app_enabled_SptTest_AllergenManager As Boolean
    Public app_enabled_WalkTest_SixMWD_ATS As Boolean
    Public app_enabled_WalkTest_NRG As Boolean
    Public app_enabled_WalkTest_Builder As Boolean
    Public app_enabled_ProvTest_Mannitol As Boolean
    Public app_enabled_ProvTest_Methacholine_Yan As Boolean
    Public app_enabled_ProvTest_Histamine As Boolean
    Public app_enabled_ProvTest_Custom As Boolean
    Public app_enabled_ProvTest_Builder As Boolean
    Public app_enabled_CpxTest As Boolean
    Public app_enabled_HastTest As Boolean
    Public app_enabled_TrendReport As Boolean

    Public app_enabled_PredictedsManager As Boolean
    Public app_enabled_DbReports_Activity As Boolean

    Public app_enabled_Reporting_PdfBatchSave As Boolean

#End Region

#Region "User interface options"
    'Trend
    Public trend_include_ParametersWithNoResults As Boolean = True
    Public trend_include_Spirometry As Boolean = True
    Public trend_include_Tlco As Boolean = True
    Public trend_include_LVs As Boolean = True
    Public trend_include_Prov As Boolean = True
    Public trend_include_Mrp As Boolean = True
    Public trend_include_Feno As Boolean = True

#End Region




#Region "List preference methods"
    Public Const sDefault As String = "  <DEFAULT>"

    Public Function Save_Field(ByVal PK_Field As Integer, ByVal txt As String) As Boolean

        Dim sql As New StringBuilder
        sql.Clear()
        Select Case PK_Field
            Case 0  'new entry
                sql.Append("INSERT INTO prefs_fields (fieldname) VALUES('" & txt & "')")
                If cDAL.Insert_Record(sql.ToString) > 0 Then Return True Else Return False
            Case Else
                sql.Append("UPDATE prefs_fields SET fieldname='" & txt & "' WHERE field_id=" & PK_Field)
                If cDAL.Update_Record(sql.ToString) Then Return True Else Return False
        End Select

    End Function

    Public Function Delete_Field(ByVal PK_Field As Integer) As Boolean

        If cDAL.Delete_Record(PK_Field, eTables.Prefs_fields) Then Return True Else Return False

    End Function

    Public Function Save_FieldOption(ByVal PK_Field As Integer, ByVal PK_FieldOption As Integer, ByVal txt As String) As Integer
        'Choose between new and edit record with passed 'PK_FieldOption'. If zero then new

        Dim sql As New StringBuilder
        txt = Replace(txt, "'", "''")   'sql needs ' doubled in queries

        sql.Clear()
        Select Case PK_FieldOption
            Case 0  'new entry
                sql.Append("INSERT INTO prefs_fielditems (field_id, fielditem) VALUES(" & PK_Field & ", '" & txt & "')")
                Dim ID As Integer = cDAL.Insert_Record(sql.ToString)
                If ID > 0 Then Return ID Else Return 0
            Case Else
                sql.Append("UPDATE prefs_fielditems SET fielditem='" & txt & "' WHERE prefs_id=" & PK_FieldOption)
                If cDAL.Update_Record(sql.ToString) Then Return PK_FieldOption Else Return 0
        End Select

    End Function

    Public Function Delete_FieldOption(ByVal PK_FieldOption As Integer) As Boolean

        If cDAL.Delete_Record(PK_FieldOption, eTables.Prefs_fielditems) Then Return True Else Return False

    End Function

    Public Function Get_FieldList() As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As String = ""

        sql = "SELECT * FROM prefs_fields"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            For Each r As DataRow In Ds.Tables(0).Rows
                d.Add(r.Item("field_id"), r.Item("fieldname"))
            Next
        End If

        Return d

    End Function

    Public Function get_fieldID_fromFieldName(name As String) As String

        Dim sql As String = "SELECT field_id FROM prefs_fields WHERE fieldname='" & name & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count = 1 Then
            Return Ds.Tables(0).Rows(0)("field_id")
        Else
            Return Nothing
        End If

    End Function

    Public Function Get_FieldItemsForFieldID(ByVal fieldID As String, Optional ByVal AddDefault As Boolean = True) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim DefaultOptionID As Integer = Me.Get_DefaultFieldOptionID(fieldID)
        Dim sql As String = "SELECT prefs_fielditems.prefs_id, prefs_fielditems.fielditem FROM prefs_fields INNER JOIN prefs_fielditems ON prefs_fields.field_id = prefs_fielditems.field_id "
        sql = sql & "WHERE  (prefs_fields.field_id = N'" & fieldID & "') "
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If AddDefault And DefaultOptionID > 0 Then
            For Each r As DataRow In Ds.Tables(0).Rows
                If r.Item("prefs_ID") = DefaultOptionID Then
                    d.Add(r.Item("prefs_id"), r.Item("fielditem") & sDefault)
                Else
                    d.Add(r.Item("prefs_id"), r.Item("fielditem"))
                End If
            Next
        Else
            For Each r As DataRow In Ds.Tables(0).Rows
                d.Add(r.Item("prefs_id"), r.Item("fielditem"))
            Next
        End If
        Ds = Nothing

        Return d

    End Function

    Public Function Get_FieldOptionTextForID(ByRef FieldOptionID As Integer) As String

        Dim sql As String = "SELECT fielditem FROM prefs_fielditems WHERE prefs_id=" & FieldOptionID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Ds.Tables(0).Rows.Count > 0 Then
            Return Ds.Tables(0).Rows(0)("fielditem")
        Else
            Return ""
        End If

    End Function

    Public Function Get_DefaultFieldOptionID(ByRef fieldID As String) As Integer

        Dim DefaultID As Integer = 0
        Dim sql As String = "SELECT  prefs_fields.default_fielditem_id FROM prefs_fields "
        sql = sql & "WHERE  (prefs_fields.field_id = '" & fieldID & "')"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Ds.Tables(0).Rows.Count > 0 Then
            Select Case IsDBNull(Ds.Tables(0).Rows(0)("default_fielditem_id"))
                Case True : DefaultID = 0
                Case False : DefaultID = Ds.Tables(0).Rows(0)("default_fielditem_id")
            End Select

            Return DefaultID
        Else
            Return 0
        End If

    End Function

    Public Function Set_DefaultFieldOptionID(ByVal FieldID As Integer, ByVal FieldOptionID As Integer) As Boolean
        'If the FieldOptionID is already the default then turn it off

        Dim ID As String = ""

        If FieldOptionID = cPrefs.Get_DefaultFieldOptionID(FieldID) Then
            ID = "NULL"
        Else
            ID = Str(FieldOptionID)
        End If

        Dim sql As String = "UPDATE prefs_fields SET default_fielditem_id =" & ID & " WHERE field_id=" & FieldID
        If cDAL.Update_Record(sql) Then Return True Else Return False

    End Function
#End Region

End Class
