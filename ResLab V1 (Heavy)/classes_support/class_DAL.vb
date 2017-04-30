

Imports System
Imports System.Data
Imports System.Text
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.IO
Imports System.Drawing.Imaging
Imports ResLab_V1_Heavy.cDatabaseInfo


Public Class class_DataAccessLayer

    Public Enum eDatabaseType
        MySQL
        MicrosoftAccess
        SQLServer
    End Enum

    Public Enum eParameterType
        Intger
        Chars
        VarChars
    End Enum

    Public Enum MarkRecordsForDeletion
        MarkAsDelete = 1
        MarkAsUndelete = 0
    End Enum

    Public Function Get_DBType() As eDatabaseType

        Dim dbtype As String = My.Settings.Item(My.Settings.Item("Selected_DB_type"))
        Dim z As eDatabaseType = [Enum].Parse(GetType(eDatabaseType), dbtype)
        Return z

    End Function

    Public Function Get_ConnString() As String

        Dim CnString As String = My.Settings.Item(My.Settings.Item("Selected_ConnectionString"))
        Return CnString

    End Function

    Public Function Execute_NonQuery(ByVal sql As String) As Boolean

        Dim Cn As DbConnection = Me.Create_Connection()
        Dim Cmd As DbCommand = Me.Create_Command(sql, Cn)
        Cmd.Connection = Cn
        Try
            Cn.Open()
            Cmd.ExecuteNonQuery()
            Cn.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Update_Record(ByVal SqlString As String) As Boolean

        'SQL requires \ char to be doubled because it is an escape character
        SqlString = Replace(SqlString, "\", "\\")

        Dim Cn As DbConnection = Me.Create_Connection()
        Dim Cmd As DbCommand = Me.Create_Command(SqlString, Cn)
        Cmd.Connection = Cn
        Try
            Cn.Open()
            Cmd.ExecuteNonQuery()
            Cn.Close()
            Return True
        Catch ex As Exception
            MsgBox("Error saving record in cDAL.Update_Record" & vbNewLine & Err.Description & vbNewLine & SqlString)
            Return False
        End Try

    End Function

    Public Function Insert_Record(ByVal SqlString As String) As Long

        Dim Cn As DbConnection = Me.Create_Connection()
        Dim NewIdentityValue As Int32 = 0
        Dim cmdInsert As DbCommand = Me.Create_Command(SqlString, Cn)

        cmdInsert.Connection = Cn
        Try
            Cn.Open()
            cmdInsert.ExecuteNonQuery()
            NewIdentityValue = Me.Get_NewIdentityValue(Cn)
            Cn.Close()
            Return NewIdentityValue
        Catch
            MsgBox("Error saving record in cDAL.Insert_Record" & vbNewLine & Err.Description & vbNewLine & "SQL statement=" & SqlString)
            Return 0
        End Try

    End Function

    Public Function Delete_Record(ByVal RecordID As Long, ByVal Tbl As eTables, Optional SelectByForeignKeyName As String = "", Optional ForeignKeyValue As Long = 0, Optional SuppressConfirmMsg As Boolean = False) As Boolean
        'If 'SelectByForeignKey' not passed then primary key field is used

        Dim Cn As DbConnection = Me.Create_Connection()
        Dim sqlString As String = Nothing

        sqlString = "DELETE FROM " & cDbInfo.table_name(Tbl) & " WHERE " & cDbInfo.primarykey(Tbl) & " = " & RecordID
        If SelectByForeignKeyName <> "" And ForeignKeyValue > 0 Then sqlString = sqlString & " AND " & SelectByForeignKeyName & " = " & ForeignKeyValue

        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : sqlString = sqlString
            Case eDatabaseType.SQLServer : sqlString = sqlString
            Case eDatabaseType.MySQL : sqlString = sqlString
        End Select

        Dim Cmd As DbCommand = Me.Create_Command(sqlString, Cn)
        Dim Da As DbDataAdapter = Me.Create_Adapter(Cmd)

        Try
            Cn.Open()
            Cmd.ExecuteNonQuery()
            Cn.Close()
            If Not SuppressConfirmMsg Then MsgBox("Record deleted.", vbOKOnly, "Delete database record")
            Return True
        Catch ex As Exception
            MsgBox("Error deleting recordin cDAL.Delete_Record" & vbNewLine & Err.Description)
            Return False
        End Try

    End Function

    Public Function Delete_Records_All(ByVal Tbl As eTables) As Boolean

        Dim Cn As DbConnection = Me.Create_Connection()
        Dim sqlString As String = Nothing

        sqlString = "DELETE FROM " & cDbInfo.table_name(Tbl)
        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : sqlString = sqlString
            Case eDatabaseType.SQLServer : sqlString = sqlString
            Case eDatabaseType.MySQL : sqlString = sqlString
        End Select

        Dim Cmd As DbCommand = Me.Create_Command(sqlString, Cn)
        Dim Da As DbDataAdapter = Me.Create_Adapter(Cmd)

        Try
            Cn.Open()
            Cmd.ExecuteNonQuery()
            Cn.Close()
            Return True
        Catch ex As Exception
            MsgBox("Error deleting all records in table in cDAL.Delete_Records_All " & Tbl & vbNewLine & Err.Description)
            Return False
        End Try

    End Function

    Public Function Get_DataAsDataset(ByVal SqlString As String) As DataSet

        Dim Cn As DbConnection = Me.Create_Connection()
        Dim Cmd As DbCommand = Me.Create_Command(SqlString, Cn)
        Dim Da As DbDataAdapter = Me.Create_Adapter(Cmd)

        If SqlString = "" Then
            Return Nothing
        End If

        Try
            Dim Ds As New DataSet("MyData")
            Da.Fill(Ds)
            Return Ds
        Catch
            MsgBox("Error in cDAL.Get_DataAsDataset" & vbNewLine & Err.Description & vbNewLine & "sql=" & SqlString)
            Return Nothing
        End Try

    End Function

    Public Function Build_InsertQuery(ByVal eTbl As eTables, ByVal D As Dictionary(Of String, String)) As String

        Dim sb As New StringBuilder
        Dim val As String = ""
        Dim kv As KeyValuePair(Of String, String) = Nothing
        Dim kv1 As KeyValuePair(Of String, String) = Nothing
        Dim PrimaryKeyName As String = cDbInfo.primarykey(eTbl).ToLower

        sb.Clear()
        sb.Append("INSERT INTO " & eTbl.ToString & " (")

        For Each kv In D
            If kv.Key.ToLower <> PrimaryKeyName Then 'Don't include primary key in the field list
                sb.Append(kv.Key & ", ")
            End If
        Next
        sb.Remove(sb.Length - 2, 2)

        sb.Append(") VALUES (")
        For Each kv1 In D
            If kv1.Key.ToLower <> PrimaryKeyName Then 'Don't include primary key value in the field list
                If kv1.Value = Nothing Or kv1.Value = "NULL" Then
                    sb.Append("NULL,")
                Else
                    'Change single apostrophes in text to '' eg 'O'Brien' to  'O''Brien'
                    If Left(kv1.Value, 1) = "'" And Right(kv1.Value, 1) = "'" Then
                        val = Replace(Mid(kv1.Value, 2, Len(kv1.Value) - 2), "'", "''")
                        val = "'" & val & "'"
                    Else
                        val = kv1.Value
                    End If
                    sb.Append(val & ",")
                End If
            End If
        Next
        sb.Remove(sb.Length - 1, 1)
        sb.Append(");")

        Return sb.ToString

    End Function

    Public Function Build_UpdateQuery(ByVal eTbl As eTables, ByVal D As Dictionary(Of String, String)) As String

        Dim sb As New StringBuilder
        'Dim R As New class_RftFields
        'Dim T As New class_ProvTestFields
        'Dim TD As New class_ProvTestDataFields
        'Dim E As New class_pred_EquationFields
        'Dim S As New class_pred_SourceFields
        'Dim P As New class_Pref_PredFields
        Dim val As String = ""
        'Dim p1 As Integer = 0
        'Dim p2 As Integer = 0
        Dim PrimaryKeyName As String = cDbInfo.primarykey(eTbl).ToLower

        sb.Clear()
        sb.Append("UPDATE " & eTbl.ToString & " SET ")
        For Each kv As KeyValuePair(Of String, String) In D
            If kv.Key.ToLower <> PrimaryKeyName Then 'Don't include primary key in the field list
                If kv.Value = Nothing Then
                    sb.Append(kv.Key & " = " & "NULL, ")
                Else
                    'Change single apostrophes in text to '' eg 'O'Brien' to  'O''Brien'
                    If Left(kv.Value, 1) = "'" And Right(kv.Value, 1) = "'" Then
                        val = Replace(Mid(kv.Value, 2, Len(kv.Value) - 2), "'", "''")
                        val = "'" & val & "'"
                    Else
                        val = kv.Value
                    End If
                    sb.Append(kv.Key & " = " & val & ", ")
                End If
            End If
        Next
        sb.Remove(sb.Length - 2, 2)
        sb.Append(" WHERE " & PrimaryKeyName & " = " & D(PrimaryKeyName) & ";")

        Return sb.ToString

    End Function

    Public Function Create_Connection() As DbConnection

        Dim Cn As DbConnection = Nothing

        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : Cn = New OleDbConnection(Get_ConnString)
            Case eDatabaseType.SQLServer : Cn = New SqlConnection(Get_ConnString)
                'Case eDatabaseType.MySQL : Cn = New MySql.Data.MySqlClient.MySqlConnection(Get_ConnString)
            Case eDatabaseType.MySQL : Cn = New OdbcConnection(Get_ConnString)
        End Select

        Return Cn

    End Function

    Public Function Create_Parameter() As DbParameter

        Dim Pr As DbParameter = Nothing

        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : Pr = Nothing
            Case eDatabaseType.SQLServer : Pr = New SqlParameter
            Case eDatabaseType.MySQL : Pr = New OdbcParameter
        End Select

        Return Pr

    End Function

    Public Function Create_Command(ByVal CommandText As String, ByVal Cn As DbConnection)

        Dim cmd As DbCommand = Nothing
        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : cmd = New OleDbCommand(CommandText, Cn)
            Case eDatabaseType.SQLServer : cmd = New SqlCommand(CommandText, Cn)
            Case eDatabaseType.MySQL : cmd = New OdbcCommand(CommandText, Cn)
        End Select

        Return cmd

    End Function

    Public Function Get_NewIdentityValue(ByVal Cn As DbConnection) As Int32

        Dim txt As String = ""

        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : txt = ""
            Case eDatabaseType.SQLServer : txt = "SELECT SCOPE_IDENTITY() AS SCOPE_IDENTITY;"
            Case eDatabaseType.MySQL : txt = "SELECT LAST_INSERT_ID() ;"
        End Select
        Dim Cmd As DbCommand = Me.Create_Command(txt, Cn)

        Try
            Dim z As Int32 = Convert.ToInt32(Cmd.ExecuteScalar())
            Return z
        Catch
            Return Nothing
        End Try

    End Function

    Public Function Create_Adapter(ByVal cmd As DbCommand) As DbDataAdapter

        Dim Da As DbDataAdapter = Nothing

        Select Case Get_DBType()
            Case eDatabaseType.MicrosoftAccess : Da = New OleDbDataAdapter(cmd)
            Case eDatabaseType.SQLServer : Da = New SqlDataAdapter(cmd)
            Case eDatabaseType.MySQL : Da = New OdbcDataAdapter(cmd)
        End Select

        Return Da

    End Function

    Public Function Create_DataReader(ByVal cmd As DbCommand) As DbDataReader

        Dim Dr As DbDataReader = Nothing

        'Select Case Get_DBType
        '    Case eDatabaseType.MicrosoftAccess : Dr = New OleDbDataReader
        '    Case eDatabaseType.SQLServer : Dr = New SqlDataReader
        '    Case eDatabaseType.MySql : Dr = New OdbcDataReader
        'End Select

        Return Dr

    End Function

    Public Function Get_image(ByVal PKID As Long, ByVal pic As PictureBox, ByVal tbl As eTables, ByVal fld As String) As Boolean

        Try
            Dim sql As String = "SELECT " & fld & " FROM " & cDbInfo.table_name(tbl) & " WHERE " & cDbInfo.primarykey(tbl) & "=" & PKID
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

            If Ds.Tables(0).Rows.Count = 1 And Ds.Tables(0).Rows(0)(fld) IsNot DBNull.Value Then
                Dim bImage As Byte() = CType(Ds.Tables(0).Rows(0)(fld), Byte())
                Dim ms As New IO.MemoryStream(bImage)
                pic.Image = Image.FromStream(ms)
                ms.Close()
                ms = Nothing
                Ds = Nothing
                bImage = Nothing
            End If
            Return True
        Catch
            MsgBox("Database error: can't load flow volume loop image", "RFT")
            Return False
        End Try

    End Function

    Public Function Get_imageAsFile(ByVal PKID As Long, ByVal tbl As eTables, ByVal fld As String, ByVal file As String) As Boolean

        Try
            Dim sql As String = "SELECT " & fld & " FROM " & cDbInfo.table_name(tbl) & " WHERE " & cDbInfo.primarykey(tbl) & "=" & PKID
            Dim Ds As DataSet = Me.Get_DataAsDataset(sql)

            If Ds.Tables(0).Rows.Count = 1 And Ds.Tables(0).Rows(0)(fld) IsNot DBNull.Value Then
                Dim bImage As Byte() = CType(Ds.Tables(0).Rows(0)(fld), Byte())
                Dim ms As New IO.MemoryStream(bImage)
                Image.FromStream(ms).Save(file)
                ms.Close()
                ms = Nothing
                Ds = Nothing
                bImage = Nothing
                Return True
            Else
                Return False
            End If

        Catch
            MsgBox("Database error: can't load flow volume loop image", "RFT")
            Return False
        End Try

    End Function

    Public Function Update_image(ByVal PKID As Long, ByVal pic As PictureBox, ByVal tbl As eTables, ByVal fld As String) As Boolean
        'Must be jpg


        Try
            'Transfer image data from picbox to stream
            Dim s As MemoryStream = New MemoryStream()
            Dim i As Image
            Dim sql As String = ""
            Dim Cn As DbConnection
            Dim cmd As DbCommand
            Dim Pr As DbParameter
            Dim CmdPlaceholder As String = ""

            If pic.Image Is Nothing Then
                sql = "UPDATE " & cDbInfo.table_name(tbl) & " SET " & fld & "= NULL WHERE " & cDbInfo.primarykey(tbl) & "=" & PKID
                Cn = Me.Create_Connection()
                cmd = Me.Create_Command(sql, Cn)
                cmd.CommandText = sql
            Else
                i = New Bitmap(pic.Image)
                i.Save(s, ImageFormat.Jpeg)
                Dim bytBLOBData(s.Length - 1) As Byte
                s.Position = 0
                s.Read(bytBLOBData, 0, s.Length)
                s.Close()
                s = Nothing

                Select Case Me.Get_DBType
                    Case eDatabaseType.MySQL : CmdPlaceholder = "?"
                    Case eDatabaseType.SQLServer : CmdPlaceholder = "@BLOBData"
                End Select
                sql = "UPDATE " & cDbInfo.table_name(tbl) & " SET " & fld & "= " & CmdPlaceholder & " WHERE " & cDbInfo.primarykey(tbl) & "=" & PKID

                Cn = Me.Create_Connection()
                cmd = Me.Create_Command(sql, Cn)
                cmd.CommandText = sql

                Pr = Me.Create_Parameter
                Pr.ParameterName = "@BLOBData"
                Pr.DbType = DbType.Binary
                Pr.Direction = ParameterDirection.Input
                Pr.Value = bytBLOBData
                cmd.Parameters.Add(Pr)
            End If

            'Save stream to DB
            Cn.Open()
            cmd.ExecuteNonQuery()
            Cn.Close()
            Return True
        Catch
            MsgBox("Database error: can't save flow volume loop image in 'class_DAL.update_image" & vbCrLf & Err.Description, "RFT")
            Return False
        End Try
        ' Else
        '  Return False
        '  End If

    End Function

    Public Function CanConnectToDb() As String

        Dim Cn As DbConnection = Me.Create_Connection
        Dim CanConnect As String = ""

        Try
            Cn.Open()
            CanConnect = "Success"
        Catch ex As Exception
            CanConnect = False
            CanConnect = ex.Message
        End Try

        Cn = Nothing
        Return CanConnect

    End Function

    Public Function is_ColumnNameInTable(name As String, eTbl As eTables) As Boolean

        Dim Sql As String = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = '" & name & "' AND TABLE_NAME ='" & cDbInfo.table_name(eTbl) & "'"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(Sql)
        If IsNothing(Ds) Then
            Return False
        ElseIf Ds.Tables(0).Rows.Count <> 1 Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Function check_duplicatesInColumn(name As String, eTbl As eTables) As Integer

        Dim Sql As String = "SELECT code FROM " & cDbInfo.table_name(eTbl) & " WHERE (code IS NOT NULL AND code <> '') GROUP BY code HAVING(COUNT(code) > 1)"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(Sql)
        If IsNothing(Ds) Then
            Return 0
        Else
            Return Ds.Tables(0).Rows.Count
        End If

    End Function

End Class


