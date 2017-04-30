Imports System.Text
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_spt

    Public Function Get_AllergensForPanel(enabled_only As Boolean, panelID As Integer) As List(Of PanelMember_AllData)

        Try

            Dim sql As New StringBuilder
            sql.Append("SELECT spt_allergens.*, spt_allergencategories.*, spt_panelmembers.* ")
            sql.Append("FROM ((spt_panels INNER JOIN spt_panelmembers ON spt_panels.panelID = spt_panelmembers.panelID) INNER JOIN spt_allergens ")
            sql.Append("ON spt_panelmembers.allergenID = spt_allergens.allergenID) INNER JOIN spt_allergencategories ON spt_allergens.categoryID = spt_allergencategories.categoryID WHERE ")

            If enabled_only Then sql.Append(" spt_allergens.enabled = 1 AND ")
            If panelID > 0 Then sql.Append(" spt_panelmembers.panelID = " & panelID & " AND ")
            Select Case Strings.Right(sql.ToString, 4)
                Case "ERE " : sql.Remove(sql.Length - 6, 6)
                Case "AND " : sql.Remove(sql.Length - 4, 4)
            End Select
            sql.Append(" ORDER BY allergenorder ASC")

            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
            If IsNothing(Ds) Then
                Return Nothing
            Else
                Dim a As New List(Of PanelMember_AllData)
                For Each row As DataRow In Ds.Tables(0).Rows
                    Dim b = New PanelMember_AllData
                    b.panelid = row("panelid")
                    b.allergenid = row("allergenID")
                    b.allergenname = row("allergenName")
                    If IsDBNull(row("allergenorder")) Then b.allergenorder = 0 Else b.allergenorder = row("allergenorder")
                    b.displayColour = row("displayColour")
                    b.memberid = row("memberid")
                    b.allergengroup = row("categoryname")
                    b.allergengroupID = row("categoryID")
                    a.Add(b)
                Next
                If a.Count = 0 Then Return Nothing Else Return a
            End If
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function Get_AllergensFromDatabase(enabled_only As Boolean, categoryID As Integer) As List(Of AllergenData)

        Dim sql As New StringBuilder
        sql.Append("SELECT spt_allergens.*, spt_allergencategories.categoryname ")
        sql.Append("FROM spt_allergens INNER JOIN spt_allergencategories ON spt_allergens.categoryID = spt_allergencategories.categoryID WHERE ")

        If enabled_only Then sql.Append(" spt_allergens.enabled = 1 AND ")
        If categoryID > 0 Then sql.Append(" spt_allergens.categoryID = " & categoryID & " AND ")
        Select Case Strings.Right(sql.ToString, 4)
            Case "ERE " : sql.Remove(sql.Length - 6, 6)
            Case "AND " : sql.Remove(sql.Length - 4, 4)
        End Select
        sql.Append(" ORDER BY allergenName")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            Dim a As New List(Of AllergenData)
            For Each row As DataRow In Ds.Tables(0).Rows
                Dim b = New AllergenData
                b.allergenid = row("allergenID")
                b.allergenname = row("allergenName")
                b.categoryid = row("categoryID")
                b.categoryname = row("categoryName")
                b.enabled = row("enabled")
                a.Add(b)
            Next
            If a.Count = 0 Then Return Nothing Else Return a
        End If

    End Function

    Public Function Get_AllergenCategories(enabled_only As Boolean) As List(Of AllergenCategoryData)

        Dim sql As String = "SELECT * FROM spt_AllergenCategories "
        If enabled_only Then sql = sql & "WHERE enabled = 1 "
        sql = sql & " ORDER BY categoryName"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            Dim a As New List(Of AllergenCategoryData)
            For Each row As DataRow In Ds.Tables(0).Rows
                Dim b = New AllergenCategoryData
                b.categoryid = row("categoryID")
                b.categoryname = row("categoryName")
                b.displaycolour = row("displayColour")
                b.enabled = row("enabled")
                a.Add(b)
            Next
            If a.Count = 0 Then Return Nothing Else Return a
        End If

    End Function

    Public Function Get_Category(categoryID As Integer) As AllergenCategoryData

        Dim sql As String = "SELECT * FROM spt_allergencategories WHERE categoryID = " & categoryID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            Dim c As New AllergenCategoryData
            c.categoryid = categoryID
            c.categoryname = Ds.Tables(0).Rows(0).Item("categoryname")
            c.displaycolour = Ds.Tables(0).Rows(0).Item("displaycolour")
            c.enabled = Ds.Tables(0).Rows(0).Item("enabled")
            Return c
            Ds = Nothing
        End If

    End Function

    Public Function Get_CategoryForAllergen(allergenID As Integer) As AllergenCategoryData

        Dim sql As String = "SELECT spt_allergencategories.* FROM spt_allergens INNER JOIN spt_allergencategories "
        sql = sql & "ON spt_allergens.categoryID = spt_allergencategories.categoryID WHERE spt_allergens.allergenID = " & allergenID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            If Ds.Tables(0).Rows.Count = 1 Then
                Dim c As New AllergenCategoryData
                c.categoryid = Ds.Tables(0).Rows(0).Item("categoryID")
                c.categoryname = Ds.Tables(0).Rows(0).Item("categoryname")
                c.displaycolour = Ds.Tables(0).Rows(0).Item("displaycolour")
                c.enabled = Ds.Tables(0).Rows(0).Item("enabled")
                Return c
                Ds = Nothing
            Else
                Return Nothing
            End If
        End If

    End Function

    Public Function Update_Category(categorydata As Dictionary(Of String, String)) As Integer

        Dim sqlString As String = ""
        Try
            sqlString = cDAL.Build_UpdateQuery(eTables.spt_allergencategories, categorydata)
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving category" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sqlString)
            Return 0
        End Try

    End Function

    Public Function Get_AllergenPanels(enabled_only As Boolean) As List(Of PanelData)

        Dim sql As String = "SELECT * FROM spt_panels  "
        If enabled_only Then sql = sql & "WHERE enabled = 1"
        sql = sql & " ORDER BY panelName"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            Dim a As New List(Of PanelData)
            For Each row As DataRow In Ds.Tables(0).Rows
                Dim b = New PanelData
                b.panelid = row("panelID")
                b.panelname = row("panelName")
                b.enabled = row("enabled")
                a.Add(b)
            Next
            If a.Count = 0 Then Return Nothing Else Return a
        End If

    End Function

    Public Function Get_PanelnameFromID(panelid As Long) As String

        Dim sql As String = "SELECT panelname FROM spt_panels WHERE panelid= " & panelid
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)("panelname")
            Else
                Return ""
            End If
        End If

    End Function

    Public Function Get_Panel(panelID As Integer) As PanelData

        Dim p As New PanelData


        Return p
    End Function
    Public Function Get_Panel(panelname As String) As PanelData

        Dim p As New PanelData


        Return p
    End Function

    Public Function Insert_Panel(paneldata As Dictionary(Of String, String)) As Integer

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.spt_panels, paneldata)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new panel" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sql)
            Return 0
        End Try

    End Function

    Public Function Delete_Panel(panelID As Integer) As Boolean

        If panelID > 0 Then
            'First the related panel members records
            cDAL.Delete_Record(panelID, eTables.spt_panelmembers, "panelID", True)
            'Now the panel record
            cDAL.Delete_Record(panelID, eTables.spt_panels)
            Return True
        Else
            Return False
        End If

    End Function

    Public Function Update_Panel(paneldata As Dictionary(Of String, String)) As Integer

        Dim sqlString As String = ""
        Try
            sqlString = cDAL.Build_UpdateQuery(eTables.spt_panels, paneldata)
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving panel" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sqlString)
            Return 0
        End Try

    End Function

    Public Function Is_PanelNameUnique(panelname As String) As Boolean

        Dim sql As String = "SELECT count(panelID) FROM spt_panels WHERE panelName='" & panelname & "'"
        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return False
        Else
            If ds.Tables(0).Rows(0).Item(0) = 0 Then Return True Else Return False
        End If

    End Function

    Public Function Is_AllergenNameUnique(allergenname As String, Optional IgnoreThisAllergenID As Integer = 0) As Boolean

        Dim sql As String = "SELECT count(allergenID) FROM spt_allergens WHERE allergenname='" & allergenname & "' "
        If IgnoreThisAllergenID > 0 Then sql = sql & " AND allergenID <> " & IgnoreThisAllergenID
        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return False
        Else
            If ds.Tables(0).Rows(0).Item(0) = 0 Then Return True Else Return False
        End If

    End Function

    Public Function Is_CategoryNameUnique(categoryname As String) As Boolean

        Dim sql As String = "SELECT count(categoryID) FROM spt_allergencategories WHERE categoryname='" & categoryname & "'"
        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return False
        Else
            If ds.Tables(0).Rows(0).Item(0) = 0 Then Return True Else Return False
        End If

    End Function

    Public Function Is_CategoryEmpty(categoryID As Integer) As Boolean

        Dim sql As String = "SELECT COUNT(spt_allergens.allergenID) FROM spt_allergens INNER JOIN spt_allergencategories "
        sql = sql & "ON spt_allergens.categoryID = spt_allergencategories.categoryID WHERE spt_allergencategories.categoryID = " & categoryID
        Dim ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If IsNothing(ds) Then
            Return False
        Else
            If ds.Tables(0).Rows(0).Item(0) = 0 Then Return True Else Return False
        End If

    End Function

    Public Function Delete_Allergen(allergenID As Integer) As Boolean

        If allergenID > 0 Then
            cDAL.Delete_Record(allergenID, eTables.spt_allergens)
            Return True
        Else
            Return False
        End If

    End Function

    Public Function Insert_Allergen(allergenData As Dictionary(Of String, String)) As Integer

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.spt_allergens, allergenData)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new allergen" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sql)
            Return 0
        End Try

    End Function

    Public Function Update_Allergen(allergendata As Dictionary(Of String, String)) As Integer

        Dim sqlString As String = ""
        Try
            sqlString = cDAL.Build_UpdateQuery(eTables.spt_allergens, allergendata)
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving allergen" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sqlString)
            Return 0
        End Try

    End Function


    Public Function Insert_Category(categorydata As Dictionary(Of String, String)) As Integer

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.spt_allergencategories, categorydata)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new category" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sql)
            Return 0
        End Try

    End Function

    Public Function Delete_Category(categoryID As Integer) As Boolean

        If categoryID > 0 Then
            'Can't delete if there are related allergen records
            If Me.Is_CategoryEmpty(categoryID) Then
                If cDAL.Delete_Record(categoryID, eTables.spt_allergencategories) Then Return True Else Return False
            Else
                MsgBox("Can't delete category while it contains allergens.", vbOKOnly, "Delete category")
                Return False
            End If
        Else
            Return False
        End If

    End Function

    Public Function Insert_PanelMember(panelMemberData As Dictionary(Of String, String)) As Integer

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.spt_panelmembers, panelMemberData)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error adding new panel member" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sql)
            Return 0
        End Try

    End Function

    Public Function Delete_PanelMember(memberID As Integer) As Boolean

        If memberID > 0 Then
            cDAL.Delete_Record(memberID, eTables.spt_panelmembers, , True)
            Return True
        Else
            MsgBox("memberid=0")
            Return False
        End If

    End Function

    Public Function Update_PanelMember(panelMemberData As Dictionary(Of String, String)) As Integer

        Dim sqlString As String = ""
        Try
            sqlString = cDAL.Build_UpdateQuery(eTables.spt_panelmembers, panelMemberData)
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving panel member" & vbNewLine & ex.Message.ToString & vbNewLine & "SQL statement: " & sqlString)
            Return 0
        End Try

    End Function

End Class
