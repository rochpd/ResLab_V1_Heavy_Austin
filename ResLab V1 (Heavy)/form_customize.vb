Imports ResLab_V1_Heavy.cDatabaseInfo


Public Class form_customize

    Public Structure RepHead
        Dim ID As Integer
        Dim Fieldname As String
        Dim Contents As String
    End Structure

    Private Sub form_customize_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Dim style_cell As DataGridViewCellStyle = New DataGridViewCellStyle()
        style_cell.Alignment = DataGridViewContentAlignment.MiddleLeft
        style_cell.ForeColor = Color.Black
        style_cell.BackColor = Color.Ivory
        grdReportHeader.Columns(1).DefaultCellStyle = style_cell
        grdReportHeader.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        grdReportHeader.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable

        Dim style_header As DataGridViewCellStyle = New DataGridViewCellStyle()
        Dim HeaderFont = New Font(grdReportHeader.Font.Name, grdReportHeader.Font.Size, FontStyle.Bold)
        style_header.Font = HeaderFont
        grdReportHeader.EnableHeadersVisualStyles = False
        style_header.ForeColor = Color.Black
        style_header.BackColor = Color.Ivory
        grdReportHeader.ColumnHeadersDefaultCellStyle = style_header

        'Get and load report header info
        Me.Load_ReportHeaderStrings(Me.grdReportHeader, Me.Get_ReportHeaderStrings)



    End Sub

    Private Function Get_ReportHeaderStrings() As RepHead()

        Dim head() As RepHead = Nothing
        Dim sql As String = "SELECT * FROM prefs_reports_strings WHERE stringid > 1"    'Obsolete field - leave for compatibility
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        Dim i As Integer = 0

        If IsNothing(Ds) Then
            Return Nothing
        Else
            For Each record As DataRow In Ds.Tables(0).Rows
                ReDim Preserve head(i)
                head(i).ID = record.Item("stringid").ToString
                head(i).Fieldname = record.Item("name").ToString
                head(i).Contents = record.Item("text").ToString
                i = i + 1
            Next
            Return head
        End If

    End Function

    Private Sub Load_ReportHeaderStrings(ByRef grd As DataGridView, info() As RepHead)

        Dim i As Integer = 0
        grd.Rows.Clear()

        For i = 0 To UBound(info)
            grd.Rows.Add(info(i).ID, info(i).Fieldname, info(i).Contents)
        Next

    End Sub

    Private Sub tsBtn_Cancel_Click(sender As System.Object, e As System.EventArgs) Handles tsBtn_Cancel.Click

        Me.Close()

    End Sub

    Private Sub Save_ReportHeaderStrings(ByRef grd As DataGridView)

        Dim i As Integer = 0
        Dim sql As String = ""
        Dim Success As Boolean = False
        Dim nFailed As Integer = 0

        'Load dic from grid
        Dim d() As Dictionary(Of String, String) = Nothing
        For i = 0 To grd.Rows.Count - 1
            ReDim Preserve d(i)
            d(i) = New Dictionary(Of String, String)
            d(i).Add("stringid", "'" & grd.Rows(i).Cells(0).Value & "'")
            d(i).Add("name", "'" & grd.Rows(i).Cells(1).Value & "'")
            d(i).Add("text", "'" & grd.Rows(i).Cells(2).Value & "'")
        Next

        'Save to database
        For Each dic In d
            sql = cDAL.Build_UpdateQuery(eTables.Prefs_reports_strings, dic)
            Success = cDAL.Update_Record(sql)
            If Not Success Then
                nFailed = nFailed + 1
            End If
        Next

        If nFailed > 0 Then
            MsgBox("An error occurred saving Customize data - some data may be lost.")
        End If

    End Sub

    Private Sub tsBtn_SaveAndClose_Click(sender As System.Object, e As System.EventArgs) Handles tsBtn_SaveAndClose.Click

        If form_password.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.Save_ReportHeaderStrings(Me.grdReportHeader)
            Me.Close()
        End If

    End Sub

    Private Sub tabpageReportHeader_Click(sender As System.Object, e As System.EventArgs) Handles tabpageReportHeader.Click

    End Sub
End Class