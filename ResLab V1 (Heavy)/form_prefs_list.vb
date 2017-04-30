Imports System.Text
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_prefs_list

    Public Const sDefault As String = "  <DEFAULT>"

    Public Structure RepHead
        Dim ID As Integer
        Dim Fieldname As String
        Dim Contents As String
    End Structure

    Dim toolTip As ToolTip = New ToolTip()

    Private Sub lstFieldOptions_MouseMove(sender As Object, e As MouseEventArgs) Handles lstFieldOptions.MouseMove
        Dim index As Integer = lstFieldOptions.IndexFromPoint(e.Location)
        If (index <> -1 AndAlso index < lstFieldOptions.Items.Count) Then
            If (toolTip.GetToolTip(lstFieldOptions) <> lstFieldOptions.Items(index).key()) Then
                toolTip.SetToolTip(lstFieldOptions, lstFieldOptions.Items(index).key())
            End If
        End If
    End Sub

    Private Sub form_prefs_list_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Load pref list of fields
        RemoveHandler lstFields.SelectedValueChanged, AddressOf lstFields_SelectedValueChanged
        cMyRoutines.Listbox_LoadItems(lstFields, cPrefs.Get_FieldList())
        AddHandler lstFields.SelectedValueChanged, AddressOf lstFields_SelectedValueChanged

        'Set up the report header grid
        Dim style_cell As DataGridViewCellStyle = New DataGridViewCellStyle()
        style_cell.Alignment = DataGridViewContentAlignment.MiddleLeft
        style_cell.ForeColor = Color.Black
        style_cell.BackColor = SystemColors.ButtonFace
        grdReportHeader.Columns(1).DefaultCellStyle = style_cell
        grdReportHeader.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        grdReportHeader.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable

        Dim style_header As DataGridViewCellStyle = New DataGridViewCellStyle()
        Dim HeaderFont = New Font(grdReportHeader.Font.Name, grdReportHeader.Font.Size, FontStyle.Bold)
        style_header.Font = HeaderFont
        grdReportHeader.EnableHeadersVisualStyles = False
        style_header.ForeColor = Color.Blue
        style_header.BackColor = SystemColors.GradientInactiveCaption
        grdReportHeader.ColumnHeadersDefaultCellStyle = style_header

        'Get and load report header info
        Me.Load_ReportHeaderStrings(Me.grdReportHeader, Me.Get_ReportHeaderStrings)

        cUser.set_access(Me)

    End Sub

    Private Sub lstFields_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Me.Load_FieldOptions()

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
            MsgBox("An error occurred saving report header items - some data may be lost.")
        End If

    End Sub

    Private Sub Load_ReportHeaderStrings(ByRef grd As DataGridView, info() As RepHead)

        Dim i As Integer = 0
        grd.Rows.Clear()

        For i = 0 To UBound(info)
            grd.Rows.Add(info(i).ID, info(i).Fieldname, info(i).Contents)
        Next

    End Sub

    Private Sub Load_FieldOptions()

        Dim kv As KeyValuePair(Of String, String) = CType(lstFields.SelectedItem, KeyValuePair(Of String, String))
        Dim FieldOptions As Dictionary(Of String, String) = cPrefs.Get_FieldItemsForFieldID(kv.Key)

        lstFieldOptions.Items.Clear()
        lstFieldOptions.ValueMember = "key"
        lstFieldOptions.DisplayMember = "value"

        cMyRoutines.Listbox_LoadItems(lstFieldOptions, FieldOptions)

    End Sub


    Private Sub tsbFieldOptions_new_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbFieldOptions_new.Click

        Dim item As KeyValuePair(Of String, String)
        Dim FoundDuplicates As Boolean = False
        Dim txt = InputBox("Enter a new field option", "Preferences")
        If txt <> "" Then
            'Check for duplicate - not allowed
            For Each item In lstFieldOptions.Items
                If Trim(txt.ToLower) = Trim(Replace(item.Value, "<DEFAULT>", "")).ToLower Then FoundDuplicates = True
            Next

            If FoundDuplicates Then
                MessageBox.Show("Duplicate field options not allowed", "Preferences", MessageBoxButtons.OK)
            Else
                Dim kv As New KeyValuePair(Of String, String)(0, txt)
                Dim kvField As KeyValuePair(Of String, String) = lstFields.SelectedItem
                If cPrefs.Save_FieldOption(kvField.Key, kv.Key, kv.Value) Then
                    lstFieldOptions.Items.Add(kv)
                End If
            End If
        End If

    End Sub

    Private Sub tsbFieldOptions_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbFieldOptions_delete.Click

        If lstFieldOptions.SelectedItems.Count > 0 Then
            If MessageBox.Show("Delete option?", "Preferences", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Dim kv As KeyValuePair(Of String, String) = lstFieldOptions.SelectedItem
                If cPrefs.Delete_FieldOption(kv.Key) Then
                    lstFieldOptions.Items.Remove(lstFieldOptions.SelectedItem)
                End If
            End If
        End If

    End Sub


    Private Sub tsbFieldOptions_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbFieldOptions_edit.Click

        If lstFieldOptions.SelectedItems.Count = 1 Then
            Dim kv As KeyValuePair(Of String, String) = lstFieldOptions.SelectedItem
            Dim SelectedIndex = lstFieldOptions.SelectedIndex
            Dim txt = InputBox("Edit field option", "Preferences", Strings.Replace(kv.Value, sDefault, ""))
            If txt <> "" Then
                Dim MatchedIndex As Integer = lstFieldOptions.FindStringExact(txt)
                If ListBox.NoMatches Or MatchedIndex = SelectedIndex Then

                    Dim kvField As KeyValuePair(Of String, String) = lstFields.SelectedItem
                    If cPrefs.Save_FieldOption(kvField.Key, kv.Key, txt) Then
                        Load_FieldOptions()   'Reload
                    End If

                Else
                    MessageBox.Show("Duplicate field options not allowed", "Preferences", MessageBoxButtons.OK)
                End If
            Else
                'nothing entered - edit cancelled
            End If
        Else
            MessageBox.Show("Select a field option to edit", "Preferences", MessageBoxButtons.OK)
        End If

    End Sub

    Private Sub tsbField_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbField_edit.Click

        If lstFields.SelectedItems.Count = 1 Then
            Dim kv As KeyValuePair(Of String, Integer) = lstFields.SelectedItem
            Dim txt = InputBox("Enter a new field option", "Preferences", kv.Key)
            If txt <> "" Then
                If lstFields.FindStringExact(txt) = ListBox.NoMatches Then
                    Dim kvNew As New KeyValuePair(Of String, Integer)(txt, kv.Value)
                    If cPrefs.Save_Field(kvNew.Value, kvNew.Key) Then
                        lstFields.Items.Remove(lstFields.SelectedItem)
                        lstFields.Items.Add(kvNew)
                    End If
                Else
                    MessageBox.Show("Duplicate field names not allowed", "Preferences", MessageBoxButtons.OK)
                End If
            End If
        Else
            MessageBox.Show("Select a field to edit", "Preferences", MessageBoxButtons.OK)
        End If

    End Sub


    Private Sub tsbField_new_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbField_new.Click

        Dim txt = InputBox("Enter a new field", "Preferences")
        If txt <> "" Then
            If lstFields.FindStringExact(txt) = ListBox.NoMatches Then
                Dim kv As New KeyValuePair(Of String, Integer)(txt, 0)
                If cPrefs.Save_Field(kv.Value, kv.Key) Then
                    lstFields.Items.Add(kv)
                    're-load list to get the new ID
                    RemoveHandler lstFields.SelectedValueChanged, AddressOf lstFields_SelectedValueChanged
                    lstFields.Items.Clear()
                    cMyRoutines.Listbox_LoadItems(lstFields, cPrefs.Get_FieldList())
                    AddHandler lstFields.SelectedValueChanged, AddressOf lstFields_SelectedValueChanged
                End If
            Else
                MessageBox.Show("Duplicate field names not allowed", "Preferences", MessageBoxButtons.OK)
            End If
        End If

    End Sub

    Private Sub tsbField_delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbField_delete.Click

        If lstFields.SelectedItems.Count = 1 Then
            If lstFieldOptions.Items.Count = 0 Then
                Dim kv As KeyValuePair(Of String, Integer) = lstFields.SelectedItem
                If cPrefs.Delete_Field(kv.Value) Then
                    lstFields.Items.Remove(lstFields.SelectedItem)
                End If
            Else
                MessageBox.Show("Delete all associated field options before deleting this field", "Preferences", MessageBoxButtons.OK)
            End If
        Else
            MessageBox.Show("Select a field to delete", "Preferences", MessageBoxButtons.OK)
        End If

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    Private Sub tsbFieldOptions_makedefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbFieldOptions_makedefault.Click

        Dim kv As KeyValuePair(Of String, String) = lstFields.SelectedItem
        Dim kv1 As KeyValuePair(Of String, String) = lstFieldOptions.SelectedItem
        Dim FieldID As String = kv.Key
        Dim FieldOptionID As String = kv1.Key

        'Set the new default
        cPrefs.Set_DefaultFieldOptionID(FieldID, FieldOptionID)

        'Reload the list
        Load_FieldOptions()

    End Sub

    Private Sub grdReportHeader_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdReportHeader.CellEndEdit

        Me.Save_ReportHeaderStrings(grdReportHeader)

    End Sub

End Class