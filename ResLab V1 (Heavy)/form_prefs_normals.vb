Imports System.Text
Imports ResLab_V1_Heavy.form_PredSourceEditor
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_prefs_normals

    Public Class RowData
        Public Structure RowItem
            Dim Name As String
            Dim Value As String
            Dim visible As Boolean
        End Structure
        Public Count As New RowItem
        Public PrefID As New RowItem
        Public Test As New RowItem
        Public TestID As New RowItem
        Public Parameter As New RowItem
        Public ParameterID As New RowItem
        Public Source As New RowItem
        Public Source_PubReference As New RowItem
        Public Source_PubYear As New RowItem
        Public SourceID As New RowItem
        Public Gender As New RowItem
        Public GenderID As New RowItem
        Public AgeGroup As New RowItem
        Public AgeGroupID As New RowItem
        Public Age_lower As New RowItem
        Public Age_upper As New RowItem
        Public Age_clipmethod As New RowItem
        Public Age_clipmethodID As New RowItem
        Public Ht_lower As New RowItem
        Public Ht_upper As New RowItem
        Public Ht_clipmethod As New RowItem
        Public Ht_clipmethodID As New RowItem
        Public Wt_lower As New RowItem
        Public Wt_upper As New RowItem
        Public Wt_clipmethod As New RowItem
        Public Wt_clipmethodID As New RowItem
        Public Ethnicity As New RowItem
        Public EthnicityID As New RowItem
        Public Ethnicity_ApplyATS1991Correction As New RowItem
        Public StartDate As New RowItem
        Public EndDate As New RowItem
        Public Equation_Class As New RowItem
        Public Equation_ClassID As New RowItem
        'following x2  used by pred preferences
        Public Equationid_normalrange As New RowItem
        Public Equationid_mpv As New RowItem
        'following x4  used by pred sources editor
        Public Equation As New RowItem
        Public Equationid As New RowItem
        Public StatType As New RowItem
        Public StatTypeid As New RowItem

        Public Sub New()
            Me.Count.Name = "#"
            Me.PrefID.Name = "prefid"
            Me.Test.Name = "test"
            Me.TestID.Name = "testid"
            Me.Parameter.Name = "parameter"
            Me.ParameterID.Name = "parameterid"
            Me.Source.Name = "source"
            Me.SourceID.Name = "sourceid"
            Me.Source_PubReference.Name = "pub_reference"
            Me.Source_PubYear.Name = "pub_year"
            Me.Gender.Name = "gender"
            Me.GenderID.Name = "genderid"
            Me.AgeGroup.Name = "agegroup"
            Me.AgeGroupID.Name = "agegroupid"
            Me.Age_lower.Name = "age_lower"
            Me.Age_upper.Name = "age_upper"
            Me.Age_clipmethod.Name = "age_clipmethod"
            Me.Age_clipmethodID.Name = "age_clipmethodid"
            Me.Ht_lower.Name = "ht_lower"
            Me.Ht_upper.Name = "ht_upper"
            Me.Ht_clipmethod.Name = "ht_clipmethod"
            Me.Ht_clipmethodID.Name = "ht_clipmethodid"
            Me.Wt_lower.Name = "wt_lower"
            Me.Wt_upper.Name = "wt_upper"
            Me.Wt_clipmethod.Name = "wt_clipmethod"
            Me.Wt_clipmethodID.Name = "wt_clipmethodid"
            Me.Ethnicity.Name = "ethnicity"
            Me.EthnicityID.Name = "ethnicityid"
            Me.Ethnicity_ApplyATS1991Correction.Name = "ethnicity correction type"
            Me.StartDate.Name = "startdate"
            Me.EndDate.Name = "enddate"
            Me.Equationid_normalrange.Name = "equationid_normalrange"
            Me.Equationid_mpv.Name = "equationid_mpv"
            Me.Equation_Class.Name = "equation_class"
            Me.Equation_ClassID.Name = "equation_classid"
            Me.Equation.Name = "equation"
            Me.Equationid.Name = "equationid"
            Me.StatType.Name = "stattype"
            Me.StatTypeid.Name = "stattypeid"

            Me.Count.visible = True
            Me.PrefID.visible = False
            Me.Test.visible = False
            Me.TestID.visible = False
            Me.Parameter.visible = True
            Me.ParameterID.visible = False
            Me.Source.visible = False
            Me.SourceID.visible = False
            Me.Source_PubReference.visible = False
            Me.Source_PubYear.visible = False
            Me.Gender.visible = True
            Me.GenderID.visible = False
            Me.AgeGroup.visible = True
            Me.AgeGroupID.visible = False
            Me.Age_lower.visible = True
            Me.Age_upper.visible = True
            Me.Age_clipmethod.visible = False
            Me.Age_clipmethodID.visible = False
            Me.Ht_lower.visible = False
            Me.Ht_upper.visible = False
            Me.Ht_clipmethod.visible = False
            Me.Ht_clipmethodID.visible = False
            Me.Wt_lower.visible = False
            Me.Wt_upper.visible = False
            Me.Wt_clipmethod.visible = False
            Me.Wt_clipmethodID.visible = False
            Me.Ethnicity.visible = True
            Me.EthnicityID.visible = False
            Me.Ethnicity_ApplyATS1991Correction.visible = True
            Me.StartDate.visible = True
            Me.EndDate.visible = True
            Me.Equationid_mpv.visible = False
            Me.Equationid_normalrange.visible = False
            Me.Equation_Class.visible = False
            Me.Equation_ClassID.visible = False
            Me.Equation.visible = False
            Me.Equationid.visible = False
            Me.StatType.visible = True
            Me.StatTypeid.visible = False
        End Sub
    End Class

    Dim flagFormLoading As Boolean = False
    Dim rd As New RowData

    Private Sub SetUp_Grid()

        Dim i As Integer = 0
        Dim Wid = 70
        grdP.ColumnCount = 30

        grdP.Columns(i).HeaderText = rd.Count.Name : grdP.Columns(i).Name = rd.Count.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 30 : i = i + 1
        grdP.Columns(i).HeaderText = rd.PrefID.Name : grdP.Columns(i).Name = rd.PrefID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Test.Name : grdP.Columns(i).Name = rd.Test.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.TestID.Name : grdP.Columns(i).Name = rd.TestID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Source.Name : grdP.Columns(i).Name = rd.Source.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.SourceID.Name : grdP.Columns(i).Name = rd.SourceID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Parameter.Name : grdP.Columns(i).Name = rd.Parameter.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.ParameterID.Name : grdP.Columns(i).Name = rd.ParameterID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Gender.Name : grdP.Columns(i).Name = rd.Gender.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.GenderID.Name : grdP.Columns(i).Name = rd.GenderID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Ethnicity.Name : grdP.Columns(i).Name = rd.Ethnicity.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.EthnicityID.Name : grdP.Columns(i).Name = rd.EthnicityID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.AgeGroup.Name : grdP.Columns(i).Name = rd.AgeGroup.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.AgeGroupID.Name : grdP.Columns(i).Name = rd.AgeGroupID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Age_lower.Name : grdP.Columns(i).Name = rd.Age_lower.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Age_upper.Name : grdP.Columns(i).Name = rd.Age_upper.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Age_clipmethod.Name : grdP.Columns(i).Name = rd.Age_clipmethod.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Age_clipmethodID.Name : grdP.Columns(i).Name = rd.Age_clipmethodID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Ht_lower.Name : grdP.Columns(i).Name = rd.Ht_lower.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Ht_upper.Name : grdP.Columns(i).Name = rd.Ht_upper.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Ht_clipmethod.Name : grdP.Columns(i).Name = rd.Ht_clipmethod.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Ht_clipmethodID.Name : grdP.Columns(i).Name = rd.Ht_clipmethodID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Wt_lower.Name : grdP.Columns(i).Name = rd.Wt_lower.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Wt_upper.Name : grdP.Columns(i).Name = rd.Wt_upper.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Wt_clipmethod.Name : grdP.Columns(i).Name = rd.Wt_clipmethod.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Wt_clipmethodID.Name : grdP.Columns(i).Name = rd.Wt_clipmethodID.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Equationid_mpv.Name : grdP.Columns(i).Name = rd.Equationid_mpv.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.Equationid_normalrange.Name : grdP.Columns(i).Name = rd.Equationid_normalrange.Name : grdP.Columns(i).Visible = False : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.StartDate.Name : grdP.Columns(i).Name = rd.StartDate.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1
        grdP.Columns(i).HeaderText = rd.EndDate.Name : grdP.Columns(i).Name = rd.EndDate.Name : grdP.Columns(i).Visible = True : grdP.Columns(i).Width = 70 : i = i + 1

        grdP.Columns(0).Width = 30

        grdP.AllowUserToResizeColumns = True

    End Sub

    Private Sub lstTests_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTests.SelectedValueChanged

        Dim Sources As Dictionary(Of String, String)
        Dim kv As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))
        Sources = cPred.Get_SourcesForTestID(kv.Key)

        lstSets.Items.Clear()
        lstParameters.Items.Clear()
        lstEthnicity.Items.Clear()
        lstAgeGroup.Items.Clear()
        lstGender.Items.Clear()
        cMyRoutines.Listbox_LoadItems(lstSets, Sources)

        If lstSets.Items.Count > 0 Then lstSets.SelectedIndex = 0

    End Sub

    Private Sub lstAgeGroup_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstAgeGroup.SelectedValueChanged


        Dim Ethnicities As Dictionary(Of String, String)
        Dim Genders As Dictionary(Of String, String)
        Dim kv As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        Dim kv1 As KeyValuePair(Of String, String) = CType(lstParameters.SelectedItem, KeyValuePair(Of String, String))
        Dim kv2 As KeyValuePair(Of String, String) = CType(lstAgeGroup.SelectedItem, KeyValuePair(Of String, String))
        Dim kv3 As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))

        Ethnicities = cPred.Get_EthnicitiesFor_Source_Parameter_Agegroup(kv.Key, kv1.Key, kv2.Key)
        lstEthnicity.Items.Clear()
        cMyRoutines.Listbox_LoadItems(lstEthnicity, Ethnicities)
        If lstEthnicity.Items.Count > 0 Then lstEthnicity.SelectedIndex = 0

        Genders = cPred.Get_GendersForTestID_SourceID(kv3.Key, kv.Key)
        lstGender.Items.Clear()
        cMyRoutines.Listbox_LoadItems(lstGender, Genders)
        If lstGender.Items.Count > 0 Then lstGender.SelectedIndex = 0

    End Sub

    Private Sub lstParameters_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstParameters.SelectedValueChanged

        Dim AgeGroups As Dictionary(Of String, String)
        Dim kv As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        Dim kv1 As KeyValuePair(Of String, String) = CType(lstParameters.SelectedItem, KeyValuePair(Of String, String))

        AgeGroups = cPred.Get_AgeGroupsForSourceParameter(kv.Key, kv1.Key)

        lstAgeGroup.Items.Clear()
        lstEthnicity.Items.Clear()

        RemoveHandler lstAgeGroup.SelectedValueChanged, AddressOf lstAgeGroup_SelectedValueChanged
        cMyRoutines.Listbox_LoadItems(lstAgeGroup, AgeGroups)
        AddHandler lstAgeGroup.SelectedValueChanged, AddressOf lstAgeGroup_SelectedValueChanged

        If lstAgeGroup.Items.Count > 0 Then lstAgeGroup.SelectedIndex = 0

    End Sub

    Private Sub lstSets_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSets.SelectedValueChanged

        Dim Parameters As Dictionary(Of String, String)
        Dim kv As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))

        Parameters = cPred.Get_ParametersForSourceID(kv.Key)

        lstParameters.Items.Clear()
        lstEthnicity.Items.Clear()
        lstAgeGroup.Items.Clear()
        lstGender.Items.Clear()

        RemoveHandler lstParameters.SelectedValueChanged, AddressOf lstParameters_SelectedValueChanged
        cMyRoutines.Listbox_LoadItems(lstParameters, Parameters)
        AddHandler lstParameters.SelectedValueChanged, AddressOf lstParameters_SelectedValueChanged

        If lstParameters.Items.Count > 0 Then lstParameters.SelectedIndex = 0

    End Sub

    Private Sub btnSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelect.Click

        Dim AlreadyExists As Boolean = False
        Dim c(0 To 29) As String

        'Check selection has been made for all lists
        If lstTests.SelectedIndex = -1 Then MsgBox("No can do." & vbCrLf & "Please select a test.") : Exit Sub
        If lstSets.SelectedIndex = -1 Then MsgBox("No can do." & vbCrLf & "Please select a source.") : Exit Sub
        If lstParameters.SelectedIndex = -1 Then MsgBox("No can do." & vbCrLf & "Please select a parameter.") : Exit Sub
        If lstAgeGroup.SelectedIndex = -1 Then MsgBox("No can do." & vbCrLf & "Please select an age group.") : Exit Sub
        If lstGender.SelectedIndex = -1 Then MsgBox("No can do." & vbCrLf & "Please select a gender.") : Exit Sub
        If lstEthnicity.SelectedIndex = -1 Then MsgBox("No can do." & vbCrLf & "Please select an ethnicity.") : Exit Sub

        'CheckBox if selection already exists in the grid and is open dated. Check against test, source, parameter, agegroup, ethnicity and end date.
        For i As Integer = 0 To grdP.RowCount - 1
            If grdP.Item(3, i).Value = lstTests.SelectedItem.key Then
                If grdP.Item(5, i).Value = lstSets.SelectedItem.key Then
                    If grdP.Item(7, i).Value = lstParameters.SelectedItem.key Then
                        If grdP.Item(9, i).Value = lstGender.SelectedItem.key Then
                            If grdP.Item(13, i).Value = lstAgeGroup.SelectedItem.key Then
                                If grdP.Item(11, i).Value = lstEthnicity.SelectedItem.key Then
                                    If grdP.Item(29, i).Value = Nothing Then
                                        AlreadyExists = True
                                        MsgBox("No can do." & vbCrLf & "Row " & i + 1 & " already exists as an open selection." & vbCrLf & "Add an end date an try again.")
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next

        'If OK then load to grid
        If Not AlreadyExists Then
            c(0) = grdP.RowCount + 1
            c(1) = ""   'prefID
            c(2) = lstTests.SelectedItem.key
            c(3) = lstTests.SelectedItem.value
            c(4) = lstSets.SelectedItem.key
            c(5) = lstSets.SelectedItem.value
            c(6) = lstParameters.SelectedItem.key
            c(7) = lstParameters.SelectedItem.value
            c(8) = lstGender.SelectedItem.key
            c(9) = lstGender.SelectedItem.value
            c(10) = lstEthnicity.SelectedItem.key
            c(11) = lstEthnicity.SelectedItem.value
            c(12) = lstAgeGroup.SelectedItem.key
            c(13) = lstAgeGroup.SelectedItem.value
            c(14) = "Age_lower"
            c(15) = "Age_upper"
            c(16) = "Age_clipmethod"
            c(17) = "Age_clipmethodID"
            c(18) = "Ht_lower"
            c(19) = "Ht_upper"
            c(20) = "Ht_clipmethod"
            c(21) = "Ht_clipmethodID"
            c(22) = "Wt_lower"
            c(23) = "Wt_upper"
            c(24) = "Wt_clipmethod"
            c(25) = "Wt_clipmethodID"
            c(26) = "Equationid_mpv"
            c(27) = "Equationid_normalrange"
            c(28) = Format(Now(), "dd/MM/yyyy")
            c(29) = ""

            Dim dic() As Dictionary(Of String, String)
            dic = cPred.Match_Equation(c(3), c(5), c(7), c(13), c(9), c(11))
            If IsNothing(dic) Then
                MsgBox("Error - can't find matching equation in database")
            Else
                'Returns equation info for each stattype
                c(14) = dic(0)("age_lower")
                c(15) = dic(0)("age_upper")
                c(16) = dic(0)("age_clipmethod")
                c(17) = dic(0)("age_clipmethodID")
                c(18) = dic(0)("ht_lower")
                c(19) = dic(0)("ht_upper")
                c(20) = dic(0)("ht_clipmethod")
                c(21) = dic(0)("ht_clipmethodID")
                c(22) = dic(0)("wt_lower")
                c(23) = dic(0)("wt_upper")
                c(24) = dic(0)("wt_clipmethod")
                c(25) = dic(0)("wt_clipmethodID")
                c(26) = ""
                c(27) = ""
                grdP.Rows.Add(c)
            End If
        End If

    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnSelect.Enabled = True
        btnRemove.Enabled = True
        grdP.ReadOnly = False
        grdP.EditMode = DataGridViewEditMode.EditOnEnter

        'Save a copy of all current prefs data to a log file in the app directory as a backup in case editing is stuffed up
        cMyRoutines.WriteToFile_DataGrid(grdP)

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        If grdP.Rows.Count = 0 Then
            MsgBox("No preferences to save", vbOKCancel, "Save preferences")
            Exit Sub
        End If

        If form_password.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim i As Integer = MsgBox("All current preference data will be over-written. Continue?", vbOKCancel, "Save preferences")
            If i = vbOK Then cPred.Update_PrefsPred(Load_MemFromScreenFields())
            Me.Close()
        End If

    End Sub

    Private Function Load_MemFromScreenFields() As Dictionary(Of String, String)()
        'Returns each grid row as an dictionary element (fieldname,value) in a dictionary array

        Dim d() As Dictionary(Of String, String) = Nothing
        Dim e() As Dictionary(Of String, String) = Nothing
        Dim EqID_mpv As Integer = 0
        Dim EqID_normalrange As Integer = 0

        'Copy all rows
        For i As Integer = 0 To grdP.Rows.Count - 1
            ReDim Preserve d(i)
            d(i) = New Dictionary(Of String, String)
            d(i) = cMyRoutines.MakeEmpty_dicPrefPred_ForSave

            'Get MPV and normal range equationID's - each pref should have both a MPV and a normal range equation in the equations table 
            e = cPred.Match_Equation(grdP("testID", i).Value, grdP("sourceID", i).Value, grdP("parameterID", i).Value, grdP("agegroupID", i).Value, grdP("genderID", i).Value, grdP("ethnicityID", i).Value)
            Select Case e.Count
                Case 2
                    If e(0)("stattype") = "MPV" Then
                        EqID_mpv = e(0)("equationid")
                        EqID_normalrange = e(1)("equationid")
                    ElseIf e(1)("stattype") = "MPV" Then
                        EqID_mpv = e(1)("equationid")
                        EqID_normalrange = e(0)("equationid")
                    Else
                        EqID_mpv = Nothing
                        EqID_normalrange = Nothing
                        MsgBox("Problem with preference in row " & i & vbCrLf & "Equations (MPV and normal range) not found. Save aborted.")
                        Return Nothing
                        Exit Function
                    End If
                Case Else
                    MsgBox("Problem with preference in row " & i & vbCrLf & "Should be 2 equations (MPV and normal range). There were " & e.Count & "found. Save aborted.")
                    Return Nothing
                    Exit Function
            End Select

            'Copy row data
            Dim col As Integer = 0
            Dim s As String = ""
            For col = 1 To grdP.ColumnCount - 1
                s = grdP.Columns(col).HeaderCell.Value
                If d(i).ContainsKey(s) Then
                    If s = "startdate" Or s = "enddate" Then
                        If IsDate(grdP(col, i).Value) Then
                            d(i)(grdP.Columns(col).HeaderCell.Value) = "'" & Format(CDate(grdP(col, i).Value), "yyyy-MM-dd") & "'"
                        Else
                            d(i)(grdP.Columns(col).HeaderCell.Value) = "NULL"
                        End If
                    ElseIf grdP.Columns(col).HeaderCell.Value = "prefid" Then   'Insert to table will fail if value for PK presnet
                        d(i)(grdP.Columns(col).HeaderCell.Value) = ""
                    ElseIf grdP.Columns(col).HeaderCell.Value = "equationid_mpv" Then
                        d(i)("equationid_mpv") = EqID_mpv
                    ElseIf grdP.Columns(col).HeaderCell.Value = "equationid_normalrange" Then
                        d(i)("equationid_normalrange") = EqID_normalrange
                    Else
                        d(i)(grdP.Columns(col).HeaderCell.Value) = "'" & grdP(col, i).Value & "'"
                    End If
                End If
            Next
        Next

        Return d

    End Function

    Private Function Load_ScreenFieldsFromMem() As Boolean

        Dim d() As Dictionary(Of String, String) = cPred.Get_PrefsPred(True)
        Dim col As Integer = 0

        grdP.Rows.Clear()

        If UBound(d) > 0 Then
            For i As Integer = 1 To UBound(d)
                grdP.Rows.Add()
                For col = 0 To grdP.ColumnCount - 1
                    If col = 0 Then
                        grdP.Item(0, grdP.Rows.Count - 1).Value = i
                    Else
                        If d(i).ContainsKey(grdP.Columns(col).HeaderCell.Value & "") Then
                            grdP.Item(col, grdP.Rows.Count - 1).Value = d(i)(grdP.Columns(col).HeaderCell.Value)
                        End If
                    End If
                Next
            Next
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub form_prefs_normals_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        flagFormLoading = True

        'Load tests listbox
        RemoveHandler lstTests.SelectedValueChanged, AddressOf lstTests_SelectedValueChanged
        cMyRoutines.Listbox_LoadItems(lstTests, cPred.Get_RefItems(RefItems.Tests))
        AddHandler lstTests.SelectedValueChanged, AddressOf lstTests_SelectedValueChanged

        'Load tests toolstrip filter combos
        RemoveHandler ToolStripComboBox_Test.SelectedIndexChanged, AddressOf ToolStripComboBox_Test_SelectedIndexChanged
        cMyRoutines.Combo_ts_LoadItemsFromList(ToolStripComboBox_Test, eTables.Pred_ref_tests)
        AddHandler ToolStripComboBox_Test.SelectedIndexChanged, AddressOf ToolStripComboBox_Test_SelectedIndexChanged

        RemoveHandler ToolStripComboBox_parameter.SelectedIndexChanged, AddressOf ToolStripComboBox_parameter_SelectedIndexChanged
        cMyRoutines.Combo_ts_LoadItemsFromList(ToolStripComboBox_parameter, eTables.Pred_ref_parameters)
        AddHandler ToolStripComboBox_parameter.SelectedIndexChanged, AddressOf ToolStripComboBox_parameter_SelectedIndexChanged


        RemoveHandler ToolStripComboBox_source.SelectedIndexChanged, AddressOf ToolStripComboBox_source_SelectedIndexChanged
        cMyRoutines.Combo_ts_LoadItemsFromList(ToolStripComboBox_source, eTables.Pred_ref_sources)
        AddHandler ToolStripComboBox_source.SelectedIndexChanged, AddressOf ToolStripComboBox_source_SelectedIndexChanged


        RemoveHandler ToolStripComboBox_Gender.SelectedIndexChanged, AddressOf ToolStripComboBox_Gender_SelectedIndexChanged
        cMyRoutines.Combo_ts_LoadItemsFromList(ToolStripComboBox_Gender, eTables.Pred_ref_genders)
        AddHandler ToolStripComboBox_Gender.SelectedIndexChanged, AddressOf ToolStripComboBox_Gender_SelectedIndexChanged

        ToolStripComboBox_EndDated.Items.AddRange(New String() {"Active", "All"})

        'Load pred prefs
        grdP.ReadOnly = True
        SetUp_Grid()
        Me.Load_ScreenFieldsFromMem()

        If lstTests.Items.Count > 0 Then lstTests.SelectedIndex = 0

    End Sub

    Private Sub SplitContainer1_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        grdP.Height = SplitContainer1.Panel2.Height - ToolStrip6.Height
    End Sub

    Private Sub Do_FilterList() Handles ToolStripComboBox_Test.SelectedIndexChanged, ToolStripComboBox_source.SelectedIndexChanged, ToolStripComboBox_parameter.SelectedIndexChanged, ToolStripComboBox_Gender.SelectedIndexChanged, ToolStripComboBox_EndDated.SelectedIndexChanged

        Dim bshow As Boolean = True

        For i As Integer = 0 To grdP.RowCount - 1
            If ToolStripComboBox_Test.SelectedIndex <> -1 And grdP.Item(2, i).Value <> ToolStripComboBox_Test.Text Then
                bshow = False
            Else
                If ToolStripComboBox_source.SelectedIndex <> -1 And grdP.Item(4, i).Value <> ToolStripComboBox_source.Text Then
                    bshow = False
                Else
                    If ToolStripComboBox_parameter.SelectedIndex <> -1 And grdP.Item(6, i).Value <> ToolStripComboBox_parameter.Text Then
                        bshow = False
                    Else
                        If ToolStripComboBox_Gender.SelectedIndex <> -1 And grdP.Item(8, i).Value <> ToolStripComboBox_Gender.Text Then
                            bshow = False
                        Else
                            If (ToolStripComboBox_EndDated.Text <> "Active") Or (ToolStripComboBox_EndDated.Text = "Active" And grdP.Item(17, i).Value = Nothing) Then
                                bshow = True
                            Else
                                bshow = False
                            End If
                        End If
                    End If
                End If
            End If
            grdP.Rows(i).Visible = bshow
        Next

    End Sub

    Private Sub ToolStripComboBox_parameter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox_parameter.SelectedIndexChanged

    End Sub

    Private Sub ToolStripComboBox_source_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox_source.SelectedIndexChanged

    End Sub

    Private Sub ToolStripComboBox_Gender_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox_Gender.SelectedIndexChanged

    End Sub

    Private Sub ToolStripButton_clearfilters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton_clearfilters.Click

        ToolStripComboBox_Gender.SelectedIndex = -1
        ToolStripComboBox_source.SelectedIndex = -1
        ToolStripComboBox_parameter.SelectedIndex = -1
        ToolStripComboBox_Test.SelectedIndex = -1
        ToolStripComboBox_EndDated.SelectedIndex = -1

    End Sub

    Private Sub tsbSource_edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSource_edit.Click

        Dim kv As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))
        Dim kv1 As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        Dim f As Form = New form_PredSourceEditor(EditModes.EditRecord, kv.Key, kv1.Key)
        f.Show()

    End Sub

    Private Sub tsbSource_new_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSource_new.Click

        Dim kv As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))
        Dim kv1 As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        Dim f As Form = New form_PredSourceEditor(EditModes.NewRecord, kv.Key, 0)
        f.Show()

    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click

        Dim i As Integer = grdP.SelectedRows(0).Index
        If i >= 0 And i <= grdP.Rows.Count - 1 Then grdP.Rows.RemoveAt(grdP.SelectedRows(0).Index)

    End Sub

    Private Sub ToolStripComboBox_Test_SelectedIndexChanged(sender As Object, e As EventArgs)
        Throw New NotImplementedException
    End Sub

End Class