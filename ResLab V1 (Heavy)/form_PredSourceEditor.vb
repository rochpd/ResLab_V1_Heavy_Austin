Imports System.Text

Public Class form_PredSourceEditor

    Public Enum EditModes
        NewRecord
        EditRecord
    End Enum

    Dim EditMode As EditModes
    Dim rd As New form_prefs_normals.RowData
    Dim flagFormLoading As Boolean = False
    Dim _SourceID As Integer = 0
    Dim _TestID As Integer = 0

    Public Sub New(ByVal Mode As EditModes, Optional ByVal TestID As Integer = 0, Optional ByVal SourceID As Integer = 0)

        InitializeComponent()
        EditMode = Mode
        _SourceID = SourceID
        _TestID = TestID

    End Sub

    Private Sub form_PredSourceEditor_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If flagFormLoading Then

            'Load cmbo with all available tests - don't load parameters until a test is selected
            cMyRoutines.Combo_LoadItems(cmboTest, cPred.Get_RefItems(RefItems.Tests))

            Select Case EditMode
                Case EditModes.NewRecord
                    txtSource.Tag = 0
                    btnCancel.Enabled = True
                    btnClose.Enabled = False
                    btnEdit.Enabled = False
                    btnSave.Enabled = True

                Case EditModes.EditRecord

                    'Set selected test
                    Dim kv As KeyValuePair(Of String, String) = cPred.Get_TestForSourceID(_SourceID)
                    cmboTest.SelectedIndex = cmboTest.FindStringExact(kv.Key)

                    'Get source info and load
                    Dim SourceInfo As form_prefs_normals.RowData = cPred.Get_SourceInfo(_SourceID, _TestID)
                    txtSource.Text = SourceInfo.Source.Value
                    txtSource.Tag = SourceInfo.SourceID.Value
                    txtReference.Text = SourceInfo.Source_PubReference.Value
                    'txtPubYear.Text = SourceInfo.Source_PubYear.Value
                    cmboEquationClass.SelectedIndex = cmboEquationClass.FindStringExact(SourceInfo.Equation_Class.Value)
                    txtEquation.Tag = ""

                    'Load equations for source and test into grid
                    Me.load_equationstogrid()

            End Select
            'btnEdit.PerformClick()

        End If

        flagFormLoading = False

    End Sub

    Private Sub load_equationstogrid()

        'Load equations for source and test into grid
        Dim Eqs() As form_prefs_normals.RowData = cPred.Get_EquationsAllFor_SourceID_TestID(_SourceID, _TestID)
        Dim c(31) As String
        grdEqs.Rows.Clear()

        For i As Integer = 0 To Eqs.GetUpperBound(0)

            Array.Clear(c, 0, c.GetUpperBound(0))

            c(0) = grdEqs.RowCount + 1
            c(1) = Eqs(i).Test.Value
            c(2) = Eqs(i).TestID.Value
            c(3) = Eqs(i).Source.Value
            c(4) = Eqs(i).SourceID.Value
            c(5) = Eqs(i).Parameter.Value
            c(6) = Eqs(i).ParameterID.Value
            c(7) = Eqs(i).Gender.Value
            c(8) = Eqs(i).GenderID.Value
            c(9) = Eqs(i).AgeGroup.Value
            c(10) = Eqs(i).AgeGroupID.Value
            c(11) = Eqs(i).Age_lower.Value
            c(12) = Eqs(i).Age_upper.Value
            c(13) = Eqs(i).Age_clipmethod.Value
            c(14) = Eqs(i).Age_clipmethodID.Value
            c(15) = Eqs(i).Ht_lower.Value
            c(16) = Eqs(i).Ht_upper.Value
            c(17) = Eqs(i).Ht_clipmethod.Value
            c(18) = Eqs(i).Ht_clipmethodID.Value
            c(19) = Eqs(i).Wt_lower.Value
            c(20) = Eqs(i).Wt_upper.Value
            c(21) = Eqs(i).Wt_clipmethod.Value
            c(22) = Eqs(i).Wt_clipmethodID.Value
            c(23) = Eqs(i).Ethnicity.Value
            c(24) = Eqs(i).EthnicityID.Value
            c(25) = Eqs(i).Equation.Value
            c(26) = Eqs(i).Equationid.Value
            c(27) = Eqs(i).StatType.Value
            c(28) = Eqs(i).StatTypeid.Value
            c(29) = Eqs(i).Equation_Class.Value
            c(30) = Eqs(i).Equation_ClassID.Value
            c(31) = Eqs(i).Ethnicity_ApplyATS1991Correction.Value
            Me.AddNewEquationInfoToGrid(c)
        Next

    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click


        'If form_password.ShowDialog() = Windows.Forms.DialogResult.OK Then

        If cmboParameter.SelectedIndex = -1 Then
            MsgBox("Please select a parameter")
            Exit Sub
        End If
        If cmboTest.SelectedIndex = -1 Then
            MsgBox("Please select a test")
            Exit Sub
        End If

        Dim c(31) As String
        c(0) = grdEqs.RowCount + 1
        c(1) = cmboTest.SelectedItem.key
        c(2) = cmboTest.SelectedItem.value
        c(3) = txtSource.Text
        c(4) = txtSource.Tag  'SourceID
        c(5) = cmboParameter.SelectedItem.key
        c(6) = cmboParameter.SelectedItem.value
        c(7) = cmboGender.SelectedItem.key
        c(8) = cmboGender.SelectedItem.value
        c(9) = cmboAgeGroup.SelectedItem.key
        c(10) = cmboAgeGroup.SelectedItem.value
        c(11) = txtAgeLower.Text
        c(12) = txtAgeUpper.Text
        c(13) = cmboAge_clipmethod.SelectedItem.key
        c(14) = cmboAge_clipmethod.SelectedItem.value
        c(15) = txtHtLower.Text
        c(16) = txtHtUpper.Text
        c(17) = cmboHt_clipmethod.SelectedItem.key
        c(18) = cmboHt_clipmethod.SelectedItem.value
        c(19) = txtWtLower.Text
        c(20) = txtWtUpper.Text
        c(21) = cmboWt_clipmethod.SelectedItem.key
        c(22) = cmboWt_clipmethod.SelectedItem.Value
        c(23) = cmboEthnicity.SelectedItem.key
        c(24) = cmboEthnicity.SelectedItem.value
        c(25) = txtEquation.Text
        If txtEquation.Tag = 0 Then c(26) = "0" Else c(26) = txtEquation.Tag
        c(27) = cmboEquationClass.SelectedItem.key
        c(28) = cmboEquationClass.SelectedItem.value
        c(29) = cmboStat.SelectedItem.key
        c(30) = cmboStat.SelectedItem.value
        Select Case chkApplyNonCaucasianCorrection.Checked
            Case True : c(31) = "ATS(1991)"
            Case False : c(31) = ""
        End Select


        'Save to database
        Dim RecordID_eq As Integer = 0
        Dim RecordID_source As Integer = 0

        Select Case EditMode
            Case EditModes.EditRecord
                RecordID_source = cPred.Update_Source(Load_MemFromScreenFields_source())
                If txtEquation.Tag = 0 Then
                    RecordID_eq = cPred.New_Equation(Load_MemFromScreenFields_eq())
                Else
                    Me.UpdateEquationInfoInGrid(grdEqs.SelectedRows(0).Index, c)
                    RecordID_eq = cPred.Update_Equation(Load_MemFromScreenFields_eq())
                End If
            Case EditModes.NewRecord
                'Check if already exists in DB 
                Dim dic() As Dictionary(Of String, String)
                dic = cPred.Match_Equation(c(2), c(4), c(6), c(10), c(8), c(24), c(30))
                If Not (dic Is Nothing) Then
                    'Already exists
                    MsgBox("No can do." & vbCrLf & "Already exists in the database")
                Else
                    'Good to save - no match in DB
                    Me.AddNewEquationInfoToGrid(c)
                    RecordID_eq = cPred.New_Equation(Load_MemFromScreenFields_eq())
                    RecordID_source = cPred.New_Source(Load_MemFromScreenFields_source())
                End If
        End Select
        ' End If

        'Refresh
        Me.load_equationstogrid()

    End Sub

    Private Function Load_MemFromScreenFields_eq() As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)

        d = cMyRoutines.MakeEmpty_dic_Pred_equation()
        Dim flds = New class_pred_EquationFields

        d(flds.Test) = "'" & cmboTest.SelectedItem.key & "'"
        d(flds.TestID) = "'" & cmboTest.SelectedItem.value & "'"
        d(flds.Source) = "'" & txtSource.Text & "'"
        d(flds.SourceID) = "'" & txtSource.Tag & "'"
        d(flds.Parameter) = "'" & cmboParameter.SelectedItem.key & "'"
        d(flds.ParameterID) = "'" & cmboParameter.SelectedItem.value & "'"
        d(flds.Gender) = "'" & cmboGender.SelectedItem.key & "'"
        d(flds.GenderID) = "'" & cmboGender.SelectedItem.value & "'"
        d(flds.AgeGroup) = "'" & cmboAgeGroup.SelectedItem.key & "'"
        d(flds.AgeGroupID) = "'" & cmboAgeGroup.SelectedItem.value & "'"
        d(flds.Age_lower) = "'" & txtAgeLower.Text & "'"
        d(flds.Age_upper) = "'" & txtAgeUpper.Text & "'"
        d(flds.Age_clipmethod) = "'" & cmboAge_clipmethod.SelectedItem.key & "'"
        d(flds.Age_clipmethodID) = "'" & cmboAge_clipmethod.SelectedItem.value & "'"
        d(flds.Ht_lower) = "'" & txtHtLower.Text & "'"
        d(flds.Ht_upper) = "'" & txtHtUpper.Text & "'"
        d(flds.Ht_clipmethod) = "'" & cmboHt_clipmethod.SelectedItem.key & "'"
        d(flds.Ht_clipmethodID) = "'" & cmboHt_clipmethod.SelectedItem.value & "'"
        d(flds.Wt_lower) = "'" & txtWtLower.Text & "'"
        d(flds.Wt_upper) = "'" & txtWtUpper.Text & "'"
        d(flds.Wt_clipmethod) = "'" & cmboWt_clipmethod.SelectedItem.key & "'"
        d(flds.Wt_clipmethodID) = "'" & cmboWt_clipmethod.SelectedItem.value & "'"
        d(flds.Ethnicity) = "'" & cmboEthnicity.SelectedItem.key & "'"
        d(flds.EthnicityID) = "'" & cmboEthnicity.SelectedItem.value & "'"
        d(flds.Equation) = "'" & txtEquation.Text & "'"
        If txtEquation.Tag = 0 Then d(flds.EquationID) = "0" Else d(flds.EquationID) = "'" & txtEquation.Tag & "'"
        d(flds.EquationType) = "'" & cmboEquationClass.SelectedItem.key & "'"
        d(flds.EquationTypeID) = "'" & cmboEquationClass.SelectedItem.value & "'"
        d(flds.StatType) = "'" & cmboStat.SelectedItem.key & "'"
        d(flds.StatTypeID) = "'" & cmboStat.SelectedItem.value & "'"
        Select Case chkApplyNonCaucasianCorrection.Checked
            Case True : d(flds.EthnicityCorrectionType) = "'ATS(1991)'"
            Case False : d(flds.EthnicityCorrectionType) = "'" & "" & "'"
        End Select
        d(flds.Lastedit) = "'" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
        d(flds.LasteditBy) = "'" & "" & "'"

        Return d

    End Function

    Private Function Load_MemFromScreenFields_source() As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)

        d = cMyRoutines.MakeEmpty_dic_Pred_source()
        Dim flds = New class_pred_SourceFields

        If txtSource.Tag = "" Then d(flds.SourceID) = "0" Else d(flds.SourceID) = "'" & txtSource.Tag & "'"
        d(flds.Source) = "'" & txtSource.Text & "'"
        d(flds.Pub_Reference) = "'" & txtReference.Text & "'"
        'd(flds.Pub_Year) = "'" & txtPubYear.Text & "'"
        d(flds.Lastedit) = "'" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "'"
        d(flds.LasteditBy) = "'" & "" & "'"

        Return d

    End Function

    Private Sub form_PredSourceEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        flagFormLoading = True

        cMyRoutines.Combo_LoadItems(cmboEthnicity, cPred.Get_RefItems(RefItems.Ethnicities))
        cMyRoutines.Combo_LoadItems(cmboAgeGroup, cPred.Get_RefItems(RefItems.AgeGroups))
        cMyRoutines.Combo_LoadItems(cmboEquationClass, cPred.Get_RefItems(RefItems.EquationTypes))
        cMyRoutines.Combo_LoadItems(cmboStat, cPred.Get_RefItems(RefItems.StatisticTypes))
        cMyRoutines.Combo_LoadItems(cmboGender, cPred.Get_RefItems(RefItems.Genders))
        cMyRoutines.Combo_LoadItems(cmboAge_clipmethod, cPred.Get_RefItems(RefItems.ClipMethods))
        cMyRoutines.Combo_LoadItems(cmboHt_clipmethod, cPred.Get_RefItems(RefItems.ClipMethods))
        cMyRoutines.Combo_LoadItems(cmboWt_clipmethod, cPred.Get_RefItems(RefItems.ClipMethods))

        SetUp_Grid()

    End Sub


    Private Sub SetUp_Grid()

        Dim i As Integer = 0

        grdEqs.ColumnCount = 32

        grdEqs.Columns(i).HeaderText = rd.Count.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Test.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.TestID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Source.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.SourceID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Parameter.Name : grdEqs.Columns(i).Name = rd.Parameter.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.ParameterID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Gender.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.GenderID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.AgeGroup.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.AgeGroupID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Age_lower.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Age_upper.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Age_clipmethod.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Age_clipmethodID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Ht_lower.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Ht_upper.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Ht_clipmethod.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Ht_clipmethodID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Wt_lower.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Wt_upper.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Wt_clipmethod.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Wt_clipmethodID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Ethnicity.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.EthnicityID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Equation.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Equationid.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.StatType.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.StatTypeid.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Equation_Class.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Equation_ClassID.Name : i = i + 1
        grdEqs.Columns(i).HeaderText = rd.Ethnicity_ApplyATS1991Correction.Name : i = i + 1


        i = 0
        grdEqs.Columns(i).Width = 30
        grdEqs.Columns(i).Visible = rd.Count.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Test.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.TestID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Source.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.SourceID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Parameter.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.ParameterID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Gender.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.GenderID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.AgeGroup.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.AgeGroupID.visible : i = i + 1
        grdEqs.Columns(i).Width = 80
        grdEqs.Columns(i).Visible = rd.Age_lower.visible : i = i + 1
        grdEqs.Columns(i).Width = 80
        grdEqs.Columns(i).Visible = rd.Age_upper.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Age_clipmethod.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Age_clipmethodID.visible : i = i + 1
        grdEqs.Columns(i).Width = 80
        grdEqs.Columns(i).Visible = rd.Ht_lower.visible : i = i + 1
        grdEqs.Columns(i).Width = 80
        grdEqs.Columns(i).Visible = rd.Ht_upper.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Ht_clipmethod.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Ht_clipmethodID.visible : i = i + 1
        grdEqs.Columns(i).Width = 80
        grdEqs.Columns(i).Visible = rd.Wt_lower.visible : i = i + 1
        grdEqs.Columns(i).Width = 80
        grdEqs.Columns(i).Visible = rd.Wt_upper.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Wt_clipmethod.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Wt_clipmethodID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Ethnicity.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.EthnicityID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Equation.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Equationid.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Equation_Class.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Equation_ClassID.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.StatType.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.StatTypeid.visible : i = i + 1
        grdEqs.Columns(i).Visible = rd.Ethnicity_ApplyATS1991Correction.visible : i = i + 1

        grdEqs.AllowUserToResizeColumns = True


    End Sub


    Private Sub AddNewEquationInfoToGrid(ByVal c() As String)

        grdEqs.Rows.Add()
        For i As Integer = 0 To grdEqs.Columns.Count - 1
            grdEqs.Item(i, grdEqs.Rows.Count - 1).Value = c(i)
        Next

    End Sub

    Private Sub UpdateEquationInfoInGrid(Row As Integer, ByVal c() As String)

        For i As Integer = 0 To grdEqs.Columns.Count - 1
            grdEqs.Item(i, Row).Value = c(i)
        Next

    End Sub



    Private Sub Load_equationFromGrid(ByVal RowIndex As Integer)

        cmboParameter.SelectedIndex = cmboParameter.FindStringExact(grdEqs.Item(5, RowIndex).Value)
        cmboGender.SelectedIndex = cmboGender.FindStringExact(grdEqs.Item(7, RowIndex).Value)
        cmboEthnicity.SelectedIndex = cmboEthnicity.FindStringExact(grdEqs.Item(23, RowIndex).Value)
        cmboAgeGroup.SelectedIndex = cmboAgeGroup.FindStringExact(grdEqs.Item(9, RowIndex).Value)
        txtAgeLower.Text = grdEqs.Item(11, RowIndex).Value
        txtAgeUpper.Text = grdEqs.Item(12, RowIndex).Value
        cmboAge_clipmethod.SelectedIndex = cmboAge_clipmethod.FindStringExact(grdEqs.Item(13, RowIndex).Value)
        txtHtLower.Text = grdEqs.Item(15, RowIndex).Value
        txtHtUpper.Text = grdEqs.Item(16, RowIndex).Value
        cmboHt_clipmethod.SelectedIndex = cmboHt_clipmethod.FindStringExact(grdEqs.Item(17, RowIndex).Value)
        txtWtLower.Text = grdEqs.Item(19, RowIndex).Value
        txtWtUpper.Text = grdEqs.Item(20, RowIndex).Value
        cmboWt_clipmethod.SelectedIndex = cmboWt_clipmethod.FindStringExact(grdEqs.Item(21, RowIndex).Value)
        txtEquation.Text = grdEqs.Item(25, RowIndex).Value
        txtEquation.Tag = grdEqs.Item(26, RowIndex).Value  'equationID
        cmboEquationClass.SelectedIndex = cmboEquationClass.FindStringExact(grdEqs.Item(29, RowIndex).Value)
        cmboStat.SelectedIndex = cmboStat.FindStringExact(grdEqs.Item(27, RowIndex).Value)
        Select Case grdEqs.Item(31, RowIndex).Value
            Case "ATS(1991)" : chkApplyNonCaucasianCorrection.Checked = True
            Case "" : chkApplyNonCaucasianCorrection.Checked = False
        End Select

    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click

        Me.Close()

    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click

        If cmboTest.Items.Count = 1 Then
            cmboTest.SelectedIndex = 0
        End If

        btnCancel.Enabled = True
        btnClose.Enabled = False
        btnEdit.Enabled = False
        btnSave.Enabled = True
        btnNew.Enabled = True

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        btnCancel.Enabled = False
        btnClose.Enabled = True
        btnEdit.Enabled = True
        btnSave.Enabled = False
        btnNew.Enabled = False
    End Sub



    Private Sub btnNew_Click(sender As Object, e As System.EventArgs) Handles btnNew.Click

        txtEquation.Tag = 0
        'txtEquation.Text = ""
        cmboParameter.SelectedIndex = -1
        'txtAgeLower.Text = ""
        'txtAgeUpper.Text = ""
        'txtHtLower.Text = ""
        'txtHtUpper.Text = ""
        'txtWtLower.Text = ""
        'txtWtUpper.Text = ""
        'cmboAge_clipmethod.SelectedIndex = -1
        'cmboHt_clipmethod.SelectedIndex = -1
        'cmboWt_clipmethod.SelectedIndex = -1
        'cmboAgeGroup.SelectedIndex = -1
        'cmboEthnicity.SelectedIndex = -1
        'cmboGender.SelectedIndex = -1
        'cmboStat.SelectedIndex = -1
        'chkApplyNonCaucasianCorrection.Checked = False


    End Sub

    Private Sub grdEqs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEqs.Click

        If grdEqs.SelectedRows.Count = 1 Then
            Me.Load_equationFromGrid(grdEqs.SelectedCells(0).RowIndex)
        End If

    End Sub


    Private Sub cmboTest_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmboTest.SelectedIndexChanged

        Dim i As Integer = 0

        'Load parameter listbox with all available parameters for selected test
        Dim kv As KeyValuePair(Of String, String) = CType(cmboTest.SelectedItem, KeyValuePair(Of String, String))
        Dim d As Dictionary(Of String, String) = cPred.Get_ParametersForTestID(kv.Key)

        lstParams.Items.Clear()
        lstParams.DisplayMember = "value"
        For Each kv In d
            lstParams.Items.Add(kv)
        Next

        'Check the checkbox for parameters selected for source
        Dim Parameters As Dictionary(Of String, String) = cPred.Get_ParametersForSourceID(_SourceID)
        cmboParameter.Items.Clear()
        cmboParameter.Text = ""
        cmboParameter.DisplayMember = "value"
        For i = 0 To lstParams.Items.Count - 1
            If Parameters.ContainsKey(lstParams.Items(i).key) Then
                lstParams.SetItemChecked(i, True)
                'Add to equation parameter combo as well
                cmboParameter.Items.Add(lstParams.Items(i))
            End If
        Next

    End Sub

 

    Private Sub lstParams_ItemCheck(sender As Object, e As System.Windows.Forms.ItemCheckEventArgs) Handles lstParams.ItemCheck
        'If a parameter is unchecked, and equations exist for this parameter, don't allow unchecking until
        '  all existing equations for this parameter are deleted.

        If Not flagFormLoading Then
            Dim item As ItemCheckEventArgs = e
            Select Case item.NewValue
                Case CheckState.Unchecked
                    'Check for existing equations in grid for unchecked parameter 
                    Dim i As Integer = 0
                    Dim j As Integer = 0
                    For i = 0 To grdEqs.Rows.Count - 1
                        If grdEqs.Item("parameter", i).Value = lstParams.Items(item.Index).key Then
                            j = j + 1
                        End If
                    Next
                    If j > 0 Then
                        Dim Msg As New StringBuilder
                        Msg.Append("No can do." & vbCrLf)
                        Msg.Append("Equations (x" & j & ") already exist for " & lstParams.Items(item.Index).key & vbCrLf)
                        Msg.Append("All equations for this parameter must first be deleted before deselecting.")
                        MsgBox(Msg.ToString)
                        e.NewValue = e.CurrentValue
                    Else
                        cmboParameter.Items.RemoveAt(item.Index)
                    End If



                Case CheckState.Checked
                    'Add item to parameters combo
                    cmboParameter.Items.Add(lstParams.Items(item.Index))
            End Select
        End If

    End Sub


End Class