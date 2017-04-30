Imports System.Text
Imports System.Collections.Generic

Public Class form_Pred

    Dim flagFormLoading As Boolean = False

    Private Sub form_Pred_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If flagFormLoading Then
            





        End If

        flagFormLoading = False

    End Sub

    Private Sub form_Pred_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        flagFormLoading = True
        'module_initialise.InitialiseApp()

        SplitContainer1.SplitterDistance = 400

        'Load tests listbox
        RemoveHandler lstTests.SelectedValueChanged, AddressOf lstTests_SelectedValueChanged
        cMyRoutines.Listbox_LoadItems(lstTests, cPred.Get_RefItems(RefItems.Tests))
        AddHandler lstTests.SelectedValueChanged, AddressOf lstTests_SelectedValueChanged

        'RemoveHandler cmboGender.SelectedValueChanged, AddressOf cmboGender_SelectedValueChanged
        cMyRoutines.Combo_LoadItems(cmboGender, cPred.Get_RefItems(RefItems.Genders))
        'AddHandler cmboGender.SelectedValueChanged, AddressOf cmboGender_SelectedValueChanged

        'Defaults
        txtDOB.Text = "01/01/1970"
        txtTestDate.Text = "01/01/2014"
        txtHt.Text = CStr(170)
        txtWt.Text = CStr(70)
        cmboGender.SelectedIndex = 0

    End Sub

    Private Sub lstTests_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTests.SelectedValueChanged

        'Clear options
        lstSets.Items.Clear()
        lstParameters.Items.Clear()
        grdEqs.Rows.Clear()
        cmboEthnicity.Items.Clear()

        'Get sources for selected test
        Dim kv As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))
        cMyRoutines.Listbox_LoadItems(lstSets, cPred.Get_SourcesForTestID(kv.Key))

    End Sub

    Private Sub lstParameters_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstParameters.SelectedValueChanged

        Dim Eqs() As Dictionary(Of String, String)
        Dim InputData As Pred_InputData = Nothing

        Dim kvP As KeyValuePair(Of String, String) = CType(lstParameters.SelectedItem, KeyValuePair(Of String, String))
        Dim kvS As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        Dim kvT As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))

        InputData.SourceID = kvS.Key
        InputData.ParameterID = kvP.Key

        Eqs = cPred.Get_EquationsForSourceID_ParameterID(InputData, InputData.ParameterID)

        grdEqs.Rows.Clear()
        For i As Integer = 1 To UBound(Eqs)
            grdEqs.Rows.Add()
            grdEqs("EquationID", i - 1).Value = Eqs(i)("EquationID")
            grdEqs("Source", i - 1).Value = kvS.Key
            grdEqs("Test", i - 1).Value = kvT.Key
            grdEqs("Parameter", i - 1).Value = Eqs(i)("Parameter")
            grdEqs("Statistic", i - 1).Value = Eqs(i)("StatType")
            grdEqs("Gender", i - 1).Value = Eqs(i)("Gender")
            grdEqs("Ethnicity", i - 1).Value = Eqs(i)("Ethnicity")
            grdEqs("AgeGroup", i - 1).Value = Eqs(i)("AgeGroup")
            grdEqs("MinAge", i - 1).Value = Eqs(i)("Age_lower")
            grdEqs("MaxAge", i - 1).Value = Eqs(i)("Age_upper")
            grdEqs("Equation", i - 1).Value = Eqs(i)("Equation")
            grdEqs("EthnicityCorrectionType", i - 1).Value = Eqs(i)("EthnicityCorrectionType")
        Next

    End Sub

    Private Sub lstSets_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSets.SelectedIndexChanged

        Dim Parameters As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)

        Dim kv As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        Parameters = cPred.Get_ParametersForSourceID(kv.Key)

        lstParameters.Items.Clear()
        cmboEthnicity.Items.Clear()
        grdEqs.Rows.Clear()

        'Load relevant ethnicities for source
        kv = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
        cMyRoutines.Combo_LoadItems(cmboEthnicity, cPred.Get_EthnicitiesForSource(kv.Key))
        If cmboEthnicity.Items.Count > 0 Then cmboEthnicity.SelectedIndex = 0

        'Load parameters for source/test
        RemoveHandler lstParameters.SelectedValueChanged, AddressOf lstParameters_SelectedValueChanged
        cMyRoutines.Listbox_LoadItems(lstParameters, Parameters)
        AddHandler lstParameters.SelectedValueChanged, AddressOf lstParameters_SelectedValueChanged

    End Sub

    Private Sub btnCalculate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCalculate.Click

        Dim dic As New Dictionary(Of Integer, Single)
        Dim demo As Pred_demo = Nothing

        MsgBox("Under construction, apologies.")
        Exit Sub


       
        'Check selections all OK
        If lstTests.SelectedItems.Count = 0 Then
            MsgBox("Please select a test", vbOKOnly, "ResLab (Lite)")
            Exit Sub
        End If     
        If lstSets.SelectedItems.Count = 0 Then
            MsgBox("Please select a source", vbOKOnly, "ResLab (Lite)")
            Exit Sub
        End If
        If lstParameters.SelectedItems.Count = 0 Then
            MsgBox("Please select a parameter", vbOKOnly, "ResLab (Lite)")
            Exit Sub
        End If


        demo.Age = cMyRoutines.Calc_Age(CDate(txtDOB.Text), CDate(txtTestDate.Text))
        demo.EthnicityID = cMyRoutines.Combo_GetValue(cmboEthnicity)
        demo.GenderID = cMyRoutines.Combo_GetValue(cmboGender)
        demo.Htcm = txtHt.Text
        demo.Wtkg = txtWt.Text


        dic = cPred.Calc_PredForParameter(cMyRoutines.Listbox_GetValue(lstSets), cMyRoutines.Listbox_GetValue(lstParameters), demo)
        If IsNothing(dic) Then

        Else
            txtParameter.Text = lstParameters.Text
            txtUnits.Text = cPred.Get_UnitsForParamID(cMyRoutines.Listbox_GetValue(lstParameters))
            If dic.ContainsKey(StatType.MPV) Then txtMPV.Text = dic(StatType.MPV) Else txtMPV.Text = "---"
            If dic.ContainsKey(StatType.LLN) Then txtLLN.Text = dic(StatType.LLN) Else txtLLN.Text = "---"
            If dic.ContainsKey(StatType.ULN) Then txtULN.Text = dic(StatType.ULN) Else txtULN.Text = "---"
            If dic.ContainsKey(StatType.Range) Then txtRange.Text = dic(StatType.Range) Else txtRange.Text = "---"
        End If

    End Sub

    Private Sub SplitContainer1_SplitterMoved(ByVal sender As Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer1.SplitterMoved
        grdEqs.Height = SplitContainer1.Panel2.Height - ToolStrip5.Height
        grdEqs.Top = ToolStrip5.Height
    End Sub


    Private Sub tsbSource_edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSource_edit.Click

        If lstSets.SelectedItems.Count = 1 Then
            Dim kv As KeyValuePair(Of String, String) = CType(lstSets.SelectedItem, KeyValuePair(Of String, String))
            Dim kv1 As KeyValuePair(Of String, String) = CType(lstTests.SelectedItem, KeyValuePair(Of String, String))
            Dim f As New form_PredSourceEditor(form_PredSourceEditor.EditModes.EditRecord, kv1.Key, kv.Key)
            f.Show()
        End If

    End Sub

    Private Sub tsbSource_new_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSource_new.Click

        Dim f As New form_PredSourceEditor(form_PredSourceEditor.EditModes.NewRecord)
        f.Show()

    End Sub

    Private Sub grdEqs_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdEqs.CellContentClick

    End Sub

    Private Sub SplitContainer2_Panel2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer2.Panel2.Paint

    End Sub

    Private Sub SplitContainer2_Panel1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer2.Panel1.Paint

    End Sub
End Class