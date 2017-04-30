Imports System.Reflection
Imports System.Linq.Expressions


Public Class form_spt_panelbuilder

    Private _Show_CopyToSptButton As Boolean

    Public Sub New(ByVal Show_CopyToSptButton As Boolean)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _Show_CopyToSptButton = Show_CopyToSptButton
        ToolStripButton_CopyPanelToTest.Visible = Show_CopyToSptButton
        ToolStripButton_CopyPanelToTest2.Visible = Show_CopyToSptButton

    End Sub

    Private Sub form_spt_panelbuilder_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.Load_panels()
        Me.Load_AllergenCategories()

        'SplitContainer1.SplitterDistance = Me.Width / 2
        'SplitContainer_Allergens.SplitterDistance = SplitContainer1.Panel2.Width / 2
        'SplitContainer_Panels.SplitterDistance = SplitContainer1.Panel1.Width / 2

    End Sub

    Private Sub Load_panels()

        Dim lvi As ListViewItem
        Dim p As Generic.List(Of PanelData) = cSpt.Get_AllergenPanels(False)

        lvPanels.Clear()
        lvPanels.View = View.Details
        lvPanels.FullRowSelect = True
        lvPanels.GridLines = False
        lvPanels.HeaderStyle = ColumnHeaderStyle.None

        ' Create columns for the items and subitems
        lvPanels.Columns.Add("", 0, HorizontalAlignment.Left)
        lvPanels.Columns.Add("panelID", 0, HorizontalAlignment.Left)
        lvPanels.Columns.Add("panelName", lvPanels.Width - 5, HorizontalAlignment.Left)
        If Not IsNothing(p) Then
            For i As Integer = 0 To p.Count - 1
                lvi = New ListViewItem
                lvi.SubItems.Add(p(i).panelid)
                lvi.SubItems.Add(p(i).panelname)
                lvPanels.Items.Add(lvi)
            Next (i)
        End If


    End Sub

    Private Sub Load_panelmembers()

        If lvPanels.SelectedItems.Count > 0 Then
            Try
                Dim panelID As Integer = CLng(lvPanels.Items(lvPanels.SelectedIndices(0)).SubItems(1).Text)
                If panelID > 0 Then
                    Dim a As List(Of PanelMember_AllData) = cSpt.Get_AllergensForPanel(False, panelID)
                    lvPanelAllergens.Clear()
                    If Not IsNothing(a) Then
                        Dim lvi As ListViewItem
                        lvPanelAllergens.View = View.Details
                        lvPanelAllergens.FullRowSelect = True
                        lvPanelAllergens.GridLines = False
                        lvPanelAllergens.HeaderStyle = ColumnHeaderStyle.None

                        ' Create columns for the items and subitems
                        lvPanelAllergens.Columns.Add("", 0, HorizontalAlignment.Left)
                        lvPanelAllergens.Columns.Add("allergenID", 0, HorizontalAlignment.Left)
                        lvPanelAllergens.Columns.Add("allergenName", lvPanelAllergens.Width - 5, HorizontalAlignment.Left)
                        lvPanelAllergens.Columns.Add("displayColour", 0, HorizontalAlignment.Left)
                        lvPanelAllergens.Columns.Add("memberid", 0, HorizontalAlignment.Left)
                        For i As Integer = 0 To a.Count - 1
                            lvi = New ListViewItem
                            lvi.SubItems.Add(a(i).allergenid)
                            lvi.SubItems.Add(a(i).allergenname)
                            lvi.SubItems.Add(a(i).displayColour)
                            lvi.ForeColor = Color.FromArgb(a(i).displayColour)
                            lvi.SubItems.Add(a(i).memberid)
                            lvPanelAllergens.Items.Add(lvi)
                        Next i
                    End If
                End If
            Catch
                MsgBox("Error in Load_panelmembers" & vbNewLine & Err.Description)
            End Try
        End If

    End Sub

    Private Sub Save_panelmembers()

        Dim panelID As Integer = lvPanels.SelectedItems(0).SubItems(1).Text
        For i = 0 To lvPanelAllergens.Items.Count - 1
            Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptPanelMember_TableDataOnly
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelMember_TableDataOnly).memberid)) = lvPanelAllergens.Items(i).SubItems(4).Text
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelMember_TableDataOnly).panelid)) = lvPanels.SelectedItems(0).SubItems(1).Text
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelMember_TableDataOnly).allergenid)) = lvPanelAllergens.Items(i).SubItems(1).Text
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelMember_TableDataOnly).allergenorder)) = i + 1
            If lvPanelAllergens.Items(i).SubItems(4).Text = "0" Then
                cSpt.Insert_PanelMember(dicP)
            Else
                cSpt.Update_PanelMember(dicP)
            End If
            dicP = Nothing
        Next i

    End Sub

    Private Sub Load_AllergensFromDatabase(Optional categoryID As Integer = 0)

        Try
            If lvCategories.SelectedItems.Count = 1 Then
                Dim lvi As ListViewItem
                If categoryID = 0 Then categoryID = lvCategories.SelectedItems(0).SubItems(1).Text
                Dim a As List(Of AllergenData) = cSpt.Get_AllergensFromDatabase(False, categoryID)
                lvAllergens.Clear()
                lvAllergens.View = View.Details
                lvAllergens.FullRowSelect = True
                lvAllergens.GridLines = False
                lvAllergens.HeaderStyle = ColumnHeaderStyle.None

                ' Create columns for the items and subitems
                lvAllergens.Columns.Add("", 0, HorizontalAlignment.Left)
                lvAllergens.Columns.Add("allergenID", 0, HorizontalAlignment.Left)
                lvAllergens.Columns.Add("allergenName", lvAllergens.Width - 5, HorizontalAlignment.Left)

                If Not IsNothing(a) Then
                    For i = 0 To a.Count - 1
                        lvi = New ListViewItem
                        lvi.SubItems.Add(a(i).allergenid)
                        lvi.SubItems.Add(a(i).allergenname)
                        lvAllergens.Items.Add(lvi)
                    Next
                    a = Nothing
                End If
                lvAllergens.ForeColor = lvCategories.SelectedItems(0).ForeColor
            End If
        Catch
            MsgBox("Error in Load_AllergensFromDatabase" & vbNewLine & Err.Description)
        End Try

    End Sub

    Private Sub Load_AllergenCategories()

        Try
            Dim a As List(Of AllergenCategoryData) = cSpt.Get_AllergenCategories(False)
            Dim lvi As ListViewItem

            lvCategories.Clear()
            lvCategories.View = View.Details
            lvCategories.FullRowSelect = True
            lvCategories.GridLines = False
            lvCategories.HeaderStyle = ColumnHeaderStyle.None

            ' Create columns for the items and subitems
            lvCategories.Columns.Add("", 0, HorizontalAlignment.Left)
            lvCategories.Columns.Add("categoryID", 0, HorizontalAlignment.Left)
            lvCategories.Columns.Add("categoryName", lvCategories.Width - 5, HorizontalAlignment.Left)
            lvCategories.Columns.Add("displayColor", 0, HorizontalAlignment.Left)
            lvCategories.Columns.Add("enabled", 0, HorizontalAlignment.Left)
            For i As Integer = 0 To a.Count - 1
                lvi = New ListViewItem
                lvi.SubItems.Add(a(i).categoryid)
                lvi.SubItems.Add(a(i).categoryname)
                lvi.SubItems.Add(a(i).displaycolour)
                lvi.SubItems.Add(a(i).enabled)
                lvi.ForeColor = Color.FromArgb(a(i).displaycolour)
                lvCategories.Items.Add(lvi)
            Next (i)
        Catch
            MsgBox("Error in Load_AllergenCategories" & vbNewLine & Err.Description)
        End Try

    End Sub

    Private Function Update_AllergenCategory() As Integer

        Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptCategory
        dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).categoryid)) = CType(lvCategories.SelectedItems(0).SubItems(1).Text, Integer)
        dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).categoryname)) = "'" & lvCategories.SelectedItems(0).SubItems(2).Text & "'"
        dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).displaycolour)) = "'" & lvCategories.SelectedItems(0).SubItems(3).Text & "'"
        dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).enabled)) = "'" & lvCategories.SelectedItems(0).SubItems(4).Text & "'"
        Dim id As Integer = cSpt.Update_Category(dicP)
        Return id

    End Function

    Private Function New_AllergenCategory() As Integer
        Return Nothing
    End Function

    Private Sub ToolStripButton_newpanel_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_newpanel.Click

        Dim Response As String = ""
        Dim Unique As Boolean = False

        Do
            Response = InputBox("Enter a panel name", "Create new panel")
            If Response <> "" Then
                Unique = cSpt.Is_PanelNameUnique(Response)
                If Not Unique Then
                    MsgBox("Panel name already exists, must be unique", vbOKOnly, "Create new panel")
                End If
            End If
        Loop Until Response = "" Or Unique

        If Unique Then   'Save the panel
            Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptPanel
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelData).panelid)) = 0
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelData).panelname)) = "'" & Response & "'"
            dicP(cMyRoutines.GetPropertyName(Function() (New PanelData).enabled)) = "'" & True & "'"
            cSpt.Insert_Panel(dicP)
            Me.Load_panels()
            If lvPanels.Items.Count > 0 Then lvPanels.FindItemWithText(Response).Selected = True
        End If

    End Sub

    Private Sub ToolStripButton_deletepanel_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_deletepanel.Click

        Select Case lvPanels.SelectedItems.Count
            Case 0
                MsgBox("Please select a panel to delete.", vbOKOnly, "Delete panel")
            Case 1
                Dim Response As Integer = 0
                Response = MsgBox("You are about to permanently delete this panel. Continue?", vbYesNo, "Delete database record")
                If Response = vbYes Then
                    Try
                        Dim panelID As Integer = CType(lvPanels.Items(lvPanels.SelectedIndices(0)).SubItems(1).Text, Integer)
                        If cSpt.Delete_Panel(panelID) Then
                            Me.Load_panels()
                            If lvPanels.Items.Count > 0 Then lvPanels.Items(0).Selected = True
                        Else
                            MsgBox("Problem deleting panel.", vbOKOnly, "Delete panel")
                        End If
                    Catch
                        MsgBox("Problem deleting panel - aborted.", vbOKOnly, "Delete panel")
                    End Try
                End If
        End Select

    End Sub

    Private Sub ToolStripButton_editpanel_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_editpanel.Click

        Dim Response As String = ""
        Dim Unique As Boolean = False

        If lvPanels.SelectedIndices.Count = 1 Then
            Do
                Response = InputBox("Edit panel name", "Edit panel", lvPanels.SelectedItems(0).SubItems(2).Text)
                If Response <> "" Then
                    Unique = cSpt.Is_PanelNameUnique(Response)
                    If Not Unique Then
                        MsgBox("Panel name already exists, must be unique", vbOKOnly, "Edit panel")
                    End If
                End If
            Loop Until Response = "" Or Unique

            If Unique Then   'Save the panel
                Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptPanel
                dicP(cMyRoutines.GetPropertyName(Function() (New PanelData).panelid)) = lvPanels.SelectedItems(0).SubItems(1).Text
                dicP(cMyRoutines.GetPropertyName(Function() (New PanelData).panelname)) = "'" & Response & "'"
                dicP(cMyRoutines.GetPropertyName(Function() (New PanelData).enabled)) = "'" & True & "'"
                cSpt.Update_Panel(dicP)
                Dim CurrentIndex As Integer = lvPanels.SelectedIndices(0)
                Me.Load_panels()
                lvPanels.Items(CurrentIndex).Selected = True
            End If
        Else
            MsgBox("Please select a panel to edit", vbOKOnly, "Edit panel")
        End If

    End Sub

    Private Sub lvPanels_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles lvPanels.SelectedIndexChanged

        Me.Load_panelmembers()

    End Sub

    Private Sub lvCategories_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles lvCategories.SelectedIndexChanged

        Me.Load_AllergensFromDatabase()

    End Sub

    Private Sub ToolStripButton_selectColor_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_selectColor.Click

        Dim MyDialog As New ColorDialog()
        ' Keeps the user from selecting a custom color.
        MyDialog.AllowFullOpen = False
        ' Allows the user to get help. (The default is false.)
        MyDialog.ShowHelp = True
        ' Sets the initial color select to the current text color,
        MyDialog.Color = lvCategories.SelectedItems(0).SubItems(2).ForeColor

        ' Update the text box color if the user clicks OK 
        If (MyDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            lvCategories.SelectedItems(0).ForeColor = MyDialog.Color
            lvCategories.SelectedItems(0).SubItems(3).Text = MyDialog.Color.ToArgb
            lvAllergens.ForeColor = MyDialog.Color
            'Save the colour to the database
            Me.Update_AllergenCategory()
            'Refresh selected panel allergens list
            Me.Load_panelmembers()
        End If

    End Sub


    Private Sub ToolStripButton_newCat_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_newCat.Click

        Dim Response As String = ""
        Dim Unique As Boolean = False

        Do
            Response = InputBox("Enter a category name", "Create new category")
            If Response <> "" Then
                Unique = cSpt.Is_CategoryNameUnique(Response)
                If Not Unique Then
                    MsgBox("Category name already exists, must be unique", vbOKOnly, "Create new category")
                End If
            End If
        Loop Until Response = "" Or Unique

        If Unique Then   'Save the panel
            Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptCategory
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).categoryid)) = 0
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).categoryname)) = "'" & Response & "'"
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).displaycolour)) = "'" & Color.Black.Name & "'"
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).enabled)) = "'" & True & "'"
            cSpt.Insert_Category(dicP)
            Me.Load_AllergenCategories()
            If lvCategories.Items.Count > 0 Then lvCategories.FindItemWithText(Response).Selected = True
        End If

    End Sub

    Private Sub ToolStripButton_editAllergen_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_editAllergen.Click

        If lvAllergens.SelectedItems.Count = 1 Then
            Dim allergenID As Integer = CType(lvAllergens.Items(lvAllergens.SelectedIndices(0)).SubItems(1).Text, Integer)
            Dim allergenName As String = lvAllergens.Items(lvAllergens.SelectedIndices(0)).SubItems(2).Text
            Dim f As New form_spt_allergen(allergenID, allergenName)
            If f.ShowDialog = DialogResult.OK Then
                Me.Load_AllergensFromDatabase(f.categoryID)
            End If
        Else
            MsgBox("Please selected an allergen to edit", vbOKOnly, "Edit allergen")
        End If
    End Sub

    Private Sub ToolStripButton_newAllergen_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_newAllergen.Click

        Dim f As New form_spt_allergen(0, , lvCategories.SelectedItems(0).SubItems(2).Text)
        If f.ShowDialog = DialogResult.OK Then
            Me.Load_AllergensFromDatabase(f.categoryID)
        End If

    End Sub

    Private Sub ToolStripButton_deleteAllergen_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_deleteAllergen.Click

        Select Case lvAllergens.SelectedItems.Count
            Case 0
                MsgBox("Please select an allergen to delete.", vbOKOnly, "Delete allergen")
            Case 1
                Dim Response As Integer = 0
                Response = MsgBox("You are about to permanently delete this allergen. Continue?", vbYesNo, "Delete database record")
                If Response = vbYes Then
                    Try
                        Dim allergenID As Integer = CType(lvAllergens.Items(lvAllergens.SelectedIndices(0)).SubItems(1).Text, Integer)
                        If cSpt.Delete_Allergen(allergenID) Then
                            Me.Load_AllergensFromDatabase(CType(lvCategories.Items(lvCategories.SelectedIndices(0)).SubItems(1).Text, Integer))
                            If lvAllergens.Items.Count > 0 Then lvAllergens.Items(0).Selected = True
                        Else
                            MsgBox("Problem deleting allergen.", vbOKOnly, "Delete allergen")
                        End If
                    Catch
                        MsgBox("Problem allergen - aborted.", vbOKOnly, "Delete allergen")
                    End Try
                End If
        End Select


    End Sub

    Private Sub ToolStripButton_deleteCat_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_deleteCat.Click

        Select Case lvCategories.SelectedItems.Count
            Case 0
                MsgBox("Please select a category to delete.", vbOKOnly, "Delete category")
            Case 1
                Dim Response As Integer = 0
                Response = MsgBox("You are about to permanently delete this category. Continue?", vbYesNo, "Delete database record")
                If Response = vbYes Then
                    Try
                        Dim categoryID As Integer = CType(lvCategories.Items(lvCategories.SelectedIndices(0)).SubItems(1).Text, Integer)
                        'Can't delete if there are related allergen records
                        If cSpt.Is_CategoryEmpty(categoryID) Then
                            If cSpt.Delete_Category(categoryID) Then
                                Me.Load_AllergenCategories()
                                If lvCategories.Items.Count > 0 Then lvCategories.Items(0).Selected = True
                            Else
                                MsgBox("Problem deleting category.", vbOKOnly, "Delete category")
                            End If
                        Else
                            MsgBox("Can't delete category while it contains allergens.", vbOKOnly, "Delete category")
                        End If
                    Catch
                        MsgBox("Problem deleting category - aborted.", vbOKOnly, "Delete category")
                    End Try
                End If
        End Select

    End Sub

    Private Sub ToolStripButton_editCat_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_editCat.Click

        Dim Response As String = ""
        Dim Unique As Boolean = False

        If lvCategories.SelectedIndices.Count = 1 Then
            Do
                Response = InputBox("Edit category name", "Edit category", lvCategories.SelectedItems(0).SubItems(2).Text)
                If Response <> "" Then
                    Unique = cSpt.Is_CategoryNameUnique(Response)
                    If Not Unique Then
                        MsgBox("Category name already exists, must be unique", vbOKOnly, "Edit category")
                    End If
                End If
            Loop Until Response = "" Or Unique

            If Unique Then   'Save the category
                Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptCategory
                dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).categoryid)) = lvCategories.SelectedItems(0).SubItems(1).Text
                dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).categoryname)) = "'" & Response & "'"
                dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).displaycolour)) = "'" & lvCategories.SelectedItems(0).SubItems(3).Text & "'"
                dicP(cMyRoutines.GetPropertyName(Function() (New AllergenCategoryData).enabled)) = "'" & True & "'"
                cSpt.Update_Category(dicP)
                Dim CurrentIndex As Integer = lvCategories.SelectedIndices(0)
                Me.Load_AllergenCategories()
                lvCategories.Items(CurrentIndex).Selected = True
            End If
        Else
            MsgBox("Please select a category to edit", vbOKOnly, "Edit category")
        End If

    End Sub

    Private Sub ToolStripButton_add_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_add.Click
        'Add an allergen to panel

        Dim i As Integer = 0
        Dim Found As Boolean = False

        'Panel selected and allergen selected?
        If lvPanels.SelectedItems.Count = 1 Then
            If lvAllergens.SelectedItems.Count = 1 Then
                'Allergen unique to panel?
                For i = 0 To lvPanelAllergens.Items.Count - 1
                    If lvAllergens.SelectedItems(0).SubItems(2).Text = lvPanelAllergens.Items(i).SubItems(2).Text Then
                        Found = True
                        Exit For
                    End If
                Next
                If Found Then
                    MsgBox("Allergen already in panel.", vbOKOnly, "Add allergen")
                Else
                    'Add allergen to panel above current selected allergen 
                    Dim lvi = New ListViewItem
                    lvi.SubItems.Add(lvAllergens.SelectedItems(0).SubItems(1).Text)
                    lvi.SubItems.Add(lvAllergens.SelectedItems(0).SubItems(2).Text)
                    lvi.SubItems.Add(lvAllergens.ForeColor.ToArgb)
                    lvi.SubItems.Add(0)  'New memberID

                    lvi.ForeColor = lvAllergens.ForeColor

                    If lvPanelAllergens.SelectedItems.Count = 1 Then
                        lvPanelAllergens.Items.Insert(lvPanelAllergens.SelectedIndices(0), lvi)
                    Else
                        lvPanelAllergens.Items.Add(lvi)
                    End If

                    'Update database with new panel member and update all existing (likely have new order)
                    Me.Save_panelmembers()

                    'Re-load panel
                    Me.Load_panelmembers()

                End If
            Else
                MsgBox("Please select an allergen", vbOKOnly, "Add allergen")
            End If
        Else
            MsgBox("Please select a panel", vbOKOnly, "Add allergen")
        End If


    End Sub

    Private Sub ToolStripButton_remove_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_remove.Click

        If lvPanelAllergens.SelectedItems.Count = 1 Then
            Dim memberID As Integer = CType(lvPanelAllergens.Items(lvPanelAllergens.SelectedIndices(0)).SubItems(4).Text, Integer)
            Dim CurrentIndex As Integer = lvPanelAllergens.SelectedIndices(0)

            'Remove from listview
            lvPanelAllergens.Items.RemoveAt(CurrentIndex)

            'Remove from database
            If cSpt.Delete_PanelMember(memberID) Then
                If CurrentIndex >= 0 And CurrentIndex <= lvPanelAllergens.Items.Count - 1 Then
                    lvPanelAllergens.Items(CurrentIndex).Selected = True
                Else
                    If lvPanelAllergens.Items.Count > 0 Then lvPanelAllergens.Items(0).Selected = True
                End If
            Else
                MsgBox("Problem deleting allergen from panel.", vbOKOnly, "Remove allergen")
            End If

            'Save the panel to update allergen order
            Me.Save_panelmembers()
        End If

    End Sub

    Private Sub ToolStripButton_upAllergen_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton_upAllergen.Click

        If lvPanelAllergens.SelectedItems.Count = 1 Then
            Dim CurrentIndex As Integer = lvPanelAllergens.SelectedItems(0).Index
            If CurrentIndex > 0 Then
                Dim lvi As ListViewItem = lvPanelAllergens.Items(lvPanelAllergens.SelectedItems(0).Index)
                lvPanelAllergens.Items.RemoveAt(lvPanelAllergens.SelectedItems(0).Index)
                lvPanelAllergens.Items.Insert(CurrentIndex - 1, lvi)
                lvPanelAllergens.Items(CurrentIndex - 1).Selected = True
                'Re-save panel to database
                Me.Save_panelmembers()
            End If
        End If

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click

        If lvPanelAllergens.SelectedItems.Count = 1 Then
            Dim CurrentIndex As Integer = lvPanelAllergens.SelectedItems(0).Index
            If CurrentIndex < lvPanelAllergens.Items.Count - 1 Then
                Dim lvi As ListViewItem = lvPanelAllergens.Items(lvPanelAllergens.SelectedItems(0).Index)
                lvPanelAllergens.Items.RemoveAt(lvPanelAllergens.SelectedItems(0).Index)
                lvPanelAllergens.Items.Insert(CurrentIndex + 1, lvi)
                lvPanelAllergens.Items(CurrentIndex + 1).Selected = True
                'Re-save panel to database
                Me.Save_panelmembers()
            End If
        End If

    End Sub


    Private Sub ToolStripButton_closeform_Click(sender As Object, e As System.EventArgs) Handles ToolStripButton_closeform.Click

        DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()

    End Sub


    Private Sub ToolStripButton_CopyPanelToTest2_Click(sender As Object, e As System.EventArgs) Handles ToolStripButton_CopyPanelToTest2.Click

        If lvPanels.SelectedItems.Count = 1 Then
            DialogResult = Windows.Forms.DialogResult.OK
        Else
            MsgBox("Select a panel to copy.", vbOKOnly, "Copy panel to SPT")
        End If

    End Sub

    Private Sub ToolStripButton_CopyPanelToTest_Click(sender As Object, e As System.EventArgs) Handles ToolStripButton_CopyPanelToTest.Click

        If lvPanels.SelectedItems.Count = 1 Then
            DialogResult = Windows.Forms.DialogResult.OK
        Else
            MsgBox("Select a panel to copy.", vbOKOnly, "Copy panel to SPT")
        End If

    End Sub
End Class