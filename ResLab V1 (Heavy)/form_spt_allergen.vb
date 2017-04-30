Imports System.Windows.Forms

Public Class form_spt_allergen

    Dim _allergenID As Integer
    Dim _allergenName As String
    Dim _categoryName As String

    Public ReadOnly Property categoryID As Integer
        Get
            categoryID = cmboCategory.SelectedValue
        End Get
    End Property

    Public Sub New(allergenID As Integer, Optional allergenName As String = "", Optional categoryName As String = "")
        'allergenID passed=0 if new allergen

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _allergenID = allergenID
        _allergenName = allergenName
        _categoryName = categoryName

    End Sub


    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()

    End Sub

    Private Sub form_spt_allergen_new_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load


        Dim c As List(Of AllergenCategoryData) = cSpt.Get_AllergenCategories(True)
        If Not IsNothing(c) Then
            cmboCategory.Items.Clear()
            cmboCategory.DataSource = c
            cmboCategory.DisplayMember = "categoryname"
            cmboCategory.ValueMember = "categoryid"
        End If

        Select Case _allergenID
            Case 0
                Me.Text = "New allergen"
                lblInfo.Text = "Add a new allergen to the database ....."
                cmboCategory.SelectedIndex = cmboCategory.FindStringExact(_categoryName)
            Case Else
                Me.Text = "Edit allergen"
                lblInfo.Text = "Edit allergen name and/or category ....."
                txtAllergenName.Text = _allergenName
                Dim categoryname As String = cSpt.Get_CategoryForAllergen(_allergenID).categoryname
                Dim i As Integer = cmboCategory.FindStringExact(categoryname)
                cmboCategory.SelectedIndex = i
        End Select
        txtAllergenName.Select()

    End Sub

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click

        If cSpt.Is_AllergenNameUnique(txtAllergenName.Text, _allergenID) Then
            Dim dicP As Dictionary(Of String, String) = cMyRoutines.MakeEmpty_dic_sptAllergen
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenData).allergenid)) = _allergenID
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenData).allergenname)) = "'" & txtAllergenName.Text & "'"
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenData).categoryid)) = cmboCategory.SelectedValue
            dicP(cMyRoutines.GetPropertyName(Function() (New AllergenData).enabled)) = "'True'"
            If _allergenID = 0 Then cSpt.Insert_Allergen(dicP) Else cSpt.Update_Allergen(dicP)
        Else
            MsgBox("Allergen name already exists, must be unique", vbOKOnly, "Allergen database")
            txtAllergenName.Focus()
        End If
        Me.Close()

    End Sub

End Class
