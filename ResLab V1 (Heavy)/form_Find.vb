

Public Class form_Find

    Public flagLoading As Boolean = False

    Public Sub New(ByVal Surname As String, ByVal FirstName As String, ByVal UR As String)

        ' This call is required by the designer.
        InitializeComponent()

        Initialise()
        txtSurname.Text = Surname
        txtFirstname.Text = FirstName
        txtUR.Text = UR

        If Surname <> "" Or FirstName <> "" Or UR <> "" Then Find_Patient()

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Select_Patient(ByVal PatientID As Long)

        Dim f As New form_Demographics(Me)
        f.Tag = PatientID
        f.Show()
        Me.Close()

    End Sub

    Private Sub Find_Patient()

        Dim PatientID() As Long
        Dim tempDate As String

        Me.Cursor = Cursors.WaitCursor

        If txtUR.Text <> "" Then
            PatientID = cPt.Get_PatientIDFromUR(txtUR.Text)
        Else
            PatientID = cPt.Find_PtByNameEtc(txtSurname.Text, cmboSurnameBy.Text, txtFirstname.Text, cmboFirstnameBy.Text, txtDOB.Text, cmboGender.Text)
        End If


        Select Case UBound(PatientID)
            Case 0
                MsgBox("No matches found")
                If txtUR.Text <> "" Then txtUR.Focus() Else txtSurname.Focus()
            Case 1
                Dim f As New form_Demographics(Me)
                f.Tag = PatientID(1)
                f.Show()
            Case Else
                If Me.Visible = False Then Me.Visible = True

                'More than one match, load grid
                Dim d As New class_DemographicFields
                grdPts.Visible = True
                grdPts.Rows.Clear()
                For i As Integer = 1 To UBound(PatientID)
                    Dim dicD As Dictionary(Of String, String) = cPt.Get_Demographics(PatientID(i))
                    If cMyRoutines.IsRealDate(dicD(d.DOB)) Then tempDate = Format(dicD(d.DOB), "Short Date") Else tempDate = ""
                    grdPts.Rows.Add(dicD(d.PatientID), i, dicD(d.Surname), dicD(d.Firstname), dicD(d.UR), tempDate, dicD(d.Gender))
                Next
        End Select

        Me.Cursor = Cursors.Arrow

    End Sub

    Private Sub form_Find_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If Me.Tag = "Close" Then Me.Close()
        flagLoading = False

    End Sub

    Private Sub form_Find_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        flagLoading = True
        Initialise()

    End Sub

    Private Sub Initialise()

        lblUR.Text = gURlabel
        grdPts.Columns("UR").HeaderText = gURlabel

        'Load combo options
        Dim items As Array = System.Enum.GetNames(GetType(eFindBy))
        For Each item As String In items
            cmboFirstnameBy.Items.Add(item)
            cmboSurnameBy.Items.Add(item)
        Next
        cmboFirstnameBy.SelectedIndex = cmboFirstnameBy.FindString(eFindBy.First_part.ToString)
        cmboSurnameBy.SelectedIndex = cmboSurnameBy.FindString(eFindBy.Exact.ToString)

        cMyRoutines.Combo_LoadItems(cmboGender, cPred.Get_RefItems(RefItems.Genders))
        cmboGender.Items.RemoveAt(cmboGender.FindStringExact("Male, Female"))

        grdPts.Visible = False

    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click

        Me.Find_Patient()

    End Sub

    Private Sub btnSelect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelect.Click

        If grdPts.SelectedCells.Count > 0 Then Me.Select_Patient(grdPts.Rows(grdPts.SelectedCells(0).RowIndex).Cells(0).Value)

    End Sub

    Private Sub grdPts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPts.DoubleClick

        Me.Select_Patient(grdPts.Rows(grdPts.SelectedCells(0).RowIndex).Cells(0).Value)

    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click

        txtUR.Text = ""
        txtSurname.Text = ""
        txtFirstname.Text = ""
        txtDOB.Text = "__/__/____"
        cmboGender.SelectedIndex = -1
        grdPts.Rows.Clear()
        grdPts.Visible = False

    End Sub


    Private Sub txtUR_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUR.KeyPress

        If Asc(e.KeyChar) = (Keys.Return) And txtUR.Text <> "" Then
            Me.Find_Patient()
        End If

    End Sub

End Class