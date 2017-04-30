

Public Class form_DuplicatePatients

    Dim Formloading As Boolean = False
    Dim _CallingForm As Form

    Public Sub New(PatientIDs() As Long, CallingForm As Form)
        InitializeComponent()
        _CallingForm = CallingForm
        LoadGrid(PatientIDs)
    End Sub

    Private Sub LoadGrid(IDs As Long())

        Dim i As Integer = 0
        Dim d As New Dictionary(Of String, String)
        Dim demo As New class_DemographicFields
        Dim address As String = ""

        For i = 0 To UBound(IDs)
            d = cPt.Get_Demographics(IDs(i))
            address = d(demo.Address_1) & ", " & d(demo.Suburb) & ", " & d(demo.PostCode)
            grd.Rows.Add(d(demo.PatientID), d(demo.Surname) & ", " & d(demo.Firstname), d(demo.DOB), d(demo.Gender), address, d(demo.Phone_mobile), d(demo.Phone_home))
        Next
        grd.Rows.Add("0", "<Add new patient record>", "", "", "", "", "", "")
        grd.ClearSelection()

    End Sub


    Private Sub btnSelect_Click(sender As Object, e As System.EventArgs)

        If grd.SelectedRows.Count = 0 Then
            MsgBox("Please select a row")
        ElseIf grd.SelectedRows.Count = 1 Then
            _CallingForm.Tag = grd.SelectedRows(0).Cells(0).Value   'pass back selected patientid
            Me.Close()
        End If

    End Sub

    Private Sub form_DuplicatePatients_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated

        If Formloading Then
            Formloading = False
            Dim Msg As String = "One or more existing patient records match the details entered." & vbCrLf
            Msg = Msg & "Please select the matching patient or add new record."
            MsgBox(Msg, vbOKOnly, "Duplicates found")
        End If

    End Sub

    Private Sub form_DuplicatePatients_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Formloading = True

    End Sub


    Private Sub btnContinue_Click(sender As Object, e As System.EventArgs) Handles btnContinue.Click

        _CallingForm.Tag = grd.SelectedCells(0).Value

    End Sub

End Class