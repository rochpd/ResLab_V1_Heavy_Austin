Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class form_UpdateReportStatus

    Public Sub New(ByVal CurrentReportStatus As String)

        InitializeComponent()
        lblCurrentReportStatus.Text = CurrentReportStatus

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub Label82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label82.Click

    End Sub

    Private Sub Label80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label80.Click

    End Sub



    Private Sub form_UpdateReportStatus_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        cMyRoutines.Combo_LoadItemsFromList(cmboUpdate, eTables.List_ReportStatuses, True)
        cmboUpdate.SelectedIndex = cmboUpdate.FindString(lblCurrentReportStatus.Text) + 1
    End Sub
End Class