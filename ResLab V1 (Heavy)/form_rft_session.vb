Imports ResLab_V1_Heavy.cDatabaseInfo


Public Class form_rft_session

    Private _formloading As Boolean
    Private _patientID As Long
    Public _sessionID As Nullable(Of Long)
    Private _testID As Nullable(Of Long)
    Private _test_tbl As eTables
    Private _callingform As Form
    Private _isNewSession As Boolean
    Private _isNewTest As Boolean
    Private _defaulttestdate As String
    Public _cancelButton_enable As Boolean = False

    Public Sub New(patientID As Long, testID As Nullable(Of Long), test_tbl As eTables, callingForm As Form, cancelButton_enable As Boolean)
        'testID=0 for new test, otherwise testID refers to an exisitng test

        InitializeComponent()
        _patientID = patientID
        _callingform = callingForm
        _testID = testID
        _test_tbl = test_tbl
        _cancelButton_enable = cancelButton_enable

        If testID = 0 Then _isNewTest = True Else _isNewTest = False

        Select Case Me._isNewTest
            Case True : Me._sessionID = Math.Abs(cRfts.Get_SessionToUse_newtest(_patientID, Date.Today))
            Case False : Me._sessionID = cRfts.Get_SessionToUse_existingtest(_testID, _test_tbl)
        End Select

        Select Case Me._sessionID
            Case 0
                _isNewSession = True
                _defaulttestdate = Format(Date.Today, "dd/MM/yyyy")
            Case vbNull
                _isNewSession = True
                _defaulttestdate = ""   'Format(Date.Today, "dd/MM/yyyy")
            Case Else
                _isNewSession = False
                _defaulttestdate = ""
        End Select


        'Load combo options
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboSmoke_hx, "Smoking_history", _isNewSession)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboAdmission, "AdmissionStatus", _isNewSession)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboReqMO_name, "ReferringMO name", _isNewSession)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboReportCopyTo, "ReportCopyTo", _isNewSession)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboBilledTo, "Billed to", _isNewSession)
        cMyRoutines.Combo_LoadItemsFromPrefs(cmboBillingMO_name, "BillingMO", _isNewSession)
        cmboReqMO_HealthService.Items.AddRange(cHs.get_list_Healthservices(, , True).ToArray)

        'Get and load relevant demographic data to screen fields (readonly)
        Dim dicD As Dictionary(Of String, String) = cPt.get_demographics(_patientID)
        Dim flds As New class_DemographicFields
        If IsDate(dicD(flds.DOB)) Then txtDOB.Text = dicD(flds.DOB) Else txtDOB.Text = "__/__/____"
        txtGender.Text = dicD(flds.Gender)
        txtRace.Text = dicD(flds.Race_forRfts)
        txtMedicareNo.Text = dicD(flds.Medicare_No)
        txtMedicareExpires.Text = dicD(flds.Medicare_expirydate)

        'Get and load session data to screen fields
        If Me._isNewSession Then
            txtTestDate.Text = Me._defaulttestdate
        Else
            Me.load_textboxes(_sessionID)
            SetVisibleCore(False)
        End If

    End Sub

    Private Sub form_rft_session_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        btnCancel.Enabled = _cancelButton_enable
    End Sub

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)

        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If
        MyBase.SetVisibleCore(value)

    End Sub

    Private Sub form_rft_session_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub load_textboxes(sessionid As Long)

        Dim dicS As Dictionary(Of String, String) = cRfts.get_rft_session(sessionid)
        Dim f As New class_Rft_RoutineAndSessionFields

        Me._defaulttestdate = Format(CDate(dicS(f.TestDate)), "dd/MM/yyyy") 'store original testdate in case edited

        txtTestDate.Text = Format(CDate(dicS(f.TestDate)), "dd/MM/yyyy")
        txtHt.Text = Format(Val(dicS(f.Height)), "###.0")
        txtWt.Text = Format(Val(dicS(f.Weight)), "###.0")
        cmboSmoke_hx.Text = dicS(f.Smoke_hx)
        txtSmoke_cigsperday.Text = dicS(f.Smoke_cigsperday)
        txtSmoke_packyears.Text = dicS(f.Smoke_packyears)
        txtSmoke_yrs.Text = dicS(f.Smoke_yearssmoked)
        cmboAdmission.Text = dicS(f.AdmissionStatus)
        txtClinicalNote.Text = dicS(f.Req_clinicalnotes)
        If IsDate(dicS(f.Req_date)) Then txtRequestDate.Text = Format(CDate(dicS(f.Req_date)), "dd/MM/yyyy")
        cmboReqMO_name.Text = dicS(f.Req_name)
        cmboReqMO_address.Text = dicS(f.Req_address)
        txtReqMO_email.Text = dicS(f.Req_email)
        txtReqMO_fax.Text = dicS(f.Req_fax)
        txtReqMO_phone.Text = dicS(f.Req_phone)
        txtReqMO_pn.Text = dicS(f.Req_providernumber)
        cmboReqMO_HealthService.Text = dicS(f.Req_healthservice)
        cmboReportCopyTo.Text = dicS(f.Report_copyto)
        cmboBilledTo.Text = dicS(f.Billing_billedto)
        cmboBillingMO_name.Text = dicS(f.Billing_billingMO)
        cmboBillingMO_pn.Text = dicS(f.Billing_billingMOproviderno)

        txtTestDate.Focus()

    End Sub

    Private Sub clear_textboxes()

        txtHt.Text = ""
        txtWt.Text = ""
        cmboSmoke_hx.Text = ""
        txtSmoke_cigsperday.Text = ""
        txtSmoke_packyears.Text = ""
        txtSmoke_yrs.Text = ""
        cmboAdmission.Text = ""
        txtClinicalNote.Text = ""
        txtRequestDate.Text = "__/__/____"
        cmboReqMO_name.Text = ""
        cmboReqMO_address.Text = ""
        txtReqMO_email.Text = ""
        txtReqMO_fax.Text = ""
        txtReqMO_phone.Text = ""
        txtReqMO_pn.Text = ""
        cmboReqMO_HealthService.Text = ""
        cmboReportCopyTo.Text = ""
        cmboBilledTo.Text = ""
        cmboBillingMO_name.Text = ""
        cmboBillingMO_pn.Text = ""

    End Sub

    Private Sub btnContinue_Click(sender As System.Object, e As System.EventArgs) Handles btnContinue.Click


        'Do any data validation checks
        'If test date has been changed, check again if there is an existing session to attach to.
        '  If yes, use it, otherwise create new session

        Dim response As Integer = 0
        Dim Msg As String = ""

        If IsDate(txtTestDate.Text) Then
            If txtTestDate.Text <> _defaulttestdate Then     'Do a text comparison - default can be empty

                Dim sID As Long = cRfts.Exists_RftSessionForDate(_patientID, CDate(txtTestDate.Text))
                Select Case sID
                    Case 0  'new session
                        Msg = "No existing session for " & txtTestDate.Text & "." & vbCrLf & "Create new session and continue?"
                        response = MsgBox(Msg, vbYesNo, "Warning")
                        Select Case response
                            Case vbYes
                                response = MsgBox("Clear loaded session data?", vbYesNo, "Test session")
                                If response = vbYes Then
                                    'Clear text boxes
                                    Me.clear_textboxes()
                                End If
                                _sessionID = 0
                                _isNewSession = True
                                txtHt.Focus()
                            Case vbNo
                                'Set test date back to original
                                txtTestDate.Text = _defaulttestdate
                                txtTestDate.Focus()
                        End Select
                    Case Else
                        Msg = "An existing session for " & txtTestDate.Text & " found." & vbCrLf & "Attach test to this session and continue?"
                        response = MsgBox(Msg, vbYesNo, "Warning")
                        Select Case response
                            Case vbYes
                                MsgBox("Session data for " & txtTestDate.Text & " retrieved.", vbOKOnly, "Test session")
                                'Load session data
                                Me.load_textboxes(sID)
                                _sessionID = sID
                                _isNewSession = False
                                Me._defaulttestdate = txtTestDate.Text
                                DialogResult = Windows.Forms.DialogResult.OK
                                Me.Visible = False
                            Case vbNo
                                'Set test date back to original
                                txtTestDate.Text = _defaulttestdate
                                txtTestDate.Focus()
                        End Select
                End Select
            Else
                DialogResult = Windows.Forms.DialogResult.OK
                Me.Visible = False
            End If
        Else
            MsgBox("Please enter a valid test date.", vbOKOnly, "Test session")
            txtTestDate.Focus()
        End If

    End Sub

    Private Sub Label6_DoubleClick(sender As Object, e As System.EventArgs)

        If Not IsDate(txtRequestDate.Text) Then txtRequestDate.Text = Now

    End Sub

    Private Sub txtSmoke_cigsperday_TextChanged(sender As System.Object, e As System.EventArgs)

        txtSmoke_packyears.Text = cMyRoutines.calc_PackYears(txtSmoke_cigsperday.Text, txtSmoke_yrs.Text)

    End Sub

    Private Sub txtSmoke_yrs_TextChanged(sender As System.Object, e As System.EventArgs)

        txtSmoke_packyears.Text = cMyRoutines.calc_PackYears(txtSmoke_cigsperday.Text, txtSmoke_yrs.Text)

    End Sub


    Private Sub cmboAdmission_DropDownClosed(sender As Object, e As System.EventArgs)
        cmboSmoke_hx.Focus()
    End Sub

    Private Sub cmboSmoke_hx_DropDownClosed(sender As Object, e As System.EventArgs)
        txtSmoke_cigsperday.Focus()
    End Sub

    Private Sub cmboReqMO_name_DropDownClosed(sender As Object, e As System.EventArgs)
        txtReqMO_pn.Focus()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnPick_Click(sender As System.Object, e As System.EventArgs) Handles btnPick.Click

    End Sub

End Class