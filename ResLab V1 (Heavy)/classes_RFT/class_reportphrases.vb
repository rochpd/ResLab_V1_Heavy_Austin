Imports System.Text

Public Enum eRftSeverityTypes
    Spiro
    TLCO
    Hyperinflation
    Restriction
End Enum

Public Enum eAutoreport_testgroups
    DoNotAutoreport
    RoutineRft
    Provocation
    Cpx
    Spt
    Walk_tests
    Hast
End Enum

Public Class class_reportphrases

    Public Function get_phrases(groupID As Integer) As Dictionary(Of Integer, String)

        Dim d As New Dictionary(Of Integer, String)
        Dim sql As New StringBuilder

        sql.Append("SELECT prefs_ID, fielditem FROM prefs_fielditems ")
        sql.Append("WHERE field_id = " & groupID)

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                With Ds.Tables(0)
                    For Each record As DataRow In .Rows
                        d.Add(record("prefs_ID"), record("fielditem"))
                    Next
                End With
            End If
        End If

        Return d

    End Function

    Public Function get_phrasegroups(tests As eAutoreport_testgroups) As Dictionary(Of Integer, String)

        Dim d As New Dictionary(Of Integer, String)
        Dim sql As New StringBuilder
        Dim GroupName As String = ""
        Dim s As New StringBuilder

        Select Case tests
            Case eAutoreport_testgroups.Cpx
                s.Append("fieldname='Reportphrases_CPX'")
            Case eAutoreport_testgroups.DoNotAutoreport
            Case eAutoreport_testgroups.Hast
                s.Append("fieldname='Reportphrases_HAST'")
            Case eAutoreport_testgroups.Provocation
                s.Append("fieldname='Reportphrases_Spirometry'")
                s.Append(" OR fieldname LIKE 'Reportphrases_Provocation_%'")
            Case eAutoreport_testgroups.RoutineRft
                s.Append("fieldname='Reportphrases_Spirometry' OR ")
                s.Append("fieldname='Reportphrases_TLCO' OR ")
                s.Append("fieldname='Reportphrases_LV' OR ")
                s.Append("fieldname='Reportphrases_MRP' OR ")
                s.Append("fieldname='Reportphrases_FeNO' OR ")
                s.Append("fieldname='Reportphrases_SpO2'")
            Case eAutoreport_testgroups.Walk_tests
                s.Append("fieldname='Reportphrases_WalkTests'")
            Case eAutoreport_testgroups.Spt
                s.Append("fieldname='Reportphrases_SPT'")
        End Select

        sql.Append("SELECT * FROM prefs_fields WHERE " & s.ToString)
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If IsNothing(Ds) Then
            Return Nothing
        Else
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                With Ds.Tables(0)
                    For Each record As DataRow In .Rows
                        GroupName = Replace(record("fieldname"), "Reportphrases_", "")
                        d.Add(record("field_id"), GroupName)
                    Next
                End With
            End If
        End If

        Return d

    End Function

    Public Function Save_phrase(groupID As Integer, phraseID As Integer, txt As String) As Integer

        Dim ID As Integer = cPrefs.Save_FieldOption(groupID, phraseID, txt)
        Return ID

    End Function

    Public Sub Delete_phrase(phraseID As Integer)

        cPrefs.Delete_FieldOption(phraseID)

    End Sub

    Public Function AutoInterpret_rft(testGroup As eAutoreport_testgroups, ByVal Demo As Pred_demo, ByVal R As Dictionary(Of String, String)) As String
        'runs the auto interpreting algorithm
        'Dic R contains RFT results

        Dim AutoReport As String = ""
        Dim f As New class_Rft_RoutineAndSessionFields

        Dim p As New Dictionary(Of String, String)     'ParameterID|StatTypeID, result
        p = cPred.Get_PredValues(Demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)

        Dim kcoRange, tlcRange, spo2Range As String
        Dim fev1LLN, vcLLN, fvcLLN, ferLLN, TLCOlln, vaLLN, kcoRangeLo, kcoRangeHi, tlcRangeLo, tlcRangeHi, rvLLN, rvtlcULN, phlln, phuln, po2lln, pco2lln, pco2uln, aapo2uln, spo2lln, spo2uln, miplln, meplln As Single
        Dim pFEV1, pFVC, pVC, pFER, pTLCO, pVA, pTLC As Single

        If p.ContainsKey("FEV1|LLN") Then fev1LLN = p("FEV1|LLN")
        If p.ContainsKey("FEV1|MPV") Then pFEV1 = p("FEV1|MPV")
        If p.ContainsKey("FVC|LLN") Then fvcLLN = p("FVC|LLN")
        If p.ContainsKey("FVC|MPV") Then pFVC = p("FVC|MPV")
        If p.ContainsKey("VC|LLN") Then vcLLN = p("VC|LLN")
        If p.ContainsKey("VC|MPV") Then pVC = p("VC|MPV")
        If p.ContainsKey("FER|LLN") Then ferLLN = p("FER|LLN")
        If p.ContainsKey("FER|MPV") Then pFER = p("FER|MPV")
        If p.ContainsKey("TLCO|LLN") Then TLCOlln = p("TLCO|LLN")
        If p.ContainsKey("TLCO|MPV") Then pTLCO = p("TLCO|MPV")
        If p.ContainsKey("VA|LLN") Then vaLLN = p("VA|LLN")
        If p.ContainsKey("VA|MPV") Then pVA = p("VA|MPV")
        If p.ContainsKey("KCO|Range") Then
            kcoRange = p("KCO|Range")
            kcoRangeLo = Left(kcoRange, InStr(kcoRange, "[TO]") - 1)
            kcoRangeHi = Mid(kcoRange, InStr(kcoRange, "[TO]") + 5)
        End If
        If p.ContainsKey("TLC|Range") Then
            tlcRange = p("TLC|Range")
            tlcRangeLo = Left(tlcRange, InStr(tlcRange, "[TO]") - 1)
            tlcRangeHi = Mid(tlcRange, InStr(tlcRange, "[TO]") + 5)
        End If
        If p.ContainsKey("TLC|MPV") Then pTLC = p("TLC|MPV")
        If p.ContainsKey("RV|ULN") Then rvLLN = p("RV|ULN")
        If p.ContainsKey("RV/TLC|ULN") Then rvtlcULN = p("RV/TLC|ULN")
        If p.ContainsKey("PH|LLN") Then phlln = p("PH|LLN")
        If p.ContainsKey("PH|ULN") Then phuln = p("PH|ULN")
        If p.ContainsKey("PO2|LLN") Then po2lln = p("PO2|LLN")
        If p.ContainsKey("PCO2|LLN") Then pco2lln = p("PCO2|LLN")
        If p.ContainsKey("PCO2|ULN") Then pco2uln = p("PCO2|ULN")
        If p.ContainsKey("AAPO2|ULN") Then aapo2uln = p("AAPO2|ULN")
        If p.ContainsKey("SpO2|Range") Then
            spo2Range = p("SpO2|Range")
            spo2lln = Left(spo2Range, InStr(spo2Range, "[TO]") - 1)
            spo2uln = Mid(spo2Range, InStr(spo2Range, "[TO]") + 5)
        End If
        If p.ContainsKey("MIP|LLN") Then miplln = p("MIP|LLN")
        If p.ContainsKey("MEP|LLN") Then meplln = p("MEP|LLN")


        'First do ventilatory function------------------------------------------------------------------
        'If only postBD spiro, use these to report
        If testGroup = eAutoreport_testgroups.RoutineRft Or testGroup = eAutoreport_testgroups.Provocation Then

            Dim EffEeeVee As String = ""
            Dim VeeCee As String = ""
            Dim postBDSpiro As Boolean = False
            Dim preBDSpiro As Boolean = False
            Dim prepostBDSpiro As Boolean = False

            If R(f.R_bl_Fev1) <> "" And (R(f.R_bl_Fvc) <> "" Or R(f.R_bl_Vc) <> "") Then
                preBDSpiro = True
                EffEeeVee = R(f.R_bl_Fev1)
                If Val(R(f.R_bl_Fvc)) > Val(R(f.R_bl_Vc)) Then VeeCee = R(f.R_bl_Fvc) Else VeeCee = R(f.R_bl_Vc)
            End If
            If R.ContainsKey(f.R_post_Fev1) Then    'Prov has no post keys
                If R(f.R_post_Fev1) <> "" And (R(f.R_post_Fvc) <> "" Or R(f.R_post_Vc) <> "") Then
                    postBDSpiro = True
                    EffEeeVee = R(f.R_post_Fev1)
                    If Val(R(f.R_post_Fvc)) > Val(R(f.R_post_Vc)) Then VeeCee = R(f.R_post_Fvc) Else VeeCee = R(f.R_post_Vc)
                End If
            End If
            prepostBDSpiro = preBDSpiro And postBDSpiro

                If Val(EffEeeVee) <> 0 Then
                    If Val(Format(100 * EffEeeVee / VeeCee, "###")) > Val(ferLLN) Then
                        If Val(VeeCee) >= Val(vcLLN) Then
                            AutoReport = "Ventilatory function is within normal limits"
                        Else
                            AutoReport = "Spirometry suggests a " & Me.GetSeverity(eRftSeverityTypes.Spiro, Format(100 * EffEeeVee / Format(pFEV1, "#.##"), "###")) & " restrictive ventilatory defect"
                        End If
                    Else
                        If Val(VeeCee) >= Val(vcLLN) Then
                            AutoReport = "There is a " & Me.GetSeverity(eRftSeverityTypes.Spiro, Format(100 * EffEeeVee / Format(pFEV1, "#.##"), "###")) & " obstructive ventilatory defect"
                        Else
                            AutoReport = "There is a " & Me.GetSeverity(eRftSeverityTypes.Spiro, Format(100 * EffEeeVee / Format(pFEV1, "#.##"), "###")) & " mixed obstructive/restrictive ventilatory defect"
                        End If
                    End If

                    'now for bd response---------------------------------------------------------------------------------
                    If testGroup = eAutoreport_testgroups.RoutineRft Then
                        If prepostBDSpiro Then
                            Dim deltfev1 As Single, deltfvc As Single, deltvc As Single
                            Dim FEV1response As Boolean = False, FVCresponse As Boolean = False, VCresponse As Boolean = False
                            Dim bdcomment As String = ""

                            If R(f.R_bl_Fev1) <> "" And R(f.R_post_Fev1) <> "" Then deltfev1 = Val(R(f.R_post_Fev1)) - Val(R(f.R_bl_Fev1))
                            If R(f.R_bl_Fvc) <> "" And R(f.R_post_Fvc) <> "" Then deltfvc = Val(R(f.R_post_Fvc)) - Val(R(f.R_bl_Fvc))
                            If R(f.R_bl_Vc) <> "" And R(f.R_post_Vc) <> "" Then deltvc = Val(R(f.R_post_Vc)) - Val(R(f.R_bl_Vc))
                            If deltfev1 >= 0.2 Then FEV1response = True
                            If deltfvc >= 0.2 Then FVCresponse = True
                            If deltvc >= 0.2 Then VCresponse = True
                            If Val(R(f.R_bl_Fev1)) <> 0 And FEV1response Then If deltfev1 / R(f.R_bl_Fev1) >= 0.12 Then FEV1response = True Else FEV1response = False
                            If Val(R(f.R_bl_Fvc)) <> 0 And FVCresponse Then If deltfvc / R(f.R_bl_Fvc) >= 0.12 Then FVCresponse = True Else FVCresponse = False
                            If Val(R(f.R_bl_Vc)) <> 0 And VCresponse Then If deltvc / R(f.R_bl_Vc) >= 0.12 Then VCresponse = True Else VCresponse = False
                            If FEV1response Or FVCresponse Or VCresponse Then bdcomment = " and a significant bronchodilator response." Else bdcomment = " with no significant bronchodilator response on this occasion."
                            If (R(f.R_bl_Fev1) = "" And R(f.R_bl_Vc) = "" And R(f.R_bl_Fvc) = "") Or (R(f.R_post_Fev1) = "" And R(f.R_post_Vc) = "" And R(f.R_post_Fvc) = "") Then bdcomment = ""
                            If Val(R(f.R_bl_Fev1)) < 1 And deltfev1 >= 0.1 Then bdcomment = " together with a small but possibly clinically useful bronchodilator response."
                            If Val(R(f.R_bl_Fev1)) <> 0 Then If deltfev1 / R(f.R_bl_Fev1) > 0.2 Then bdcomment = " together with a highly significant bronchodilator response that is suggestive of asthma."

                            'if normal baseline and +ve response then need special comment
                            If AutoReport = "Ventilatory function is within normal limits" And (FEV1response Or FVCresponse Or VCresponse) Then
                                AutoReport = "Although baseline spirometric values are within normal limits there is a significant bronchodilator response indicating some airflow obstruction."
                            Else
                                AutoReport = AutoReport & bdcomment
                            End If
                        End If
                    End If
                    If Right(AutoReport, 1) <> "." Then AutoReport = AutoReport & ". " Else AutoReport = AutoReport & " "

                End If
            End If

            'Then TLCO----------------------------------------------------------------------------
            If testGroup = eAutoreport_testgroups.RoutineRft Then
            Dim TeeLCO, KceeO As Single
            Dim AddedBit, TLCOcomment As String

            If Val(R(f.R_Bl_Tlco)) <> 0 Then
                If Val(R(f.R_Bl_Hb)) = 0 Then
                    TeeLCO = R(f.R_Bl_Tlco)
                    KceeO = R(f.R_Bl_Kco)
                    AddedBit = ", uncorrected for haemoglobin,"
                Else
                    Dim HbFactor As Single = cMyRoutines.calc_HbFac(R(f.R_Bl_Hb), Demo.DOB, Demo.Gender, Demo.TestDate)
                    TeeLCO = R(f.R_Bl_Tlco) * HbFactor
                    KceeO = R(f.R_Bl_Kco) * HbFactor
                    AddedBit = ", corrected for haemoglobin,"
                End If
                If Val(TeeLCO) <= Val(TLCOlln) Then
                    TLCOcomment = "Carbon monoxide transfer factor" & AddedBit
                    TLCOcomment = TLCOcomment & " is " & Me.GetSeverity(eRftSeverityTypes.TLCO, Format(100 * TeeLCO / Format(pTLCO, "##.#"))) & " reduced "
                    If Val(R(f.R_Bl_Va)) > Val(vaLLN) Then
                        TLCOcomment = TLCOcomment & "indicating lung parenchymal and/or pulmonary vascular dysfunction. "
                    Else
                        If KceeO > Val(kcoRangeHi) Then
                            TLCOcomment = TLCOcomment & "due to loss of alveolar volume rather than lung parenchymal or pulmonary vascular dysfunction. "
                        Else
                            TLCOcomment = TLCOcomment & "at least in part due to loss of alveolar volume, but also indicating lung parenchymal and/or pulmonary vascular dysfunction. "
                        End If
                    End If
                    AutoReport = AutoReport & TLCOcomment
                Else
                    AutoReport = AutoReport & "Carbon monoxide transfer factor is within normal limits. "
                End If
            End If
            If AutoReport = "Ventilatory function is within normal limits. Carbon monoxide transfer factor is within normal limits. " Then
                AutoReport = "Ventilatory function and carbon monoxide transfer factor are within normal limits. "
            End If


            'Do lung volumes-----------------------------------------------------------------------------------------------------------
            Dim LVcomment As String

            If R(f.R_Bl_Tlc) <> "" And R(f.R_Bl_Frc) <> "" And R(f.R_Bl_Rv) <> "" Then
                If Val(R(f.R_Bl_Tlc)) < Val(tlcRangeLo) Then
                    If Val(R(f.R_Bl_Rv)) < Val(rvLLN) Then
                        LVcomment = "Plethysmographic lung volumes are " & GetSeverity(eRftSeverityTypes.Restriction, 100 * Val(R(f.R_Bl_Tlc)) / Val(pTLC)) & " reduced confirming a restrictive pulmonary defect. "
                    Else
                        LVcomment = "Total lung capacity is " & GetSeverity(eRftSeverityTypes.Restriction, 100 * Val(R(f.R_Bl_Tlc)) / Val(pTLC)) & " reduced confirming a restrictive pulmonary defect. "
                    End If
                Else
                    If Val(R(f.R_Bl_Tlc)) > Val(tlcRangeHi) Then
                        If 100 * Val(R(f.R_Bl_Rv)) / Val(R(f.R_Bl_Tlc)) > rvtlcULN Then
                            LVcomment = "Plethysmographic lung volumes indicate " & GetSeverity(eRftSeverityTypes.Hyperinflation, 100 * Val(R(f.R_Bl_Tlc)) / Val(pTLC)) & " hyperinflation with gas trapping. "
                        Else
                            LVcomment = "Plethysmographic lung volumes indicate " & GetSeverity(eRftSeverityTypes.Hyperinflation, 100 * Val(R(f.R_Bl_Tlc)) / Val(pTLC)) & " hyperinflation but without an increased RV/TLC ratio. "
                        End If
                    Else
                        If Val(R(f.R_Bl_Tlc)) < Val(tlcRangeLo) + 0.25 Then
                            LVcomment = "Total lung capacity is towards the lower limit of the normal range raising the possibility of a restrictive pulmonary defect. "
                        Else
                            If Val(100 * R(f.R_Bl_Rv) / R(f.R_Bl_Tlc)) > rvtlcULN Then
                                LVcomment = "Whilst TLC is within the normal range, the raised RV/TLC ratio indicates some gas trapping. "
                            Else
                                LVcomment = "Plethysmographic lung volumes are within normal limits. "
                            End If
                        End If
                    End If
                End If
                AutoReport = AutoReport & LVcomment
            End If

            'Do ABGs------------------------------------------------------------------------------------------------------------
            Dim abgComment As String

            If R(f.R_abg1_pao2) <> "" And R(f.R_abg1_paco2) <> "" And R(f.R_abg1_ph) <> "" Then
                'Do PO2
                If Val(R(f.R_abg1_pao2)) < po2lln Then
                    If R(f.R_abg1_fio2) = "Room air" Then
                        If Val(cMyRoutines.Calculate_AaPO2(R(f.R_abg1_pao2), R(f.R_abg1_paco2))) > Val(aapo2uln) Then
                            abgComment = "Arterial blood gases reveal hypoxaemia with a widened (A-a)PO2 gradient"
                        Else
                            abgComment = "Arterial blood gases reveal hypoxaemia but with a normal (A-a)PO2 gradient"
                        End If
                    Else
                        abgComment = "Arterial blood gases reveal hypoxaemia"
                    End If
                Else
                    If R(f.R_abg1_fio2) = "Room air" Then
                        If Val(cMyRoutines.Calculate_AaPO2(R(f.R_abg1_pao2), R(f.R_abg1_paco2))) > Val(aapo2uln) Then
                            abgComment = "Arterial blood gases reveal normoxaemia but with a widened (A-a)PO2 gradient"
                        Else
                            abgComment = "Arterial blood gases reveal normoxaemia with a normal (A-a)PO2 gradient"
                        End If
                    Else
                        abgComment = "Arterial blood gases reveal normoxaemia"
                    End If
                End If
                'Do PCO2
                If Val(R(f.R_abg1_paco2)) < Val(pco2lln) Then
                    abgComment = abgComment & ", alveolar hyperventilation"
                ElseIf Val(R(f.R_abg1_paco2)) > Val(pco2uln) Then
                    abgComment = abgComment & ", alveolar hypoventilation"
                End If
                'Do pH
                If Val(R(f.R_abg1_ph)) > phuln Then
                    abgComment = abgComment & " and alkalosis. "
                ElseIf Val(R(f.R_abg1_ph)) < phlln Then
                    abgComment = abgComment & " and acidosis. "
                Else
                    abgComment = abgComment & " and a normal pH. "
                End If
                If Right(abgComment, 2) <> ". " Then
                    abgComment = abgComment & ". "
                End If
                AutoReport = AutoReport & abgComment
            End If

            'Do SpO2-----------------------------------------------------------------------------------------------
            Dim spo2Comment As String

            If R(f.R_SpO2_1) <> "" Then
                If Val(R(f.R_SpO2_1)) < spo2lln Then
                    spo2Comment = "Oxygen saturation by pulse oximetry is reduced. "
                Else
                    spo2Comment = "Oxygen saturation by pulse oximetry is within the normal range. "
                End If
                AutoReport = AutoReport & spo2Comment
            End If

            'Do MRPs---------------------------------------------------------------------------------------------------
            Dim mrpComment As String = ""

            If R(f.R_Bl_Mip) <> "" And Val(R(f.R_Bl_Mip)) > miplln And R(f.R_Bl_Mep) <> "" And Val(R(f.R_Bl_Mep)) > meplln Then
                mrpComment = "Maximal respiratory pressures are within the normal range."
            ElseIf R(f.R_Bl_Mip) <> "" And Val(R(f.R_Bl_Mip)) <= miplln And R(f.R_Bl_Mep) <> "" And Val(R(f.R_Bl_Mep)) <= meplln Then
                mrpComment = "Maximal respiratory pressures are reduced indicating respiratory muscle and diaphragmatic weakness."
            ElseIf R(f.R_Bl_Mip) <> "" And Val(R(f.R_Bl_Mip)) <= miplln Then
                mrpComment = "Maximal inspiratory pressure is reduced indicating diaphragmatic weakness."
            ElseIf R(f.R_Bl_Mep) <> "" And Val(R(f.R_Bl_Mep)) <= meplln Then
                mrpComment = "Maximal expiratory pressure is reduced indicating respiratory muscle weakness."
            End If
            AutoReport = AutoReport & mrpComment
        End If

        Return AutoReport

    End Function

    Public Function GetSeverity(ByVal Type As eRftSeverityTypes, ByVal Value As Single) As String

        Select Case Type
            Case eRftSeverityTypes.Spiro 'FEV1ppn
                Select Case Value
                    Case Is >= 70 : Return "mild"
                    Case Is >= 60 : Return "moderate"
                    Case Is >= 50 : Return "moderately severe"
                    Case Is >= 35 : Return "severe"
                    Case Is > 0 : Return "very severe"
                    Case Else : Return ""
                End Select
            Case eRftSeverityTypes.TLCO  'TLCOppn
                Select Case Value
                    Case Is > 60 : Return "mildly"
                    Case Is > 40 : Return "moderately"
                    Case Is > 0 : Return "severely"
                    Case Else : Return ""
                End Select
            Case eRftSeverityTypes.Hyperinflation
                Select Case Value   'TLCppn
                    Case Is > 150 : Return "marked"
                    Case Is > 135 : Return "moderate"
                    Case Is > 120 : Return "mild"
                    Case Else : Return ""
                End Select
            Case eRftSeverityTypes.Restriction
                Select Case Value    'TLCppn
                    Case Is > 75 : Return "mildly"
                    Case Is > 60 : Return "moderately"
                    Case Is > 0 : Return "severely"
                    Case Else : Return ""
                End Select
            Case Else
                Return ""
        End Select

    End Function

End Class
