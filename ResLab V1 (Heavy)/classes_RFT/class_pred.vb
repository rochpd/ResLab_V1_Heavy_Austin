Option Strict Off
Option Explicit Off
Imports Microsoft.VisualBasic
Imports System
Imports System.Math
Imports System.Text
Imports System.Data
Imports System.Data.Common
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports ResLab_V1_Heavy.class_DataAccessLayer
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_Pred

    Public Enum eLoadNormalsMode
        UseCurrentPrefs
        UseSourcesSavedWithTest
        UseSourcesInUseAtTestDate
    End Enum
    Public Enum eParameterRangeCheck
        NotSupplied = 0
        OutOfRange = 1
        OK = 2
    End Enum
    Public Enum eOOR    'pred parameter value out of range
        high
        low
        inrange
        param_NotInEquation
    End Enum

    Public Structure LMS_Coefficients
        Dim q0 As Double, q1 As Double
        Dim p0 As Double, p1 As Double, p2 As Double, p3 As Double, p4 As Double, p5 As Double
        Dim a0 As Double, a1 As Double, a2 As Double, a3 As Double, a4 As Double, a5 As Double, a6 As Double
    End Structure
    Public Structure LMS_Equations
        Dim M_equation As String
        Dim L_equation As String
        Dim S_equation As String
        Dim LLN As String
        Dim Zscore As String
    End Structure
    Public Structure LMS_Splines
        Dim LSpline As Double
        Dim MSpline As Double
        Dim SSpline As Double
    End Structure
    Public Structure ParameterInfo_UseToCalc
        Dim Age_ForCalc As Single
        Dim Age_InEquation As class_Pred.eOOR
        Dim Ht_ForCalc As Single
        Dim Ht_InEquation As class_Pred.eOOR
        Dim Wt_ForCalc As Single
        Dim Wt_InEquation As class_Pred.eOOR
        Dim GenderID_ForCalc As Long
        Dim EthnicityID_ForCalc As Long
        Dim Ethnicity_ForCalc As String
    End Structure
    Public Structure RangeCheckResult
        Dim ParamValueToUse As Single
        Dim RangeResult As class_Pred.eOOR
    End Structure

    Public Function Get_Pred(ByVal aKey As String, ByVal d As Dictionary(Of String, String), ByVal DecPlaces As Integer) As String

        Dim Places As String = StrDup(DecPlaces, "0")
        Dim s As String = ""
        Dim rangeLo As Single, rangeHi As Single

        If Not IsNothing(d) Then
            If d.ContainsKey(aKey) Then
                If Val(d(aKey)) > 0 Then
                    If InStr(aKey, "LLN") > 0 Then
                        s = "> " & Format(Val(d(aKey)), "0." & Places)
                    ElseIf InStr(aKey, "ULN") > 0 Then
                        s = "< " & Format(Val(d(aKey)), "0." & Places)
                    ElseIf InStr(aKey, "Range") > 0 Then
                        rangeHi = Val(Mid(d(aKey), InStr(d(aKey), "[TO]") + 5))
                        rangeLo = Left(d(aKey), InStr(d(aKey), "[TO]") - 1)
                        If rangeHi > 0 And rangeLo > 0 Then
                            s = Format(rangeLo, "0." & Places) & " - " & Format(rangeHi, "0." & Places)
                        Else
                            s = ""
                        End If
                    End If
                Else
                    s = ""
                End If
            Else
                'MsgBox("Predicted equations covering this testdate were not found for " & Left(aKey, InStr(aKey, "|") - 1), vbOKOnly, "Warning")
                s = "---"
            End If
        Else
            s = "---"
        End If

        Return s

    End Function

    Public Function Format_Pred(ByVal Result As Double, ByVal DecPlaces As Integer) As String

        Dim fstring As String = ""

        If Result = 0 Then
            Return ""
        Else
            Select Case DecPlaces
                Case 0 : fstring = " (0)"
                Case 1 : fstring = " (0.0)"
                Case 2 : fstring = " (0.00)"
                Case 3 : fstring = " (0.000)"
            End Select
            Return Format(Result, fstring)
        End If

    End Function

    Public Function Calc_Predicted(ByVal Demo As Pred_demo, ByVal ParamID As Integer, ByVal SourceID As Integer, ByVal Eq As String) As String

        Dim p As New Mathos.Parser.MathParser
        Dim result As Double = 0
        Dim Eqs As New Dictionary(Of Integer, String)

        If Eq.Contains("_NA") Then
            Return ""
        Else
            'Create the parser variables
            p.LocalVariables.Add("Age", Demo.Age)
            p.LocalVariables.Add("Ht", Demo.Htcm)
            p.LocalVariables.Add("Wt", Demo.Wtkg)

            If Eq.Contains("_MPV") Then
                Eqs = Me.Get_EquationsForParameter(ParamID, SourceID, Demo)
                Dim mpv_eq As String = Eqs(StatType.MPV)
                Dim _mpv As Double = p.Parse(mpv_eq)
            End If

            Try
                result = p.Parse(Eq)
                Return Me.Format_Pred(result, Me.Get_DecPlacesForParamID(ParamID))
            Catch ex As Exception
                MsgBox("Error: " & ex.Message)
                Return ""
            End Try
        End If

    End Function

    Public Function Get_ParamForID(ByVal ParamID As Integer) As String

        Dim sql As String = "SELECT description FROM Pred_Ref_Parameters WHERE code=" & ParamID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count = 1 Then
            Return Ds.Tables(0).Rows(0).Item("Parameter")
        Else
            Return ""
        End If

    End Function

    Public Function Get_UnitsForParamID(ByVal ParamID As Integer) As String

        Dim sql As String = "SELECT Units FROM Pred_Ref_Parameters WHERE code=" & ParamID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count = 1 Then
            Return Ds.Tables(0).Rows(0).Item("Units")
        Else
            Return ""
        End If

    End Function

    Public Function Get_DecPlacesForParamID(ByVal ParamID As Integer) As Integer

        Dim sql As String = "SELECT DecimalPlaces FROM Pred_Ref_Parameters WHERE code=" & ParamID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count = 1 Then
            Return Ds.Tables(0).Rows(0).Item("DecimalPlaces")
        Else
            Return ""
        End If

    End Function

    Public Function Match_Equation(ByVal TestID As String, ByVal SourceID As String, ByVal ParameterID As String, ByVal AgeGroupID As String, ByVal GenderID As String, ByVal EthnicityID As String, Optional ByVal StatTypeID As String = "") As Dictionary(Of String, String)()

        Dim sql As String = ""
        Dim row As DataRow

        sql = "SELECT Pred_equations.* FROM Pred_equations "
        sql = sql & "WHERE SourceID=" & SourceID
        sql = sql & " AND TestID=" & TestID
        sql = sql & " AND ParameterID=" & ParameterID
        sql = sql & " AND AgeGroupID=" & AgeGroupID
        sql = sql & " AND GenderID=" & GenderID
        sql = sql & " AND EthnicityID=" & EthnicityID
        If StatTypeID <> "" Then sql = sql & " AND StatTypeID=" & StatTypeID

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        Dim dic(-1) As Dictionary(Of String, String)

        With Ds.Tables(0)
            If .Rows.Count = 0 Then
                Return Nothing
            Else
                For Each row In .Rows
                    ReDim Preserve dic(UBound(dic) + 1)
                    dic(UBound(dic)) = New Dictionary(Of String, String)
                    dic(UBound(dic)).Add("age_lower", row("age_lower") & "")
                    dic(UBound(dic)).Add("age_upper", row("age_upper") & "")
                    dic(UBound(dic)).Add("age_clipmethod", row("age_clipmethod") & "")
                    dic(UBound(dic)).Add("age_clipmethodID", row("age_clipmethodID") & "")
                    dic(UBound(dic)).Add("ht_lower", row("ht_lower") & "")
                    dic(UBound(dic)).Add("ht_upper", row("ht_upper") & "")
                    dic(UBound(dic)).Add("ht_clipmethod", row("ht_clipmethod") & "")
                    dic(UBound(dic)).Add("ht_clipmethodID", row("ht_clipmethodID") & "")
                    dic(UBound(dic)).Add("wt_lower", row("wt_lower") & "")
                    dic(UBound(dic)).Add("wt_upper", row("wt_upper") & "")
                    dic(UBound(dic)).Add("wt_clipmethod", row("wt_clipmethod") & "")
                    dic(UBound(dic)).Add("wt_clipmethodID", row("wt_clipmethodID") & "")
                    dic(UBound(dic)).Add("equationid", row("EquationID") & "")
                    dic(UBound(dic)).Add("stattypeid", row("StatTypeID") & "")
                    dic(UBound(dic)).Add("stattype", row("StatType") & "")
                Next
                Return dic
            End If
        End With

    End Function

    Public Function Get_EquationIDsForSourceID(ByRef SourceID As Integer) As Integer()

        Dim sql As String = ""
        Dim EqIDs() As Integer

        ReDim EqIDs(0)

        sql = "SELECT Pred_equations.EquationID FROM Pred_sourceXparameter INNER JOIN Pred_equations "
        sql = sql & "ON Pred_sourceXparameter.sXpID = Pred_equations.sXpID "
        sql = sql & "WHERE Pred_sourceXparameter.SourceID)=" & CStr(SourceID)

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        For Each r As DataRow In Ds.Tables(0).Rows
            ReDim Preserve EqIDs(EqIDs.GetUpperBound(0) + 1)
            EqIDs(EqIDs.GetUpperBound(0)) = r.Item("EquationID")
        Next

        Return EqIDs

    End Function

    Public Function Get_EquationInfoForID(ByRef EquationID As Long) As Dictionary(Of String, String)

        Dim sql As String = ""
        Dim d As New Dictionary(Of String, String)

        sql = "SELECT * FROM Pred_equations WHERE EquationID=" & EquationID

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Not IsNothing(Ds) Then
            Dim r As DataRow = Ds.Tables(0).Rows(0)
            For Each c As DataColumn In Ds.Tables(0).Columns
                If IsDBNull(r(c.ColumnName)) Then
                    d.Add(c.ColumnName, "")
                Else
                    d.Add(c.ColumnName, CStr(r(c.ColumnName)))
                End If
            Next
            Return d
        Else
            Return Nothing
        End If

    End Function



    Public Sub LoadPPN(ByVal txt_ppn As MaskedTextBox, ByVal txt_result As MaskedTextBox, ByVal aKey As String, ByVal d As Dictionary(Of String, String))

        If d.ContainsKey(aKey) AndAlso Val(txt_result.Text) > 0 Then
            txt_ppn.Text = Format(100 * Val(txt_result.Text) / Val(d(aKey)), "(###)")
        Else
            txt_ppn.Text = ""
        End If

    End Sub

    Public Sub LoadLbl(ByVal lbl As MaskedTextBox, ByVal aKey As String, ByVal d As Dictionary(Of String, String), ByVal DecPlaces As Integer)

        lbl.Text = Me.Get_Pred(aKey, d, DecPlaces)

    End Sub

    Public Function Get_EquationsForParameter(ByVal ParamID As Integer, ByVal SourceID As Integer, ByVal Demo As Pred_demo) As Dictionary(Of Integer, String)  'StatID, equation

        Dim sql As New StringBuilder
        Dim Eqs As New Dictionary(Of Integer, String)

        sql.Clear()
        sql.Append("SELECT Pred_equations.* ")
        sql.Append("FROM (Pred_Ref_tests INNER JOIN Pred_sourceXtest ON Pred_Ref_tests.code = Pred_sourceXtest.TestID) INNER JOIN Pred_equations ON Pred_sourceXtest.SourceID = Pred_equations.SourceID ")
        sql.Append("WHERE Pred_equations.SourceID=" & SourceID)
        sql.Append(" AND Pred_equations.ParameterID=" & ParamID)
        sql.Append(" AND Pred_equations.GenderID=" & Demo.GenderID)
        sql.Append(" AND Pred_equations.Age_lower <=" & Demo.Age)
        sql.Append(" AND Pred_equations.Age_upper >=" & Demo.Age)
        sql.Append(" AND Pred_equations.EthnicityID=" & Demo.EthnicityID)
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        For Each r As DataRow In Ds.Tables(0).Rows
            Eqs.Add(r.Item("StatTypeID"), r.Item("Equation").ToString)
        Next

        Return Eqs

    End Function

    Public Function Get_AgeGroupsForSourceParameter(ByVal SourceID As Integer, ByVal ParameterID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT Pred_Ref_agegroups.description, Pred_Ref_agegroups.code FROM (Pred_Ref_sources INNER JOIN Pred_equations ON ")
        sql.Append("Pred_Ref_sources.code = Pred_equations.SourceID) INNER JOIN Pred_Ref_agegroups ON ")
        sql.Append("Pred_equations.AgeGroup = Pred_Ref_agegroups.description ")
        sql.Append("WHERE(((Pred_Ref_sources.code) = " & SourceID & ") And ((Pred_equations.ParameterID) = " & ParameterID & ")) ")
        sql.Append("GROUP BY Pred_Ref_agegroups.description, Pred_Ref_agegroups.code;")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            For Each r As DataRow In Ds.Tables(0).Rows
                d.Add(r.Item("code"), r.Item("description"))
            Next
        End If

        Return d

    End Function

    Public Function Get_EthnicitiesFor_Source_Parameter_Agegroup(ByVal SourceID As Integer, ByVal ParameterID As Integer, ByVal AgeGroupID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT pred_ref_ethnicities.description, pred_ref_ethnicities.code FROM ")
        sql.Append("pred_ref_ethnicities INNER JOIN Pred_equations ON  pred_ref_ethnicities.code = Pred_equations.EthnicityID ")
        sql.Append("WHERE (Pred_equations.SourceID = " & SourceID & ") AND (Pred_equations.ParameterID = " & ParameterID & ") AND (Pred_equations.AgeGroupID = " & AgeGroupID & ")")
        sql.Append(" GROUP BY pred_ref_ethnicities.description, pred_ref_ethnicities.code")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        If Not IsNothing(Ds) Then
            For Each r As DataRow In Ds.Tables(0).Rows
                d.Add(r.Item("code"), r.Item("description"))
            Next
        End If

        Return d

    End Function

    Public Function Calc_PredForParameter(ByVal SourceID As Integer, ByVal ParameterID As Integer, ByVal demo As Pred_demo) As Dictionary(Of Integer, Single)



        MsgBox("Under construction, apologies.")
        Return Nothing
        Exit Function


        'Dim sql As New StringBuilder
        'Dim d As New Dictionary(Of Integer, Single)    'StatTypeID, result
        'Dim ds As DataSet
        'Dim result, MPV
        'Dim Equation As String

        'sql.Clear()
        'sql.Append("SELECT pred_equations.* FROM pred_equations ")
        'sql.Append("WHERE SourceID = " & SourceID & " AND ParameterID = " & ParameterID)
        'sql.Append(" AND genderid = " & demo.GenderID)
        'sql.Append(" AND ethnicityid = " & demo.EthnicityID)
        ''sql.Append(" AND " & p.Age_ForCalc & " >= Age_lower AND " & p.Age_ForCalc & " <= Age_upper")

        'ds = cDAL.Get_DataAsDataset(sql.ToString)
        'If IsNothing(ds) Then
        '    Return Nothing
        'Else
        '    Select Case ds.Tables(0).Rows.Count
        '        Case 0      'No equations found
        '        Case Else
        '            'Calculate preds 
        '            For Each row As DataRow In ds.Tables(0).Rows
        '                Select Case row.Item("EquationType")
        '                    Case "RMS"
        '                        Equation = row.Item("Equation")
        '                        If InStr(Equation, "_MPV") > 0 Then
        '                            'Find MPV equation in the dataset
        '                            Dim dsCopy As DataSet = ds
        '                            For Each r As DataRow In dsCopy.Tables(0).Rows
        '                                If r.Item("StatType") = "MPV" Then
        '                                    MPV = cPred.ParseEquation(p, r.Item("Equation"))
        '                                    Equation = Replace(Equation, "_MPV", MPV.ToString)
        '                                    Exit For
        '                                End If
        '                            Next
        '                            dsCopy = Nothing
        '                        End If
        '                        result = cPred.ParseEquation(p, Equation)
        '                        d.Add(row.Item("StatTypeID"), result)
        '                    Case "LMS"






        '                End Select
        '            Next
        '    End Select
        '    ds = Nothing
        'End If

        'Return d

    End Function

    Private Function Check_parameter_in_range(equation As String, param_name As String, param_value As Single, hi As Single, lo As Single, clipmethod As String) As RangeCheckResult
        'Used by Check_equation_limits

        Dim Rangecheck As RangeCheckResult
        Dim ValueToUse As Single
        Dim Range As eOOR

        'Is this parameter in this equation
        If InStr(equation, param_name) > 0 Then
            If param_value < lo Then
                Range = eOOR.low
                Select Case clipmethod
                    Case "Use min/max value" : ValueToUse = lo
                    Case "Do not calculate predicteds" : ValueToUse = Nothing
                    Case "Extrapolate" : ValueToUse = param_value
                End Select

            ElseIf param_value > hi Then
                Range = eOOR.high
                Select Case clipmethod
                    Case "Use min/max value" : ValueToUse = hi
                    Case "Do not calculate predicteds" : ValueToUse = Nothing
                    Case "Extrapolate" : ValueToUse = param_value
                End Select
            Else
                Range = eOOR.inrange
                ValueToUse = param_value
            End If
        Else
            'Parameter not used in this equation
            Range = eOOR.param_NotInEquation
            ValueToUse = param_value    'needed elsewhere
        End If

        Rangecheck.ParamValueToUse = ValueToUse
        Rangecheck.RangeResult = Range
        Return Rangecheck

    End Function

    Public Function Check_equation_limits(EqInfo As DataRow, demo As Pred_demo) As ParameterInfo_UseToCalc
        'Used by Get_PredValues
        ' The equation itself is needed here to see what variables are in the equation

        Dim ParamInfo As New ParameterInfo_UseToCalc
        Dim Rangecheck As RangeCheckResult
        Dim d As Dictionary(Of String, String)

        d = Me.Get_EquationInfoForID(EqInfo.Item("equationID_mpv"))

        Rangecheck = Me.Check_parameter_in_range(d("Equation"), "Age", demo.Age, EqInfo.Item("age_upper"), EqInfo.Item("age_lower"), EqInfo.Item("age_clipmethod"))
        ParamInfo.Age_ForCalc = Rangecheck.ParamValueToUse
        ParamInfo.Age_InEquation = Rangecheck.RangeResult

        Rangecheck = Me.Check_parameter_in_range(d("Equation"), "Ht", demo.Htcm, EqInfo.Item("Ht_upper"), EqInfo.Item("Ht_lower"), EqInfo.Item("Ht_clipmethod"))
        ParamInfo.Ht_ForCalc = Rangecheck.ParamValueToUse
        ParamInfo.Ht_InEquation = Rangecheck.RangeResult

        Rangecheck = Me.Check_parameter_in_range(d("Equation"), "Wt", demo.Wtkg, EqInfo.Item("Wt_upper"), EqInfo.Item("Wt_lower"), EqInfo.Item("Wt_clipmethod"))
        ParamInfo.Wt_ForCalc = Rangecheck.ParamValueToUse
        ParamInfo.Wt_InEquation = Rangecheck.RangeResult

        ParamInfo.GenderID_ForCalc = demo.GenderID
        ParamInfo.EthnicityID_ForCalc = demo.EthnicityID
        ParamInfo.Ethnicity_ForCalc = demo.Ethnicity

        Return ParamInfo

    End Function

    Public Function Get_PredValues(ByVal Demo As Pred_demo, ByVal Method As eLoadNormalsMode) As Dictionary(Of String, String)

        Dim sql As New StringBuilder
        Dim sql_daterange As String = "", sql_ethnicity As String = "", sql_gender As String = ""
        Dim equation_mpv As String = "", equation_normalrange As String = "", equation As String = ""
        Dim equationID_mpv As Long, equationID_normalrange As Long

        Dim ds As DataSet, ds1 As DataSet, ds2 As DataSet
        Dim test As DataRow, param As DataRow, r As DataRow
        Dim d As New Dictionary(Of String, String)    'ParameterID|StatID, result
        Dim result As Single = 0, mpv As Single, normalrange As String
        Dim ps() As ParameterInfo_UseToCalc, p As ParameterInfo_UseToCalc = Nothing
        Dim rFound As Boolean
        Dim i As Integer

        'Construct sql clauses 
        sql_gender = "(preferences_pred.GenderID = " & Demo.GenderID & " OR preferences_pred.GenderID = 3) "
        sql_ethnicity = "((preferences_pred.EthnicityID = " & Demo.EthnicityID & " OR preferences_pred.EthnicityID = 8) OR (EthnicityCorrectionType <>'' AND EthnicityCorrectionType IS NOT NULL)) "
        Select Case Method
            Case class_Pred.eLoadNormalsMode.UseCurrentPrefs : sql_daterange = " EndDate IS NULL "
            Case class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate : sql_daterange = " CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) >= StartDate AND CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) <= (CASE WHEN EndDate IS NULL THEN CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) ELSE enddate END) "
        End Select

        'Get the tests which have pred equations selected
        sql.Append("SELECT testID FROM preferences_pred WHERE " & sql_daterange & " GROUP BY testID")
        ds = cDAL.Get_DataAsDataset(sql.ToString)

        If (Not IsNothing(ds) AndAlso ds.Tables(0).Rows.Count > 0) Then
            For Each test In ds.Tables(0).Rows
                'Get the parameters with equations for each test
                sql.Clear()
                sql.Append("SELECT parameterID, count(parameterID) AS num FROM preferences_pred WHERE testID=" & test("testID") & " AND " & sql_daterange & " GROUP BY parameterID")
                ds1 = cDAL.Get_DataAsDataset(sql.ToString)
                If (Not IsNothing(ds1) AndAlso ds1.Tables(0).Rows.Count > 0) Then
                    'Get MPV/range equations and calculate value for each parameter/stat in turn - ignore age, ht and wt limits initially
                    For Each param In ds1.Tables(0).Rows
                        'Get and prepare the equation
                        sql.Clear()
                        sql.Append("SELECT preferences_pred.*, pred_equations.EthnicityCorrectionType FROM preferences_pred INNER JOIN pred_equations ON pred_equations.equationID = preferences_pred.equationID_mpv ")
                        sql.Append("WHERE " & sql_gender & " AND " & sql_ethnicity & " AND " & sql_daterange & " AND (preferences_pred.testID=" & test("testID") & ") AND (preferences_pred.parameterID=" & param("parameterID") & ")")
                        ds2 = cDAL.Get_DataAsDataset(sql.ToString)
                        If (Not IsNothing(ds2) AndAlso ds2.Tables(0).Rows.Count > 0) Then
                            Select Case ds2.Tables(0).Rows.Count
                                Case 1   'Simple case of MPV/range equations found for parameter
                                    r = ds2.Tables(0).Rows(0)
                                    equationID_mpv = r("equationID_mpv")
                                    equationID_normalrange = r("equationID_normalrange")
                                    p = Me.Check_equation_limits(r, Demo)
                                Case Is > 1     'Complex case - occurs when a source has separate equations for sections of the age range eg ECCS for TLCO
                                    'First check if age fits inside the range of one of the MPV equations
                                    rFound = False : r = Nothing
                                    ReDim ps(0 To ds2.Tables(0).Rows.Count - 1)
                                    For i = 0 To ds2.Tables(0).Rows.Count - 1
                                        ps(i) = Me.Check_equation_limits(ds2.Tables(0).Rows(i), Demo)
                                        If ps(i).Age_InEquation = eOOR.inrange Then
                                            rFound = True   'Found, great! use it                                         
                                            r = ds2.Tables(0).Rows(i)
                                        End If
                                    Next

                                    'If found in range, great! use it
                                    If rFound Then
                                        equationID_mpv = r("equationID_mpv")
                                        equationID_normalrange = r("equationID_normalrange")
                                        p = Me.Check_equation_limits(r, Demo)
                                    Else
                                        'If out of range, find the relevant equation and ? extrapolate etc
                                        Dim lo As New SortedDictionary(Of Single, Integer)
                                        Dim hi As New SortedDictionary(Of Single, Integer)
                                        For i = 0 To ds2.Tables(0).Rows.Count - 1
                                            lo.Add(ds2.Tables(0).Rows(i)("age_lower"), i)
                                            hi.Add(ds2.Tables(0).Rows(i)("age_upper"), i)
                                        Next
                                        If Demo.Age < lo.First.Key Then
                                            r = ds2.Tables(0).Rows(lo.First.Value)
                                        ElseIf Demo.Age > hi.Last.Key Then
                                            r = ds2.Tables(0).Rows(hi.Last.Value)
                                        End If
                                        equationID_mpv = r("equationID_mpv")
                                        equationID_normalrange = r("equationID_normalrange")
                                        p = Me.Check_equation_limits(r, Demo)
                                    End If
                            End Select

                            'Do MPV - Get the actual equation - need to substitute the calculated MPV rather than the MPV equation so that ethnicty correction can be applied (if selected) just to mpv and not the SEE
                            Dim dEq_mpv As Dictionary(Of String, String)
                            dEq_mpv = Me.Get_EquationInfoForID(equationID_mpv)
                            equation_mpv = dEq_mpv("Equation")
                            Select Case dEq_mpv("EquationType")
                                Case "RMS"
                                    mpv = Me.ParseEquation(Demo, equation_mpv)
                                    mpv = Me.Do_EthnicityCorrection(p.EthnicityID_ForCalc, mpv, dEq_mpv("ParameterID"), dEq_mpv("EthnicityCorrectionType") & "")
                                    d.Add(dEq_mpv("Parameter") & "|" & dEq_mpv("StatType"), mpv.ToString)
                                Case "LMS"
                                    result = Me.Calculate_LMS_NotMTParser(p, dEq_mpv("Parameter"), dEq_mpv("SourceID"), dEq_mpv("StatType"), dEq_mpv("Age_upper"), dEq_mpv("Age_lower"))
                                    d.Add(dEq_mpv("Parameter") & "|" & dEq_mpv("StatType"), result.ToString)
                            End Select

                            'Do normal range
                            Dim dEq_normalrange As Dictionary(Of String, String)
                            dEq_normalrange = Me.Get_EquationInfoForID(equationID_normalrange)
                            equation_normalrange = Me.Do_MPVsubstitution(dEq_normalrange("Equation"), mpv)
                            Select Case dEq_normalrange("EquationType")
                                Case "RMS"
                                    Select Case dEq_normalrange("StatType")
                                        Case "Range"
                                            'Need to isolate and calculate the 2 ends of the range
                                            RangeHi = Mid(equation_normalrange, InStr(equation_normalrange, "[TO]") + 5)
                                            RangeLo = Left(equation_normalrange, InStr(equation_normalrange, "[TO]") - 1)
                                            resultLo = Me.ParseEquation(Demo, RangeLo)
                                            resultHi = Me.ParseEquation(Demo, RangeHi)
                                            normalrange = resultLo.ToString & " [TO] " & resultHi.ToString
                                        Case Else       'LLN, ULN
                                            result = Me.ParseEquation(Demo, equation_normalrange)
                                            normalrange = result.ToString
                                    End Select
                                    d.Add(dEq_normalrange("Parameter") & "|" & dEq_normalrange("StatType"), normalrange)
                                Case "LMS"
                                    result = Me.Calculate_LMS_NotMTParser(p, dEq_normalrange("Parameter"), dEq_normalrange("SourceID"), dEq_normalrange("StatType"), dEq_normalrange("Age_upper"), dEq_normalrange("Age_lower"))
                                    normalrange = result.ToString
                                    d.Add(dEq_normalrange("Parameter") & "|" & dEq_normalrange("StatType"), normalrange)
                            End Select
                        Else
                            'Equation(s) for parameter not found in prefs table
                            r = Nothing
                            equationID_mpv = 0
                            equationID_normalrange = 0
                            p = Nothing
                        End If
                        ds2 = Nothing
                    Next
                End If
                ds1 = Nothing
            Next
        End If

        Return d

    End Function

    Public Function Get_PredValues1(ByVal Demo As Pred_demo, ByVal Method As eLoadNormalsMode) As Dictionary(Of String, String)
        'Returns a dictionary of predicted values (MPV and normal range) 
        'Method' determines which pred sourcesare used
        'Nothing is returned for tests/parameters where equations aren't found.

        Try
            Dim sql As New StringBuilder
            Dim d As New Dictionary(Of String, String)    'ParameterID|StatID, result
            Dim result, resultLo, resultHi
            Dim Equation As String
            Dim RangeLo As String = "", RangeHi As String = ""

            'Build the relevant query to return pred equations to use
            Select Case Method
                Case class_Pred.eLoadNormalsMode.UseCurrentPrefs
                    sql.Clear()
                    sql.Append("SELECT pred_equations.* ")
                    sql.Append("FROM preferences_pred INNER JOIN pred_equations ON preferences_pred.EquationID_MPV = pred_equations.EquationID ")
                    sql.Append("WHERE EndDate IS NULL ")
                    sql.Append("AND (pred_equations.GenderID = " & Demo.GenderID & " OR pred_equations.Gender = 'Male, Female') ")
                    sql.Append("AND ((pred_equations.EthnicityID = " & Demo.EthnicityID & " OR pred_equations.EthnicityID = 8) OR (EthnicityCorrectionType<>'' AND EthnicityCorrectionType IS NOT NULL)) ")
                    'sql.Append("AND " & Demo.Age & " >= CAST(pred_equations.age_lower AS DECIMAL(10,3)) AND " & Demo.Age & " <= CAST(pred_equations.age_upper AS DECIMAL(10,3)) ")
                    sql.Append(" UNION ALL ")
                    sql.Append("SELECT pred_equations.* ")
                    sql.Append("FROM preferences_pred INNER JOIN pred_equations ON preferences_pred.EquationID_NormalRange = pred_equations.EquationID ")
                    sql.Append("WHERE EndDate IS NULL ")
                    sql.Append("AND (pred_equations.GenderID = " & Demo.GenderID & " OR pred_equations.Gender = 'Male, Female') ")
                    sql.Append("AND ((pred_equations.EthnicityID = " & Demo.EthnicityID & " OR pred_equations.EthnicityID = 8) OR (EthnicityCorrectionType<>'' AND EthnicityCorrectionType IS NOT NULL)) ")
                    'sql.Append("AND " & Demo.Age & " >= CAST(pred_equations.age_lower AS DECIMAL(10,3)) AND " & Demo.Age & " <= CAST(pred_equations.age_upper AS DECIMAL(10,3)) ")
                Case class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate
                    sql.Clear()
                    sql.Append("SELECT pred_equations.* ")
                    sql.Append("FROM preferences_pred INNER JOIN pred_equations ON preferences_pred.EquationID_MPV = pred_equations.EquationID ")
                    sql.Append("WHERE CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) >= StartDate AND CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) <= (CASE WHEN EndDate IS NULL THEN CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) ELSE enddate END) ")
                    sql.Append("AND (pred_equations.GenderID = " & Demo.GenderID & " OR pred_equations.Gender = 'Male, Female') ")
                    sql.Append("AND ((pred_equations.EthnicityID = " & Demo.EthnicityID & " OR pred_equations.EthnicityID = 8) OR (EthnicityCorrectionType<>'' AND EthnicityCorrectionType IS NOT NULL)) ")
                    'sql.Append("AND " & Demo.Age & " >= CAST(pred_equations.age_lower AS DECIMAL(10,3)) AND " & Demo.Age & " <= CAST(pred_equations.age_upper AS DECIMAL(10,3)) ")
                    sql.Append(" UNION ALL ")
                    sql.Append("SELECT pred_equations.* ")
                    sql.Append("FROM preferences_pred INNER JOIN pred_equations ON preferences_pred.EquationID_NormalRange = pred_equations.EquationID ")
                    sql.Append("WHERE CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) >= StartDate AND CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) <= (CASE WHEN EndDate IS NULL THEN CAST('" & Format(Demo.TestDate, "yyyy-MM-dd") & "' AS DATE) ELSE enddate END) ")
                    sql.Append("AND (pred_equations.GenderID = " & Demo.GenderID & " OR pred_equations.Gender = 'Male, Female') ")
                    sql.Append("AND ((pred_equations.EthnicityID = " & Demo.EthnicityID & " OR pred_equations.EthnicityID = 8) OR (EthnicityCorrectionType<>'' AND EthnicityCorrectionType IS NOT NULL)) ")
                    'sql.Append("AND " & Demo.Age & " >= CAST(pred_equations.age_lower AS DECIMAL(10,3)) AND " & Demo.Age & " <= CAST(pred_equations.age_upper AS DECIMAL(10,3)) ")

            End Select




            Dim ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
            If Not IsNothing(ds) Then

                Select Case ds.Tables(0).Rows.Count
                    Case 0      'No equations found
                        MsgBox("Predicted equations selections for this patient not found", vbOKOnly, "Warning")
                    Case Else    'Calculate preds for all equations found
                        For Each row As DataRow In ds.Tables(0).Rows

                            Select Case row.Item("EquationType")
                                Case "RMS"
                                    Equation = Do_ParameterSubstitutions1(row.Item("Parameter"), row.Item("Equation"), ds)
                                    If Equation <> Nothing Then   'Do parameter substitutions eg _MPV

                                        'Dim ParameterCheck As eParameterRangeCheck = Me.Do_ParameterOutOfRangeCheck(row, Demo)
                                        'If ParameterCheck = eParameterRangeCheck.OK Or ParameterCheck = eParameterRangeCheck.OutOfRange Then   'Check if all parameter values present and in-range eg Ht, Age etc

                                        'Limit weight if ROCA TLCO preds - 94.7kg for males and 86.6kg for females
                                        'Dim demo_copy As Pred_demo = Demo
                                        'If row.Item("SourceID") = 4 And row.Item("TestID") = 2 Then
                                        '    demo_copy.Wtkg = Demo.Wtkg
                                        '    If Demo.GenderID = 1 And Demo.Wtkg > 94.7 Then demo_copy.Wtkg = 94.7
                                        '    If Demo.GenderID = 2 And Demo.Wtkg > 86.6 Then demo_copy.Wtkg = 86.6
                                        'End If

                                        Select Case row.Item("StatType")
                                            Case "Range"
                                                'Need to isolate and calculate the 2 ends of the range
                                                RangeHi = Mid(Equation, InStr(Equation, "[TO]") + 5)
                                                RangeLo = Left(Equation, InStr(Equation, "[TO]") - 1)
                                                resultLo = Me.ParseEquation1(Demo, RangeLo)
                                                resultHi = Me.ParseEquation1(Demo, RangeHi)
                                                result = resultLo.ToString & " [TO] " & resultHi.ToString
                                            Case Else
                                                result = Me.ParseEquation1(Demo, Equation)
                                                result = Me.Do_EthnicityCorrection(Demo.EthnicityID, result, row.Item("parameterID"), row.Item("EthnicityCorrectionType") & "")
                                        End Select
                                        'Else
                                        '    result = ""
                                        'End If
                                    Else
                                        result = ""
                                    End If
                                    d.Add(row.Item("Parameter") & "|" & row.Item("StatType"), result.ToString)

                                Case "LMS"
                                    result = Me.Calculate_LMS_NotMTParser1(Demo, row.Item("Parameter"), row.Item("SourceID"), row.Item("StatType"), row.Item("age_upper"), row.Item("age_lower"))
                                    d.Add(row.Item("Parameter") & "|" & row.Item("StatType"), result.ToString)

                            End Select
                        Next
                End Select
                Return d
            Else
                Return Nothing
            End If

        Catch ex As Exception
            MsgBox("Error in class_pred.Get_PredValues" & vbCrLf & Err.Description)
            Return Nothing
        End Try


    End Function

    Public Function Do_EthnicityCorrection(EthnicityID As Integer, result As Single, ParameterID As Integer, EthnicityCorrectionType As String) As Single

        Try
            If EthnicityID <> 1 Then    'ie not caucasian
                Select Case EthnicityCorrectionType
                    Case "ATS(1991)"
                        Return result * Me.Get_EthnicityCorrection(ParameterID, EthnicityCorrectionType)
                    Case Else
                        Return result
                End Select
            Else
                Return result
            End If

        Catch ex As Exception
            MsgBox("Error in class_pred.Do_EthnicityCorrection" & vbCrLf & Err.Description)
            Return result
        End Try

    End Function

    Public Function Get_EthnicityCorrection(ParameterID As Integer, EthnicityCorrectionType As String) As Single

        Try
            Dim sql As String = "SELECT factor FROM pred_noncaucasian_corrections WHERE parameterID=" & ParameterID & " AND reference='" & EthnicityCorrectionType & "'"
            Dim ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
            Dim Factor As Single = 1.0

            If IsNothing(ds) Then
                Return 1.0
            Else
                If ds.Tables(0).Rows.Count = 1 Then
                    Factor = ds.Tables(0).Rows(0).Item(0)
                    If IsNumeric(Factor) Then
                        If Factor > 0 And Factor < 2 Then
                            Return Factor
                        Else
                            Return 1.0
                        End If
                    Else
                        Return 1.0
                    End If
                Else
                    Return 1.0
                End If
            End If
            ds = Nothing
        Catch ex As Exception
            MsgBox("Error in class_pred.Get_EthnicityCorrection" & vbCrLf & Err.Description)
            Return 1.0
        End Try

    End Function

    Public Function Do_ParameterOutOfRangeCheck(ByVal EquationInfo As DataRow, ByVal Demo As Pred_demo) As eParameterRangeCheck
        'Finds all parameter labels in equation and checks that an in-range value has been supplied
        'Returns true if all OK

        Dim names() As String = {"Age", "Ht", "Wt"}
        Dim values() As Single = {Val(Demo.Age), Val(Demo.Htcm), Val(Demo.Wtkg)}
        Dim upperlimit() As Single = {100, 200, 300}
        Dim lowerlimit() As Single = {1, 20, 20}
        Dim OK As Boolean = False
        Dim NotSupplied As Boolean = False
        Dim OutOfRange As Boolean = False


        Dim d As Dictionary(Of String, String) = Me.Get_EquationInfoForID(EquationInfo("EquationID"))
        If d Is Nothing Then
            Return eParameterRangeCheck.NotSupplied     'Cant find equation info record - should never happen
        Else
            upperlimit(0) = Val(EquationInfo("Age_upper") & "")
            lowerlimit(0) = Val(EquationInfo("Age_lower") & "")
            upperlimit(1) = Val(EquationInfo("Ht_upper") & "")
            lowerlimit(1) = Val(EquationInfo("Ht_lower") & "")
            upperlimit(2) = Val(EquationInfo("Wt_upper") & "")
            lowerlimit(2) = Val(EquationInfo("Wt_lower") & "")
        End If

        For i As Integer = 0 To UBound(names)
            If InStr(EquationInfo("equation"), names(i)) > 0 Then
                Select Case values(i)
                    Case 0, Is = Nothing : NotSupplied = True
                    Case Is < lowerlimit(i) : OutOfRange = True
                    Case Is > upperlimit(i) : OutOfRange = True
                    Case Else
                End Select
            End If
        Next

        If NotSupplied Then
            Return eParameterRangeCheck.NotSupplied     'one or more parameters not supplied, could also be some out of range
        ElseIf OutOfRange Then
            Return eParameterRangeCheck.OutOfRange
        Else
            Return eParameterRangeCheck.OK
        End If


    End Function

    Public Function Do_MPVsubstitution(ByVal Equation As String, mpv_result As Single) As String
        'Finds in-place '_MPV' labels in normal range equations, gets that MPV equation and replaces the label with the mpv result
        'Returns nothing if substitute equation not found, otherwise returns the complete final equation.
        'Assume that this label only found in normal range equations


        'Nothing to do
        If InStr(Equation, "_MPV") = 0 Then
            Return Equation
            Exit Function
        End If

        'Something to do
        Try
            Dim tempEq As String = Equation
            tempEq = Replace(tempEq, "_MPV", "(" & mpv_result & ")")
            Return tempEq

        Catch ex As Exception
            MsgBox("Error in 'class_pred.Do_ParameterSubstitutions'." & vbCrLf & ex.Message, vbOKOnly, "Unexpected error")
            Return Equation

        End Try

    End Function

    Public Function Do_ParameterSubstitutions1(ByVal ParameterName As String, ByVal Equation As String, ByVal Ds As DataSet) As String
        'Finds in-place parameter labels that represent another equation, gets that equation and replaces the label with the equation
        'Returns nothing if substitute equation not found, otherwise returns the complete final equation.
        'Assume that the equation set for a given parameter will only have 2 equations  -> MPV and range

        Dim Found As Boolean = False
        Dim tempEq As String = Equation

        Try
            'Find FVC equation in the dataset
            If InStr(tempEq, "_FVC.MPV") > 0 Then
                For Each r As DataRow In Ds.Tables(0).Rows
                    If r.Item("StatType") = "MPV" And r.Item("parameter") = "FVC" Then
                        tempEq = Replace(tempEq, "_FVC.MPV", "(" & r.Item("Equation") & ")")
                        Found = True
                        Exit For
                    End If
                Next
            End If

            If InStr(tempEq, "_FVC.LLN") > 0 Then
                For Each r As DataRow In Ds.Tables(0).Rows
                    If r.Item("StatType") = "LLN" And r.Item("parameter") = "FVC" Then
                        tempEq = Replace(tempEq, "_FVC.LLN", "(" & r.Item("Equation") & ")")
                        Found = True
                        Exit For
                    End If
                Next
            End If

            'Look for MPV equation in the dataset
            If InStr(tempEq, "_MPV") > 0 Then
                For Each r As DataRow In Ds.Tables(0).Rows
                    If r.Item("StatType") = "MPV" And ParameterName = r.Item("parameter") Then
                        tempEq = Replace(tempEq, "_MPV", "(" & r.Item("Equation") & ")")
                        Found = True
                        Exit For
                    End If
                Next
            End If

            If Found Then Return tempEq Else Return Equation

        Catch ex As Exception
            MsgBox("Error in 'class_pred.Do_ParameterSubstitutions'." & vbCrLf & ex.Message, vbOKOnly, "Unexpected error")
            Return Equation
        End Try

    End Function

    Private Function myIsDBNull(ByVal LookupValue) As Double

        If IsDBNull(LookupValue) Then
            Return Nothing
        Else
            Return LookupValue
        End If

    End Function

    Public Function Get_LMS_Coefficients(ByVal SourceID As Long, ByVal Parameter As String, ByVal Gender As String) As LMS_Coefficients

        Dim C As LMS_Coefficients = Nothing
        Dim sql As String = ""
        Dim Ds As DataSet = Nothing

        Try
            sql = "SELECT * FROM pred_lms_coeff WHERE SourceID=" & SourceID & " AND Parameter='" & Parameter & "' AND Gender='" & Gender & "'"
            Ds = cDAL.Get_DataAsDataset(sql)
            With Ds.Tables(0)
                If .Rows.Count = 1 Then
                    C.a0 = Me.myIsDBNull(.Rows(0)("Intercept_M"))
                    C.a1 = Me.myIsDBNull(.Rows(0)("Height_M"))
                    C.a2 = Me.myIsDBNull(.Rows(0)("Age_M"))
                    C.a3 = Me.myIsDBNull(.Rows(0)("AfrAm_M"))
                    C.a4 = Me.myIsDBNull(.Rows(0)("NEastAsia_M"))
                    C.a5 = Me.myIsDBNull(.Rows(0)("SEastAsia_M"))
                    C.a6 = Me.myIsDBNull(.Rows(0)("OtherMixed_M"))
                    C.p0 = Me.myIsDBNull(.Rows(0)("Intercept_S"))
                    C.p1 = Me.myIsDBNull(.Rows(0)("Age_S"))
                    C.p2 = Me.myIsDBNull(.Rows(0)("AfrAm_S"))
                    C.p3 = Me.myIsDBNull(.Rows(0)("NEastAsia_S"))
                    C.p4 = Me.myIsDBNull(.Rows(0)("SEastAsia_S"))
                    C.p5 = Me.myIsDBNull(.Rows(0)("OtherMixed_S"))
                    C.q0 = Me.myIsDBNull(.Rows(0)("Intercept_L"))
                    C.q1 = Me.myIsDBNull(.Rows(0)("Age_L"))
                Else
                    'Shouldn't happen
                End If
            End With
            Ds = Nothing
            Return C
        Catch ex As Exception
            MsgBox("Error in class_Pred.Get_LMS_Coefficients " & vbCrLf & ex.Message)
            Return Nothing
        End Try

    End Function

    Public Function Get_LMS_Equations(ByVal SourceID As Long) As LMS_Equations

        Dim E As LMS_Equations = Nothing
        Dim sql As String = ""
        Dim Ds As DataSet = Nothing

        Return Nothing

        '+++++Hardwire the GLI equations until the parser log() problem is sorted+++++++++++++++++



        sql = "SELECT * FROM pred_lms_equations WHERE SourceID=" & SourceID
        Ds = cDAL.Get_DataAsDataset(sql)
        With Ds.Tables(0)
            If .Rows.Count = 1 Then
                E.M_equation = .Rows(0)("M_equation")
                E.L_equation = .Rows(0)("L_equation")
                E.S_equation = .Rows(0)("S_equation")
                E.LLN = .Rows(0)("LLN_equation")
                E.Zscore = .Rows(0)("Zscore_equation")
            Else
                'Shouldn't happen
            End If
        End With
        Ds = Nothing

        Return E

    End Function

    Public Function Get_LMS_Splines(ByVal Parameter As String, ByVal p As ParameterInfo_UseToCalc, maxAge As Single, minAge As Single) As LMS_Splines
        'Retrieve the 2 x spline values that bracket 'Demo.Age' and return linearly interpolated values

        Dim S As LMS_Splines
        Dim sql1 As String = "", sql2 As String = "", sql3 As String = "", sql4 As String = ""
        Dim Ds As DataSet = Nothing
        Dim Mup As Double = 0, Lup As Double = 0, Sup As Double = 0
        Dim Mlo As Double = 0, Llo As Double = 0, Slo As Double = 0

        Dim Gender As String = cMyRoutines.Lookup_list_ByCode(p.GenderID_ForCalc, eTables.Pred_ref_genders)
        Dim M_Field As String = Parameter & "_MSpline_" & Gender
        Dim L_Field As String = Parameter & "_LSpline_" & Gender
        Dim S_Field As String = Parameter & "_SSpline_" & Gender

        'Set Age to 1 dec place
        Dim Age As Single = Round(p.Age_ForCalc, 1)

        'Build queries
        sql1 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age=" & minAge
        sql2 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age=" & maxAge
        Select Case cDAL.Get_DBType
            Case eDatabaseType.MySQL
                sql3 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age >" & Age & " ORDER BY Age ASC LIMIT 1"
                sql4 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age <" & Age & " ORDER BY Age DESC LIMIT 1"
            Case eDatabaseType.SQLServer, eDatabaseType.MicrosoftAccess
                sql3 = "SELECT TOP 1 Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age >" & Age & " ORDER BY Age ASC"
                sql4 = "SELECT TOP 1 Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age <" & Age & " ORDER BY Age DESC"
        End Select

        'Get data
        Select Case Age
            Case minAge
                Ds = cDAL.Get_DataAsDataset(sql1)
                If Ds.Tables(0).Rows.Count = 1 Then
                    S.MSpline = Ds.Tables(0).Rows(0)(M_Field)
                    S.LSpline = Ds.Tables(0).Rows(0)(L_Field)
                    S.SSpline = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If

            Case maxAge
                Ds = cDAL.Get_DataAsDataset(sql2)
                If Ds.Tables(0).Rows.Count = 1 Then
                    S.MSpline = Ds.Tables(0).Rows(0)(M_Field)
                    S.LSpline = Ds.Tables(0).Rows(0)(L_Field)
                    S.SSpline = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If
            Case minAge + 0.001 To maxAge - 0.001
                'Get the next higher age
                Ds = cDAL.Get_DataAsDataset(sql3)
                If Ds.Tables(0).Rows.Count = 1 Then
                    Mup = Ds.Tables(0).Rows(0)(M_Field)
                    Lup = Ds.Tables(0).Rows(0)(L_Field)
                    Sup = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If
                Ds = Nothing

                'Get the next lower age
                Ds = cDAL.Get_DataAsDataset(sql4)
                If Ds.Tables(0).Rows.Count = 1 Then
                    Mlo = Ds.Tables(0).Rows(0)(M_Field)
                    Llo = Ds.Tables(0).Rows(0)(L_Field)
                    Slo = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If

                'Interpolate
                S.MSpline = Mlo + (Mup - Mlo) / 2
                S.LSpline = Llo + (Lup - Llo) / 2
                S.SSpline = Slo + (Sup - Slo) / 2

            Case Else
                'Age out of range 
        End Select

        Return S
        Ds = Nothing


    End Function

    Public Function Get_LMS_Splines1(ByVal Parameter As String, ByVal Demo As Pred_demo, maxAge As Single, minAge As Single) As LMS_Splines
        'Retrieve the 2 x spline values that bracket 'Demo.Age' and return linearly interpolated values

        Dim S As LMS_Splines
        Dim sql1 As String = "", sql2 As String = "", sql3 As String = "", sql4 As String = ""
        Dim Ds As DataSet = Nothing
        Dim Mup As Double = 0, Lup As Double = 0, Sup As Double = 0
        Dim Mlo As Double = 0, Llo As Double = 0, Slo As Double = 0

        Dim Gender As String = cMyRoutines.Lookup_list_ByCode(Demo.GenderID, eTables.Pred_ref_genders)
        Dim M_Field As String = Parameter & "_MSpline_" & Gender
        Dim L_Field As String = Parameter & "_LSpline_" & Gender
        Dim S_Field As String = Parameter & "_SSpline_" & Gender

        'Set Age to 1 dec place
        Dim Age As Single = Round(Demo.Age, 1)

        'Build queries
        sql1 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age=3.0"
        sql2 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age=95.0"
        Select Case cDAL.Get_DBType
            Case eDatabaseType.MySQL
                sql3 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age >" & Age & " ORDER BY Age ASC LIMIT 1"
                sql4 = "SELECT Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age <" & Age & " ORDER BY Age DESC LIMIT 1"
            Case eDatabaseType.SQLServer, eDatabaseType.MicrosoftAccess
                sql3 = "SELECT TOP 1 Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age >" & Age & " ORDER BY Age ASC"
                sql4 = "SELECT TOP 1 Age, " & M_Field & ", " & L_Field & ", " & S_Field & " FROM pred_GLI_Lookup WHERE Age <" & Age & " ORDER BY Age DESC"
        End Select

        'Get data
        Select Case Age
            Case minAge
                Ds = cDAL.Get_DataAsDataset(sql1)
                If Ds.Tables(0).Rows.Count = 1 Then
                    S.MSpline = Ds.Tables(0).Rows(0)(M_Field)
                    S.LSpline = Ds.Tables(0).Rows(0)(L_Field)
                    S.SSpline = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If

            Case maxAge
                Ds = cDAL.Get_DataAsDataset(sql2)
                If Ds.Tables(0).Rows.Count = 1 Then
                    S.MSpline = Ds.Tables(0).Rows(0)(M_Field)
                    S.LSpline = Ds.Tables(0).Rows(0)(L_Field)
                    S.SSpline = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If
            Case minAge + 0.001 To maxAge - 0.001
                'Get the next higher age
                Ds = cDAL.Get_DataAsDataset(sql3)
                If Ds.Tables(0).Rows.Count = 1 Then
                    Mup = Ds.Tables(0).Rows(0)(M_Field)
                    Lup = Ds.Tables(0).Rows(0)(L_Field)
                    Sup = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If
                Ds = Nothing

                'Get the next lower age
                Ds = cDAL.Get_DataAsDataset(sql4)
                If Ds.Tables(0).Rows.Count = 1 Then
                    Mlo = Ds.Tables(0).Rows(0)(M_Field)
                    Llo = Ds.Tables(0).Rows(0)(L_Field)
                    Slo = Ds.Tables(0).Rows(0)(S_Field)
                Else
                    'Shouldn't happen
                End If

                'Interpolate
                S.MSpline = Mlo + (Mup - Mlo) / 2
                S.LSpline = Llo + (Lup - Llo) / 2
                S.SSpline = Slo + (Sup - Slo) / 2

            Case Else
                'Age out of range 
        End Select

        Return S
        Ds = Nothing


    End Function

    Public Function Get_EquationsForSourceID_ParameterID(ByVal InputData As Pred_InputData, ByVal ParameterID As Integer) As Dictionary(Of String, String)()
        'Get equations based on -
        '   SourceID
        '   TestID
        '   GenderID
        '   AgeGroupID
        '   EthnicityID

        Dim sql As New StringBuilder
        Dim d() As Dictionary(Of String, String)
        Dim Age As Single = Val(cMyRoutines.Calc_Age(InputData.DOB, InputData.TestDate))

        If ParameterID = 0 Then
            MsgBox("Problem selecting parameter")
            Return Nothing
        End If

        ReDim d(0)

        sql.Clear()
        sql.Append("SELECT Pred_Ref_parameters.TestID, Pred_Ref_parameters.description, Pred_Ref_parameters.code AS ParID, Pred_Ref_parameters.Units, Pred_equations.* ")
        sql.Append("FROM Pred_Ref_parameters INNER JOIN Pred_equations ON Pred_Ref_parameters.code = Pred_equations.ParameterID ")
        sql.Append("WHERE Pred_equations.SourceID=" & InputData.SourceID)
        sql.Append(" AND Pred_Ref_parameters.code=" & ParameterID)
        If InputData.GenderID > 0 Then sql.Append(" AND Pred_equations.GenderID=" & InputData.GenderID)
        If InputData.Age > 0 Then sql.Append(" AND Age_lower <= " & Age & " AND Age_upper >= " & Age)
        If InputData.EthnicityID > 0 Then sql.Append(" AND Pred_equations.EthnicityID=" & InputData.EthnicityID)

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        For Each r As DataRow In Ds.Tables(0).Rows
            ReDim Preserve d(d.GetUpperBound(0) + 1)
            d(d.GetUpperBound(0)) = New Dictionary(Of String, String)
            d(d.GetUpperBound(0)).Add("EquationID", r.Item("EquationID"))
            d(d.GetUpperBound(0)).Add("Equation", r.Item("Equation"))
            d(d.GetUpperBound(0)).Add("Units", r.Item("Units"))
            d(d.GetUpperBound(0)).Add("TestID", r.Item("TestID"))
            d(d.GetUpperBound(0)).Add("StatType", r.Item("StatType"))
            d(d.GetUpperBound(0)).Add("StatTypeID", r.Item("StatTypeID"))
            d(d.GetUpperBound(0)).Add("EquationType", r.Item("EquationType") & "")
            d(d.GetUpperBound(0)).Add("EquationTypeID", r.Item("EquationTypeID") & "")
            d(d.GetUpperBound(0)).Add("Parameter", r.Item("Parameter") & "")
            d(d.GetUpperBound(0)).Add("ParameterID", r.Item("ParID") & "")
            d(d.GetUpperBound(0)).Add("Gender", r.Item("Gender") & "")
            d(d.GetUpperBound(0)).Add("Age_lower", r.Item("Age_lower") & "")
            d(d.GetUpperBound(0)).Add("Age_upper", r.Item("Age_upper") & "")
            d(d.GetUpperBound(0)).Add("Age_clipmethod", r.Item("Age_clipmethod") & "")
            d(d.GetUpperBound(0)).Add("Ht_lower", r.Item("Ht_lower") & "")
            d(d.GetUpperBound(0)).Add("Ht_upper", r.Item("Ht_upper") & "")
            d(d.GetUpperBound(0)).Add("Ht_clipmethod", r.Item("Ht_clipmethod") & "")
            d(d.GetUpperBound(0)).Add("Wt_lower", r.Item("Wt_lower") & "")
            d(d.GetUpperBound(0)).Add("Wt_upper", r.Item("Wt_upper") & "")
            d(d.GetUpperBound(0)).Add("Wt_clipmethod", r.Item("Wt_clipmethod") & "")
            d(d.GetUpperBound(0)).Add("Ethnicity", r.Item("Ethnicity") & "")
            d(d.GetUpperBound(0)).Add("EthnicityCorrectionType", r.Item("EthnicityCorrectionType") & "")
            d(d.GetUpperBound(0)).Add("AgeGroup", r.Item("AgeGroup") & "")
        Next

        Return d

    End Function

    Public Function Get_EquationsAllFor_SourceID_TestID(ByVal SourceID As Integer, ByVal TestID As Integer) As form_prefs_normals.RowData()

        Dim sql As New StringBuilder
        Dim e() As form_prefs_normals.RowData
        ReDim e(-1)

        sql.Clear()
        sql.Append("SELECT * FROM Pred_equations WHERE SourceID =" & SourceID & " AND TestID =" & TestID)

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        For Each r As DataRow In Ds.Tables(0).Rows
            ReDim Preserve e(e.GetUpperBound(0) + 1)
            e(e.GetUpperBound(0)) = New form_prefs_normals.RowData
            e(e.GetUpperBound(0)).Test.Value = r.Item("Test")
            e(e.GetUpperBound(0)).TestID.Value = r.Item("TestID")
            e(e.GetUpperBound(0)).Parameter.Value = r.Item("Parameter") & ""
            e(e.GetUpperBound(0)).ParameterID.Value = r.Item("ParameterID") & ""
            e(e.GetUpperBound(0)).Source.Value = r.Item("Source") & ""
            e(e.GetUpperBound(0)).SourceID.Value = r.Item("SourceID") & ""
            e(e.GetUpperBound(0)).Gender.Value = r.Item("Gender") & ""
            e(e.GetUpperBound(0)).GenderID.Value = r.Item("GenderID") & ""
            e(e.GetUpperBound(0)).AgeGroup.Value = r.Item("AgeGroup") & ""
            e(e.GetUpperBound(0)).AgeGroupID.Value = r.Item("AgeGroupID") & ""
            e(e.GetUpperBound(0)).Age_lower.Value = r.Item("Age_lower") & ""
            e(e.GetUpperBound(0)).Age_upper.Value = r.Item("Age_upper") & ""
            e(e.GetUpperBound(0)).Age_clipmethod.Value = r.Item("Age_clipmethod") & ""
            e(e.GetUpperBound(0)).Ht_lower.Value = r.Item("ht_lower") & ""
            e(e.GetUpperBound(0)).Ht_upper.Value = r.Item("ht_upper") & ""
            e(e.GetUpperBound(0)).Ht_clipmethod.Value = r.Item("ht_clipmethod") & ""
            e(e.GetUpperBound(0)).Wt_lower.Value = r.Item("wt_lower") & ""
            e(e.GetUpperBound(0)).Wt_upper.Value = r.Item("wt_upper") & ""
            e(e.GetUpperBound(0)).Wt_clipmethod.Value = r.Item("wt_clipmethod") & ""
            e(e.GetUpperBound(0)).Ethnicity.Value = r.Item("Ethnicity") & ""
            e(e.GetUpperBound(0)).EthnicityID.Value = r.Item("EthnicityID") & ""
            e(e.GetUpperBound(0)).Ethnicity_ApplyATS1991Correction.Value = r.Item("EthnicityCorrectionType") & ""
            e(e.GetUpperBound(0)).StartDate.Value = ""
            e(e.GetUpperBound(0)).Equation.Value = r.Item("Equation")
            e(e.GetUpperBound(0)).Equationid.Value = r.Item("Equationid")
            e(e.GetUpperBound(0)).StatType.Value = r.Item("StatType")
            e(e.GetUpperBound(0)).StatTypeid.Value = r.Item("StatTypeID")
            e(e.GetUpperBound(0)).Equation_Class.Value = r.Item("EquationType") & ""
            e(e.GetUpperBound(0)).Equation_ClassID.Value = r.Item("EquationTypeID") & ""
        Next

        Return e

    End Function

    Public Function Get_StatNameFromEnum(ByRef Stat As module_defs.StatType) As String

        Dim s As String = ""

        Try
            Select Case Stat
                Case module_defs.StatType.LLN : s = "LLN"
                Case module_defs.StatType.MPV : s = "MPV"
                Case module_defs.StatType.ULN : s = "ULN"
                Case module_defs.StatType.Range : s = "RANGE"
                Case Else : s = ""
            End Select
            Return s
        Catch
            Return ""
        End Try

    End Function

    Public Function Get_RefItems(ByVal RefType As RefItems) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As String = ""
        Dim Tbl As String = ""
        Dim PK As String = ""
        Dim Description As String = ""

        Select Case RefType
            Case RefItems.AgeGroups : Tbl = "Pred_Ref_AgeGroups"            ': PK = "Code" : Description = "Description"
            Case RefItems.Ethnicities : Tbl = "Pred_Ref_Ethnicities"        ': PK = "Code" : Description = "Description"
            Case RefItems.Genders : Tbl = "Pred_Ref_Genders"                ': PK = "Code" : Description = "Description"
            Case RefItems.Parameters : Tbl = "Pred_Ref_Parameters"          ': PK = "Code" : Description = "Description"
            Case RefItems.Sources : Tbl = "Pred_Ref_Sources"                ': PK = "code" : Description = "Description"
            Case RefItems.EquationTypes : Tbl = "Pred_Ref_EquationTypes"    ': PK = "Code" : Description = "Description"
            Case RefItems.StatisticTypes : Tbl = "Pred_Ref_StatTypes"       ': PK = "Code" : Description = "Description"
            Case RefItems.Tests : Tbl = "Pred_Ref_Tests"                    ': PK = "code" : Description = "Description"
            Case RefItems.Variables : Tbl = "Pred_Ref_Variables"            ': PK = "Code" : Description = "Description"
            Case RefItems.ClipMethods : Tbl = "Pred_Ref_ClipMethods"        ': PK = "Code" : Description = "Description"
        End Select

        sql = "SELECT Code, Description FROM " & Tbl
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            For Each r As DataRow In Ds.Tables(0).Rows
                d.Add(r.Item("code"), r.Item("Description"))
            Next
        End If

        Return d

    End Function

    Public Function Get_SourceInfo(ByVal SourceID As Integer, ByVal TestID As Integer) As form_prefs_normals.RowData

        Dim sql As New StringBuilder
        Dim s As New form_prefs_normals.RowData
        Dim found As Boolean = False

        sql.Clear()
        sql.Append("SELECT DISTINCT Pred_Ref_sources.description, Pred_Ref_sources.code, Pred_Ref_sources.pub_reference, Pred_Ref_sources.pub_year, pred_equations.equationtypeID, pred_equations.equationtype, pred_equations.testid ")
        sql.Append("FROM (pred_equations RIGHT OUTER JOIN Pred_Ref_sources ON pred_equations.SourceID = Pred_Ref_sources.code) ")
        sql.Append("WHERE Pred_Ref_sources.code=" & CStr(SourceID))

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        With Ds.Tables(0)
            If .Rows.Count = 1 Then
                If IsDBNull(.Rows(0)("TestID")) Then
                    found = True
                ElseIf .Rows(0)("TestID") = TestID Then
                    found = True
                Else
                    found = False
                End If
                If found Then
                    s.Source.Value = .Rows(0)("description")
                    s.SourceID.Value = .Rows(0)("code")
                    s.Equation_Class.Value = .Rows(0)("equationtype") & ""
                    s.Source_PubReference.Value = .Rows(0)("Pub_Reference") & ""
                    s.Source_PubYear.Value = .Rows(0)("Pub_year") & ""
                End If
            End If
        End With

        Return s

    End Function

    Public Function Get_SourcesForTestID(ByRef TestID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT Pred_Ref_tests.*, Pred_Ref_sources.code as source_code, Pred_Ref_sources.description as source_description ")
        sql.Append("FROM (Pred_sourceXtest INNER JOIN Pred_Ref_sources ON Pred_sourceXtest.SourceID = Pred_Ref_sources.code) INNER JOIN Pred_Ref_tests ON Pred_sourceXtest.TestID = Pred_Ref_tests.code ")
        sql.Append("WHERE (((Pred_Ref_tests.code)=")
        sql.Append(CStr(TestID) & "));")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        For Each r As DataRow In Ds.Tables(0).Rows
            d.Add(r.Item("source_code"), r.Item("source_description"))
        Next

        Return d

    End Function

    Public Function Get_TestForSourceID(ByRef SourceID As Integer) As KeyValuePair(Of String, String)
        'Assume that only one test per source

        Dim kv As KeyValuePair(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT Pred_Ref_tests.* ")
        sql.Append("FROM (Pred_sourceXtest INNER JOIN Pred_Ref_sources ON Pred_sourceXtest.SourceID = Pred_Ref_sources.code) INNER JOIN Pred_Ref_tests ON Pred_sourceXtest.TestID = Pred_Ref_tests.code ")
        sql.Append("WHERE (((Pred_sourceXtest.SourceID)=")
        sql.Append(CStr(SourceID) & "));")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        Select Case Ds.Tables(0).Rows.Count
            Case 0 : Return Nothing
            Case 1
                kv = New KeyValuePair(Of String, String)(Ds.Tables(0).Rows(0).Item("description"), Ds.Tables(0).Rows(0).Item("code"))
                Return kv
            Case Else : Return Nothing
        End Select

    End Function

    Public Function Get_EthnicityForID(ByRef EthnicityID As Integer) As String

        Dim sql As String = "SELECT description FROM Pred_Ref_ethnicities WHERE code=" & EthnicityID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If (Not IsNothing(Ds) AndAlso Ds.Tables(0).Rows.Count = 1) Then
            Return Ds.Tables(0).Rows(0)("description")
        Else
            Return ""
        End If

    End Function

    Public Function Get_EthnicitiesForSource(ByRef SourceID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT Pred_Ref_ethnicities.Description, Pred_Ref_ethnicities.code  ")
        sql.Append("FROM Pred_Ref_ethnicities INNER JOIN (Pred_Ref_sources INNER JOIN Pred_sourceXethnicity ")
        sql.Append("ON Pred_Ref_sources.code = Pred_sourceXethnicity.SourceID) ON Pred_Ref_ethnicities.Code = Pred_sourceXethnicity.EthnicityID ")
        sql.Append("WHERE (((Pred_Ref_sources.code)=")
        sql.Append(CStr(SourceID) & "));")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        For Each r As DataRow In Ds.Tables(0).Rows
            d.Add(r.Item("code"), r.Item("description"))
        Next

        Return d

    End Function

    Public Function Get_ParametersForSourceID(ByRef SourceID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT DISTINCT Pred_Ref_parameters.description, Pred_Ref_parameters.code ")
        sql.Append("FROM Pred_Ref_sources INNER JOIN (Pred_Ref_parameters INNER JOIN Pred_sourceXparameter ON ")
        sql.Append("Pred_Ref_parameters.code = Pred_sourceXparameter.ParamID) ON Pred_Ref_sources.code = Pred_sourceXparameter.SourceID ")
        sql.Append("WHERE (((Pred_Ref_sources.code)=")
        sql.Append(CStr(SourceID) & "));")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        d.Clear()
        For Each r As DataRow In Ds.Tables(0).Rows
            d.Add(r.Item("code"), r.Item("description"))
        Next

        Return d

    End Function

    Public Function Get_ParametersForTestID(ByRef TestID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT Pred_Ref_parameters.description, Pred_Ref_parameters.code ")
        sql.Append("FROM Pred_Ref_tests INNER JOIN Pred_Ref_parameters ON Pred_Ref_tests.code = Pred_Ref_parameters.TestID ")
        sql.Append("WHERE Pred_Ref_tests.code=")
        sql.Append(CStr(TestID) & ";")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        d.Clear()
        For Each r As DataRow In Ds.Tables(0).Rows
            d.Add(r.Item("code"), r.Item("description"))
        Next

        Return d

    End Function

    Public Function Get_ParametersForTestID_SourceID(ByRef TestID As Integer, ByVal SourceID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT Pred_Ref_parameters.* ")
        sql.Append("FROM (Pred_Ref_tests INNER JOIN Pred_sourceXtest ON Pred_Ref_tests.code = Pred_sourceXtest.TestID) INNER JOIN Pred_Ref_parameters ON Pred_Ref_tests.code = Pred_Ref_parameters.TestID ")
        sql.Append("WHERE Pred_Ref_tests.code=" & CStr(TestID))
        sql.Append(" AND Pred_sourceXtest.SourceID=" & CStr(SourceID))

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        d.Clear()
        For Each r As DataRow In Ds.Tables(0).Rows
            d.Add(r.Item("code"), r.Item("description"))
        Next

        Return d

    End Function

    Public Function Get_GendersForTestID_SourceID(ByRef TestID As Integer, ByVal SourceID As Integer) As Dictionary(Of String, String)

        Dim d As New Dictionary(Of String, String)
        Dim sql As New StringBuilder

        sql.Clear()
        sql.Append("SELECT pred_ref_genders.description, pred_ref_genders.code ")
        sql.Append("FROM pred_ref_tests INNER JOIN (((pred_ref_genders INNER JOIN pred_sourcexgender ON pred_ref_genders.code = pred_sourcexgender.GenderID) ")
        sql.Append("INNER JOIN pred_ref_sources ON pred_sourcexgender.SourceID = pred_ref_sources.code) ")
        sql.Append("INNER JOIN pred_sourcextest ON pred_ref_sources.code = pred_sourcextest.SourceID) ON ")
        sql.Append("pred_ref_tests.code = pred_sourcextest.TestID ")
        sql.Append("WHERE Pred_Ref_tests.code=" & CStr(TestID))
        sql.Append(" AND pred_ref_sources.code=" & CStr(SourceID))

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)
        d.Clear()
        For Each r As DataRow In Ds.Tables(0).Rows
            d.Add(r.Item("code"), r.Item("description"))
        Next

        Return d

    End Function

    Public Function New_Equation(ByVal d As Dictionary(Of String, String)) As Integer
        'Returns new PKID 

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.Pred_equations, d)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new predicted equation" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

    End Function

    Public Function Update_Equation(ByVal d As Dictionary(Of String, String)) As Integer
        'Returns current PKID 

        Dim sqlString As String = cDAL.Build_UpdateQuery(eTables.Pred_equations, d)

        Try
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving predicted equation" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Public Function New_Source(ByVal d As Dictionary(Of String, String)) As Integer
        'Returns new PKID 

        'Build insert query
        Dim sql As String = cDAL.Build_InsertQuery(eTables.Pred_ref_sources, d)

        'Apply insert
        Try
            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error creating new predicted equation" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

    End Function

    Public Function Update_Source(ByVal d As Dictionary(Of String, String)) As Integer
        'Returns current PKID 

        Dim sqlString As String = cDAL.Build_UpdateQuery(eTables.Pred_ref_sources, d)

        Try
            Dim ReturnValue As Boolean = cDAL.Update_Record(sqlString)
            Return ReturnValue
        Catch ex As Exception
            MsgBox("Error saving predicted source" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

    End Function

    Public Function Update_PrefsPred(ByVal P() As Dictionary(Of String, String)) As Boolean
        'Care must be taken using this routine ..
        'Assumes -
        '1. All records (open and end-dated) are loaded to the grid
        '2. Editing takes place eg end-date old selections and add new selections
        '3. Not allowed to un-enddate an end-dated record (allowed to change the date though)
        '4. Up to the user to ensure that temporal selections do not leave date gaps that create periods not covered by any selection
        '5. Validity check performed by software to detect overlapping selections (not allowed)
        '6. Save operation completely overwrites the existing table data with the grid contents
        'So - make sure the old history of enddated prefs and current open prefs in the grid is correct BEFORE before saving!!



        Dim sql As String = ""
        Dim Success As Boolean = False

        Try
            'First mark as delete all existing records 
            sql = "UPDATE " & cDbInfo.table_name(eTables.Preferences_pred) & " SET MarkForDeletion = 1"
            Success = cDAL.Execute_NonQuery(sql)
            'Insert new records
            If Success Then
                For Each Pref As Dictionary(Of String, String) In P
                    sql = cDAL.Build_InsertQuery(eTables.Preferences_pred, Pref)
                    cDAL.Insert_Record(sql)
                Next
            Else
                MsgBox("Error updating normal values preferences" & vbNewLine & "Unable to delete old data")
                Return False
            End If

            'Delete the marked records
            sql = "DELETE FROM " & cDbInfo.table_name(eTables.Preferences_pred) & " WHERE MarkForDeletion = 1"
            If cDAL.Execute_NonQuery(sql) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("Error updating normal values preferences" & vbNewLine & ex.Message.ToString)
            Return False
        End Try

        Return True

    End Function

    Public Function Get_PrefsPred(ByVal IncludeEndDatedRecords As Boolean) As Dictionary(Of String, String)()

        Try

            Dim sql As String = ""

            Select Case IncludeEndDatedRecords
                Case True : sql = "SELECT * FROM preferences_pred"
                Case False : sql = "SELECT * FROM preferences_pred WHERE enddate IS NULL"
            End Select

            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            Dim dicR() As Dictionary(Of String, String)

            ReDim dicR(0)
            dicR(0) = New Dictionary(Of String, String)

            If Not IsNothing(Ds) Then
                With Ds.Tables(0)
                    Dim p As New class_Pref_PredFields
                    For Each r As DataRow In .Rows
                        ReDim Preserve dicR(UBound(dicR) + 1)

                        dicR(UBound(dicR)) = cMyRoutines.MakeEmpty_dicPrefPred_ForLoad()
                        dicR(UBound(dicR))(p.PrefID) = r(p.PrefID)
                        dicR(UBound(dicR))(p.Age_lower) = r(p.Age_lower) & ""
                        dicR(UBound(dicR))(p.Age_upper) = r(p.Age_upper) & ""
                        dicR(UBound(dicR))(p.Age_clipmethod) = r(p.Age_clipmethod) & ""
                        dicR(UBound(dicR))(p.Age_clipmethodID) = r(p.Age_clipmethodID) & ""
                        dicR(UBound(dicR))(p.ht_lower) = r(p.ht_lower) & ""
                        dicR(UBound(dicR))(p.ht_upper) = r(p.ht_upper) & ""
                        dicR(UBound(dicR))(p.ht_clipmethod) = r(p.ht_clipmethod) & ""
                        dicR(UBound(dicR))(p.ht_clipmethodID) = r(p.ht_clipmethodID) & ""
                        dicR(UBound(dicR))(p.wt_lower) = r(p.wt_lower) & ""
                        dicR(UBound(dicR))(p.wt_upper) = r(p.wt_upper) & ""
                        dicR(UBound(dicR))(p.wt_clipmethod) = r(p.wt_clipmethod) & ""
                        dicR(UBound(dicR))(p.wt_clipmethodID) = r(p.wt_clipmethodID) & ""

                        dicR(UBound(dicR))(p.AgeGroupID) = r(p.AgeGroupID)
                        dicR(UBound(dicR))(p.AgeGroup) = cMyRoutines.Lookup_list_ByCode(r(p.AgeGroupID), eTables.Pred_ref_agegroups)
                        If IsDate(r(p.StartDate)) Then dicR(UBound(dicR))(p.StartDate) = r(p.StartDate) Else dicR(UBound(dicR))(p.StartDate) = Nothing
                        If IsDate(r(p.EndDate)) Then dicR(UBound(dicR))(p.EndDate) = r(p.EndDate) Else dicR(UBound(dicR))(p.EndDate) = Nothing
                        dicR(UBound(dicR))(p.EthnicityID) = r(p.EthnicityID)
                        dicR(UBound(dicR))(p.Ethnicity) = cMyRoutines.Lookup_list_ByCode(r(p.EthnicityID), eTables.Pred_ref_ethnicities)
                        dicR(UBound(dicR))(p.GenderID) = r(p.GenderID)
                        dicR(UBound(dicR))(p.Gender) = cMyRoutines.Lookup_list_ByCode(r(p.GenderID), eTables.Pred_ref_genders)
                        If IsDate(r(p.Lastedit)) Then dicR(UBound(dicR))(p.Lastedit) = r(p.Lastedit) Else dicR(UBound(dicR))(p.Lastedit) = Nothing
                        dicR(UBound(dicR))(p.LasteditBy) = r(p.LasteditBy) & ""
                        dicR(UBound(dicR))(p.ParameterID) = r(p.ParameterID)
                        dicR(UBound(dicR))(p.Parameter) = cMyRoutines.Lookup_list_ByCode(r(p.ParameterID), eTables.Pred_ref_parameters)
                        dicR(UBound(dicR))(p.PrefID) = r(p.PrefID)
                        dicR(UBound(dicR))(p.SourceID) = r(p.SourceID)
                        dicR(UBound(dicR))(p.Source) = cMyRoutines.Lookup_list_ByCode(r(p.SourceID), eTables.Pred_ref_sources)
                        dicR(UBound(dicR))(p.TestID) = r(p.TestID)
                        dicR(UBound(dicR))(p.Test) = cMyRoutines.Lookup_list_ByCode(r(p.TestID), eTables.Pred_ref_tests)
                        If IsDBNull(r(p.EquationID_mpv)) Then dicR(UBound(dicR))(p.EquationID_mpv) = Nothing Else dicR(UBound(dicR))(p.EquationID_mpv) = r(p.EquationID_mpv)
                        If IsDBNull(r(p.EquationID_NormalRange)) Then dicR(UBound(dicR))(p.EquationID_NormalRange) = Nothing Else dicR(UBound(dicR))(p.EquationID_NormalRange) = r(p.EquationID_NormalRange)
                    Next
                End With
            End If

            Return dicR

        Catch ex As Exception
            MsgBox("Error in class_pred.Get_PrefsPred" & vbCrLf & Err.Description)
            Return Nothing

        End Try

    End Function

    Public Function Get_PrefSources_AsCodedString(ByVal Method As eLoadNormalsMode, ByVal demo As Pred_demo) As String
        'Coded as - TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|......

        'Build the relevant query to return pred equations to use
        Dim sql As New StringBuilder
        Select Case Method
            Case class_Pred.eLoadNormalsMode.UseCurrentPrefs
                sql.Clear()
                sql.Append("SELECT DISTINCT TestID, SourceID, EthnicityID, AgeGroupID FROM preferences_pred WHERE EndDate IS NULL")
            Case class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate
                sql.Clear()
                sql.Append("SELECT  DISTINCT TestID, SourceID, EthnicityID, AgeGroupID  FROM preferences_pred ")
                sql.Append("WHERE CAST('" & Format(demo.TestDate, "yyyy-MM-dd") & "' AS DATE) >= StartDate AND CAST('" & Format(demo.TestDate, "yyyy-MM-dd") & "' AS DATE) <= (CASE WHEN EndDate IS NULL THEN CAST('" & Format(demo.TestDate, "yyyy-MM-dd") & "' AS DATE) ELSE enddate END) ")
                'Case class_Pred.eLoadNormalsMode.UseSourcesSavedWithTest
                '    ' TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|......
                '    'Get all the sourceIDs and include in query
                '    Dim SourceIDs() As Long = Me.Decode_SourcesStringToSourceIDs(demo.SourcesString)
                '    If Not IsNothing(SourceIDs) Then
                '        sql.Clear()
                '        sql.Append("SELECT * FROM pred_equations WHERE ")
                '        For i As Integer = 0 To UBound(SourceIDs)
                '            sql.Append("SourceID=" & SourceIDs(i) & " OR ")
                '        Next
                '        sql.Remove(sql.Length - 4, 4)
                '    Else
                '        Return ""
                '    End If
        End Select

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql.ToString)

        If Ds Is Nothing Then
            Return ""
        Else
            With Ds.Tables(0)
                Dim p As New class_Pref_PredFields
                sql.Clear()
                For Each r As DataRow In .Rows
                    sql.Append("testid=" & r("TestID") & ",")
                    sql.Append("sourceid=" & r("SourceID") & ",")
                    sql.Append("ethnicityid=" & r("EthnicityID") & ",")
                    sql.Append("agegroupid=" & r("AgeGroupID") & ",|")
                Next
            End With
        End If

        Return sql.ToString

    End Function

    Public Function Get_PrefSources_AsFormattedString(ByVal TestDate As Date, ByVal demo As Pred_demo) As String
        'Used for new tests
        'Convert ID's to descriptions and format nicely for display on data entry page 

        Dim s As New StringBuilder
        s.Clear()
        s.Append("SELECT DISTINCT TestID, SourceID, AgeGroupID FROM preferences_pred ")
        's.Append("WHERE " & demo.Age & " >= CAST(age_lower AS DECIMAL(10,1)) AND " & demo.Age & " <= CAST(age_upper AS DECIMAL(10,1)) ")
        s.Append("WHERE CAST('" & Format(TestDate, "yyyy-MM-dd") & "' AS DATE) >= StartDate ")
        s.Append("AND CAST('" & Format(TestDate, "yyyy-MM-dd") & "' AS DATE) <= (CASE WHEN EndDate IS NULL ")
        s.Append(" THEN CAST('" & Format(TestDate, "yyyy-MM-dd") & "' AS DATE) ELSE enddate END) ")

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(s.ToString)
        If Ds Is Nothing Then
            Return "Not recorded"
        Else
            With Ds.Tables(0)
                s.Clear()
                For Each r As DataRow In .Rows
                    s.Append(cMyRoutines.Lookup_list_ByCode(r("TestID"), eTables.Pred_ref_tests))
                    s.Append(":  ")
                    s.Append(cMyRoutines.Lookup_list_ByCode(r("SourceID"), eTables.Pred_ref_sources))
                    s.Append(",   Agegroup: ")
                    s.Append(cMyRoutines.Lookup_list_ByCode(r("AgeGroupID"), eTables.Pred_ref_agegroups) & vbCrLf)
                Next
                If demo.EthnicityID > 1 Then
                    s.Append("Predicteds corrected for ethnicity")
                End If
                Return s.ToString
            End With
        End If

    End Function

    Public Function Decode_SourcesStringForDisplay(ByVal s As String) As String
        'Input string is the pred sources info saved in the database with each test coded as-
        '   TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|......
        'Convert ID's to descriptions and format nicely for display on data entry page 

        Dim sb As New StringBuilder
        Dim ss() As String
        Dim TestID As Integer = 0, SourceID As Integer = 0, AgeGroupID As Integer = 0
        Dim j As Integer = 0, k As Integer = 0
        Dim Source As String = ""

        sb.Clear()

        If s = "" Then
            Return "Not recorded"
        Else
            ss = s.Split("|")
            For i As Integer = 0 To UBound(ss) - 1
                j = 1
                j = InStr(j, ss(i), "testid")
                j = InStr(j, ss(i), "=")
                k = InStr(j, ss(i), ",")
                TestID = Val(Mid(ss(i), j + 1, k - j - 1))

                j = InStr(j, ss(i), "sourceid")
                j = InStr(j, ss(i), "=")
                k = InStr(j, ss(i), ",")
                SourceID = Val(Mid(ss(i), j + 1, k - j - 1))

                Source = cMyRoutines.Lookup_list_ByCode(TestID, eTables.Pred_ref_tests) & ":" & cMyRoutines.Lookup_list_ByCode(SourceID, eTables.Pred_ref_sources)

                'Get rid of duplicates
                If InStr(sb.ToString, Source) > 0 Then
                    'Skip
                Else
                    sb.Append(Source & vbCrLf)
                End If

            Next
            Return sb.ToString
        End If

    End Function

    Public Function Decode_SourcesStringToSourceIDs(ByVal s As String) As Long()
        'Input string is the pred sources info saved in the database with each test coded as-
        '   TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|TestID=x,SourceID=x,EthnicityID=x,AgeGroupID=x|......
        'Output an array of sourceIDs

        Dim ss() As String
        Dim SourceID As Long = 0
        Dim j As Integer = 0, k As Integer = 0
        Dim SourceIDs() As Long = Nothing

        If s = "" Then
            Return Nothing
        Else
            ss = s.Split("|")
            For i As Integer = 0 To UBound(ss) - 1
                j = 1
                j = InStr(j, ss(i), "sourceid")
                j = InStr(j, ss(i), "=")
                k = InStr(j, ss(i), ",")
                SourceID = Val(Mid(ss(i), j + 1, k - j - 1))

                If i = 0 Then
                    ReDim SourceIDs(0)
                    SourceIDs(0) = SourceID
                Else
                    'Save if not a duplicate
                    If Not SourceIDs.Contains(SourceID) Then
                        ReDim Preserve SourceIDs(UBound(SourceIDs) + 1)
                        SourceIDs(UBound(SourceIDs)) = SourceID
                    End If
                End If
            Next
            Return SourceIDs
        End If

    End Function

    Public Function ParseEquation(ByVal Demo As Pred_demo, ByVal Eq As String) As Single
        'https://github.com/MathosProject/Mathos-Parser

        Try

            Dim Parser As New Mathos.Parser.MathParser
            Dim result As Double = 0

            If Eq.Contains("_NA") Then
                Return ""
            Else
                Parser.LocalVariables.Add("Age", Demo.Age)
                Parser.LocalVariables.Add("Ht", Demo.Htcm)
                Parser.LocalVariables.Add("Wt", Demo.Wtkg)
                Parser.LocalVariables.Add("vo2", Demo.vo2)

                Try
                    'Bug in parser - doesn't like the equation starting with a negative sign as first char in string
                    Eq = " (" & Eq & ")"
                    result = Parser.Parse(Eq)
                    Return CSng(result)
                Catch ex As Exception
                    MsgBox("Error: " & ex.Message)
                    Return vbEmpty
                End Try
            End If

        Catch ex As Exception
            MsgBox("Error in class_pred.ParseEquation" & vbCrLf & ex.Message)
            Return vbEmpty
        End Try

    End Function


    Public Function ParseEquation1(ByVal paramInfo As Pred_demo, ByVal Eq As String) As Single
        'https://github.com/MathosProject/Mathos-Parser

        Try

            Dim Parser As New Mathos.Parser.MathParser
            Dim result As Double = 0

            If Eq.Contains("_NA") Then
                Return ""
            Else
                Parser.LocalVariables.Add("Age", paramInfo.Age)
                Parser.LocalVariables.Add("Ht", paramInfo.Htcm)
                Parser.LocalVariables.Add("Wt", paramInfo.Wtkg)

                Try
                    'Bug in parser - doesn't like the equation starting with a negative sign as first char in string
                    Eq = " (" & Eq & ")"
                    result = Parser.Parse(Eq)
                    Return CSng(result)
                Catch ex As Exception
                    MsgBox("Error: " & ex.Message)
                    Return vbEmpty
                End Try
            End If

        Catch ex As Exception
            MsgBox("Error in 'class_pred.ParseEquation'." & vbCrLf & ex.Message, vbOKOnly, "Unexpected error")
            Return vbEmpty
        End Try

    End Function

    Public Function ParseEquation_LMS(ByVal Demo As Pred_demo, ByVal Sp As LMS_Splines, ByVal Eq As LMS_Equations, ByVal Co As LMS_Coefficients, ByVal SourceID As Integer, ByVal StatType As String) As Single
        'https://github.com/MathosProject/Mathos-Parser

        Try

            Dim P As New Mathos.Parser.MathParser

            Select Case SourceID
                Case 2  'GLI 2014

                    Dim gBlack As Integer = 0, gNEAsia As Integer = 0, gSEAsia As Integer = 0, gOther As Integer = 0
                    Dim Gender As String = cMyRoutines.Lookup_list_ByCode(Demo.GenderID, eTables.Pred_ref_genders)
                    Select Case Gender
                        Case "Caucasian" : gBlack = 0 : gNEAsia = 0 : gSEAsia = 0 : gOther = 0
                        Case "NE Asian" : gNEAsia = 1
                        Case "SE Asian" : gSEAsia = 1
                        Case "African American" : gBlack = 1
                        Case "Other Mixed" : gOther = 1
                    End Select

                    'Create the parser variables
                    P.LocalVariables.Add("Age", Demo.Age)
                    P.LocalVariables.Add("Ht", Demo.Htcm)
                    P.LocalVariables.Add("Wt", Demo.Wtkg)
                    P.LocalVariables.Add("q0", Co.q0)
                    P.LocalVariables.Add("q1", Co.q1)
                    P.LocalVariables.Add("p0", Co.p0)
                    P.LocalVariables.Add("p1", Co.p1)
                    P.LocalVariables.Add("p2", Co.p2)
                    P.LocalVariables.Add("p3", Co.p3)
                    P.LocalVariables.Add("p4", Co.p4)
                    P.LocalVariables.Add("p5", Co.p5)
                    P.LocalVariables.Add("a0", Co.a0)
                    P.LocalVariables.Add("a1", Co.a1)
                    P.LocalVariables.Add("a2", Co.a2)
                    P.LocalVariables.Add("a3", Co.a3)
                    P.LocalVariables.Add("a4", Co.a4)
                    P.LocalVariables.Add("a5", Co.a5)
                    P.LocalVariables.Add("a6", Co.a6)
                    P.LocalVariables.Add("LSpline", Sp.LSpline)
                    P.LocalVariables.Add("mSpline", Sp.MSpline)
                    P.LocalVariables.Add("sSpline", Sp.SSpline)
                    P.LocalVariables.Add("Black", gBlack)
                    P.LocalVariables.Add("NEAsia", gNEAsia)
                    P.LocalVariables.Add("SEAsia", gSEAsia)
                    P.LocalVariables.Add("Other", gOther)

                    Try
                        'Bug in parser - doesn't like the equation starting with a negative sign as first char in string

                        Dim L As Double = P.Parse(" (" & Eq.L_equation & ")")
                        Dim M As Double = P.Parse(" (" & Eq.M_equation & ")")
                        Dim S As Double = P.Parse(" (" & Eq.S_equation & ")")
                        Dim LLN As Double = P.Parse(" (" & Eq.LLN & ")")

                        Select Case StatType
                            Case "MPV" : Return M
                            Case "LLN" : Return LLN
                            Case Else : Return vbEmpty
                        End Select

                    Catch ex As Exception
                        MsgBox("Error: " & ex.Message)
                        Return vbEmpty
                    End Try

                Case Else
                    Return vbEmpty
            End Select

        Catch ex As Exception
            MsgBox("Error in class_pred.ParseEquation_LMS" & vbCrLf & ex.Message)
            Return vbEmpty
        End Try

    End Function

    Public Function Calculate_LMS_NotMTParser(ByVal p As ParameterInfo_UseToCalc, ByVal Parameter As String, ByVal SourceID As Integer, ByVal StatType As String, maxAge As Single, minAge As Single) As Single


        Select Case SourceID
            Case 2  'GLI 2014
                If Parameter = "VC" Then
                    Parameter = "FVC"
                End If


                Dim gBlack As Integer = 0, gNEAsia As Integer = 0, gSEAsia As Integer = 0, gOther As Integer = 0
                Dim Gender As String = cMyRoutines.Lookup_list_ByCode(p.GenderID_ForCalc, eTables.Pred_ref_genders)
                Select Case p.Ethnicity_ForCalc
                    Case "Caucasian" : gBlack = 0 : gNEAsia = 0 : gSEAsia = 0 : gOther = 0
                    Case "NE Asian" : gNEAsia = 1
                    Case "SE Asian" : gSEAsia = 1
                    Case "African American" : gBlack = 1
                    Case "Other Mixed" : gOther = 1
                End Select

                Dim Sp As LMS_Splines = Me.Get_LMS_Splines(Parameter, p, maxAge, minAge)
                'Dim Eq As LMS_Equations = Me.Get_LMS_Equations(SourceID)
                Dim Co As LMS_Coefficients = Me.Get_LMS_Coefficients(SourceID, Parameter, Gender)

                Try
                    Dim L = Co.q0 + Co.q1 * Log(p.Age_ForCalc) + Sp.LSpline
                    Dim M = Exp(Co.a0 + Co.a1 * Log(p.Ht_ForCalc) + Co.a2 * Log(p.Age_ForCalc) + Co.a3 * gBlack + Co.a4 * gNEAsia + Co.a5 * gSEAsia + Co.a6 * gOther + Sp.MSpline)
                    Dim S = Exp(Co.p0 + Co.p1 * Log(p.Age_ForCalc) + Co.p2 * gBlack + Co.p3 * gNEAsia + Co.p4 * gSEAsia + Co.p5 * gOther + Sp.SSpline)
                    Dim LLN = Exp(Log(M) + Log(1 - 1.645 * L * S) / L)

                    If Parameter = "FER" Then   'Express FER as %
                        M = 100 * M
                        LLN = 100 * LLN
                    End If

                    Select Case StatType
                        Case "MPV" : Return M
                        Case "LLN" : Return LLN
                        Case Else : Return vbEmpty
                    End Select

                Catch ex As Exception
                    MsgBox("Error in 'class_pred.Calculate_LMS_NotMTParser'." & ex.Message, vbOKOnly, "Unexpected error")
                    Return vbEmpty
                End Try

            Case Else
                Return vbEmpty
        End Select

    End Function

    Public Function Calculate_LMS_NotMTParser1(ByVal Demo As Pred_demo, ByVal Parameter As String, ByVal SourceID As Integer, ByVal StatType As String, maxAge As Single, minAge As Single) As Single


        Select Case SourceID
            Case 2  'GLI 2014
                If Parameter = "VC" Then
                    Parameter = "FVC"
                End If


                Dim gBlack As Integer = 0, gNEAsia As Integer = 0, gSEAsia As Integer = 0, gOther As Integer = 0
                Dim Gender As String = cMyRoutines.Lookup_list_ByCode(Demo.GenderID, eTables.Pred_ref_genders)
                Select Case Demo.Ethnicity
                    Case "Caucasian" : gBlack = 0 : gNEAsia = 0 : gSEAsia = 0 : gOther = 0
                    Case "NE Asian" : gNEAsia = 1
                    Case "SE Asian" : gSEAsia = 1
                    Case "African American" : gBlack = 1
                    Case "Other Mixed" : gOther = 1
                End Select

                Dim Sp As LMS_Splines = Me.Get_LMS_Splines1(Parameter, Demo, maxAge, minAge)
                'Dim Eq As LMS_Equations = Me.Get_LMS_Equations(SourceID)
                Dim Co As LMS_Coefficients = Me.Get_LMS_Coefficients(SourceID, Parameter, Gender)

                Try
                    Dim L = Co.q0 + Co.q1 * Log(Demo.Age) + Sp.LSpline
                    Dim M = Exp(Co.a0 + Co.a1 * Log(Demo.Htcm) + Co.a2 * Log(Demo.Age) + Co.a3 * gBlack + Co.a4 * gNEAsia + Co.a5 * gSEAsia + Co.a6 * gOther + Sp.MSpline)
                    Dim S = Exp(Co.p0 + Co.p1 * Log(Demo.Age) + Co.p2 * gBlack + Co.p3 * gNEAsia + Co.p4 * gSEAsia + Co.p5 * gOther + Sp.SSpline)
                    Dim LLN = Exp(Log(M) + Log(1 - 1.645 * L * S) / L)

                    If Parameter = "FER" Then   'Express FER as %
                        M = 100 * M
                        LLN = 100 * LLN
                    End If

                    Select Case StatType
                        Case "MPV" : Return M
                        Case "LLN" : Return LLN
                        Case Else : Return vbEmpty
                    End Select

                Catch ex As Exception
                    MsgBox("Error in 'class_pred.Calculate_LMS_NotMTParser'." & ex.Message, vbOKOnly, "Unexpected error")
                    Return vbEmpty
                End Try

            Case Else
                Return vbEmpty
        End Select

    End Function

    Public Function Get_Pred_cpet_values(d As Pred_demo, method As class_Pred.eLoadNormalsMode) As Dictionary(Of String, String)
        'Returns empty string if no value


        Dim p As New Dictionary(Of String, String)  'Parameter|Stat, equation


        'Austin adult equations at 15/7/2016---------------------------------------------------------
        'HRmax
        Select Case d.GenderID
            Case 1 : p.Add("hrmax|mpv", 207 - 0.78 * d.Age) : p.Add("hrmax|lln", (207 - 0.78 * d.Age) - 26)
            Case 2 : p.Add("hrmax|mpv", 209 - 0.86 * d.Age) : p.Add("hrmax|lln", (209 - 0.86 * d.Age) - 23)
        End Select
        'VO2max
        Select Case d.GenderID
            Case 1 : p.Add("vo2max|mpv", 0.023 * d.Htcm - 0.031 * d.Age - 0.332 + 0.0117 * d.Wtkg) : p.Add("vo2max|lln", (0.023 * d.Htcm - 0.031 * d.Age - 0.332 + 0.0117 * d.Wtkg) - 1.24)
            Case 2 : p.Add("vo2max|mpv", 0.0158 * d.Htcm - 0.027 * d.Age + 0.207 + 0.009 * d.Wtkg) : p.Add("vo2max|lln", (0.0158 * d.Htcm - 0.027 * d.Age + 0.207 + 0.009 * d.Wtkg) - 0.7)
        End Select
        'VO2/kg_max
        p.Add("vo2kg_max|mpv", "---") : p.Add("vo2kg_max|lln", "---")
        'VCO2max
        Select Case d.GenderID
            Case 1 : p.Add("vco2max|mpv", "---") : p.Add("vco2max|lln", "---")
            Case 2 : p.Add("vco2max|mpv", "---") : p.Add("vco2max|lln", "---")
        End Select
        'O2pulseMax
        If Val(p("vo2max|mpv")) > 0 And Val(p("hrmax|mpv")) > 0 Then
            p.Add("o2pulsemax|mpv", 1000 * p("vo2max|mpv") / p("hrmax|mpv")) : p.Add("o2pulsemax|lln", "")
        Else
            p.Add("o2pulsemax|mpv", "---") : p.Add("o2pulsemax|lln", "---")
        End If
        'Wmax
        Select Case d.GenderID
            Case 1 : p.Add("wmax|mpv", 74.47 * 0.83 * (1 - 0.007 * d.Age) * (0.01 * d.Htcm) ^ 2.7) : p.Add("wmax|lln", "---")
            Case 2 : p.Add("wmax|mpv", 77.42 * 0.62 * (1 - 0.007 * d.Age) * (0.01 * d.Htcm) ^ 2.7) : p.Add("wmax|lln", "---")
        End Select
        'VEmax
        If d.FEV1 > 0 And d.FVC > 0 Then
            If d.FEV1 < 1.6 And d.FEV1 / d.FVC < 0.5 Then
                p.Add("vemax|mpv", 37.5 * d.FEV1) : p.Add("vemax|lln", "---")
            Else
                p.Add("vemax|mpv", 35 * d.FEV1) : p.Add("vemax|lln", "---")
            End If
        Else
            p.Add("vemax|mpv", "---") : p.Add("vemax|lln", "---")
        End If
        'VTmax
        If d.FVC > 0 Then
            p.Add("vtmax|mpv", 0.67 * d.FVC - 0.64) : p.Add("vtmax|lln", (0.67 * d.FVC - 0.64) - 1.64 * 0.54)
        Else
            p.Add("vtmax|mpv", "---") : p.Add("vtmax|lln", "---")
        End If
        '---------------------------------------------------------------------------------------------------------------

        Return p

    End Function


    Public Function Get_Pred_cpet_reference_equations(d As Pred_demo, method As class_Pred.eLoadNormalsMode) As Dictionary(Of String, String)

        Dim p As New Dictionary(Of String, String)  'Parameter|Stat, equation

        'Austin adult equations at 15/7/2016

        'VO2 vs Workload(watts)
        p.Add("vo2load|uln", "(0.0103 × Workload  + 0.62) + 0.236")
        p.Add("vo2load|lln", "(0.0103 × Workload  + 0.62) - 0.236")

        If d.Age >= 18 Then
            'Ve vs VO2
            Select Case d.GenderID
                Case eGenders.Male
                    p.Add("vevo2|uln", "1.7*vo2^3 + 3.6*vo2^2 + 5.5*vo2 + 14.2")
                    p.Add("vevo2|lln", "5.7*vo2^2 + 4.5*vo2 + 7.7")
                Case eGenders.Female
                    p.Add("vevo2|uln", "9.8*vo2^4 - 50.4*vo2^3 + 97.4*vo2^2 - 51.8*vo2 + 19.6")
                    p.Add("vevo2|lln", "7.1*vo2^4 - 43.1*vo2^3 + 94.3*vo2^2 - 58.6*vo2 + 20.0")
                Case eGenders.NotKnown
                    p.Add("vevo2|uln", "")
                    p.Add("vevo2|lln", "")
            End Select

            'HR vs VO2
            Select Case d.GenderID
                Case eGenders.Male
                    p.Add("hrvo2|uln", "(65.64 + 35.46*vo2) + 15")
                    p.Add("hrvo2|lln", "(65.64 + 35.46*vo2) - 15")
                Case eGenders.Female
                    p.Add("hrvo2|uln", "(69.93 + 50.2*vo2) + 14")
                    p.Add("hrvo2|lln", "(69.93 + 50.2*vo2) - 14")
                Case eGenders.NotKnown
                    p.Add("hrvo2|uln", "")
                    p.Add("hrvo2|lln", "")
            End Select
        Else
            p.Add("vevo2|uln", "")
            p.Add("vevo2|lln", "")
            p.Add("hrvo2|uln", "")
            p.Add("hrvo2|lln", "")
        End If

        Return p

    End Function

End Class