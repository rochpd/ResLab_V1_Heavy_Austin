Public Class class_challenge

    Private _Height As Integer
    Private _Width As Integer
    Private _p As ProtocolData

    Public Structure DoseData
        Public testdataID As Integer
        Public doseID As Integer
        Public time_min As Integer
        Public canskip As Boolean
        Public dosenumber As Integer
        Public dose_cumulative As String
        Public dose_discrete As String
        Public xaxislabel As String
        Public result As String
        Public response As String
    End Structure

    Public Structure ProtocolData
        Public protocolID As Integer
        Public title As String
        Public description As String
        Public parameter As String
        Public parameter_units As String
        Public parameter_response As String
        Public agent As String
        Public agent_units As String
        Public post_drug As String
        Public pd As String
        Public pd_thresh As String
        Public pd_decimalplaces As String
        Public pd_dose_effect As String
        Public pd_method As String
        Public pd_method_reference As String
        Public plot_title As String
        Public plot_xscaling_type As String
        Public plot_xtitle As String
        Public plot_ymin As String
        Public plot_ymax As String
        Public plot_ystep As String
        Public doses() As DoseData
        Public provid As Long
    End Structure
    Public Structure ProtocolMenuData
        Public Label As String
        Public protocolID As Integer
    End Structure

    Private Function yS(ByVal y As Single) As Single
        'Converts Y in user units to Y in pixels

        Dim yFac As Single = _Height / 100
        Dim yOff As Single = 30
        Dim scaledY As Single = (130 - y) * yFac
        Return scaledY

    End Function

    Private Function xS(ByVal x As Single) As Single

        Dim xFac As Single = _Width / 130
        Dim xOff As Single = -10
        Dim scaledX = (x - xOff) * xFac
        Return scaledX

    End Function

    Public Function load_provdataStructure_from_dictionaries(ByVal T As Dictionary(Of String, String), ByVal TD() As Dictionary(Of String, String)) As class_challenge.ProtocolData
        'T contains the once only test data
        'TD contains an array of the dose data

        If IsNothing(T) Or IsNothing(TD) Then
            'should never happen
            Return Nothing
        Else
            Dim p As class_challenge.ProtocolData = Nothing
            Dim f1 As New class_fields_ProvAndSession
            Dim f2 As New class_ProvTestDataFields

            p.protocolID = T(f1.ProtocolID)
            p.title = T(f1.Protocol_title)
            p.pd_thresh = T(f1.Protocol_threshold) & ""
            p.pd_decimalplaces = T(f1.Protocol_pd_decimalplaces) & ""
            p.pd_method = T(f1.Protocol_method) & ""
            p.agent_units = T(f1.Protocol_doseunits) & ""
            p.agent = T(f1.Protocol_drug) & ""
            p.parameter = T(f1.Protocol_parameter) & ""
            p.parameter_units = T(f1.Protocol_parameter_units) & ""
            p.parameter_response = T(f1.Protocol_parameter_response) & ""
            p.pd_dose_effect = T(f1.Protocol_dose_effect) & ""
            p.pd_method_reference = T(f1.Protocol_method_reference) & ""
            p.post_drug = T(f1.Protocol_post_drug) & ""

            p.plot_title = T(f1.plot_title) & ""
            p.plot_xtitle = T(f1.plot_xtitle) & ""
            p.plot_ymax = T(f1.plot_ymax) & ""
            p.plot_ymin = T(f1.plot_ymin) & ""
            p.plot_ystep = T(f1.plot_ystep) & ""
            p.plot_xscaling_type = T(f1.plot_xscaling_type) & ""

            Dim i As Integer = 0
            Dim d As Dictionary(Of String, String)
            ReDim p.doses(TD.Count - 1)
            For Each d In TD
                p.doses(i).canskip = CBool(Val(d(f2.dose_canskip)))
                p.doses(i).dose_cumulative = d(f2.dose_cumulative) & ""
                p.doses(i).dose_discrete = d(f2.dose_discrete) & ""
                p.doses(i).testdataID = d(f2.testdataid)
                p.doses(i).doseID = d(f2.doseid)
                p.doses(i).dosenumber = d(f2.dose_number) & ""
                If d(f2.dose_time_min) & "" = "" Then p.doses(i).time_min = vbEmpty Else p.doses(i).time_min = d(f2.dose_time_min)
                p.doses(i).xaxislabel = d(f2.xaxis_label) & ""
                p.doses(i).result = d(f2.result) & ""
                p.doses(i).response = d(f2.response) & ""
                i = i + 1
            Next
            Return p
        End If

    End Function

    Public Function Get_Protocols(ByVal RestrictToEnabled As Boolean) As ProtocolMenuData()
        'Returns data to populate main menu list of available protocols

        Dim sql As String = ""
        Dim p() As ProtocolMenuData
        Dim i As Integer = 0

        Select Case RestrictToEnabled
            Case True : sql = "SELECT protocolID, p_menulabel FROM prov_protocols WHERE enabled=1"
            Case False : sql = "SELECT protocolID, p_menulabel FROM prov_protocols"
        End Select
        sql = sql & " ORDER BY p_menulabel"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            ReDim p(Ds.Tables(0).Rows.Count - 1)
            i = 0
            For Each r As DataRow In Ds.Tables(0).Rows
                p(i).protocolID = r("protocolID")
                p(i).Label = r("p_menulabel")
                i = i + 1
            Next
            Return p
        Else
            Return Nothing
        End If

    End Function

    Public Function Get_ProtocolProperties(ByVal ProtocolID As Integer) As ProtocolData

        Dim p As ProtocolData = Nothing
        Dim f1 As New class_fields_ProvAndSession
        Dim f2 As New class_ProvTestDataFields

        Dim sql As String = "SELECT * FROM prov_protocols WHERE protocolID=" & ProtocolID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            For Each r As DataRow In Ds.Tables(0).Rows
                p.title = r(f1.Protocol_title) & ""
                p.agent = r(f1.Protocol_drug) & ""
                p.agent_units = r(f1.Protocol_doseunits) & ""
                p.parameter = r(f1.Protocol_parameter) & ""
                p.parameter_response = r(f1.Protocol_parameter_response) & ""
                p.parameter_units = r(f1.Protocol_parameter_units) & ""
                p.pd_method = r(f1.Protocol_method) & ""
                p.pd_method_reference = r(f1.Protocol_method_reference) & ""
                p.pd_thresh = r(f1.Protocol_threshold) & ""
                p.pd_decimalplaces = r(f1.Protocol_pd_decimalplaces) & ""
                p.pd_dose_effect = r(f1.Protocol_dose_effect) & ""
                p.plot_title = r(f1.plot_title) & ""
                p.plot_xtitle = r(f1.plot_xtitle) & ""
                p.plot_ymax = r(f1.plot_ymax) & ""
                p.plot_ymin = r(f1.plot_ymin) & ""
                p.plot_ystep = r(f1.plot_ystep) & ""
                p.plot_xscaling_type = r(f1.plot_xscaling_type) & ""
                p.protocolID = r(f1.ProtocolID)
            Next
            Ds = Nothing

            sql = "SELECT * FROM prov_protocol_doseschedule WHERE protocolID=" & ProtocolID
            Ds = cDAL.Get_DataAsDataset(sql)
            If Ds.Tables(0).Rows.Count > 0 Then
                ReDim p.doses(Ds.Tables(0).Rows.Count - 1)
                Dim i As Integer = 0
                For Each r As DataRow In Ds.Tables(0).Rows
                    p.doses(i).canskip = r(f2.dose_canskip)
                    p.doses(i).dose_cumulative = r(f2.dose_cumulative) & ""
                    p.doses(i).dose_discrete = r(f2.dose_discrete) & ""
                    p.doses(i).doseID = r(f2.doseid)
                    p.doses(i).dosenumber = r(f2.dose_number) & ""
                    If r(f2.dose_time_min) & "" = "" Then p.doses(i).time_min = vbEmpty Else p.doses(i).time_min = r(f2.dose_time_min)
                    p.doses(i).xaxislabel = r(f2.xaxis_label) & ""
                    i = i + 1
                Next
                Ds = Nothing
            End If

            Return p
        Else
            Return Nothing
        End If

    End Function

    Public Function Draw_ProvocationGraph(ByVal P As class_challenge.ProtocolData, ByVal Ht As Integer, ByVal Width As Integer, ByVal fnt As Font) As Bitmap
        'Assume:
        ' 1 x baseline result (first in array)
        ' 1 x control dose = 0 (second)
        ' 1 x post result (last)
        ' variable number of doses per protocol
        ' response is pre-calculated and passed in (can be referenced to either baseline or control)
        Try

            _Width = Width
            _Height = Ht
            Dim flds As New class_challenge.ProtocolData
            Dim i As Integer
            Dim ymin As Integer = CInt(P.plot_ymin)
            Dim ymax As Integer = CInt(P.plot_ymax)
            Dim ystep As Integer = CInt(P.plot_ystep)
            Dim xstep As Single = 90 / (P.doses.Count - 3) 'x axis step for body of graph

            Dim img = New Bitmap(Width, Ht)
            Dim g As Graphics = Graphics.FromImage(img)
            Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Black, 0.1)

            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
            g.InterpolationMode = Drawing2D.InterpolationMode.High

            'Numeric scaling - doesn't operate on text
            g.ScaleTransform(Width / 130, -Ht / 100)
            g.TranslateTransform(10, -130)

            'Do horizontal grid
            'myPen.DashStyle = Drawing2D.DashStyle.Dot
            'For i = ymin To ymax Step ystep
            'g.DrawLine(myPen, 10, i, 100, i)
            'Next

            'Do horizontal lines
            myPen.DashStyle = Drawing2D.DashStyle.Dash
            For i = ymin To ymax Step ystep
                g.DrawLine(myPen, 0, i, 1, i)
                g.DrawLine(myPen, 10, i, 100, i)
                g.DrawLine(myPen, 110, i, 109, i)
            Next

            'Draw vertical lines
            Dim X, Y As Single
            g.DrawLine(myPen, 0, 50, 0, 120)
            g.DrawLine(myPen, 110, ymin, 110, ymax)
            For i = 1 To P.doses.Count - 2
                X = 10 + (i - 1) * xstep
                g.DrawLine(myPen, X, ymin, X, ymax)
            Next

            'Labels
            TextRenderer.DrawText(g, P.plot_title, fnt, New Point(xS(40), yS(128)), Color.Black)
            TextRenderer.DrawText(g, P.plot_xtitle, fnt, New Point(xS(40), yS(40)), Color.Black)
            TextRenderer.DrawText(g, P.doses(0).xaxislabel, fnt, New Point(xS(-2), yS(48)), Color.Black)
            TextRenderer.DrawText(g, P.doses(UBound(P.doses)).xaxislabel, fnt, New Point(xS(106), yS(48)), Color.Black)
            For i = ymin To ymax Step ystep
                TextRenderer.DrawText(g, i.ToString, fnt, New Point(xS(-8), yS(i + 2)), Color.Black)
            Next

            'Dim fnt_xaxislabels As New Font(fnt.Name, 8, FontStyle.Bold)
            For i = 1 To P.doses.Count - 2
                X = 10 + (i - 1) * xstep
                TextRenderer.DrawText(g, P.doses(i).xaxislabel, fnt, New Point(xS(X - 2), yS(48)), Color.Black)
            Next

            'Plot the significant fall line
            myPen.Color = Color.Blue : myPen.Width = 0.5 : myPen.DashStyle = Drawing2D.DashStyle.Dash
            g.DrawLine(myPen, 10, 100 - CInt(P.pd_thresh), 100, 100 - CInt(P.pd_thresh))

            'Plot the data
            Dim fnt_plot As New Font(fnt.Name, fnt.SizeInPoints, FontStyle.Bold)
            myPen.Color = Color.Red : myPen.Width = 0.5 : myPen.DashStyle = Drawing2D.DashStyle.Solid
            Dim xbit = -1.4, ybit = 2.8
            If P.doses(1).result <> "" Then

                'Plot b/l and post points
                If P.doses(0).response <> "" Then TextRenderer.DrawText(g, "o", fnt_plot, New Point(xS(0 + xbit), yS(Val(P.doses(0).response) + ybit)), Color.Red)
                If P.doses(UBound(P.doses)).response <> "" Then TextRenderer.DrawText(g, "o", fnt_plot, New Point(xS(110 + xbit), yS(Val(P.doses(UBound(P.doses)).response) + ybit)), Color.Red)

                'Plot body of graph
                Dim prevX As Single = 10
                Dim prevY As Single = P.doses(1).response

                For i = 1 To P.doses.Count - 2
                    If P.doses(i).response <> "" Then
                        X = (i - 1) * xstep + 10
                        Y = Val(P.doses(i).response)
                        TextRenderer.DrawText(g, "o", fnt_plot, New Point(xS(X + xbit), yS(Y + ybit)), Color.Red)
                        g.DrawLine(myPen, prevX, prevY, X, Y)
                        prevX = X
                        prevY = Y
                    End If
                Next i

                'Connect last dose result to postBD result
                If P.doses(UBound(P.doses)).response <> "" Then
                    myPen.DashStyle = Drawing2D.DashStyle.Dash
                    g.DrawLine(myPen, X, Y, 110, CSng(P.doses(UBound(P.doses)).response))
                End If

                'Connect baseline with first dose result 
                g.DrawLine(myPen, 0, CSng(P.doses(0).response), 10, CSng(P.doses(1).response))

            End If

            Return img

        Catch
            MsgBox("Error in Draw_provocationGraph" & vbNewLine & Err.Description)
            Return Nothing

        End Try

    End Function

    Public Function Calculate_PDx(ByVal P As class_challenge.ProtocolData) As String
        'Calculates the provocative dose for a given percent fall

        Dim FallOfX As Boolean
        Dim i As Integer
        Dim dose_After As Integer
        Dim dose_last As Integer
        Dim dose(UBound(P.doses)) As String
        Dim pd As String
        Dim A As Single, B As Single, C As Single, D As Single, M As Single
        Dim flds As New class_fields_ProvAndSession

        'Get thresholds, exit if no good
        Dim Thresh_fall As Single = Val(P.pd_thresh)
        Dim Thresh_level As Single = 0
        Select Case P.pd_method_reference
            Case "Control"
                If P.doses(1).response = "" Then pd = "" : GoTo outahere
                Thresh_level = Val(P.doses(1).response) - Thresh_fall
            Case "Baseline"
                If P.doses(0).response = "" Then pd = "" : GoTo outahere
                Thresh_level = Val(P.doses(0).response) - Thresh_fall
        End Select

        'Set up doses array
        For i = 0 To UBound(P.doses)
            If P.pd_dose_effect = "Cumulative" Then
                dose(i) = P.doses(i).dose_cumulative
            ElseIf P.pd_dose_effect = "Discrete" Then
                dose(i) = P.doses(i).dose_discrete
            Else
                MsgBox("Can't calculate PD, problem with dose values")
                pd = "" : GoTo outahere
            End If
        Next

        'Find the last dose given
        dose_last = 1
        For i = UBound(dose) - 1 To 1 Step -1
            If P.doses(i).response <> "" Then
                dose_last = i
                Exit For
            End If
        Next
        If dose_last = 1 Then pd = "" : GoTo outahere

        'Find the doses that bracket the defined fall in FEV1 - begin with first dose <= thresh_level
        FallOfX = False
        For i = 1 To dose_last
            If P.doses(i).response <> "" Then   'skip skipped doses
                If Val(P.doses(i).response) = Thresh_level Then
                    pd = "= " & dose(i)
                    GoTo outahere
                ElseIf Val(P.doses(i).response) < Thresh_level Then
                    FallOfX = True
                    dose_After = i
                    Exit For
                End If
            End If
        Next

        'If defined fall not achieved
        If Not FallOfX Then
            pd = "> " & dose(dose_last)
            GoTo outahere
        Else
            Select Case P.pd_method_reference   'If fall after first dose
                Case "Control" : If dose_After = 2 Then pd = "< " & dose(dose_After) : GoTo outahere
                Case "Baseline" : If dose_After = 1 Then pd = "< " & dose(dose_After) : GoTo outahere
            End Select

            'Interpolate PD
            A = Math.Log10(Val(dose(dose_After))) / Math.Log10(10)
            C = Math.Log10(Val(dose(dose_After - 1))) / Math.Log10(10)
            B = Val((P.doses(dose_After).response))
            D = Val((P.doses(dose_After - 1).response))
            M = (B - D) / (A - C)
            pd = cMyRoutines.fmt(10 ^ ((Thresh_level - (B - M * A)) / M), CInt(P.pd_decimalplaces))
            GoTo outahere
        End If

outahere: Return pd & " " & P.agent_units

    End Function


End Class
