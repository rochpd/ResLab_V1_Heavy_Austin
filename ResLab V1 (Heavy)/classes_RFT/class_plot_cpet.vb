Imports System.Drawing
Imports System.Drawing.Printing
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms.DataVisualization.Charting

Public Class class_plot_cpet

    Public Enum eCpetPlots
        Ve_VO2
        HrSpO2_Vo2


    End Enum

    Public Structure plot_referenceline
        Public xData() As Single
        Public yData() As Single
        Public linestyle As Drawing2D.DashStyle
    End Structure
    Public Structure plot_reference_equation
        Public equation As String
        Public xParameter As String
        Public yParameter As String
        Public linestyle As Drawing2D.DashStyle
    End Structure

    Public Structure cpet_plot_properties
        Public plot_width As Integer
        Public plot_height As Integer

        Public plot_yAxesCount As Integer
        Public plot_font As Font

        Public xData() As Single
        Public y1_1_Data() As Single
        Public y1_2_Data() As Single
        Public y2Data() As Single

        Public xAxis_label As String
        Public xAxis_scale_auto As Boolean
        Public xAxis_min As String
        Public xAxis_max As String
        Public xAxis_tickinterval As String

        Public y1Axis_label_1 As String
        Public y1Axis_label_2 As String
        Public y1Axis_scale_auto As Boolean
        Public y1Axis_min As String
        Public y1Axis_max As String
        Public y1Axis_tickinterval As String
        Public y1Axis_symbol_1 As MarkerStyle
        Public y1Axis_symbolcolour_1 As Color
        Public y1Axis_symbol_2 As MarkerStyle
        Public y1Axis_symbolcolour_2 As Color
        Public y1Axis_plot_ref_lines() As plot_referenceline
        Public y1Axis_plot_ref_equations() As plot_reference_equation

        Public y2Axis_enabled As Boolean
        Public y2Axis_label As String
        Public y2Axis_scale_auto As Boolean
        Public y2Axis_min As String
        Public y2Axis_max As String
        Public y2Axis_tickinterval As String
        Public y2Axis_symbol As MarkerStyle
        Public y2Axis_symbolcolour As Color
        Public y2Axis_plot_ref_lines() As plot_referenceline
        Public y2Axis_plot_ref_equations() As plot_reference_equation

    End Structure

    Public Function Get_plotproperties_VeVO2(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties
        Dim j As Integer = 0

        p.plot_yAxesCount = 1
        p.y2Axis_enabled = False

        p.xAxis_label = "VO2 (L/min)"
        p.xAxis_min = 0
        p.xAxis_max = 4
        p.xAxis_tickinterval = 1
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "VE (L/min)"
        p.y1Axis_min = 0
        p.y1Axis_max = 160
        p.y1Axis_tickinterval = 40
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        If Not IsNothing(pred_values) Then
            j = 0
            'Add VE max line
            If Val(pred_values("vemax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, pred_values("vo2max|mpv")}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("vemax|mpv"), pred_values("vemax|mpv")}
                j = j + 1
            End If
            'Add VO2 max line
            If Val(pred_values("vo2max|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                If IsNumeric(pred_values("vo2max|mpv")) Then p.y1Axis_plot_ref_lines(j).xData = {pred_values("vo2max|mpv"), pred_values("vo2max|mpv")}
                If IsNumeric(pred_values("vemax|mpv")) Then p.y1Axis_plot_ref_lines(j).yData = {p.y1Axis_min, pred_values("vemax|mpv")}
                j = j + 1
            End If

            'Get ref equations
            Dim e As Dictionary(Of String, String) = cPred.Get_Pred_cpet_reference_equations(pred_demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)
            ReDim Preserve p.y1Axis_plot_ref_equations(0 To 1)

            p.y1Axis_plot_ref_equations(0).equation = e("vevo2|uln")
            p.y1Axis_plot_ref_equations(0).xParameter = "vo2"
            p.y1Axis_plot_ref_equations(0).yParameter = "ve"
            p.y1Axis_plot_ref_equations(0).linestyle = Drawing2D.DashStyle.Solid

            p.y1Axis_plot_ref_equations(1).equation = e("vevo2|lln")
            p.y1Axis_plot_ref_equations(1).xParameter = "vo2"
            p.y1Axis_plot_ref_equations(1).yParameter = "ve"
            p.y1Axis_plot_ref_equations(1).linestyle = Drawing2D.DashStyle.Solid

        End If

        Return p

    End Function

    Public Function Get_plotproperties_VeVCO2(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties
        Dim j As Integer = 0

        p.plot_yAxesCount = 1
        p.y2Axis_enabled = False

        p.xAxis_label = "VCO2 (L/min)"
        p.xAxis_min = 0
        p.xAxis_max = 4
        p.xAxis_tickinterval = 1
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "VE (L/min)"
        p.y1Axis_min = 0
        p.y1Axis_max = 160
        p.y1Axis_tickinterval = 40
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        If Not IsNothing(pred_values) Then
            j = 0
            'Add VE max line
            If Val(pred_values("vemax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, p.xAxis_max}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("vemax|mpv"), pred_values("vemax|mpv")}
                j = j + 1
            End If
        End If

        Return p

    End Function

    Public Function Get_plotproperties_VtVe(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties
        Dim j As Integer = 0

        p.plot_yAxesCount = 1
        p.y2Axis_enabled = False

        p.xAxis_label = "VE (L/min)"
        p.xAxis_min = 0
        p.xAxis_max = 160
        p.xAxis_tickinterval = 40
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "Vt (L)"
        p.y1Axis_min = 0
        p.y1Axis_max = 3.0
        p.y1Axis_tickinterval = 1.0
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        If Not IsNothing(pred_values) Then
            j = 0
            'Add VE max line
            If Val(pred_values("vemax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {pred_values("vemax|mpv"), pred_values("vemax|mpv")}
                p.y1Axis_plot_ref_lines(j).yData = {p.y1Axis_min, p.y1Axis_max}
                j = j + 1
            End If
        End If

        Return p

    End Function

    Public Function Get_plotproperties_ve_time(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 1
        p.y2Axis_enabled = False

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 15
        p.xAxis_tickinterval = 3
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "VE (L/min)"
        p.y1Axis_min = 0
        p.y1Axis_max = 160
        p.y1Axis_tickinterval = 40
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        If Not IsNothing(pred_values) Then
            Dim j As Integer = 0
            'Add VE max line
            If Val(pred_values("vemax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, p.xAxis_max}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("vemax|mpv"), pred_values("vemax|mpv")}
                j = j + 1
            End If
        End If

        Return p

    End Function

    Public Function Get_plotproperties_HrSPo2_VO2(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 2
        p.y2Axis_enabled = True

        p.xAxis_label = "VO2 (L/min)"
        p.xAxis_min = 0
        p.xAxis_max = 4
        p.xAxis_tickinterval = 1
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "HR (bpm)"
        p.y1Axis_min = 0
        p.y1Axis_max = 200
        p.y1Axis_tickinterval = 40
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        p.y2Axis_label = "SpO2 (%)"
        p.y2Axis_min = 0
        p.y2Axis_max = 100
        p.y2Axis_tickinterval = 20
        p.y2Axis_scale_auto = False
        p.y2Axis_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.y2Axis_symbolcolour = Color.Red

        If Not IsNothing(pred_values) Then
            Dim j As Integer = 0
            'Add HR max line
            If Val(pred_values("hrmax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, pred_values("vo2max|mpv")}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("hrmax|mpv"), pred_values("hrmax|mpv")}
                j = j + 1
            End If
            'Add VO2 max line
            If Val(pred_values("vo2max|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {pred_values("vo2max|mpv"), pred_values("vo2max|mpv")}
                p.y1Axis_plot_ref_lines(j).yData = {p.y1Axis_min, pred_values("hrmax|mpv")}
                j = j + 1
            End If

            'Get ref equations
            Dim e As Dictionary(Of String, String) = cPred.Get_Pred_cpet_reference_equations(pred_demo, class_Pred.eLoadNormalsMode.UseSourcesInUseAtTestDate)
            ReDim Preserve p.y1Axis_plot_ref_equations(0 To 1)

            p.y1Axis_plot_ref_equations(0).equation = e("hrvo2|uln")
            p.y1Axis_plot_ref_equations(0).xParameter = "vo2"
            p.y1Axis_plot_ref_equations(0).yParameter = "hr"
            p.y1Axis_plot_ref_equations(0).linestyle = Drawing2D.DashStyle.Solid

            p.y1Axis_plot_ref_equations(1).equation = e("hrvo2|lln")
            p.y1Axis_plot_ref_equations(1).xParameter = "vo2"
            p.y1Axis_plot_ref_equations(1).yParameter = "hr"
            p.y1Axis_plot_ref_equations(1).linestyle = Drawing2D.DashStyle.Solid
        End If

        Return p

    End Function

    Public Function Get_plotproperties_Hr_O2Pulse_time(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 2
        p.y2Axis_enabled = True

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 15
        p.xAxis_tickinterval = 3
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "HR (bpm)"
        p.y1Axis_min = 0
        p.y1Axis_max = 200
        p.y1Axis_tickinterval = 40
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        p.y2Axis_label = "O2 PULSE (ml/beat)"
        p.y2Axis_min = 0
        p.y2Axis_max = 30
        p.y2Axis_tickinterval = 5
        p.y2Axis_scale_auto = False
        p.y2Axis_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.y2Axis_symbolcolour = Color.Red

        If Not IsNothing(pred_values) Then
            Dim j As Integer = 0
            'Add HR max line
            If Val(pred_values("hrmax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, 0.9 * (p.xAxis_max - p.xAxis_min) + p.xAxis_min}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("hrmax|mpv"), pred_values("hrmax|mpv")}
                j = j + 1
            End If

        End If

        Return p

    End Function

    Public Function Get_plotproperties_Hr_VCO2_VO2(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String), vo2_rer_min As Single, vco2_rer_min As Single) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 2
        p.y2Axis_enabled = True

        p.xAxis_label = "VO2 (L/min)"
        p.xAxis_min = 0
        p.xAxis_max = 4
        p.xAxis_tickinterval = 1
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "HR (bpm)"
        p.y1Axis_min = 0
        p.y1Axis_max = 200
        p.y1Axis_tickinterval = 40
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        p.y2Axis_label = "VCO2 (L/min)"
        p.y2Axis_min = 0
        p.y2Axis_max = 4
        p.y2Axis_tickinterval = 1
        p.y2Axis_scale_auto = False
        p.y2Axis_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.y2Axis_symbolcolour = Color.Red

        If Not IsNothing(pred_values) Then
            Dim j As Integer = 0
            'Add HR max line
            If Val(pred_values("hrmax|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, 0.9 * (p.xAxis_max - p.xAxis_min) + p.xAxis_min}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("hrmax|mpv"), pred_values("hrmax|mpv")}
                j = j + 1
            End If
        End If

        'Do v-slope ref line
        Dim gradient = 1
        Dim yIntercept = vco2_rer_min - vo2_rer_min



        Return p

    End Function

    Public Function Get_plotproperties_RER_time(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 1
        p.y2Axis_enabled = False

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 15
        p.xAxis_tickinterval = 3
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "RER"
        p.y1Axis_min = 0
        p.y1Axis_max = 1.6
        p.y1Axis_tickinterval = 0.2
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        Return p

    End Function

    Public Function Get_plotproperties_vo2_vco2_work_time(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 2
        p.y2Axis_enabled = True

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 15
        p.xAxis_tickinterval = 3
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "VO2"
        p.y1Axis_label_2 = "VCO2"
        p.y1Axis_min = 0
        p.y1Axis_max = 4
        p.y1Axis_tickinterval = 1
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue
        p.y1Axis_symbol_2 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_2 = Color.Red

        p.y2Axis_label = "WORK (watt)"
        p.y2Axis_min = 0
        p.y2Axis_max = 400
        p.y2Axis_tickinterval = 100
        p.y2Axis_scale_auto = False
        p.y2Axis_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.y2Axis_symbolcolour = Color.Black

        If Not IsNothing(pred_values) Then
            Dim j As Integer = 0
            'Add VO2 max line
            If Val(pred_values("vo2max|mpv")) > 0 Then
                ReDim Preserve p.y1Axis_plot_ref_lines(j)
                p.y1Axis_plot_ref_lines(j).xData = {p.xAxis_min, 0.9 * (p.xAxis_max - p.xAxis_min) + p.xAxis_min}
                p.y1Axis_plot_ref_lines(j).yData = {pred_values("vo2max|mpv"), pred_values("vo2max|mpv")}
                j = j + 1
            End If
        End If


        Return p

    End Function

    Public Function Get_plotproperties_vevo2_vevco2_time(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 2
        p.y2Axis_enabled = True

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 15
        p.xAxis_tickinterval = 3
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "VEVO2"
        p.y1Axis_min = 0
        p.y1Axis_max = 80
        p.y1Axis_tickinterval = 20
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue

        p.y2Axis_label = "VEVCO2"
        p.y2Axis_min = 0
        p.y2Axis_max = 80
        p.y2Axis_tickinterval = 20
        p.y2Axis_scale_auto = False
        p.y2Axis_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.y2Axis_symbolcolour = Color.Red

        Return p

    End Function

    Public Function Get_plotproperties_peto2_petco2_spo2_time(pred_demo As Pred_demo, pred_values As Dictionary(Of String, String)) As cpet_plot_properties

        Dim p As New cpet_plot_properties

        p.plot_yAxesCount = 2
        p.y2Axis_enabled = True

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 15
        p.xAxis_tickinterval = 3
        p.xAxis_scale_auto = False

        p.y1Axis_label_1 = "PetO2"
        p.y1Axis_label_2 = "PetCO2"
        p.y1Axis_min = 0
        p.y1Axis_max = 140
        p.y1Axis_tickinterval = 20
        p.y1Axis_scale_auto = False
        p.y1Axis_symbol_1 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_1 = Color.MediumBlue
        p.y1Axis_symbol_2 = DataVisualization.Charting.MarkerStyle.Circle
        p.y1Axis_symbolcolour_2 = Color.Red

        p.y2Axis_label = "SpO2"
        p.y2Axis_min = 75
        p.y2Axis_max = 100
        p.y2Axis_tickinterval = 5
        p.y2Axis_scale_auto = False
        p.y2Axis_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.y2Axis_symbolcolour = Color.Black

        Return p

    End Function

    Public Sub Remove_plot(ch As Chart)
        If Not IsNothing(ch) Then ch = Nothing
    End Sub

    Public Function Create_plot(pred_demo As Pred_demo, p As cpet_plot_properties) As System.Windows.Forms.DataVisualization.Charting.Chart

        Dim fnt1 As New Font("Arial", 7, FontStyle.Regular)
        Dim newSeries As Series
        Dim ch = New System.Windows.Forms.DataVisualization.Charting.Chart
        Dim ChartArea1 = New ChartArea
        Dim i As Integer, j As Integer

        ch.ChartAreas.Add(ChartArea1)
        ch.Dock = System.Windows.Forms.DockStyle.Fill
        ch.Location = New System.Drawing.Point(0, 0)
        ch.Name = "cpetGraph"

        ch.ChartAreas.Clear()
        ch.Series.Clear()
        ch.Legends.Clear()
        ch.Titles.Clear()

        'Do y1_1 series________________________________________________
        newSeries = New Series
        newSeries.YAxisType = AxisType.Primary
        newSeries.Name = "series_y1_1"
        newSeries.ChartType = SeriesChartType.Point
        newSeries.Points.DataBindXY(p.xData, p.y1_1_Data)
        newSeries.MarkerStyle = p.y1Axis_symbol_1
        newSeries.MarkerColor = Color.White
        newSeries.MarkerBorderWidth = 1
        newSeries.MarkerBorderColor = p.y1Axis_symbolcolour_1
        newSeries.MarkerSize = 4
        ch.Series.Add(newSeries)

        'Do y1_2 series________________________________________________
        If p.y1Axis_label_2 <> "" Then
            newSeries = New Series
            newSeries.YAxisType = AxisType.Primary
            newSeries.Name = "series_y1_2"
            newSeries.ChartType = SeriesChartType.Point
            newSeries.Points.DataBindXY(p.xData, p.y1_2_Data)
            newSeries.MarkerStyle = p.y1Axis_symbol_2
            newSeries.MarkerColor = Color.White
            newSeries.MarkerBorderWidth = 1
            newSeries.MarkerBorderColor = p.y1Axis_symbolcolour_2
            newSeries.MarkerSize = 4
            ch.Series.Add(newSeries)
        End If

        'Draw ref lines 
        If Not IsNothing(p.y1Axis_plot_ref_lines) Then
            For i = 0 To UBound(p.y1Axis_plot_ref_lines)
                If Not (IsNothing(p.y1Axis_plot_ref_lines(i).xData) Or IsNothing(p.y1Axis_plot_ref_lines(i).yData)) Then
                    newSeries = New Series
                    newSeries.ChartType = SeriesChartType.Line
                    newSeries.MarkerStyle = MarkerStyle.None
                    newSeries.Points.DataBindXY(p.y1Axis_plot_ref_lines(i).xData, p.y1Axis_plot_ref_lines(i).yData)
                    For j = 0 To newSeries.Points.Count - 1
                        newSeries.Points(j).Color = p.y1Axis_symbolcolour_1
                        newSeries.Points(j).BorderDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                        newSeries.Points(j).BorderWidth = 1
                    Next
                    ch.Series.Add(newSeries)
                End If
            Next
        End If

        'Draw ref equations
        If Not IsNothing(p.y1Axis_plot_ref_equations) Then
            For i = 0 To UBound(p.y1Axis_plot_ref_equations)
                newSeries = New Series
                newSeries.ChartType = SeriesChartType.Point
                newSeries.MarkerStyle = MarkerStyle.Circle
                newSeries.MarkerSize = 1

                'Create data series from equation
                Dim xmin_plot As Single = 0, xmax_plot As Single = 0 ', ymin_plot As Single = 0, ymax_plot As Single = 0
                Dim x As New List(Of Single), y As New List(Of Single)
                Dim xValue As Single, yValue As Single
                xmin_plot = p.xAxis_min + 0.05 * (p.xAxis_max - p.xAxis_min)
                xmax_plot = p.xAxis_max - 0.05 * (p.xAxis_max - p.xAxis_min)

                For xValue = xmin_plot To xmax_plot Step (xmax_plot - xmin_plot) / 200
                    pred_demo.vo2 = xValue
                    yValue = cPred.ParseEquation(pred_demo, p.y1Axis_plot_ref_equations(i).equation)
                    If yValue > p.y1Axis_max Then Exit For
                    x.Add(xValue)
                    y.Add(yValue)
                Next

                newSeries.Points.DataBindXY(x, y)

                For j = 0 To newSeries.Points.Count - 1
                    newSeries.Points(j).Color = Color.Black
                    newSeries.Points(j).BorderDashStyle = p.y1Axis_plot_ref_equations(i).linestyle
                    newSeries.Points(j).BorderWidth = 1
                Next
                ch.Series.Add(newSeries)
            Next
        End If



        'Setup the chart area
        ch.ChartAreas.Add("y1")
        With ch.ChartAreas("y1")
            .Position = New ElementPosition(1, 10, 99, 90)

            'Axis titles
            Dim y1_1_t As New Title
            Dim y1_2_t As New Title
            Dim y2_1_t As New Title
            y1_1_t.Text = p.y1Axis_label_1
            y1_1_t.ForeColor = p.y1Axis_symbolcolour_1
            y1_1_t.Alignment = ContentAlignment.MiddleLeft
            y1_1_t.Position.Auto = False
            y1_1_t.Position = New ElementPosition(1, 0, 50, 10)
            y1_1_t.Font = fnt1
            ch.Titles.Add(y1_1_t)

            If p.y1Axis_label_2 <> "" Then
                y1_2_t.Text = p.y1Axis_label_2
                y1_2_t.ForeColor = p.y1Axis_symbolcolour_2
                y1_2_t.Alignment = ContentAlignment.MiddleLeft
                y1_2_t.Position.Auto = False
                y1_2_t.Position = New ElementPosition(15, 0, 50, 10)
                y1_2_t.Font = fnt1
                ch.Titles.Add(y1_2_t)
            End If

            y2_1_t.Text = p.y2Axis_label
            y2_1_t.ForeColor = p.y2Axis_symbolcolour
            y2_1_t.Alignment = ContentAlignment.MiddleRight
            y2_1_t.Position.Auto = False
            y2_1_t.Position = New ElementPosition(0, 0, 100, 10)
            y2_1_t.Font = fnt1
            ch.Titles.Add(y2_1_t)


            'Setup x axis
            .AxisX.Title = p.xAxis_label
            .AxisX.MajorGrid.Enabled = False
            .AxisX.TitleFont = fnt1
            .AxisX.LabelStyle.Font = fnt1
            .AxisX.IsLabelAutoFit = False
            If p.xAxis_scale_auto Then
                .AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount
                .AxisX.Minimum = Double.NaN
                .AxisX.Maximum = Double.NaN
            Else
                .AxisX.IntervalAutoMode = IntervalAutoMode.FixedCount
                .AxisX.Interval = p.xAxis_tickinterval
                .AxisX.Minimum = p.xAxis_min
                .AxisX.Maximum = p.xAxis_max
            End If
            .AxisY.MajorGrid.Enabled = False
            .AxisY.LabelStyle.Font = fnt1
            .AxisY.IsLabelAutoFit = False
            If p.y1Axis_scale_auto Then
                .AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount
                .AxisY.Minimum = Double.NaN
                .AxisY.Maximum = Double.NaN
            Else
                .AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
                .AxisY.Interval = p.y1Axis_tickinterval
                .AxisY.Minimum = p.y1Axis_min
                .AxisY.Maximum = p.y1Axis_max
            End If

            'Do y2 series________________________________________________   
            If p.y2Axis_enabled Then

                newSeries = New Series
                newSeries.YAxisType = AxisType.Secondary
                newSeries.Name = "series_y2"
                newSeries.ChartType = SeriesChartType.Point
                If Not IsNothing(p.y2Data) Then
                    newSeries.Points.DataBindXY(p.xData, p.y2Data)
                End If
                newSeries.MarkerStyle = p.y2Axis_symbol
                newSeries.MarkerColor = Color.White
                newSeries.MarkerBorderWidth = 1
                newSeries.MarkerBorderColor = p.y2Axis_symbolcolour
                newSeries.MarkerSize = 4
                ch.Series.Add(newSeries)

                'Set ref lines series
                If Not IsNothing(p.y2Axis_plot_ref_lines) Then
                    For i = 0 To UBound(p.y2Axis_plot_ref_lines)
                        newSeries = New Series()
                        newSeries.ChartType = SeriesChartType.Line
                        newSeries.MarkerStyle = MarkerStyle.None
                        newSeries.Points.DataBindXY(p.y2Axis_plot_ref_lines(i).xData, p.y2Axis_plot_ref_lines(i).yData)
                        ch.Series.Add(newSeries)
                    Next
                End If
                'Setup y2 axis
                .AxisY2.MajorGrid.Enabled = False
                .AxisY2.LabelStyle.Font = fnt1
                .AxisY2.LabelStyle.ForeColor = Color.Black
                .AxisY2.IsLabelAutoFit = False

                If p.y2Axis_scale_auto Then
                    .AxisY2.IntervalAutoMode = IntervalAutoMode.VariableCount
                    .AxisY2.Minimum = Double.NaN
                    .AxisY2.Maximum = Double.NaN
                Else
                    .AxisY2.IntervalAutoMode = IntervalAutoMode.FixedCount
                    .AxisY2.Interval = p.y2Axis_tickinterval
                    .AxisY2.Minimum = p.y2Axis_min
                    .AxisY2.Maximum = p.y2Axis_max
                End If
            End If

        End With

        Return ch

    End Function

End Class
