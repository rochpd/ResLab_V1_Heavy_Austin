Imports System.Windows.Forms.DataVisualization.Charting

Public Class class_walktest_plot

    Public Structure walk_protocoldata
        Public protocolID As Long
        Public name As String
        Public description As String
        Public pMenuLabel As String
        Public timepoint_interval_min As Single
        Public nTimePoints_rest As Integer
        Public nTimePoints_walk As Integer
        Public nTimePoints_post As Integer
        Public walkmode As String
        Public var_time As String
        Public var_speed As String
        Public var_grade As String
        Public var_fio2 As String
        Public var_spo2 As String
        Public var_hr As String
        Public var_dyspnoea As String
        Public enabled As String
    End Structure

    Public Structure walk_trialdata
        Public trialID As Integer
        Public walkID As Long
        Public trial_number As String
        Public trial_label As String
        Public trial_distance As String
        Public timepoints() As walk_timepointdata
        Public trial_timeofday As String
    End Structure

    Public Structure walk_timepointdata
        Public timepointID As Integer
        Public trialID As Long
        Public time_label As String
        Public time_minute As String
        Public time_speed_kph As String
        Public time_gradient As String
        Public time_spo2 As String
        Public time_hr As String
        Public time_o2flow As String
        Public time_dyspnoea As String
    End Structure

    Public Structure walk_testdata
        Public walkID As Integer
        Public PatientID As String
        Public sessionID As String
        Public TestDate As String
        Public TestTime As String
        Public TestType As String
        Public Lab As String
        Public Scientist As String
        Public AdmissionStatus As String
        Public BDStatus As String
        Public Height As String
        Public Weight As String
        Public Smoke_hx As String
        Public Smoke_yearssmoked As String
        Public Smoke_cigsperday As String
        Public Smoke_packyears As String
        Public Req_name As String
        Public Req_address As String
        Public Req_fax As String
        Public Req_email As String
        Public Report_copyto As String
        Public Req_clinicalnotes As String
        Public TechnicalNote As String
        Public Report_text As String
        Public Report_status As String
        Public Report_reportedby As String
        Public Report_reporteddate As String
        Public Report_verifiedby As String
        Public Report_verifieddate As String
        Public Pred_SourceIDs As String
        Public LastUpdate As String
        Public ProtocolID As Integer
        Public trials() As walk_trialdata
    End Structure

    Public Structure walk_SummaryResults
        Dim trialnumber As Integer
        Dim FiO2_rest As String
        Dim FiO2_exercise As String
        Dim FiO2_recovery As String
        Dim SpO2_rest As String
        Dim SpO2_exercise As String
        Dim SpO2_recovery As String
        Dim HR_rest As String
        Dim HR_exercise As String
        Dim HR_recovery As String
        Dim Speed_rest As String
        Dim Speed_exercise As String
        Dim Speed_recovery As String
        Dim Gradient_rest As String
        Dim Gradient_exercise As String
        Dim Gradient_recovery As String
        Dim Dyspnoea_rest As String
        Dim Dyspnoea_exercise As String
        Dim Dyspnoea_recovery As String
        Dim Distance_exercise As String
    End Structure


    Public Structure plot_referenceline
        Public xData As List(Of Single)
        Public yData() As List(Of Single)
        Public linestyle As Drawing2D.DashStyle
    End Structure

    Public Structure walk_plot_properties
        Public plot_width As Integer
        Public plot_height As Integer

        Public plot_yAxesCount As Integer
        Public plot_font As Font

        Public xData() As List(Of Single)
        Public yData() As List(Of Single)

        Public xAxis_label As String
        Public xAxis_scale_auto As Boolean
        Public xAxis_min As String
        Public xAxis_max As String
        Public xAxis_tickinterval As String

        Public yAxis_label As String
        Public yAxis_min As String
        Public yAxis_max As String
        Public yAxis_tickinterval As String
        Public yAxis_series_1_symbol As MarkerStyle
        Public yAxis_series_1_symbolcolour As Color
        Public yAxis_series_2_symbol As MarkerStyle
        Public yAxis_series_2_symbolcolour As Color

    End Structure

    Public Function Convert_XYdatatoArray(w() As walk_trialdata) As List(Of Single)()
        'Return array contains 4xarrays
        '(0)=trial_1 X data
        '(0)=trial_1 Y data
        '(0)=trial_2 X data
        '(0)=trial_2 Y data

        'Get values in Y column - skip non-numeric Y values and corresponding X values
        'Assume complete set of X vaues present

        Dim a(3) As List(Of Single)
        a(0) = New List(Of Single)
        a(1) = New List(Of Single)
        a(2) = New List(Of Single)
        a(3) = New List(Of Single)

        For i = 0 To 1
            For j = 0 To UBound(w(i).timepoints)
                If IsNumeric(w(i).timepoints(j).time_spo2) Then
                    a(i * 2).Add(CSng(w(i).timepoints(j).time_minute))
                    a(i * 2 + 1).Add(CSng(w(i).timepoints(j).time_spo2))
                End If
            Next
        Next

        Return a

    End Function

    Public Function Get_plotproperties_walk() As walk_plot_properties

        Dim p As New walk_plot_properties

        p.xAxis_label = "Time (min)"
        p.xAxis_min = 0
        p.xAxis_max = 9
        p.xAxis_tickinterval = 1
        p.xAxis_scale_auto = False

        p.yAxis_label = "SpO2 (%)"
        p.yAxis_min = 75
        p.yAxis_max = 100
        p.yAxis_tickinterval = 5
        p.yAxis_series_1_symbol = DataVisualization.Charting.MarkerStyle.Circle
        p.yAxis_series_1_symbolcolour = Color.Blue
        p.yAxis_series_2_symbol = DataVisualization.Charting.MarkerStyle.Star5
        p.yAxis_series_2_symbolcolour = Color.Red

        Return p

    End Function

    Public Function Create_plot(p As walk_plot_properties) As System.Windows.Forms.DataVisualization.Charting.Chart

        Dim fnt1 As New Font("Arial", 9, FontStyle.Regular)

        Dim ch = New System.Windows.Forms.DataVisualization.Charting.Chart
        ch.Name = "walkGraph"
        ch.ChartAreas.Clear()
        ch.Series.Clear()
        ch.Legends.Clear()
        ch.Titles.Clear()
        ch.Size = New Size(450, 270)

        'Setup the chart area
        Dim leg As New Legend
        leg.Name = "myLeg"
        leg.Position.Auto = True
        leg.Font = fnt1
        ch.Legends.Add(leg)

        Dim e1 As New ElementPosition, e2 As New ElementPosition
        ch.Titles.Add("Recovery")
        ch.Titles(0).Text = "Recovery"
        e1.X = 65 : e1.Y = 3
        ch.Titles(0).Position = e1

        ch.Titles.Add("Walk")
        ch.Titles(1).Text = "Walk"
        e2.X = 35 : e2.Y = 3
        ch.Titles(1).Position = e2

        Dim ChartArea1 = New ChartArea
        ChartArea1.Name = "ChartArea1"
        ch.ChartAreas.Add(ChartArea1)
        With ch.ChartAreas("ChartArea1")
            'Setup x axis
            .AxisX.Title = p.xAxis_label
            .AxisX.MajorGrid.Enabled = True
            .AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash
            .AxisX.TitleFont = fnt1
            .AxisX.LabelStyle.Font = fnt1
            .AxisX.IsLabelAutoFit = False
            .AxisX.IntervalAutoMode = IntervalAutoMode.FixedCount
            .AxisX.Interval = p.xAxis_tickinterval
            .AxisX.Minimum = p.xAxis_min
            .AxisX.Maximum = p.xAxis_max
            'Setup y axis
            .AxisY.Title = p.yAxis_label
            .AxisY.MajorGrid.Enabled = True
            .AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash
            .AxisY.TitleFont = fnt1
            .AxisY.LabelStyle.Font = fnt1
            .AxisY.IsLabelAutoFit = False
            .AxisY.IntervalAutoMode = IntervalAutoMode.FixedCount
            .AxisY.Interval = p.yAxis_tickinterval
            .AxisY.Minimum = p.yAxis_min
            .AxisY.Maximum = p.yAxis_max
        End With


        'Do trial 1 series________________________________________________
        Dim s As Series
        s = New Series
        s.YAxisType = AxisType.Primary
        s.Name = "trial_1"
        s.ChartType = SeriesChartType.Line
        s.ChartArea = "ChartArea1"
        s.Legend = "myLeg"
        s.IsVisibleInLegend = True
        s.LegendText = "Trial 1"
        s.Points.DataBindXY(p.xData(0), p.yData(0))
        s.MarkerStyle = p.yAxis_series_1_symbol
        s.MarkerSize = 5
        s.MarkerColor = p.yAxis_series_1_symbolcolour
        ch.Series.Add(s)
        ch.Series("trial_1").Color = Color.Blue
        ch.Series("trial_1").BorderWidth = 1
        ch.Series("trial_1").BorderDashStyle = ChartDashStyle.DashDot

        'Do trial 2 series________________________________________________
        s = New Series
        s.YAxisType = AxisType.Primary
        s.Name = "trial_2"
        s.ChartType = SeriesChartType.Line
        s.Legend = "myLeg"
        s.ChartArea = "ChartArea1"
        s.IsVisibleInLegend = True
        s.LegendText = "Trial 2"
        s.Points.DataBindXY(p.xData(1), p.yData(1))
        s.MarkerStyle = p.yAxis_series_2_symbol
        s.MarkerSize = 8
        s.MarkerColor = p.yAxis_series_2_symbolcolour
        ch.Series.Add(s)
        ch.Series("trial_2").Color = Color.Red
        ch.Series("trial_2").BorderWidth = 1

        'Do walk-recovery vertical bar as third series________________________________________________
        s = New Series
        s.YAxisType = AxisType.Primary
        s.Name = "bar"
        s.ChartType = SeriesChartType.Line
        s.ChartArea = "ChartArea1"
        s.IsVisibleInLegend = False
        s.Points.DataBindXY(New List(Of Integer) From {6.5, 6.5}, New List(Of Integer) From {p.yAxis_min, p.yAxis_max})
        s.MarkerStyle = MarkerStyle.None
        ch.Series.Add(s)
        ch.Series("bar").Color = Color.Black
        ch.Series("bar").BorderWidth = 2

        Return (ch)

    End Function


End Class
