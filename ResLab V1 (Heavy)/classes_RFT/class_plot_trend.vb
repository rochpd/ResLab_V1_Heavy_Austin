Imports System.Drawing
Imports System.Drawing.Printing
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms.DataVisualization.Charting

Public Class class_plot_trend

    Public Enum eYValueType
        Absolute
        PercentChange
        PercentPredicted
    End Enum
    Public Enum eReferenceType
        EarliestTest
        SelectedTest
    End Enum
    Public Structure plotproperties
        Dim ShowDataLabels As Boolean
        Dim PlotYValueAs As eYValueType
        Dim yRef_Type As eReferenceType
        Dim yRef_Index As Integer
        Dim yRef_Date As String
        Dim Autoscale As Boolean
        Dim Scale_Ymin As Double
        Dim Scale_Ymax As Double
        Dim GridCanMultiselect As Boolean
        Dim Seriesdata As plot_seriesdata
    End Structure
    Public Structure plot_seriesdata
        Dim x() As Nullable(Of Date)
        Dim y(,) As Nullable(Of Double)
        Dim yLabels() As String
    End Structure

    
    Public Function Create_trend_plot(p As plotproperties) As Chart

        If IsNothing(p.Seriesdata.x) Or IsNothing(p.Seriesdata.y) Then
            Return Nothing
        End If

        Dim x() As Nullable(Of Date) = p.Seriesdata.x
        Dim y(,) As Nullable(Of Double) = p.Seriesdata.y

        'Setup new chart
        Dim chartT As New Chart
        chartT.Name = "chartT"
        chartT.Dock = System.Windows.Forms.DockStyle.Fill

        Dim T = chartT.Titles.Add("1")
        chartT.Titles(0).Text = "Lung Function Trend Graph"
        T.Font = New System.Drawing.Font("Arial", 12, FontStyle.Regular)

        Dim leg As New Legend
        leg.Name = "main_legend"
        leg.Font = New Font("Arial", 10)
        chartT.Legends.Add(leg)

        Dim chartArea1 = New ChartArea
        chartArea1.Name = "main_chartarea"
        chartArea1.AxisX.LabelStyle.Font = New Font("Arial", 8, FontStyle.Regular)
        chartArea1.AxisY.LabelStyle.Font = New Font("Arial", 8, FontStyle.Regular)
        chartArea1.AxisX.IsLabelAutoFit = True
        chartArea1.AxisY.TitleFont = T.Font
        Select Case p.PlotYValueAs
            Case eYValueType.Absolute : chartArea1.AxisY.Title = "Absolute units"
            Case eYValueType.PercentChange
                Select Case p.yRef_Type
                    Case eReferenceType.EarliestTest
                        chartArea1.AxisY.Title = "% Change (ref = earliest)"
                    Case eReferenceType.SelectedTest
                        chartArea1.AxisY.Title = "% Change (ref = " & p.yRef_Date & ")"
                End Select
        End Select
        chartT.ChartAreas.Add(chartArea1)

        'Axes scaling
        If p.Autoscale Then
            ' Set auto minimum and maximum values.
            chartT.ChartAreas("main_chartarea").AxisY.Minimum = [Double].NaN
            chartT.ChartAreas("main_chartarea").AxisY.Maximum = [Double].NaN
        Else
            ' Set manual minimum and maximum values.
            chartT.ChartAreas("main_chartarea").AxisY.Minimum = p.Scale_Ymin
            chartT.ChartAreas("main_chartarea").AxisY.Maximum = p.Scale_Ymax
        End If




        'Plot each series
        Dim newSeries As Series
        Dim xx As List(Of Date), yy As List(Of Double)
        Dim y_refvalue As Nullable(Of Double)


        For i = 0 To UBound(p.Seriesdata.y, 1)
            newSeries = New Series()
            newSeries.Name = p.Seriesdata.yLabels(i)
            newSeries.Legend = "main_legend"
            newSeries.LegendText = newSeries.Name
            newSeries.IsVisibleInLegend = True
            newSeries.ChartType = SeriesChartType.Line

            'Make temp x and y data arrays with empty values removed
            xx = New List(Of Date) : yy = New List(Of Double)
            For j = 0 To UBound(p.Seriesdata.y, 2)
                If Not IsNothing(y(i, j)) Then
                    xx.Add(x(j))
                    yy.Add(y(i, j))
                End If
            Next

            'Transform data if necessary
            If p.PlotYValueAs = eYValueType.PercentChange Then
                If IsNumeric(p.Seriesdata.y(i, p.yRef_Index)) Then
                    y_refvalue = p.Seriesdata.y(i, p.yRef_Index)
                    yy = Me.transform_Yvalues(yy, y_refvalue, p)
                Else
                    MsgBox("No valid reference value for " & p.Seriesdata.yLabels(i) & vbCrLf & "Can't be plotted as % change")
                    yy.Clear()
                End If
            End If

            'If the series has data then do it
            If yy.Count > 0 And xx.Count > 0 Then
                newSeries.Points.DataBindXY(xx, yy)
                newSeries.MarkerStyle = MarkerStyle.Circle
                newSeries.Label = "#VALY"

                'Set datapoint properties
                For Each pt In newSeries.Points
                    If Not p.ShowDataLabels Then pt.Label = String.Empty
                Next

                chartT.Series.Add(newSeries)
            End If


        Next
        Return chartT



    End Function

    Private Function transform_Yvalues(y As List(Of Double), ref_value As Nullable(Of Double), p As plotproperties) As List(Of Double)

        Dim yT As New List(Of Double)

        If IsNothing(ref_value) Then
            Return Nothing
        End If

        Select Case p.PlotYValueAs
            Case eYValueType.Absolute 'return array unchanged
                yT = y
            Case eYValueType.PercentChange
                For Each v In y
                    yT.Add(Math.Round(100 * v / CDec(ref_value), 0))
                Next
            Case eYValueType.PercentPredicted

        End Select

        Return yT

    End Function

    Public Function Create_trend_table(patientID As Long, CanMultiSelect As Boolean) As DataGridView

        Dim grd As New DataGridView
        Dim TestDate As String = "", TestTime As String = ""
        Dim d As Dictionary(Of String, String) = Nothing
        Dim d1 As Dictionary(Of String, String) = Nothing
        Dim rf As New class_Rft_RoutineAndSessionFields
        Dim rd As New class_DemographicFields
        Dim rowheaders() As String = {"Date", "Time", "Tests", "Weight (kg)", "BMI (kg/m2)", "Baseline FEV1 (L)", "Baseline FVC (L)", "Baseline VC (L)", "Baseline FER (%)", "Post BD FEV1 (L)", _
                               "Post BD FVC (L)", "Post BD VC (L)", "Post BD FER (%)", "VA (L)", "TLCO (ml/min/mmHg)", "KCO", "TLCO Hb (ml/min/mmHg)", "KCO Hb", "TLC (L)", "FRC (L)", "RV (L)", "RV/TLC (%)", "MIP (cmH2O)", "MEP (cmH2O)", "FeNO (ppb)", "Technical note", "Report"}
        Dim j As Integer = 0
        Dim fnt As New Font("Segoe UI", 8, FontStyle.Regular)

        'Set up the grid
        grd.RowHeadersDefaultCellStyle.Font = fnt
        grd.DefaultCellStyle.Font = fnt

        grd.Dock = DockStyle.Fill
        grd.MultiSelect = CanMultiSelect
        grd.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        grd.AllowUserToAddRows = False
        grd.AllowUserToDeleteRows = False
        grd.AllowUserToResizeRows = False
        grd.AllowUserToResizeColumns = True
        grd.ColumnHeadersVisible = False
        grd.ColumnCount = 1
        grd.RowHeadersVisible = True
        grd.RowCount = UBound(rowheaders) + 1
        grd.Rows(0).Frozen = True
        grd.Rows(1).Frozen = True
        grd.Rows(2).Frozen = True
        grd.Rows(0).DefaultCellStyle.BackColor = SystemColors.ButtonFace
        grd.Rows(1).DefaultCellStyle.BackColor = SystemColors.ButtonFace
        grd.Rows(2).DefaultCellStyle.BackColor = SystemColors.ButtonFace
        grd.RowHeadersWidth = 160
        grd.RowHeadersDefaultCellStyle.BackColor = SystemColors.ButtonFace


        For i As Integer = 0 To UBound(rowheaders)
            grd.Rows(i).HeaderCell.Value = rowheaders(i)
        Next

        'Get selected demo's
        d1 = cPt.Get_Demographics(patientID)

        'Get the RFTs and load the grid
        Dim dicR() As Dictionary(Of String, String) = cRfts.Get_rfts_ForTrend(patientID, eSortMode.Descending)
        If Not (dicR Is Nothing) Then
            For i As Integer = 0 To UBound(dicR)
                Select Case dicR(i)("tbl")
                    Case "rft_routine"
                        If i > 0 Then grd.ColumnCount = grd.ColumnCount + 1
                        grd.Columns(grd.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
                        grd.Columns(grd.ColumnCount - 1).Width = 70
                        d = cRfts.Get_rft_byRftID(dicR(i)("ID"))
                        j = 0

                        If IsDate(d(rf.TestDate)) Then TestDate = Format(CDate(d(rf.TestDate)), "dd/MM/yy") Else TestDate = ""
                        If IsDate(d(rf.TestTime)) Then TestTime = Format(CDate(d(rf.TestTime)), "HH:mm") Else TestTime = ""

                        grd.Item(grd.ColumnCount - 1, j).Value = TestDate
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = TestTime
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.TestType)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.Weight)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = cMyRoutines.calc_BMI(d(rf.Height), d(rf.Weight))
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_bl_Fev1)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_bl_Fvc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_bl_Vc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = cMyRoutines.Calc_Fer(d(rf.R_bl_Fev1), d(rf.R_bl_Fvc), d(rf.R_bl_Vc))
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_post_Fev1)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_post_Fvc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_post_Vc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = cMyRoutines.Calc_Fer(d(rf.R_post_Fev1), d(rf.R_post_Fvc), d(rf.R_post_Vc))
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Va)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Tlco)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Kco)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = Me.Format_TlcoHbCorr(d(rf.TestDate), d(rf.R_Bl_Tlco), d(rf.R_Bl_Hb), d1(rd.DOB), d1(rd.Gender))
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = Me.Format_TlcoHbCorr(d(rf.TestDate), d(rf.R_Bl_Kco), d(rf.R_Bl_Hb), d1(rd.DOB), d1(rd.Gender))
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Tlc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Frc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Rv)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_RvTlc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Mip)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_Mep)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_Bl_FeNO)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.TechnicalNotes)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.Report_text)
                    Case "prov_test"
                        If i > 0 Then grd.ColumnCount = grd.ColumnCount + 1
                        grd.Columns(grd.ColumnCount - 1).SortMode = DataGridViewColumnSortMode.NotSortable
                        grd.Columns(grd.ColumnCount - 1).Width = 70
                        d = cRfts.Get_prov_test_session(dicR(i)("ID"))
                        j = 0
                        If IsDate(d(rf.TestDate)) Then TestDate = Format(CDate(d(rf.TestDate)), "dd/MM/yy") Else TestDate = ""
                        If IsDate(d(rf.TestTime)) Then TestTime = Format(CDate(d(rf.TestTime)), "HH:mm") Else TestTime = ""

                        grd.Item(grd.ColumnCount - 1, j).Value = TestDate.ToString
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = TestTime
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.TestType)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.Weight)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = cMyRoutines.calc_BMI(d(rf.Height), d(rf.Weight))
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_bl_Fev1)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_bl_Fvc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.R_bl_Vc)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = cMyRoutines.Calc_Fer(d(rf.R_bl_Fev1), d(rf.R_bl_Fvc), d(rf.R_bl_Vc))
                        j = 23
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.TechnicalNotes)
                        j = j + 1 : grd.Item(grd.ColumnCount - 1, j).Value = d(rf.Report_text)
                    Case Else
                        'Don't add any other test types to trend
                End Select
            Next
        End If

        'Alternate colour of tests
        For i = 3 To 26
            Select Case i
                Case 3 To 4 : grd.Rows(i).DefaultCellStyle.BackColor = Color.LightYellow
                Case 5 To 12 : grd.Rows(i).DefaultCellStyle.BackColor = Color.White
                Case 13 To 17 : grd.Rows(i).DefaultCellStyle.BackColor = Color.LightYellow
                Case 18 To 21 : grd.Rows(i).DefaultCellStyle.BackColor = Color.White
                Case 22 To 23 : grd.Rows(i).DefaultCellStyle.BackColor = Color.LightYellow
                Case 24 : grd.Rows(i).DefaultCellStyle.BackColor = Color.White
                Case 25 To 26 : grd.Rows(i).DefaultCellStyle.BackColor = Color.LightYellow
            End Select
        Next

        grd.ClearSelection()

        Return grd

    End Function

    Private Function Format_TlcoHbCorr(ByVal TestDate, ByVal Result, ByVal Hb, ByVal DOB, ByVal Gender) As String

        If IsDate(TestDate) And IsDate(DOB) And Val(Result) > 0 And Gender <> "" Then
            Dim Fac As Single = cMyRoutines.calc_HbFac(Hb, DOB, Gender, TestDate)
            If Fac > 0 Then
                Dim TLHb As Single = Result * Fac
                Return Format(TLHb, "#0.0")
            Else
                Return ""
            End If
        Else
            Return ""
        End If

    End Function

   
End Class
