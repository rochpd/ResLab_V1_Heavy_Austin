Imports System.Text
Imports System.Drawing
Imports System.Drawing.Printing
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms.DataVisualization.Charting
Imports ResLab_V1_Heavy.class_walktest_plot

Public Class class_walktest

#Region "Variable defs"

    Private _testdata As class_walktest_plot.walk_testdata
    Private _patientID As Long
    Private _testID As Long
    Private _sessionID As Long

    Public Structure ProtocolMenuData
        Public Label As String
        Public protocolID As Integer
    End Structure

#End Region


#Region "Enum defs"
    Public Enum eWalkModes
        Treadmill
        Floor
    End Enum

    Public Enum eWalkProtocols
        NRG
        Standard6mwd
    End Enum

#End Region


#Region "Methods"

    Public Function Get_Protocols(ByVal RestrictToEnabled As Boolean) As ProtocolMenuData()
        'Returns data to populate main menu list of available protocols

        Dim sql As String = ""
        Dim p() As ProtocolMenuData
        Dim i As Integer = 0

        Select Case RestrictToEnabled
            Case True : sql = "SELECT protocolID, pmenulabel FROM r_walktests_protocols WHERE enabled=1"
            Case False : sql = "SELECT protocolID, pmenulabel FROM r_walktests_protocols"
        End Select
        sql = sql & " ORDER BY pmenulabel"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            ReDim p(Ds.Tables(0).Rows.Count - 1)
            i = 0
            For Each r As DataRow In Ds.Tables(0).Rows
                p(i).protocolID = r("protocolID")
                p(i).Label = r("pmenulabel")
                i = i + 1
            Next
            Return p
        Else
            Return Nothing
        End If

    End Function

    Public Function get_ProtocolData(protocolID As Long) As class_walktest_plot.walk_protocoldata

        Dim sql As String = "SELECT * FROM r_walktests_protocols WHERE protocolID=" & protocolID
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        Dim p As New walk_protocoldata

        If Ds Is Nothing Then
            Return Nothing
        Else
            With Ds.Tables(0)
                p.protocolID = .Rows(0)("protocolID")
                p.name = .Rows(0)("name")
                p.description = .Rows(0)("description")
                p.pMenuLabel = .Rows(0)("pMenuLabel")
                p.timepoint_interval_min = .Rows(0)("timepoint_interval_min")
                p.nTimePoints_rest = .Rows(0)("nTimePoints_rest")
                p.nTimePoints_walk = .Rows(0)("nTimePoints_walk")
                p.nTimePoints_post = .Rows(0)("nTimePoints_post")
                p.walkmode = .Rows(0)("walkmode")
                p.var_time = .Rows(0)("var_time")
                p.var_speed = .Rows(0)("var_speed")
                p.var_grade = .Rows(0)("var_grade")
                p.var_fio2 = .Rows(0)("var_fio2")
                p.var_spo2 = .Rows(0)("var_spo2")
                p.var_hr = .Rows(0)("var_hr")
                p.var_dyspnoea = .Rows(0)("var_dyspnoea")
                p.enabled = .Rows(0)("enabled")
            End With
        End If

        Return p

    End Function

    Public Function Get_ProtocolList(ByVal RestrictToEnabled As Boolean) As ProtocolMenuData()
        'Returns data to populate main menu list of available protocols

        Dim sql As String = ""
        Dim p() As ProtocolMenuData
        Dim i As Integer = 0

        Select Case RestrictToEnabled
            Case True : sql = "SELECT protocolID, pMenulabel FROM r_walktests_protocols WHERE enabled=1"
            Case False : sql = "SELECT protocolID, pMenulabel FROM r_walktests_protocols"
        End Select
        sql = sql & " ORDER BY pmenulabel"

        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Ds.Tables(0).Rows.Count > 0 Then
            ReDim p(Ds.Tables(0).Rows.Count - 1)
            i = 0
            For Each r As DataRow In Ds.Tables(0).Rows
                p(i).protocolID = r("protocolID")
                p(i).Label = r("pmenulabel")
                i = i + 1
            Next
            Return p
        Else
            Return Nothing
        End If

    End Function

    Public Function create_walkgrid(protocolinfo As walk_protocoldata, IsNewTrial As Boolean) As DataGridView

        Dim i As Integer = 0
        Dim g As New DataGridView
        Dim CellStyle_header As New DataGridViewCellStyle
        Dim CellStyle_body As New DataGridViewCellStyle

        CellStyle_header.Alignment = DataGridViewContentAlignment.MiddleLeft
        CellStyle_header.BackColor = SystemColors.Control
        CellStyle_header.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))
        CellStyle_header.ForeColor = Drawing.Color.Black

        CellStyle_body.Alignment = DataGridViewContentAlignment.MiddleLeft
        CellStyle_body.BackColor = Drawing.Color.White
        CellStyle_body.Font = New Font("Microsoft Sans Serif", 8, FontStyle.Regular, GraphicsUnit.Point, CType(0, Byte))
        CellStyle_body.ForeColor = Drawing.Color.Black


        g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        g.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        g.AllowUserToAddRows = False
        g.AllowUserToDeleteRows = False
        g.AllowUserToResizeRows = False
        g.BorderStyle = BorderStyle.None
        g.DefaultCellStyle = CellStyle_body
        g.Dock = DockStyle.Fill
        g.MultiSelect = False
        g.ReadOnly = False
        g.RowHeadersVisible = False
        g.ScrollBars = ScrollBars.Both
        g.SelectionMode = DataGridViewSelectionMode.CellSelect
        g.ColumnHeadersDefaultCellStyle = CellStyle_header

        'Create columns
        g.Columns.Clear()
        Dim col_names = New List(Of String) From {"levelid", "time", "stage", "speed", "grade", "fio2", "spo2", "hr", "dyspnoea"}
        Dim col_headertext = New List(Of String) From {"levelID", "Time(min)", "Stage", "Speed(kph)", "Grade(%)", "FiO2(L/min)", "SpO2(%)", "HR(bpm)", "Dysp(Borg)"}
        Dim col_visible = New List(Of Boolean) From {False, True, True, CBool(protocolinfo.var_speed), CBool(protocolinfo.var_grade), CBool(protocolinfo.var_fio2), CBool(protocolinfo.var_spo2), CBool(protocolinfo.var_hr), CBool(protocolinfo.var_dyspnoea)}
        Dim col_width = New List(Of Integer) From {80, 80, 80, 90, 80, 90, 80, 80, 90}
        Dim c As DataGridViewTextBoxColumn
        For i = 0 To col_names.Count - 1
            c = New DataGridViewTextBoxColumn
            c.Name = col_names(i)
            c.HeaderText = col_headertext(i)
            c.Visible = col_visible(i)
            c.Width = col_width(i)
            c.SortMode = DataGridViewColumnSortMode.NotSortable
            g.Columns.Add(c)
        Next
        g.Columns(1).DefaultCellStyle.BackColor = SystemColors.ButtonFace
        g.Columns(2).DefaultCellStyle.BackColor = SystemColors.ButtonFace

        'Create rows
        g.Rows.Clear()
        Dim r As Integer = 0
        Dim timevalue As Single = protocolinfo.timepoint_interval_min
        '  Rests
        For i = 1 To protocolinfo.nTimePoints_rest
            r = r + 1
            g.RowCount = r
            If IsNewTrial Then g.Item(0, g.RowCount - 1).Value = 0
            g.Item(1, g.RowCount - 1).Value = Format(Math.Round((r - 1) * protocolinfo.timepoint_interval_min, 1), "#0.0")
            g.Item(2, g.RowCount - 1).Value = "Rest"
        Next i
        '  Walk
        For i = 1 To protocolinfo.nTimePoints_walk
            r = r + 1
            g.RowCount = r
            If IsNewTrial Then g.Item(0, g.RowCount - 1).Value = 0
            g.Item(1, g.RowCount - 1).Value = Format(Math.Round((r - 1) * protocolinfo.timepoint_interval_min, 1), "#0.0")
            g.Item(2, g.RowCount - 1).Value = "Walk " & g.Item(1, g.RowCount - 1).Value
        Next i
        '  Post
        For i = 1 To protocolinfo.nTimePoints_post
            r = r + 1
            g.RowCount = r
            If IsNewTrial Then g.Item(0, g.RowCount - 1).Value = 0
            g.Item(1, g.RowCount - 1).Value = Format(Math.Round((r - 1) * protocolinfo.timepoint_interval_min, 1), "#0.0")
            g.Item(2, g.RowCount - 1).Value = "Post " & g.Item(1, g.RowCount - 1).Value
        Next i


        Return g

    End Function

    Public Function Calc_distance(g As DataGridView, protocolID As Long) As String
        'Based on time (col 1) and speed (col 3) columns
        'Assume a speed in at least the first row

        Dim protocoldata As class_walktest_plot.walk_protocoldata = Me.get_ProtocolData(protocolID)
        Dim startrow As Integer = protocoldata.nTimePoints_rest
        Dim endrow As Integer = protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
        Dim cT As Integer = 1
        Dim cS As Integer = 3
        Dim speed As Single = 0
        Dim dist As Single = 0

        If Not IsNumeric(g(cT, startrow).Value) Then
            MsgBox("First walk row must have a valid spedd value", vbOKOnly, "ResLab")
            Return ""
        End If

        For i = startrow To endrow
            'Update speed if new value in grid
            If IsNumeric(g(cS, i).Value) Then speed = Val(g(cS, i).Value)
            'Do distance calc
            dist = dist + speed * protocoldata.timepoint_interval_min * 1000 / 60
        Next

        Return Format(dist, "###")

    End Function

    Public Function array_getminvalue(a() As Single) As String

        If IsNothing(a) Or UBound(a) < 0 Then
            Return ""
        Else
            Array.Sort(a)
            Return a(LBound(a)).ToString
        End If

    End Function

    Public Function array_getmaxvalue(a() As Single) As String

        If IsNothing(a) Or UBound(a) < 0 Then
            Return ""
        Else
            Array.Sort(a)
            Return a(UBound(a)).ToString
        End If

    End Function

    Public Function Calculate_SummaryResults(w As walk_trialdata(), protocolID As Long) As walk_SummaryResults()

        Dim s(1) As walk_SummaryResults
        Dim td As walk_trialdata
        Dim i As Integer
        Dim a() As Single
        Dim trial As Integer
        Dim min As String = "", max As String = ""
        Dim protocoldata As class_walktest_plot.walk_protocoldata = Me.get_ProtocolData(protocolID)

        For Each td In w

            trial = td.trial_number - 1
            s(trial).trialnumber = td.trial_number

            'SPO2
            'rest spo2 - use last rest value
            s(trial).SpO2_rest = td.timepoints(protocoldata.nTimePoints_rest - 1).time_spo2
            'walk spo2 - nadir
            ReDim a(-1)
            For i = protocoldata.nTimePoints_rest To protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
                If Val(td.timepoints(i).time_spo2) > 0 Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_spo2)
                End If
            Next
            s(trial).SpO2_exercise = Me.array_getminvalue(a)
            'recovery spo2 - peak
            ReDim a(-1)
            For i = protocoldata.nTimePoints_walk To protocoldata.nTimePoints_walk + protocoldata.nTimePoints_post - 1
                If Val(td.timepoints(i).time_spo2) > 0 Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_spo2)
                End If
            Next
            s(trial).SpO2_recovery = Me.array_getminvalue(a)

            'HR
            'rest hr - use last rest value
            s(trial).HR_rest = td.timepoints(protocoldata.nTimePoints_rest - 1).time_hr
            'walk hr - peak
            ReDim a(-1)
            For i = protocoldata.nTimePoints_rest To protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
                If Val(td.timepoints(i).time_hr) > 0 Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_hr)
                End If
            Next
            s(trial).HR_exercise = Me.array_getmaxvalue(a)
            'recovery hr - nadir
            ReDim a(-1)
            For i = protocoldata.nTimePoints_walk To protocoldata.nTimePoints_walk + protocoldata.nTimePoints_post - 1
                If Val(td.timepoints(i).time_hr) > 0 Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_hr)
                End If
            Next
            s(trial).HR_recovery = Me.array_getmaxvalue(a)

            'DYSPNOEA
            'rest dyspnoea - use last rest value
            s(trial).Dyspnoea_rest = td.timepoints(protocoldata.nTimePoints_rest - 1).time_dyspnoea
            'walk dyspnoea - peak
            ReDim a(-1)
            For i = protocoldata.nTimePoints_rest To protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
                If Val(td.timepoints(i).time_dyspnoea) > 0 Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_dyspnoea)
                End If
            Next
            s(trial).Dyspnoea_exercise = Me.array_getmaxvalue(a)
            'recovery dyspnoea - nadir
            ReDim a(-1)
            For i = protocoldata.nTimePoints_walk To protocoldata.nTimePoints_walk + protocoldata.nTimePoints_post - 1
                If Val(td.timepoints(i).time_dyspnoea) > 0 Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_dyspnoea)
                End If
            Next
            s(trial).Dyspnoea_recovery = Me.array_getmaxvalue(a)


            'FiO2
            'rest FiO2 - use last rest value
            If td.timepoints(protocoldata.nTimePoints_rest - 1).time_o2flow = "0" Then s(trial).FiO2_rest = "Air" Else s(trial).FiO2_rest = td.timepoints(protocoldata.nTimePoints_rest - 1).time_o2flow
            'walk FiO2 
            ReDim a(-1)
            For i = protocoldata.nTimePoints_rest To protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
                If td.timepoints(i).time_o2flow <> "" Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_o2flow)
                End If
            Next
            min = Me.array_getminvalue(a)
            max = Me.array_getmaxvalue(a)
            If min = max Then
                If min = "0" Then s(trial).FiO2_exercise = "Air" Else s(trial).FiO2_exercise = min
            Else
                s(trial).FiO2_exercise = min & "-" & max
            End If
            'recovery FiO2 
            ReDim a(-1)
            For i = protocoldata.nTimePoints_walk To protocoldata.nTimePoints_walk + protocoldata.nTimePoints_post - 1
                If td.timepoints(i).time_o2flow <> "" Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_o2flow)
                End If
            Next
            min = Me.array_getminvalue(a)
            max = Me.array_getmaxvalue(a)
            If min = max Then
                If min = "0" Then s(trial).FiO2_recovery = "Air" Else s(trial).FiO2_recovery = min
            Else
                s(trial).FiO2_recovery = min & "-" & max
            End If

            'Speed
            ReDim a(-1)
            For i = protocoldata.nTimePoints_rest To protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
                If td.timepoints(i).time_speed_kph <> "" Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_speed_kph)
                End If
            Next
            min = Me.array_getminvalue(a)
            max = Me.array_getmaxvalue(a)
            If min = max Then
                s(trial).Speed_exercise = min
            Else
                s(trial).Speed_exercise = min & "-" & max
            End If

            'Gradient
            ReDim a(-1)
            For i = protocoldata.nTimePoints_rest To protocoldata.nTimePoints_rest + protocoldata.nTimePoints_walk - 1
                If td.timepoints(i).time_gradient <> "" Then
                    If UBound(a) = -1 Then ReDim a(0) Else ReDim Preserve a(UBound(a) + 1)
                    a(UBound(a)) = Val(td.timepoints(i).time_gradient)
                End If
            Next
            min = Me.array_getminvalue(a)
            max = Me.array_getmaxvalue(a)
            If min = max Then
                s(trial).Gradient_exercise = min
            Else
                s(trial).Gradient_exercise = min & "-" & max
            End If

            'Distance
            s(trial).Distance_exercise = td.trial_distance
        Next

        Return s

    End Function

    Public Function Draw_WalkGraph(ByVal w() As walk_trialdata, ByVal fnt As Font) As Bitmap

        Try

            Const Ht As Integer = 500
            Const Width As Integer = 900



            Dim img = New Bitmap(Width, Ht)
            Dim g As Graphics = Graphics.FromImage(img)
            Dim myPen As New System.Drawing.Pen(System.Drawing.Color.Black, 0.1)

            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
            g.InterpolationMode = Drawing2D.InterpolationMode.High

            'Canvas area (ie graph plus clear border for labels
            Dim xmin As Integer = -2
            Dim xmax As Integer = 10
            Dim ymin As Integer = 70
            Dim ymax As Integer = 105

            'Graph axes min max values
            Dim x1 As Integer = 0
            Dim x2 As Integer = 9
            Dim y1 As Integer = 75
            Dim y2 As Integer = 100
            Dim xstep As Integer = 1
            Dim ystep As Integer = 5

            'Numeric scaling - doesn't operate on text
            Dim scaleX As Single = 0
            Dim scaleY As Single = 0

            'g.ScaleTransform(Width / 130, -Ht / 100)
            'g.TranslateTransform(10, -130)

            g.ScaleTransform(Width / 90, -Ht / 50)
            g.TranslateTransform(0, -130)

            g.DrawLine(myPen, 0, 0, 100, 100)

            'Do horizontal lines
            myPen.DashStyle = Drawing2D.DashStyle.Solid
            For i = 0 To 100 Step 20
                g.DrawLine(myPen, 10, i, 100, i)
            Next

            'Draw vertical lines
            'Dim X, Y As Single
            'g.DrawLine(myPen, 0, 50, 0, 120)
            'g.DrawLine(myPen, 110, ymin, 110, ymax)
            'For i = xmin To xmax Step xstep
            '    X = 10 + (i - 1) * xstep
            '    g.DrawLine(myPen, X, ymin, X, ymax)
            'Next


            Return img



            'DRAW GRAPH
            'xmin = 20             'Co-ords (in mm) for graph
            'xmax = xmin + 40
            'ymin = IDP(ystart) + 20
            'ymax = ymin + 40
            'xsize = xmax - xmin
            'ysize = ymax - ymin
            'minSaO2 = 75

            'pdf.SetLineWidth(0.8)
            'pdf_P(pdf, "SaO2", (xmin - 16), (ymin + 10), Fs3, 3)
            'pdf_P(pdf, "(%)", (xmin - 15), (ymin + 13), Fs3, 3)

            'FiO2 = summaryData.FiO2_1_exercise
            'If Val(summaryData.FiO2_1_exercise) > 0 Then FiO2 = FiO2 & " L/min"
            'pdf_P(pdf, "Trial 1   (FiO2 = " & FiO2 & ")", (xmin), (ymin - 13), Fs3, 3)
            'pdf.SetLineDash("[] 0")
            'X = DPI(xmin + 10) : Y = DPI(ymin - 6)
            'pdf.MoveTo(X, Y) : pdf.DrawLineTo(X + 20, Y)
            'pdf.Stroke()

            'FiO2 = summaryData.FiO2_2_exercise
            'If Val(summaryData.FiO2_2_exercise) > 0 Then FiO2 = FiO2 & " L/min"
            'pdf_P(pdf, "Trial 2   (FiO2 = " & FiO2 & ")", (xmin), (ymin - 10), Fs3, 3)
            'pdf.SetLineDash("[3] 0")
            'X = DPI(xmin + 10) : Y = DPI(ymin - 3)
            'pdf.MoveTo(X, Y) : pdf.DrawLineTo(X + 20, Y)
            'pdf.Stroke()

            'pdf_P(pdf, "Time (mins)", (xmin + 0.4 * xsize), (ymax + 4), Fs3, 3)
            'pdf_P(pdf, "REST", xmin - 2, ymin - 1.3 * lHt, Fs3, 3)
            'pdf_P(pdf, "EXERCISE", xmin + 13, ymin - 1.3 * lHt, Fs3, 3)
            'pdf_P(pdf, "RECOVERY", xmin + 39, ymin - 1.3 * lHt, Fs3, 3)

            ''Draw horizontal lines and label y-axis
            'x_div_mm = 6
            'x_div_n = 9

            'pdf.SetLineWidth(0.5)
            'pdf.SetLineDash("[] 0")
            'For j = minSaO2 To 100 Step 5
            '    Y = ymax - (j - minSaO2) * ysize / (100 - minSaO2)
            '    pdf_P(pdf, Format(j, "###"), (xmin - 6), (Y - 2), Fs3, 3)

            '    x1_dpi = M.Left + DPI(xmin)
            '    x2_dpi = M.Left + DPI(xmin) - DPI(1)
            '    y1_dpi = M.Top + DPI(Y)
            '    y2_dpi = M.Top + DPI(Y)
            '    pdf.MoveTo(x1_dpi, y1_dpi) : pdf.DrawLineTo(x2_dpi, y2_dpi)
            '    pdf.Stroke()

            '    x1_dpi = M.Left + DPI(xmin) + DPI(x_div_mm / 2)
            '    x2_dpi = M.Left + DPI(xmin) + DPI(x_div_mm * 6)
            '    pdf.MoveTo(x1_dpi, y1_dpi) : pdf.DrawLineTo(x2_dpi, y2_dpi)
            '    pdf.Stroke()

            '    x1_dpi = M.Left + DPI(xmin) + DPI(x_div_mm * 6.5)
            '    x2_dpi = M.Left + DPI(xmin) + DPI(x_div_mm * 9)
            '    pdf.MoveTo(x1_dpi, y1_dpi) : pdf.DrawLineTo(x2_dpi, y2_dpi)
            '    pdf.Stroke()
            'Next j

            ''Draw vertical lines and label x-axis
            'For j = 0 To x_div_n
            '    x1_dpi = M.Left + DPI(xmin + j * x_div_mm)  'major grid lines
            '    x2_dpi = M.Left + DPI(xmin + (j + 0.5) * x_div_mm) 'minor grid lines
            '    y1_dpi = M.Top + DPI(ymin)
            '    y2_dpi = M.Top + DPI(ymax)

            '    pdf.MoveTo(x1_dpi, y1_dpi) : pdf.DrawLineTo(x1_dpi, y2_dpi)
            '    pdf.Stroke()

            '    If j <> x_div_n Then
            '        pdf.SetLineDash("[5] 0")
            '        'If j = 0 Or j = 6 Then pdf.SetLineDash ("[] 0")
            '        pdf.MoveTo(x2_dpi, y1_dpi) : pdf.DrawLineTo(x2_dpi, y2_dpi)
            '        pdf.Stroke()
            '        pdf.SetLineDash("[] 0")
            '    End If

            '    pdf_P(pdf, Format(j, "0"), xmin - 1 + j * x_div_mm, (ymax + 1), Fs3, 3)
            'Next j

            ''Plot trial 1 data
            'pdf.SetLineWidth(0.8)
            'For i = 1 To 13 'first find the first data point >0
            '    If Val(walkData.SpO2_1(i)) > 0 Then
            '        prevX_dpi = M.Left + DPI(xmin + (i - 1) * x_div_mm / 2)
            '        prevY_dpi = M.Top + DPI(ymax - (walkData.SpO2_1(i) - minSaO2) * ysize / (100 - minSaO2))
            '        pdf.DrawCircle(prevX_dpi, prevY_dpi, DPI(0.4))
            '        pdf.Stroke()
            '        Exit For
            '    End If
            'Next i
            'For j = i + 1 To 19
            '    SaO2 = Val(walkData.SpO2_1(j))
            '    If SaO2 <> 0 Then
            '        X = M.Left + DPI(xmin + (j - 1) * x_div_mm / 2)
            '        Y = M.Top + DPI(ymax - (SaO2 - minSaO2) * ysize / (100 - minSaO2))
            '        pdf.DrawCircle(X, Y, DPI(0.4))
            '        pdf.MoveTo(prevX_dpi, prevY_dpi) : pdf.DrawLineTo(X, Y)
            '        pdf.Stroke()
            '        prevX_dpi = X
            '        prevY_dpi = Y
            '    End If
            'Next j

            ''Plot trial 2 data
            'pdf.SetLineDash("[] 0")
            'For i = 1 To 13 'first find the first data point >0
            '    If Val(walkData.SpO2_2(i)) > 0 Then
            '        prevX_dpi = M.Left + DPI(xmin + (i - 1) * x_div_mm / 2)
            '        prevY_dpi = M.Top + DPI(ymax - (walkData.SpO2_2(i) - minSaO2) * ysize / (100 - minSaO2))
            '        pdf.DrawCircle(prevX_dpi, prevY_dpi, DPI(0.4))
            '        pdf.Stroke()
            '        Exit For
            '    End If
            'Next i
            'For j = i + 1 To 19
            '    SaO2 = Val(walkData.SpO2_2(j))
            '    If SaO2 <> 0 Then
            '        X = M.Left + DPI(xmin + (j - 1) * x_div_mm / 2)
            '        Y = M.Top + DPI(ymax - (SaO2 - minSaO2) * ysize / (100 - minSaO2))

            '        pdf.SetLineDash("[3] 0")
            '        pdf.MoveTo(prevX_dpi, prevY_dpi) : pdf.DrawLineTo(X, Y)
            '        pdf.Stroke()

            '        pdf.SetLineDash("[] 0")
            '        pdf.DrawCircle(X, Y, DPI(0.4))
            '        pdf.Stroke()

            '        prevX_dpi = X
            '        prevY_dpi = Y
            '    End If
            'Next j



        Catch
            MsgBox("Error in Draw_provocationGraph" & vbNewLine & Err.Description)
            Return Nothing
        End Try

    End Function

#End Region

End Class









