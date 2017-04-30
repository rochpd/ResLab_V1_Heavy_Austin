<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_MainNew
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_MainNew))
        Dim ViewerPreferences1 As Gnostice.PDFOne.Windows.PDFViewer.ViewerPreferences = New Gnostice.PDFOne.Windows.PDFViewer.ViewerPreferences()
        Dim RenderingSettings2 As Gnostice.PDFOne.RenderingSettings = New Gnostice.PDFOne.RenderingSettings()
        Dim ImageSettings2 As Gnostice.PDFOne.ImageSettings = New Gnostice.PDFOne.ImageSettings()
        Dim PdfRenderOptions2 As Gnostice.PDFOne.PDFRenderOptions = New Gnostice.PDFOne.PDFRenderOptions()
        Dim PrinterPreferences1 As Gnostice.PDFOne.PDFPrinter.PrinterPreferences = New Gnostice.PDFOne.PDFPrinter.PrinterPreferences()
        Dim RenderingSettings3 As Gnostice.PDFOne.RenderingSettings = New Gnostice.PDFOne.RenderingSettings()
        Dim ImageSettings3 As Gnostice.PDFOne.ImageSettings = New Gnostice.PDFOne.ImageSettings()
        Dim PdfRenderOptions3 As Gnostice.PDFOne.PDFRenderOptions = New Gnostice.PDFOne.PDFRenderOptions()
        Me.tsMainMenu = New System.Windows.Forms.ToolStrip()
        Me.tsButton_patient = New System.Windows.Forms.ToolStripDropDownButton()
        Me.NewPatientToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FindPatientToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsButton_rfts = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsMenuItem_NewTest = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_RoutineRft = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_BronchialChallengeTests = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_AltitudeSimulation = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_spt = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_cpx = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_WalkTests = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_EditTest = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsButton_View = New System.Windows.Forms.ToolStripDropDownButton()
        Me.ListOfTestsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListOfPatientsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TrendToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsButton_Reporting = New System.Windows.Forms.ToolStripButton()
        Me.tsButton_tools = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsMenuItem_Preferences = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NormalValuesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_PredsManager = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_dbreports = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_ActivityReport = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_SptPanelBuilder = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_peoplemanager = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_ChallengeProtocolBuilder = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuItem_about = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsButton_exit = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip()
        Me.tstxt_PatientName = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripLabel_UR = New System.Windows.Forms.ToolStripLabel()
        Me.tstxt_UR = New System.Windows.Forms.ToolStripTextBox()
        Me.toolbtnNewPt = New System.Windows.Forms.ToolStripButton()
        Me.toolbtnEditPt = New System.Windows.Forms.ToolStripButton()
        Me.toolbtnFindPt = New System.Windows.Forms.ToolStripButton()
        Me.toolbtnBookings = New System.Windows.Forms.ToolStripButton()
        Me.toolbtnReporting = New System.Windows.Forms.ToolStripButton()
        Me.toolbtnToolbox = New System.Windows.Forms.ToolStripDropDownButton()
        Me.PreferencesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListsToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.PredictedValuesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PredictedValuesEditorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MigrateJjpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.splitter = New System.Windows.Forms.SplitContainer()
        Me.splitterUnits = New System.Windows.Forms.SplitContainer()
        Me.lvUnits = New System.Windows.Forms.ListView()
        Me.tsUnits = New System.Windows.Forms.ToolStrip()
        Me.tsLblUnits = New System.Windows.Forms.ToolStripLabel()
        Me.tsBtnExpandUnits = New System.Windows.Forms.ToolStripButton()
        Me.splitterPts = New System.Windows.Forms.SplitContainer()
        Me.lvRecall = New System.Windows.Forms.ListView()
        Me.tsTests = New System.Windows.Forms.ToolStrip()
        Me.tsLblTests = New System.Windows.Forms.ToolStripLabel()
        Me.dtp = New System.Windows.Forms.DateTimePicker()
        Me.panelDates = New System.Windows.Forms.Panel()
        Me.txtEnd = New System.Windows.Forms.TextBox()
        Me.ContextMenuStrip_date = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem6 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem7 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem8 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem10 = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtStart = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tsPtList = New System.Windows.Forms.ToolStrip()
        Me.tsLblPatients = New System.Windows.Forms.ToolStripLabel()
        Me.tsbtnExpandPts = New System.Windows.Forms.ToolStripButton()
        Me.lvPatients = New System.Windows.Forms.ListView()
        Me.tabReports = New System.Windows.Forms.TabControl()
        Me.tabpageReport = New System.Windows.Forms.TabPage()
        Me.PdfViewer1 = New Gnostice.PDFOne.Windows.PDFViewer.PDFViewer()
        Me.tabpageTrend = New System.Windows.Forms.TabPage()
        Me.splitTrend = New System.Windows.Forms.SplitContainer()
        Me.tabpageTrendOptions = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkAutoscale = New System.Windows.Forms.CheckBox()
        Me.panelScale = New System.Windows.Forms.Panel()
        Me.txtyaxis_max = New System.Windows.Forms.MaskedTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtyaxis_min = New System.Windows.Forms.MaskedTextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkShowValues = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.panelChangeRef = New System.Windows.Forms.Panel()
        Me.cmboRefTestDate = New System.Windows.Forms.ComboBox()
        Me.optPlot_selected = New System.Windows.Forms.RadioButton()
        Me.optPlot_earliest = New System.Windows.Forms.RadioButton()
        Me.optPlot_ppn = New System.Windows.Forms.RadioButton()
        Me.optPlot_pcChange = New System.Windows.Forms.RadioButton()
        Me.optPlot_absolute = New System.Windows.Forms.RadioButton()
        Me.sstrip = New System.Windows.Forms.StatusStrip()
        Me.ts_lblVersion = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ts_lblBuild = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ts_lblLoginID = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ts_lblServer = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ts_lblVersionMsg = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ts_lblAccess = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ts_lblHealthService = New System.Windows.Forms.ToolStripStatusLabel()
        Me.PdfPrinter1 = New Gnostice.PDFOne.PDFPrinter.PDFPrinter()
        Me.tsMainMenu.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        CType(Me.splitter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitter.Panel1.SuspendLayout()
        Me.splitter.Panel2.SuspendLayout()
        Me.splitter.SuspendLayout()
        CType(Me.splitterUnits, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitterUnits.Panel1.SuspendLayout()
        Me.splitterUnits.Panel2.SuspendLayout()
        Me.splitterUnits.SuspendLayout()
        Me.tsUnits.SuspendLayout()
        CType(Me.splitterPts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitterPts.Panel1.SuspendLayout()
        Me.splitterPts.Panel2.SuspendLayout()
        Me.splitterPts.SuspendLayout()
        Me.tsTests.SuspendLayout()
        Me.panelDates.SuspendLayout()
        Me.ContextMenuStrip_date.SuspendLayout()
        Me.tsPtList.SuspendLayout()
        Me.tabReports.SuspendLayout()
        Me.tabpageReport.SuspendLayout()
        Me.tabpageTrend.SuspendLayout()
        CType(Me.splitTrend, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.splitTrend.SuspendLayout()
        Me.tabpageTrendOptions.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.panelScale.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.panelChangeRef.SuspendLayout()
        Me.sstrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'tsMainMenu
        '
        Me.tsMainMenu.AutoSize = False
        Me.tsMainMenu.BackColor = System.Drawing.Color.LightGray
        Me.tsMainMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsButton_patient, Me.tsButton_rfts, Me.tsButton_View, Me.tsButton_Reporting, Me.tsButton_tools, Me.tsButton_exit})
        Me.tsMainMenu.Location = New System.Drawing.Point(0, 0)
        Me.tsMainMenu.Name = "tsMainMenu"
        Me.tsMainMenu.Size = New System.Drawing.Size(1223, 27)
        Me.tsMainMenu.TabIndex = 2
        Me.tsMainMenu.Text = "ToolStrip1"
        '
        'tsButton_patient
        '
        Me.tsButton_patient.BackColor = System.Drawing.Color.LightGray
        Me.tsButton_patient.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsButton_patient.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewPatientToolStripMenuItem, Me.EditToolStripMenuItem, Me.FindPatientToolStripMenuItem})
        Me.tsButton_patient.Font = New System.Drawing.Font("Segoe UI", 8.830189!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsButton_patient.ForeColor = System.Drawing.Color.Black
        Me.tsButton_patient.Image = CType(resources.GetObject("tsButton_patient.Image"), System.Drawing.Image)
        Me.tsButton_patient.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsButton_patient.Name = "tsButton_patient"
        Me.tsButton_patient.Size = New System.Drawing.Size(57, 24)
        Me.tsButton_patient.Text = "Patient"
        '
        'NewPatientToolStripMenuItem
        '
        Me.NewPatientToolStripMenuItem.BackColor = System.Drawing.Color.Gainsboro
        Me.NewPatientToolStripMenuItem.Name = "NewPatientToolStripMenuItem"
        Me.NewPatientToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.NewPatientToolStripMenuItem.Text = "New"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'FindPatientToolStripMenuItem
        '
        Me.FindPatientToolStripMenuItem.BackColor = System.Drawing.Color.Gainsboro
        Me.FindPatientToolStripMenuItem.Name = "FindPatientToolStripMenuItem"
        Me.FindPatientToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.FindPatientToolStripMenuItem.Text = "Find"
        '
        'tsButton_rfts
        '
        Me.tsButton_rfts.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsButton_rfts.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsMenuItem_NewTest, Me.tsMenuItem_EditTest})
        Me.tsButton_rfts.Image = CType(resources.GetObject("tsButton_rfts.Image"), System.Drawing.Image)
        Me.tsButton_rfts.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsButton_rfts.Name = "tsButton_rfts"
        Me.tsButton_rfts.Size = New System.Drawing.Size(42, 24)
        Me.tsButton_rfts.Text = "Test"
        Me.tsButton_rfts.ToolTipText = "Test"
        '
        'tsMenuItem_NewTest
        '
        Me.tsMenuItem_NewTest.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsMenuItem_RoutineRft, Me.tsMenuItem_BronchialChallengeTests, Me.tsMenuItem_AltitudeSimulation, Me.tsMenuItem_spt, Me.tsMenuItem_cpx, Me.tsMenuItem_WalkTests})
        Me.tsMenuItem_NewTest.Name = "tsMenuItem_NewTest"
        Me.tsMenuItem_NewTest.Size = New System.Drawing.Size(162, 22)
        Me.tsMenuItem_NewTest.Text = "New"
        '
        'tsMenuItem_RoutineRft
        '
        Me.tsMenuItem_RoutineRft.Name = "tsMenuItem_RoutineRft"
        Me.tsMenuItem_RoutineRft.Size = New System.Drawing.Size(205, 22)
        Me.tsMenuItem_RoutineRft.Text = "Routine RFTs"
        '
        'tsMenuItem_BronchialChallengeTests
        '
        Me.tsMenuItem_BronchialChallengeTests.Name = "tsMenuItem_BronchialChallengeTests"
        Me.tsMenuItem_BronchialChallengeTests.Size = New System.Drawing.Size(205, 22)
        Me.tsMenuItem_BronchialChallengeTests.Text = "Bronchial challenge tests"
        '
        'tsMenuItem_AltitudeSimulation
        '
        Me.tsMenuItem_AltitudeSimulation.Name = "tsMenuItem_AltitudeSimulation"
        Me.tsMenuItem_AltitudeSimulation.Size = New System.Drawing.Size(205, 22)
        Me.tsMenuItem_AltitudeSimulation.Text = "Altitude simulation test"
        '
        'tsMenuItem_spt
        '
        Me.tsMenuItem_spt.Name = "tsMenuItem_spt"
        Me.tsMenuItem_spt.Size = New System.Drawing.Size(205, 22)
        Me.tsMenuItem_spt.Text = "Skin prick test"
        '
        'tsMenuItem_cpx
        '
        Me.tsMenuItem_cpx.Name = "tsMenuItem_cpx"
        Me.tsMenuItem_cpx.Size = New System.Drawing.Size(205, 22)
        Me.tsMenuItem_cpx.Text = "CPET"
        '
        'tsMenuItem_WalkTests
        '
        Me.tsMenuItem_WalkTests.Name = "tsMenuItem_WalkTests"
        Me.tsMenuItem_WalkTests.Size = New System.Drawing.Size(205, 22)
        Me.tsMenuItem_WalkTests.Text = "Walk tests"
        '
        'tsMenuItem_EditTest
        '
        Me.tsMenuItem_EditTest.Name = "tsMenuItem_EditTest"
        Me.tsMenuItem_EditTest.Size = New System.Drawing.Size(162, 22)
        Me.tsMenuItem_EditTest.Text = "Edit selected test"
        '
        'tsButton_View
        '
        Me.tsButton_View.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsButton_View.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ListOfTestsToolStripMenuItem, Me.ListOfPatientsToolStripMenuItem, Me.TrendToolStripMenuItem})
        Me.tsButton_View.Image = CType(resources.GetObject("tsButton_View.Image"), System.Drawing.Image)
        Me.tsButton_View.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsButton_View.Name = "tsButton_View"
        Me.tsButton_View.Size = New System.Drawing.Size(45, 24)
        Me.tsButton_View.Text = "View"
        Me.tsButton_View.Visible = False
        '
        'ListOfTestsToolStripMenuItem
        '
        Me.ListOfTestsToolStripMenuItem.Name = "ListOfTestsToolStripMenuItem"
        Me.ListOfTestsToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.ListOfTestsToolStripMenuItem.Text = "List of tests"
        '
        'ListOfPatientsToolStripMenuItem
        '
        Me.ListOfPatientsToolStripMenuItem.Name = "ListOfPatientsToolStripMenuItem"
        Me.ListOfPatientsToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.ListOfPatientsToolStripMenuItem.Text = "List of patients"
        '
        'TrendToolStripMenuItem
        '
        Me.TrendToolStripMenuItem.Name = "TrendToolStripMenuItem"
        Me.TrendToolStripMenuItem.Size = New System.Drawing.Size(151, 22)
        Me.TrendToolStripMenuItem.Text = "Trend"
        '
        'tsButton_Reporting
        '
        Me.tsButton_Reporting.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsButton_Reporting.Image = CType(resources.GetObject("tsButton_Reporting.Image"), System.Drawing.Image)
        Me.tsButton_Reporting.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsButton_Reporting.Name = "tsButton_Reporting"
        Me.tsButton_Reporting.Size = New System.Drawing.Size(63, 24)
        Me.tsButton_Reporting.Text = "Reporting"
        '
        'tsButton_tools
        '
        Me.tsButton_tools.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsButton_tools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsMenuItem_Preferences, Me.tsMenuItem_PredsManager, Me.tsMenuItem_dbreports, Me.tsMenuItem_SptPanelBuilder, Me.tsMenuItem_peoplemanager, Me.tsMenuItem_ChallengeProtocolBuilder, Me.tsMenuItem_about})
        Me.tsButton_tools.Image = CType(resources.GetObject("tsButton_tools.Image"), System.Drawing.Image)
        Me.tsButton_tools.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsButton_tools.Name = "tsButton_tools"
        Me.tsButton_tools.Size = New System.Drawing.Size(63, 24)
        Me.tsButton_tools.Text = "Toolbox"
        '
        'tsMenuItem_Preferences
        '
        Me.tsMenuItem_Preferences.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ListsToolStripMenuItem, Me.NormalValuesToolStripMenuItem})
        Me.tsMenuItem_Preferences.Name = "tsMenuItem_Preferences"
        Me.tsMenuItem_Preferences.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_Preferences.Text = "Preferences"
        '
        'ListsToolStripMenuItem
        '
        Me.ListsToolStripMenuItem.Name = "ListsToolStripMenuItem"
        Me.ListsToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.ListsToolStripMenuItem.Text = "Lists"
        '
        'NormalValuesToolStripMenuItem
        '
        Me.NormalValuesToolStripMenuItem.Name = "NormalValuesToolStripMenuItem"
        Me.NormalValuesToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.NormalValuesToolStripMenuItem.Text = "Predicted values"
        '
        'tsMenuItem_PredsManager
        '
        Me.tsMenuItem_PredsManager.Name = "tsMenuItem_PredsManager"
        Me.tsMenuItem_PredsManager.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_PredsManager.Text = "Predicted values editor"
        '
        'tsMenuItem_dbreports
        '
        Me.tsMenuItem_dbreports.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsMenuItem_ActivityReport})
        Me.tsMenuItem_dbreports.Name = "tsMenuItem_dbreports"
        Me.tsMenuItem_dbreports.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_dbreports.Text = "Database reports"
        '
        'tsMenuItem_ActivityReport
        '
        Me.tsMenuItem_ActivityReport.Name = "tsMenuItem_ActivityReport"
        Me.tsMenuItem_ActivityReport.Size = New System.Drawing.Size(149, 22)
        Me.tsMenuItem_ActivityReport.Text = "Activity report"
        '
        'tsMenuItem_SptPanelBuilder
        '
        Me.tsMenuItem_SptPanelBuilder.Name = "tsMenuItem_SptPanelBuilder"
        Me.tsMenuItem_SptPanelBuilder.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_SptPanelBuilder.Text = "SPT panel builder"
        '
        'tsMenuItem_peoplemanager
        '
        Me.tsMenuItem_peoplemanager.Name = "tsMenuItem_peoplemanager"
        Me.tsMenuItem_peoplemanager.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_peoplemanager.Text = "User manager"
        '
        'tsMenuItem_ChallengeProtocolBuilder
        '
        Me.tsMenuItem_ChallengeProtocolBuilder.Enabled = False
        Me.tsMenuItem_ChallengeProtocolBuilder.Name = "tsMenuItem_ChallengeProtocolBuilder"
        Me.tsMenuItem_ChallengeProtocolBuilder.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_ChallengeProtocolBuilder.Text = "Challenge protocol builder"
        '
        'tsMenuItem_about
        '
        Me.tsMenuItem_about.Name = "tsMenuItem_about"
        Me.tsMenuItem_about.Size = New System.Drawing.Size(215, 22)
        Me.tsMenuItem_about.Text = "About ResLab"
        '
        'tsButton_exit
        '
        Me.tsButton_exit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tsButton_exit.Image = CType(resources.GetObject("tsButton_exit.Image"), System.Drawing.Image)
        Me.tsButton_exit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsButton_exit.Name = "tsButton_exit"
        Me.tsButton_exit.Size = New System.Drawing.Size(29, 24)
        Me.tsButton_exit.Text = "Exit"
        '
        'ToolStrip2
        '
        Me.ToolStrip2.AutoSize = False
        Me.ToolStrip2.BackColor = System.Drawing.SystemColors.Menu
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tstxt_PatientName, Me.ToolStripLabel_UR, Me.tstxt_UR, Me.toolbtnNewPt, Me.toolbtnEditPt, Me.toolbtnFindPt, Me.toolbtnBookings, Me.toolbtnReporting, Me.toolbtnToolbox})
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 27)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(1223, 40)
        Me.ToolStrip2.TabIndex = 5
        Me.ToolStrip2.Text = "ToolStrip2"
        '
        'tstxt_PatientName
        '
        Me.tstxt_PatientName.AutoSize = False
        Me.tstxt_PatientName.BackColor = System.Drawing.Color.SteelBlue
        Me.tstxt_PatientName.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tstxt_PatientName.ForeColor = System.Drawing.Color.White
        Me.tstxt_PatientName.Name = "tstxt_PatientName"
        Me.tstxt_PatientName.Size = New System.Drawing.Size(200, 25)
        Me.tstxt_PatientName.ToolTipText = "Enter patient's surname and optionally first name/initial to search (surname, fir" &
    "stname)"
        '
        'ToolStripLabel_UR
        '
        Me.ToolStripLabel_UR.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel_UR.Name = "ToolStripLabel_UR"
        Me.ToolStripLabel_UR.Size = New System.Drawing.Size(41, 37)
        Me.ToolStripLabel_UR.Text = "MRN:"
        '
        'tstxt_UR
        '
        Me.tstxt_UR.AutoSize = False
        Me.tstxt_UR.BackColor = System.Drawing.Color.SteelBlue
        Me.tstxt_UR.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tstxt_UR.ForeColor = System.Drawing.Color.White
        Me.tstxt_UR.Name = "tstxt_UR"
        Me.tstxt_UR.Size = New System.Drawing.Size(200, 25)
        Me.tstxt_UR.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.tstxt_UR.ToolTipText = "Enter patient's number for search"
        '
        'toolbtnNewPt
        '
        Me.toolbtnNewPt.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.toolbtnNewPt.Image = CType(resources.GetObject("toolbtnNewPt.Image"), System.Drawing.Image)
        Me.toolbtnNewPt.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolbtnNewPt.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolbtnNewPt.Name = "toolbtnNewPt"
        Me.toolbtnNewPt.Size = New System.Drawing.Size(36, 37)
        Me.toolbtnNewPt.ToolTipText = "New patient"
        '
        'toolbtnEditPt
        '
        Me.toolbtnEditPt.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.toolbtnEditPt.Image = CType(resources.GetObject("toolbtnEditPt.Image"), System.Drawing.Image)
        Me.toolbtnEditPt.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolbtnEditPt.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolbtnEditPt.Name = "toolbtnEditPt"
        Me.toolbtnEditPt.Size = New System.Drawing.Size(36, 37)
        Me.toolbtnEditPt.ToolTipText = "Edit patient details"
        '
        'toolbtnFindPt
        '
        Me.toolbtnFindPt.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.toolbtnFindPt.Image = CType(resources.GetObject("toolbtnFindPt.Image"), System.Drawing.Image)
        Me.toolbtnFindPt.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolbtnFindPt.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolbtnFindPt.Name = "toolbtnFindPt"
        Me.toolbtnFindPt.Size = New System.Drawing.Size(36, 37)
        Me.toolbtnFindPt.Text = "ToolStripButton5"
        Me.toolbtnFindPt.ToolTipText = "Find patient"
        '
        'toolbtnBookings
        '
        Me.toolbtnBookings.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.toolbtnBookings.Image = CType(resources.GetObject("toolbtnBookings.Image"), System.Drawing.Image)
        Me.toolbtnBookings.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolbtnBookings.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolbtnBookings.Name = "toolbtnBookings"
        Me.toolbtnBookings.Size = New System.Drawing.Size(36, 37)
        Me.toolbtnBookings.Text = "ToolStripButton2"
        Me.toolbtnBookings.ToolTipText = "Bookings"
        '
        'toolbtnReporting
        '
        Me.toolbtnReporting.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.toolbtnReporting.Image = CType(resources.GetObject("toolbtnReporting.Image"), System.Drawing.Image)
        Me.toolbtnReporting.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolbtnReporting.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolbtnReporting.Name = "toolbtnReporting"
        Me.toolbtnReporting.Size = New System.Drawing.Size(36, 37)
        Me.toolbtnReporting.Text = "ToolStripButton2"
        Me.toolbtnReporting.ToolTipText = "Reporting"
        '
        'toolbtnToolbox
        '
        Me.toolbtnToolbox.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.toolbtnToolbox.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PreferencesToolStripMenuItem, Me.PredictedValuesEditorToolStripMenuItem, Me.MigrateJjpToolStripMenuItem})
        Me.toolbtnToolbox.Image = CType(resources.GetObject("toolbtnToolbox.Image"), System.Drawing.Image)
        Me.toolbtnToolbox.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.toolbtnToolbox.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.toolbtnToolbox.Name = "toolbtnToolbox"
        Me.toolbtnToolbox.Size = New System.Drawing.Size(45, 37)
        Me.toolbtnToolbox.Text = "ToolStripDropDownButton2"
        Me.toolbtnToolbox.ToolTipText = "Toolbox"
        Me.toolbtnToolbox.Visible = False
        '
        'PreferencesToolStripMenuItem
        '
        Me.PreferencesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ListsToolStripMenuItem1, Me.PredictedValuesToolStripMenuItem})
        Me.PreferencesToolStripMenuItem.Name = "PreferencesToolStripMenuItem"
        Me.PreferencesToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.PreferencesToolStripMenuItem.Text = "Preferences"
        '
        'ListsToolStripMenuItem1
        '
        Me.ListsToolStripMenuItem1.Name = "ListsToolStripMenuItem1"
        Me.ListsToolStripMenuItem1.Size = New System.Drawing.Size(160, 22)
        Me.ListsToolStripMenuItem1.Text = "Lists"
        '
        'PredictedValuesToolStripMenuItem
        '
        Me.PredictedValuesToolStripMenuItem.Name = "PredictedValuesToolStripMenuItem"
        Me.PredictedValuesToolStripMenuItem.Size = New System.Drawing.Size(160, 22)
        Me.PredictedValuesToolStripMenuItem.Text = "Predicted values"
        '
        'PredictedValuesEditorToolStripMenuItem
        '
        Me.PredictedValuesEditorToolStripMenuItem.Name = "PredictedValuesEditorToolStripMenuItem"
        Me.PredictedValuesEditorToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.PredictedValuesEditorToolStripMenuItem.Text = "Predicted values editor"
        '
        'MigrateJjpToolStripMenuItem
        '
        Me.MigrateJjpToolStripMenuItem.Name = "MigrateJjpToolStripMenuItem"
        Me.MigrateJjpToolStripMenuItem.Size = New System.Drawing.Size(194, 22)
        Me.MigrateJjpToolStripMenuItem.Text = "Migrate jjp"
        '
        'splitter
        '
        Me.splitter.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.splitter.Dock = System.Windows.Forms.DockStyle.Top
        Me.splitter.Location = New System.Drawing.Point(0, 67)
        Me.splitter.Name = "splitter"
        '
        'splitter.Panel1
        '
        Me.splitter.Panel1.Controls.Add(Me.splitterUnits)
        '
        'splitter.Panel2
        '
        Me.splitter.Panel2.Controls.Add(Me.tabReports)
        Me.splitter.Size = New System.Drawing.Size(1223, 585)
        Me.splitter.SplitterDistance = 417
        Me.splitter.TabIndex = 8
        '
        'splitterUnits
        '
        Me.splitterUnits.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitterUnits.Location = New System.Drawing.Point(0, 0)
        Me.splitterUnits.Margin = New System.Windows.Forms.Padding(2)
        Me.splitterUnits.Name = "splitterUnits"
        Me.splitterUnits.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitterUnits.Panel1
        '
        Me.splitterUnits.Panel1.Controls.Add(Me.lvUnits)
        Me.splitterUnits.Panel1.Controls.Add(Me.tsUnits)
        '
        'splitterUnits.Panel2
        '
        Me.splitterUnits.Panel2.Controls.Add(Me.splitterPts)
        Me.splitterUnits.Size = New System.Drawing.Size(413, 581)
        Me.splitterUnits.SplitterDistance = 69
        Me.splitterUnits.SplitterWidth = 3
        Me.splitterUnits.TabIndex = 0
        '
        'lvUnits
        '
        Me.lvUnits.BackColor = System.Drawing.Color.OldLace
        Me.lvUnits.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvUnits.CheckBoxes = True
        Me.lvUnits.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvUnits.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvUnits.HideSelection = False
        Me.lvUnits.Location = New System.Drawing.Point(0, 22)
        Me.lvUnits.MultiSelect = False
        Me.lvUnits.Name = "lvUnits"
        Me.lvUnits.Size = New System.Drawing.Size(413, 47)
        Me.lvUnits.TabIndex = 3
        Me.lvUnits.UseCompatibleStateImageBehavior = False
        '
        'tsUnits
        '
        Me.tsUnits.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.tsUnits.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsUnits.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsLblUnits, Me.tsBtnExpandUnits})
        Me.tsUnits.Location = New System.Drawing.Point(0, 0)
        Me.tsUnits.Name = "tsUnits"
        Me.tsUnits.Size = New System.Drawing.Size(413, 25)
        Me.tsUnits.TabIndex = 0
        '
        'tsLblUnits
        '
        Me.tsLblUnits.Name = "tsLblUnits"
        Me.tsLblUnits.Size = New System.Drawing.Size(34, 22)
        Me.tsLblUnits.Text = "Units"
        '
        'tsBtnExpandUnits
        '
        Me.tsBtnExpandUnits.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsBtnExpandUnits.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsBtnExpandUnits.Image = CType(resources.GetObject("tsBtnExpandUnits.Image"), System.Drawing.Image)
        Me.tsBtnExpandUnits.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsBtnExpandUnits.Name = "tsBtnExpandUnits"
        Me.tsBtnExpandUnits.Size = New System.Drawing.Size(23, 22)
        Me.tsBtnExpandUnits.Text = "ToolStripButton1"
        '
        'splitterPts
        '
        Me.splitterPts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitterPts.Location = New System.Drawing.Point(0, 0)
        Me.splitterPts.Margin = New System.Windows.Forms.Padding(2)
        Me.splitterPts.Name = "splitterPts"
        Me.splitterPts.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'splitterPts.Panel1
        '
        Me.splitterPts.Panel1.Controls.Add(Me.lvRecall)
        Me.splitterPts.Panel1.Controls.Add(Me.tsTests)
        '
        'splitterPts.Panel2
        '
        Me.splitterPts.Panel2.Controls.Add(Me.dtp)
        Me.splitterPts.Panel2.Controls.Add(Me.panelDates)
        Me.splitterPts.Panel2.Controls.Add(Me.tsPtList)
        Me.splitterPts.Panel2.Controls.Add(Me.lvPatients)
        Me.splitterPts.Size = New System.Drawing.Size(413, 509)
        Me.splitterPts.SplitterDistance = 148
        Me.splitterPts.SplitterWidth = 3
        Me.splitterPts.TabIndex = 0
        '
        'lvRecall
        '
        Me.lvRecall.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.lvRecall.BackColor = System.Drawing.Color.OldLace
        Me.lvRecall.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvRecall.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvRecall.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvRecall.Location = New System.Drawing.Point(0, 26)
        Me.lvRecall.Name = "lvRecall"
        Me.lvRecall.Size = New System.Drawing.Size(413, 122)
        Me.lvRecall.TabIndex = 0
        Me.lvRecall.UseCompatibleStateImageBehavior = False
        '
        'tsTests
        '
        Me.tsTests.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.tsTests.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsTests.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsLblTests})
        Me.tsTests.Location = New System.Drawing.Point(0, 0)
        Me.tsTests.Name = "tsTests"
        Me.tsTests.Size = New System.Drawing.Size(413, 25)
        Me.tsTests.TabIndex = 2
        Me.tsTests.Text = "Tests"
        '
        'tsLblTests
        '
        Me.tsLblTests.Image = CType(resources.GetObject("tsLblTests.Image"), System.Drawing.Image)
        Me.tsLblTests.Name = "tsLblTests"
        Me.tsLblTests.Size = New System.Drawing.Size(49, 22)
        Me.tsLblTests.Text = "Tests"
        '
        'dtp
        '
        Me.dtp.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtp.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtp.Location = New System.Drawing.Point(201, 54)
        Me.dtp.Name = "dtp"
        Me.dtp.Size = New System.Drawing.Size(98, 23)
        Me.dtp.TabIndex = 1
        Me.dtp.Visible = False
        '
        'panelDates
        '
        Me.panelDates.BackColor = System.Drawing.SystemColors.ControlLight
        Me.panelDates.Controls.Add(Me.txtEnd)
        Me.panelDates.Controls.Add(Me.txtStart)
        Me.panelDates.Controls.Add(Me.Label3)
        Me.panelDates.Controls.Add(Me.Label2)
        Me.panelDates.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelDates.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold)
        Me.panelDates.Location = New System.Drawing.Point(0, 25)
        Me.panelDates.Name = "panelDates"
        Me.panelDates.Size = New System.Drawing.Size(413, 28)
        Me.panelDates.TabIndex = 3
        '
        'txtEnd
        '
        Me.txtEnd.ContextMenuStrip = Me.ContextMenuStrip_date
        Me.txtEnd.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEnd.Location = New System.Drawing.Point(233, 2)
        Me.txtEnd.Name = "txtEnd"
        Me.txtEnd.Size = New System.Drawing.Size(81, 23)
        Me.txtEnd.TabIndex = 4
        '
        'ContextMenuStrip_date
        '
        Me.ContextMenuStrip_date.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem4, Me.ToolStripMenuItem5, Me.ToolStripMenuItem6, Me.ToolStripMenuItem7, Me.ToolStripMenuItem8, Me.ToolStripMenuItem9, Me.ToolStripMenuItem10})
        Me.ContextMenuStrip_date.Name = "ContextMenuStrip_date"
        Me.ContextMenuStrip_date.Size = New System.Drawing.Size(166, 158)
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem4.Tag = "1"
        Me.ToolStripMenuItem4.Text = "Clear date entries"
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem5.Tag = "2"
        Me.ToolStripMenuItem5.Text = "Today"
        '
        'ToolStripMenuItem6
        '
        Me.ToolStripMenuItem6.Name = "ToolStripMenuItem6"
        Me.ToolStripMenuItem6.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem6.Tag = "3"
        Me.ToolStripMenuItem6.Text = "Yesterday"
        '
        'ToolStripMenuItem7
        '
        Me.ToolStripMenuItem7.Name = "ToolStripMenuItem7"
        Me.ToolStripMenuItem7.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem7.Tag = "4"
        Me.ToolStripMenuItem7.Text = "This week"
        '
        'ToolStripMenuItem8
        '
        Me.ToolStripMenuItem8.Name = "ToolStripMenuItem8"
        Me.ToolStripMenuItem8.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem8.Tag = "5"
        Me.ToolStripMenuItem8.Text = "Last week"
        '
        'ToolStripMenuItem9
        '
        Me.ToolStripMenuItem9.Name = "ToolStripMenuItem9"
        Me.ToolStripMenuItem9.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem9.Tag = "6"
        Me.ToolStripMenuItem9.Text = "This month"
        '
        'ToolStripMenuItem10
        '
        Me.ToolStripMenuItem10.Name = "ToolStripMenuItem10"
        Me.ToolStripMenuItem10.Size = New System.Drawing.Size(165, 22)
        Me.ToolStripMenuItem10.Tag = "7"
        Me.ToolStripMenuItem10.Text = "Last month"
        '
        'txtStart
        '
        Me.txtStart.ContextMenuStrip = Me.ContextMenuStrip_date
        Me.txtStart.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStart.Location = New System.Drawing.Point(82, 2)
        Me.txtStart.Name = "txtStart"
        Me.txtStart.Size = New System.Drawing.Size(81, 23)
        Me.txtStart.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(170, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 15)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "End date"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(14, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 15)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Start date"
        '
        'tsPtList
        '
        Me.tsPtList.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.tsPtList.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsPtList.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsLblPatients, Me.tsbtnExpandPts})
        Me.tsPtList.Location = New System.Drawing.Point(0, 0)
        Me.tsPtList.Name = "tsPtList"
        Me.tsPtList.Size = New System.Drawing.Size(413, 25)
        Me.tsPtList.TabIndex = 1
        Me.tsPtList.Text = "ToolStrip4"
        '
        'tsLblPatients
        '
        Me.tsLblPatients.Image = CType(resources.GetObject("tsLblPatients.Image"), System.Drawing.Image)
        Me.tsLblPatients.Name = "tsLblPatients"
        Me.tsLblPatients.Size = New System.Drawing.Size(155, 22)
        Me.tsLblPatients.Text = "Advanced Patient Search"
        Me.tsLblPatients.ToolTipText = "Left click on date field for calendar, right click for date options"
        '
        'tsbtnExpandPts
        '
        Me.tsbtnExpandPts.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtnExpandPts.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsbtnExpandPts.Image = CType(resources.GetObject("tsbtnExpandPts.Image"), System.Drawing.Image)
        Me.tsbtnExpandPts.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtnExpandPts.Name = "tsbtnExpandPts"
        Me.tsbtnExpandPts.Size = New System.Drawing.Size(23, 22)
        Me.tsbtnExpandPts.Text = "ToolStripButton1"
        '
        'lvPatients
        '
        Me.lvPatients.BackColor = System.Drawing.Color.OldLace
        Me.lvPatients.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvPatients.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvPatients.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvPatients.Location = New System.Drawing.Point(0, 334)
        Me.lvPatients.Name = "lvPatients"
        Me.lvPatients.Size = New System.Drawing.Size(413, 24)
        Me.lvPatients.TabIndex = 4
        Me.lvPatients.UseCompatibleStateImageBehavior = False
        '
        'tabReports
        '
        Me.tabReports.Controls.Add(Me.tabpageReport)
        Me.tabReports.Controls.Add(Me.tabpageTrend)
        Me.tabReports.Controls.Add(Me.tabpageTrendOptions)
        Me.tabReports.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabReports.Font = New System.Drawing.Font("Segoe UI", 8.150944!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabReports.Location = New System.Drawing.Point(0, 0)
        Me.tabReports.Name = "tabReports"
        Me.tabReports.SelectedIndex = 0
        Me.tabReports.Size = New System.Drawing.Size(798, 581)
        Me.tabReports.TabIndex = 37
        '
        'tabpageReport
        '
        Me.tabpageReport.Controls.Add(Me.PdfViewer1)
        Me.tabpageReport.Location = New System.Drawing.Point(4, 22)
        Me.tabpageReport.Name = "tabpageReport"
        Me.tabpageReport.Padding = New System.Windows.Forms.Padding(3)
        Me.tabpageReport.Size = New System.Drawing.Size(790, 555)
        Me.tabpageReport.TabIndex = 0
        Me.tabpageReport.Text = "Test report (pdf)"
        Me.tabpageReport.UseVisualStyleBackColor = True
        '
        'PdfViewer1
        '
        Me.PdfViewer1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.PdfViewer1.BorderWidth = 10
        Me.PdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PdfViewer1.Document = Nothing
        Me.PdfViewer1.HScrollBar = Gnostice.PDFOne.Windows.PDFViewer.ScrollBarVisibility.Always
        Me.PdfViewer1.HScrollValue = 0
        Me.PdfViewer1.KeyNavigationEnabled = True
        Me.PdfViewer1.Location = New System.Drawing.Point(3, 3)
        Me.PdfViewer1.Margin = New System.Windows.Forms.Padding(14, 32, 14, 32)
        Me.PdfViewer1.Name = "PdfViewer1"
        Me.PdfViewer1.PageBufferCount = 2
        Me.PdfViewer1.PageRotationAngle = Gnostice.PDFOne.Windows.PDFViewer.RotationAngle.Zero
        Me.PdfViewer1.Password = ""
        RenderingSettings2.AnnotsRenderingRule = CType((((Gnostice.PDFOne.ItemsRenderingRule.ifPrintableTrue Or Gnostice.PDFOne.ItemsRenderingRule.ifPrintableFalse) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifHiddenFalse) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifNoViewFalse), Gnostice.PDFOne.ItemsRenderingRule)
        RenderingSettings2.FormfieldsRenderingRule = CType((((Gnostice.PDFOne.ItemsRenderingRule.ifPrintableTrue Or Gnostice.PDFOne.ItemsRenderingRule.ifPrintableFalse) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifHiddenFalse) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifNoViewFalse), Gnostice.PDFOne.ItemsRenderingRule)
        ImageSettings2.ColorMode = Gnostice.PDFOne.ColorMode.Color
        RenderingSettings2.Image = ImageSettings2
        RenderingSettings2.ItemsToRender = CType(((((Gnostice.PDFOne.ItemsToRender.Annots Or Gnostice.PDFOne.ItemsToRender.FormFields) _
            Or Gnostice.PDFOne.ItemsToRender.Images) _
            Or Gnostice.PDFOne.ItemsToRender.Text) _
            Or Gnostice.PDFOne.ItemsToRender.Shapes), Gnostice.PDFOne.ItemsToRender)
        ViewerPreferences1.RenderingSettings = RenderingSettings2
        Me.PdfViewer1.Preferences = ViewerPreferences1
        PdfRenderOptions2.SmoothImages = True
        PdfRenderOptions2.SmoothLineart = True
        PdfRenderOptions2.SmoothText = True
        Me.PdfViewer1.RenderingOptions = PdfRenderOptions2
        Me.PdfViewer1.Size = New System.Drawing.Size(784, 549)
        Me.PdfViewer1.StdZoomType = Gnostice.PDFOne.Windows.PDFViewer.StandardZoomType.ActualSize
        Me.PdfViewer1.TabIndex = 0
        Me.PdfViewer1.VScrollBar = Gnostice.PDFOne.Windows.PDFViewer.ScrollBarVisibility.Always
        Me.PdfViewer1.VScrollValue = 0
        Me.PdfViewer1.ZoomPercent = 100.0R
        '
        'tabpageTrend
        '
        Me.tabpageTrend.Controls.Add(Me.splitTrend)
        Me.tabpageTrend.Location = New System.Drawing.Point(4, 22)
        Me.tabpageTrend.Name = "tabpageTrend"
        Me.tabpageTrend.Padding = New System.Windows.Forms.Padding(3)
        Me.tabpageTrend.Size = New System.Drawing.Size(790, 555)
        Me.tabpageTrend.TabIndex = 1
        Me.tabpageTrend.Text = "Trend"
        Me.tabpageTrend.UseVisualStyleBackColor = True
        '
        'splitTrend
        '
        Me.splitTrend.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.splitTrend.Dock = System.Windows.Forms.DockStyle.Fill
        Me.splitTrend.Location = New System.Drawing.Point(3, 3)
        Me.splitTrend.Name = "splitTrend"
        Me.splitTrend.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.splitTrend.Size = New System.Drawing.Size(784, 549)
        Me.splitTrend.SplitterDistance = 259
        Me.splitTrend.TabIndex = 0
        '
        'tabpageTrendOptions
        '
        Me.tabpageTrendOptions.Controls.Add(Me.GroupBox2)
        Me.tabpageTrendOptions.Controls.Add(Me.GroupBox1)
        Me.tabpageTrendOptions.Location = New System.Drawing.Point(4, 22)
        Me.tabpageTrendOptions.Margin = New System.Windows.Forms.Padding(2)
        Me.tabpageTrendOptions.Name = "tabpageTrendOptions"
        Me.tabpageTrendOptions.Padding = New System.Windows.Forms.Padding(2)
        Me.tabpageTrendOptions.Size = New System.Drawing.Size(790, 555)
        Me.tabpageTrendOptions.TabIndex = 2
        Me.tabpageTrendOptions.Text = "Trend options"
        Me.tabpageTrendOptions.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkAutoscale)
        Me.GroupBox2.Controls.Add(Me.panelScale)
        Me.GroupBox2.Controls.Add(Me.chkShowValues)
        Me.GroupBox2.Location = New System.Drawing.Point(294, 20)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(275, 190)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Plot properties"
        '
        'chkAutoscale
        '
        Me.chkAutoscale.AutoSize = True
        Me.chkAutoscale.Checked = True
        Me.chkAutoscale.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAutoscale.Location = New System.Drawing.Point(16, 54)
        Me.chkAutoscale.Margin = New System.Windows.Forms.Padding(2)
        Me.chkAutoscale.Name = "chkAutoscale"
        Me.chkAutoscale.Size = New System.Drawing.Size(79, 17)
        Me.chkAutoscale.TabIndex = 1
        Me.chkAutoscale.Text = "Auto scale"
        Me.chkAutoscale.UseVisualStyleBackColor = True
        '
        'panelScale
        '
        Me.panelScale.Controls.Add(Me.txtyaxis_max)
        Me.panelScale.Controls.Add(Me.Label4)
        Me.panelScale.Controls.Add(Me.txtyaxis_min)
        Me.panelScale.Controls.Add(Me.Label1)
        Me.panelScale.Location = New System.Drawing.Point(99, 50)
        Me.panelScale.Margin = New System.Windows.Forms.Padding(2)
        Me.panelScale.Name = "panelScale"
        Me.panelScale.Size = New System.Drawing.Size(109, 52)
        Me.panelScale.TabIndex = 9
        Me.panelScale.Visible = False
        '
        'txtyaxis_max
        '
        Me.txtyaxis_max.Location = New System.Drawing.Point(60, 24)
        Me.txtyaxis_max.Margin = New System.Windows.Forms.Padding(2)
        Me.txtyaxis_max.Name = "txtyaxis_max"
        Me.txtyaxis_max.Size = New System.Drawing.Size(36, 22)
        Me.txtyaxis_max.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 28)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Y axis max"
        '
        'txtyaxis_min
        '
        Me.txtyaxis_min.Location = New System.Drawing.Point(60, 4)
        Me.txtyaxis_min.Margin = New System.Windows.Forms.Padding(2)
        Me.txtyaxis_min.Name = "txtyaxis_min"
        Me.txtyaxis_min.Size = New System.Drawing.Size(36, 22)
        Me.txtyaxis_min.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Y axis min"
        '
        'chkShowValues
        '
        Me.chkShowValues.AutoSize = True
        Me.chkShowValues.Checked = True
        Me.chkShowValues.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowValues.Location = New System.Drawing.Point(16, 30)
        Me.chkShowValues.Margin = New System.Windows.Forms.Padding(2)
        Me.chkShowValues.Name = "chkShowValues"
        Me.chkShowValues.Size = New System.Drawing.Size(147, 17)
        Me.chkShowValues.TabIndex = 0
        Me.chkShowValues.Text = "Show data point values"
        Me.chkShowValues.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.panelChangeRef)
        Me.GroupBox1.Controls.Add(Me.optPlot_ppn)
        Me.GroupBox1.Controls.Add(Me.optPlot_pcChange)
        Me.GroupBox1.Controls.Add(Me.optPlot_absolute)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 20)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(246, 262)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Plot results as"
        '
        'panelChangeRef
        '
        Me.panelChangeRef.Controls.Add(Me.cmboRefTestDate)
        Me.panelChangeRef.Controls.Add(Me.optPlot_selected)
        Me.panelChangeRef.Controls.Add(Me.optPlot_earliest)
        Me.panelChangeRef.Location = New System.Drawing.Point(53, 73)
        Me.panelChangeRef.Margin = New System.Windows.Forms.Padding(2)
        Me.panelChangeRef.Name = "panelChangeRef"
        Me.panelChangeRef.Size = New System.Drawing.Size(99, 80)
        Me.panelChangeRef.TabIndex = 8
        Me.panelChangeRef.Visible = False
        '
        'cmboRefTestDate
        '
        Me.cmboRefTestDate.FormattingEnabled = True
        Me.cmboRefTestDate.Location = New System.Drawing.Point(27, 52)
        Me.cmboRefTestDate.Margin = New System.Windows.Forms.Padding(2)
        Me.cmboRefTestDate.Name = "cmboRefTestDate"
        Me.cmboRefTestDate.Size = New System.Drawing.Size(71, 21)
        Me.cmboRefTestDate.TabIndex = 9
        Me.cmboRefTestDate.Visible = False
        '
        'optPlot_selected
        '
        Me.optPlot_selected.AutoSize = True
        Me.optPlot_selected.Location = New System.Drawing.Point(13, 28)
        Me.optPlot_selected.Margin = New System.Windows.Forms.Padding(2)
        Me.optPlot_selected.Name = "optPlot_selected"
        Me.optPlot_selected.Size = New System.Drawing.Size(90, 17)
        Me.optPlot_selected.TabIndex = 4
        Me.optPlot_selected.Text = "Selected test"
        Me.optPlot_selected.UseVisualStyleBackColor = True
        '
        'optPlot_earliest
        '
        Me.optPlot_earliest.AutoSize = True
        Me.optPlot_earliest.Checked = True
        Me.optPlot_earliest.Location = New System.Drawing.Point(13, 5)
        Me.optPlot_earliest.Margin = New System.Windows.Forms.Padding(2)
        Me.optPlot_earliest.Name = "optPlot_earliest"
        Me.optPlot_earliest.Size = New System.Drawing.Size(84, 17)
        Me.optPlot_earliest.TabIndex = 6
        Me.optPlot_earliest.TabStop = True
        Me.optPlot_earliest.Text = "Earliest test"
        Me.optPlot_earliest.UseVisualStyleBackColor = True
        '
        'optPlot_ppn
        '
        Me.optPlot_ppn.AutoSize = True
        Me.optPlot_ppn.Location = New System.Drawing.Point(33, 212)
        Me.optPlot_ppn.Margin = New System.Windows.Forms.Padding(2)
        Me.optPlot_ppn.Name = "optPlot_ppn"
        Me.optPlot_ppn.Size = New System.Drawing.Size(115, 17)
        Me.optPlot_ppn.TabIndex = 3
        Me.optPlot_ppn.TabStop = True
        Me.optPlot_ppn.Text = "Percent predicted"
        Me.optPlot_ppn.UseVisualStyleBackColor = True
        Me.optPlot_ppn.Visible = False
        '
        'optPlot_pcChange
        '
        Me.optPlot_pcChange.AutoSize = True
        Me.optPlot_pcChange.Location = New System.Drawing.Point(33, 50)
        Me.optPlot_pcChange.Margin = New System.Windows.Forms.Padding(2)
        Me.optPlot_pcChange.Name = "optPlot_pcChange"
        Me.optPlot_pcChange.Size = New System.Drawing.Size(104, 17)
        Me.optPlot_pcChange.TabIndex = 2
        Me.optPlot_pcChange.TabStop = True
        Me.optPlot_pcChange.Text = "Percent change"
        Me.optPlot_pcChange.UseVisualStyleBackColor = True
        '
        'optPlot_absolute
        '
        Me.optPlot_absolute.AutoSize = True
        Me.optPlot_absolute.Checked = True
        Me.optPlot_absolute.Location = New System.Drawing.Point(33, 28)
        Me.optPlot_absolute.Margin = New System.Windows.Forms.Padding(2)
        Me.optPlot_absolute.Name = "optPlot_absolute"
        Me.optPlot_absolute.Size = New System.Drawing.Size(101, 17)
        Me.optPlot_absolute.TabIndex = 1
        Me.optPlot_absolute.TabStop = True
        Me.optPlot_absolute.Text = "Absolute value"
        Me.optPlot_absolute.UseVisualStyleBackColor = True
        '
        'sstrip
        '
        Me.sstrip.AutoSize = False
        Me.sstrip.BackColor = System.Drawing.SystemColors.Control
        Me.sstrip.Font = New System.Drawing.Font("Segoe UI", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sstrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ts_lblVersion, Me.ts_lblBuild, Me.ts_lblLoginID, Me.ts_lblServer, Me.ts_lblVersionMsg, Me.ts_lblAccess, Me.ts_lblHealthService})
        Me.sstrip.Location = New System.Drawing.Point(0, 652)
        Me.sstrip.Name = "sstrip"
        Me.sstrip.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.sstrip.Size = New System.Drawing.Size(1223, 23)
        Me.sstrip.TabIndex = 9
        Me.sstrip.Text = "StatusStrip1"
        '
        'ts_lblVersion
        '
        Me.ts_lblVersion.AutoSize = False
        Me.ts_lblVersion.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblVersion.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblVersion.Name = "ts_lblVersion"
        Me.ts_lblVersion.Size = New System.Drawing.Size(180, 18)
        Me.ts_lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ts_lblBuild
        '
        Me.ts_lblBuild.AutoSize = False
        Me.ts_lblBuild.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblBuild.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblBuild.Name = "ts_lblBuild"
        Me.ts_lblBuild.Size = New System.Drawing.Size(180, 18)
        Me.ts_lblBuild.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ts_lblLoginID
        '
        Me.ts_lblLoginID.AutoSize = False
        Me.ts_lblLoginID.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblLoginID.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblLoginID.Name = "ts_lblLoginID"
        Me.ts_lblLoginID.Size = New System.Drawing.Size(180, 18)
        Me.ts_lblLoginID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ts_lblServer
        '
        Me.ts_lblServer.AutoSize = False
        Me.ts_lblServer.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblServer.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblServer.Name = "ts_lblServer"
        Me.ts_lblServer.Size = New System.Drawing.Size(180, 18)
        Me.ts_lblServer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ts_lblVersionMsg
        '
        Me.ts_lblVersionMsg.AutoSize = False
        Me.ts_lblVersionMsg.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblVersionMsg.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblVersionMsg.ForeColor = System.Drawing.Color.Red
        Me.ts_lblVersionMsg.Name = "ts_lblVersionMsg"
        Me.ts_lblVersionMsg.Size = New System.Drawing.Size(0, 18)
        Me.ts_lblVersionMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ts_lblVersionMsg.Visible = False
        '
        'ts_lblAccess
        '
        Me.ts_lblAccess.AutoSize = False
        Me.ts_lblAccess.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblAccess.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblAccess.Name = "ts_lblAccess"
        Me.ts_lblAccess.Size = New System.Drawing.Size(180, 18)
        Me.ts_lblAccess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ts_lblHealthService
        '
        Me.ts_lblHealthService.AutoSize = False
        Me.ts_lblHealthService.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ts_lblHealthService.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ts_lblHealthService.Name = "ts_lblHealthService"
        Me.ts_lblHealthService.Size = New System.Drawing.Size(180, 18)
        Me.ts_lblHealthService.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PdfPrinter1
        '
        Me.PdfPrinter1.AutoCenter = True
        Me.PdfPrinter1.AutoRotate = True
        Me.PdfPrinter1.CurrentPageNumber = 1
        Me.PdfPrinter1.Document = Nothing
        Me.PdfPrinter1.OffsetHardMargins = False
        Me.PdfPrinter1.PageScaleType = Gnostice.PDFOne.PDFPrinter.PrintScaleType.None
        Me.PdfPrinter1.PageSubRange = Gnostice.PDFOne.PDFPrinter.PrintSubRange.All
        Me.PdfPrinter1.Password = ""
        RenderingSettings3.AnnotsRenderingRule = CType((((Gnostice.PDFOne.ItemsRenderingRule.ifPrintableTrue Or Gnostice.PDFOne.ItemsRenderingRule.ifHiddenFalse) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifNoViewTrue) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifNoViewFalse), Gnostice.PDFOne.ItemsRenderingRule)
        RenderingSettings3.FormfieldsRenderingRule = CType((((Gnostice.PDFOne.ItemsRenderingRule.ifPrintableTrue Or Gnostice.PDFOne.ItemsRenderingRule.ifHiddenFalse) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifNoViewTrue) _
            Or Gnostice.PDFOne.ItemsRenderingRule.ifNoViewFalse), Gnostice.PDFOne.ItemsRenderingRule)
        ImageSettings3.ColorMode = Gnostice.PDFOne.ColorMode.Color
        RenderingSettings3.Image = ImageSettings3
        RenderingSettings3.ItemsToRender = CType(((((Gnostice.PDFOne.ItemsToRender.Annots Or Gnostice.PDFOne.ItemsToRender.FormFields) _
            Or Gnostice.PDFOne.ItemsToRender.Images) _
            Or Gnostice.PDFOne.ItemsToRender.Text) _
            Or Gnostice.PDFOne.ItemsToRender.Shapes), Gnostice.PDFOne.ItemsToRender)
        PrinterPreferences1.RenderingSettings = RenderingSettings3
        Me.PdfPrinter1.Preferences = PrinterPreferences1
        Me.PdfPrinter1.PrintOptions = CType(resources.GetObject("PdfPrinter1.PrintOptions"), System.Drawing.Printing.PrinterSettings)
        PdfRenderOptions3.SmoothImages = True
        PdfRenderOptions3.SmoothLineart = True
        PdfRenderOptions3.SmoothText = True
        Me.PdfPrinter1.RenderingOptions = PdfRenderOptions3
        Me.PdfPrinter1.ReversePageOrder = False
        Me.PdfPrinter1.SelectedPages = ""
        Me.PdfPrinter1.ShowPrintStatus = False
        '
        'Form_MainNew
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1223, 675)
        Me.Controls.Add(Me.sstrip)
        Me.Controls.Add(Me.splitter)
        Me.Controls.Add(Me.ToolStrip2)
        Me.Controls.Add(Me.tsMainMenu)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form_MainNew"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tsMainMenu.ResumeLayout(False)
        Me.tsMainMenu.PerformLayout()
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.splitter.Panel1.ResumeLayout(False)
        Me.splitter.Panel2.ResumeLayout(False)
        CType(Me.splitter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitter.ResumeLayout(False)
        Me.splitterUnits.Panel1.ResumeLayout(False)
        Me.splitterUnits.Panel1.PerformLayout()
        Me.splitterUnits.Panel2.ResumeLayout(False)
        CType(Me.splitterUnits, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitterUnits.ResumeLayout(False)
        Me.tsUnits.ResumeLayout(False)
        Me.tsUnits.PerformLayout()
        Me.splitterPts.Panel1.ResumeLayout(False)
        Me.splitterPts.Panel1.PerformLayout()
        Me.splitterPts.Panel2.ResumeLayout(False)
        Me.splitterPts.Panel2.PerformLayout()
        CType(Me.splitterPts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitterPts.ResumeLayout(False)
        Me.tsTests.ResumeLayout(False)
        Me.tsTests.PerformLayout()
        Me.panelDates.ResumeLayout(False)
        Me.panelDates.PerformLayout()
        Me.ContextMenuStrip_date.ResumeLayout(False)
        Me.tsPtList.ResumeLayout(False)
        Me.tsPtList.PerformLayout()
        Me.tabReports.ResumeLayout(False)
        Me.tabpageReport.ResumeLayout(False)
        Me.tabpageTrend.ResumeLayout(False)
        CType(Me.splitTrend, System.ComponentModel.ISupportInitialize).EndInit()
        Me.splitTrend.ResumeLayout(False)
        Me.tabpageTrendOptions.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.panelScale.ResumeLayout(False)
        Me.panelScale.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.panelChangeRef.ResumeLayout(False)
        Me.panelChangeRef.PerformLayout()
        Me.sstrip.ResumeLayout(False)
        Me.sstrip.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tsMainMenu As System.Windows.Forms.ToolStrip
    Friend WithEvents tsButton_patient As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents NewPatientToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FindPatientToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsButton_rfts As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsMenuItem_NewTest As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_RoutineRft As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_BronchialChallengeTests As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_AltitudeSimulation As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_spt As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_cpx As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_EditTest As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsButton_View As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents ListOfTestsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListOfPatientsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TrendToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsButton_Reporting As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsButton_tools As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsMenuItem_Preferences As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NormalValuesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_PredsManager As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_dbreports As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_SptPanelBuilder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_peoplemanager As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_ChallengeProtocolBuilder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_about As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsButton_exit As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents toolbtnEditPt As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolbtnFindPt As System.Windows.Forms.ToolStripButton
    Friend WithEvents splitter As System.Windows.Forms.SplitContainer
    Friend WithEvents lvRecall As System.Windows.Forms.ListView
    Friend WithEvents sstrip As System.Windows.Forms.StatusStrip
    Friend WithEvents ts_lblVersion As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ts_lblBuild As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ts_lblLoginID As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ts_lblServer As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ts_lblVersionMsg As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents PdfPrinter1 As Gnostice.PDFOne.PDFPrinter.PDFPrinter
    Friend WithEvents tabReports As System.Windows.Forms.TabControl
    Friend WithEvents tabpageReport As System.Windows.Forms.TabPage
    Friend WithEvents PdfViewer1 As Gnostice.PDFOne.Windows.PDFViewer.PDFViewer
    Friend WithEvents tabpageTrend As System.Windows.Forms.TabPage
    Friend WithEvents splitTrend As System.Windows.Forms.SplitContainer
    Friend WithEvents splitterUnits As System.Windows.Forms.SplitContainer
    Friend WithEvents splitterPts As System.Windows.Forms.SplitContainer
    Friend WithEvents tsUnits As System.Windows.Forms.ToolStrip
    Friend WithEvents tsLblUnits As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tsPtList As System.Windows.Forms.ToolStrip
    Friend WithEvents tsLblPatients As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tsTests As System.Windows.Forms.ToolStrip
    Friend WithEvents tsLblTests As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tsBtnExpandUnits As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbtnExpandPts As System.Windows.Forms.ToolStripButton
    Friend WithEvents tstxt_PatientName As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents ToolStripLabel_UR As System.Windows.Forms.ToolStripLabel
    Friend WithEvents dtp As System.Windows.Forms.DateTimePicker
    Friend WithEvents lvUnits As System.Windows.Forms.ListView
    Friend WithEvents ContextMenuStrip_date As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem9 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem10 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents panelDates As System.Windows.Forms.Panel
    Friend WithEvents txtEnd As System.Windows.Forms.TextBox
    Friend WithEvents txtStart As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lvPatients As System.Windows.Forms.ListView
    Friend WithEvents tstxt_UR As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents toolbtnReporting As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolbtnToolbox As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents toolbtnBookings As System.Windows.Forms.ToolStripButton
    Friend WithEvents PreferencesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListsToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PredictedValuesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PredictedValuesEditorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MigrateJjpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_WalkTests As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuItem_ActivityReport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tabpageTrendOptions As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents panelChangeRef As System.Windows.Forms.Panel
    Friend WithEvents optPlot_selected As System.Windows.Forms.RadioButton
    Friend WithEvents optPlot_earliest As System.Windows.Forms.RadioButton
    Friend WithEvents optPlot_ppn As System.Windows.Forms.RadioButton
    Friend WithEvents optPlot_pcChange As System.Windows.Forms.RadioButton
    Friend WithEvents optPlot_absolute As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowValues As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutoscale As System.Windows.Forms.CheckBox
    Friend WithEvents panelScale As System.Windows.Forms.Panel
    Friend WithEvents txtyaxis_max As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtyaxis_min As System.Windows.Forms.MaskedTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmboRefTestDate As System.Windows.Forms.ComboBox
    Friend WithEvents ts_lblAccess As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ts_lblHealthService As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents toolbtnNewPt As System.Windows.Forms.ToolStripButton
End Class
