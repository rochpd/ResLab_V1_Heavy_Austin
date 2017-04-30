Attribute VB_Name = "mPDF"
'PDF Creator Pilot info. Version 4 purchased from www.colorpilot.com 10/6/08
Global gCreatorPilot As CreatorPilot
Global gFlagPrintPdf As Boolean  'allows image to be reconstructed and date stamp added to bottom of report without re-saving to DB
Private gPdfFonts(1 To 10) As pdffont
Private gPageResolution As Integer
Private M As Margins
Global Const NataDisclaimer As String = "Note: NATA accredited for blood gas analysis only."
Global Const pathAustinLogo As String = "H:\Databases\Lab results\austin logo.jpg"
Global Const pathTsanzLogo As String = "H:\Databases\Lab results\tsanz logo.jpg"
Global Const pathNataLogo As String = "H:\Databases\Lab results\nata logo.jpg"
Global Const pathRcpaLogo As String = "H:\Databases\Lab results\rcpa logo.jpg"

'* added 02/12/2009 by GN for path to Austin Health Forms barcode image for Lab reports
Global Const pathRespLabForm As String = "H:\Databases\Lab results\FAH050800.jpg"
Global Const pathSleepLabForm As String = "H:\Databases\Lab results\FAH053900.jpg"
Global Const pathBarcodeURTemp As String = "H:\Databases\Lab results\BarCodeURTemp.jpg"

'SCORING RULES HISTORY
    'Studies recorded on or after 20/5/2002 use -
    '   1. Nasal pressure in lieu of thermistors for airflow signal
    '   2. Respiratory events scored according to AASM criteria (Chicago)

    'Scoring rules changed on 1/6/2011  -
    '   Resp events: Chicago to AASM 2007/ASA 2011
    '   Sleep staging: R&K to AASM 2007/ASA 2011
    '   Arousals:       ASDA 1992 to AASM 2007/ASA 2011
Enum ScoringRulesItem
    Description
    StartDate
    EndDate
    label
End Enum
Type ScoringRulesNote
    Description As String
    StartDate As Date
    EndDate As Date
    label As String
End Type
Type ScoringRules
    Before_20_5_2002 As ScoringRulesNote
    From_20_5_2002 As ScoringRulesNote
    From_01_06_2011 As ScoringRulesNote
End Type
Global ScoringRulesNotes As ScoringRules
Global ScoringRule() As ScoringRulesNote

Global gServerPdfLocation As String


Type CreatorPilot
    RegistrationName  As String
    SerialNo As String
    LicenseType As String
End Type
Type pdffont
    Description As String
    ID As Integer
End Type
Type Margins        'Page margins adjusted for non-printable page area.
    Top As Single
    Bottom As Single
    Left As Single
    RIGHT As Single
    HeaderHeight As Single
End Type
Type FlowVolImage
    Top As Single
    Left As Single
    Width As Single
    Height As Single
End Type
Type SlpSection
    SleepStats As Boolean
    SaO2 As Boolean
    Desats As Boolean
    Arousals As Boolean
    Resp As Boolean
    Graphics As Boolean
    Abg As Boolean
End Type
Type RequestFormHeaderData
    Surname As String
    FName As String
    UR As String
    DOB As Date
    Gender As String
    Address As String
    Suburb As String
    PCode As String
    Ph_Home As String
    Ph_Work As String
    Ph_Mobile As String
End Type
Enum ReportType
    rtNone = 0
    rtRFT = 1
    rtBronchChall = 2
    rtTreadmill = 3
    rtHAST = 4
    rtSkin = 5
    rtSleep = 6
    rtCpx = 7
    rtCpapClinic = 8
    rtSixMWD = 9
End Enum
Enum LabType
    ltSleep = 1
    ltResp = 2
    ltCpapClinic = 3
End Enum
Enum RequestType
    reqPsg = 1
    reqRft = 2
End Enum

Function DrawPDFReportCpxData(pdf As PDFCreatorPilotLib.PDFDocument4, R As Cpx, D As PtDemographics, yStart As Single)

Dim pVEmax As Single
Dim pVO2max As Single
Dim pVTmax As Single
Dim pHRmax As Single
Dim pWmax As Single
Dim pO2pulse As Single
Dim VTmax As Single
Dim VEmax As Single
Dim VO2max As Single
Dim HRmax As Single
Dim VCO2max As Single
Dim VEVO2max As Single
Dim O2pulsemax As Single
Dim ExerciseData(1 To 8, 1 To 60) As Single
Dim ExLoad(1 To 60) As Variant
Dim EndLoad As Integer
Dim LoadCount As Integer
Dim i As Single, j As Integer, Plop As Integer
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single, c8 As Single 'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim X As Single, Y As Single
Dim xMin As Single, xMax As Single, yMin As Single, yMax As Single, xSize As Single, ySize As Single, lHt As Single
Dim Max As Single, ybig As Single, yoffset As Single
Dim xlabel As String, ylabel As String
Dim cat1 As Single, cat2 As Single, dog1 As Single, dog2 As Single
Dim c() As Single
Dim Heading_Line1()
Dim Heading_Line2()
Dim Msg As String

On Error GoTo DrawPDFReportCpxData_Error

'Setup stuff
lHt = 3.6
SectionTop = IDP(yStart) + 1.5
Y = SectionTop
c1 = 2: c2 = c1 + 10: c3 = c2 + 50: c4 = c3 + 22: c5 = c4 + 40: c6 = c5 + 15: c7 = c6 + 23: c8 = c7 + 20
Fs1 = 9   'Test header
Fs2 = 8    'Results values, indice labels, units, normals
Fs3 = 6    'Equipment string
Fs4 = 10    'Report

pdf_P pdf, "CARDIO-PULMONARY EXERCISE TEST (Cycle)", c1, Y, Fs4, 3

'Draw max table
Y = Y + lHt * 2.5
pdf_P pdf, "RESULTS:", c2, Y - lHt * 0.5, Fs1, 3
pdf.MoveTo DPI(c2) + M.Left, DPI(Y + lHt * 2): pdf.DrawLineTo DPI(c4 + 25) + M.Left, DPI(Y + lHt * 2)
pdf.Stroke
pdf_P pdf, "VEmax", c2, (Y + lHt), Fs1, 3
pdf_P pdf, "HRmax", c2, (Y + lHt * 2), Fs1, 3
pdf_P pdf, "VO2max", c2, (Y + lHt * 3), Fs1, 3
pdf_P pdf, "VO2max", c2, (Y + lHt * 4), Fs1, 3
pdf_P pdf, "Wmax", c2, (Y + lHt * 5), Fs1, 3
pdf_P pdf, "Blood pressure -", c2, (Y + lHt * 6), Fs1, 3
pdf_P pdf, "Rest", c2 + 27, (Y + lHt * 6), Fs1, 2
pdf_P pdf, "End-exercise", c2 + 27, (Y + lHt * 7), Fs1, 2
pdf_P pdf, "Symptoms - ", c2, (Y + lHt * 8), Fs1, 3
pdf_P pdf, "Dyspnoea", c2 + 27, (Y + lHt * 8), Fs1, 2
pdf_P pdf, "Legs", c2 + 27, (Y + lHt * 9), Fs1, 2
If R.Borg_OtherSymptom <> "" Then pdf_P pdf, R.Borg_OtherSymptom, c2 + 27, (Y + lHt * 10), Fs1, 2

pdf_P pdf, "L/min", c2 + 15, (Y + lHt), Fs1, 2
pdf_P pdf, "bpm", c2 + 15, (Y + lHt * 2), Fs1, 2
pdf_P pdf, "L/min", c2 + 15, (Y + lHt * 3), Fs1, 2
pdf_P pdf, "ml/min/kg", c2 + 15, (Y + lHt * 4), Fs1, 2
pdf_P pdf, "Watts", c2 + 15, (Y + lHt * 5), Fs1, 2

pdf_P pdf, "Normal", c3 - 2, Y - lHt * 0.5, Fs1, 3
pdf_P pdf, Format(R.predMaxVE, "###"), c3, (Y + lHt), Fs1, 2
pdf_P pdf, Format(R.predMaxHR, "###"), c3, (Y + lHt * 2), Fs1, 2
pdf_P pdf, Format(R.predMaxVO2, "#.##"), c3, (Y + lHt * 3), Fs1, 2
pdf_P pdf, Format(R.predMaxW, "###"), c3, (Y + lHt * 5), Fs1, 2

pdf_P pdf, "Measured (%pred)", c4 - 5, Y - lHt * 0.5, Fs1, 3
If Val(R.predMaxVE) > 0 Then
    pdf_P pdf, prepare(R.MaxVE, R.predMaxVE), c4, (Y + lHt), Fs1, 2
Else
    pdf_P pdf, R.MaxVE, c4, (Y + lHt), Fs1, 2
End If
pdf_P pdf, prepare(R.MaxHR, R.predMaxHR), c4, (Y + lHt * 2), Fs1, 2
pdf_P pdf, prepare(R.MaxVO2, R.predMaxVO2), c4, (Y + lHt * 3), Fs1, 2
pdf_P pdf, Format(R.MaxVO2ml_kg, "##.0"), c4, (Y + lHt * 4), Fs1, 2
pdf_P pdf, prepare(Val(R.Load(UBound(R.Load))), R.predMaxW), c4, (Y + lHt * 5), Fs1, 2

'Print symptom table
pdf_P pdf, R.BP_Rest, c4, (Y + lHt * 6), Fs1, 2
pdf_P pdf, R.BP_Post, c4, (Y + lHt * 7), Fs1, 2
pdf_P pdf, R.Borg_Dyspnoea, c4, (Y + lHt * 8), Fs1, 2
pdf_P pdf, R.Borg_Legs, c4, (Y + lHt * 9), Fs1, 2
If R.Borg_Other <> "" Then pdf_P pdf, R.Borg_Other, c4, (Y + lHt * 10), Fs1, 2

'print post ex abgs if any data there
If Not (Val(R.ABG_pH & "") = 0 And Val(R.ABG_PaCO2 & "") = 0 And Val(R.ABG_PaO2 & "") = 0 And Val(R.ABG_SaO2 & "") = 0 And Val(R.ABG_HCO3 & "") = 0 And Val(R.ABG_BE & "") = 0) Then
    pdf_P pdf, "ABG (POST EX):", c5, Y - lHt * 0.5, Fs1, 3
    pdf_P pdf, "pH", c5, Y + lHt * 1, Fs1, 3
    pdf_P pdf, "PaCO2", c5, Y + lHt * 2, Fs1, 3
    pdf_P pdf, "PaO2", c5, Y + lHt * 3, Fs1, 3
    pdf_P pdf, "SaO2", c5, Y + lHt * 4, Fs1, 3
    pdf_P pdf, "HCO3", c5, Y + lHt * 5, Fs1, 3
    pdf_P pdf, "BE", c5, Y + lHt * 6, Fs1, 3
    pdf_P pdf, "mmHg", c6, Y + lHt * 2, Fs1, 2
    pdf_P pdf, "mmHg", c6, Y + lHt * 3, Fs1, 2
    pdf_P pdf, "%", c6, Y + lHt * 4, Fs1, 2
    pdf_P pdf, "mmol/L", c6, Y + lHt * 5, Fs1, 2
    pdf_P pdf, "mmol/L", c6, Y + lHt * 6, Fs1, 2
    pdf_P pdf, R.ABG_pH, c7, Y + lHt * 1, Fs1, 2
    pdf_P pdf, R.ABG_PaCO2, c7, Y + lHt * 2, Fs1, 2
    pdf_P pdf, R.ABG_PaO2, c7, Y + lHt * 3, Fs1, 2
    pdf_P pdf, R.ABG_SaO2, c7, Y + lHt * 4, Fs1, 2
    pdf_P pdf, R.ABG_HCO3, c7, Y + lHt * 5, Fs1, 2
    pdf_P pdf, R.ABG_BE, c7, Y + lHt * 6, Fs1, 2
    pdf.MoveTo DPI(c5) + M.Left, DPI(Y + lHt * 2): pdf.DrawLineTo DPI(c5 + 45) + M.Left, DPI(Y + lHt * 2)
    pdf.Stroke

End If


'Draw VE v VO2 graph *********************************************************
xMin = 30
xMax = xMin + 50
yMin = Y + lHt * 15
yMax = yMin + 50

If Val(R.MaxVE) < Val(R.predMaxVE) Then Max = Val(R.predMaxVE) Else Max = Val(R.MaxVE)

If Max < 60 Then
    ybig = 60
ElseIf Max < 100 Then
    ybig = 100
ElseIf Max < 160 Then
    ybig = 160
Else
    ybig = 200
End If
yoffset = 0

If Val(R.MaxVO2) < Val(R.predMaxVO2) Then Max = Val(R.predMaxVO2) Else Max = Val(R.MaxVO2)
If Max < 4 Then xbig = 4 Else xbig = 6
xlabel = "VO2 (L/min)"
ylabel = "VE (L/min)"

GoSub PDF_DrawGraph
 
pdf.SetLineWidth 0.5
pdf.SetLineDash "[5] 0"
If Val(R.predMaxVE) > 0 Then
    pdf.MoveTo M.Left + DPI(xMin), DPI(yMax - ySize * Val(R.predMaxVE) / ybig)
    pdf.DrawLineTo M.Left + DPI(xMax - 1), DPI(yMax - ySize * Val(R.predMaxVE) / ybig)
    pdf.Stroke
End If
If Val(R.predMaxVO2) > 0 Then
    pdf.MoveTo M.Left + DPI(xMin + xSize * Val(R.predMaxVO2) / xbig), DPI(yMin + ySize)
    pdf.DrawLineTo M.Left + DPI(xMin + xSize * Val(R.predMaxVO2) / xbig), DPI(yMin)
    pdf.Stroke
End If
pdf.SetLineDash "[] 0"

'Draw normal lines
If R.D_Gender = "Male" Then
    x1 = 0.41: y1 = 16.8: x2 = 1.11: y2 = 28.5
    GoSub PDF_DrawLine
    x1 = 1.66: y1 = 40.2
    GoSub PDF_DrawLine
    x2 = 1.93: y2 = 49
    GoSub PDF_DrawLine
    If ybig > 100 Then
        x1 = 2.36: y1 = 70.4
        GoSub PDF_DrawLine
        x2 = 3.07: y2 = 113.4
        GoSub PDF_DrawLine
    ElseIf ybig > 60 Then
        x1 = 2.36: y1 = 70.4
        GoSub PDF_DrawLine
        x2 = 2.85: y2 = 100
        GoSub PDF_DrawLine
    Else
        x1 = 2.15: y1 = 60
        GoSub PDF_DrawLine
    End If
    x1 = 0.41: y1 = 10.4
    x2 = 1.25: y2 = 22.4
    GoSub PDF_DrawLine
    x1 = 1.64: y1 = 30.7
    GoSub PDF_DrawLine
    x2 = 1.86: y2 = 35.6
    GoSub PDF_DrawLine
    x1 = 2.45: y1 = 52.5
    GoSub PDF_DrawLine
    If ybig > 60 Then
        x2 = 3.07: y2 = 75.4
        GoSub PDF_DrawLine
    Else
        x2 = 2.65: y2 = 60
        GoSub PDF_DrawLine
    End If
Else                    'VE-VO2 Female ranges
    x1 = 0.41: y1 = 14
    x2 = 1.28: y2 = 38.5
    GoSub PDF_DrawLine
    x1 = 1.51: y1 = 47.3
    GoSub PDF_DrawLine
    If ybig > 60 Then
        x2 = 1.78: y2 = 61
        GoSub PDF_DrawLine
        x1 = 2.01: y1 = 77.7
        GoSub PDF_DrawLine
        x2 = 2.15: y2 = 91.2
        GoSub PDF_DrawLine
    Else
        x2 = 1.76: y2 = 60
        GoSub PDF_DrawLine
    End If
    x1 = 0.41: y1 = 9
    x2 = 1.4: y2 = 32.1
    GoSub PDF_DrawLine
    x1 = 1.89: y1 = 45.8
    GoSub PDF_DrawLine
    If ybig > 60 Then
        x2 = 2.36: y2 = 62.5
        GoSub PDF_DrawLine
        x1 = 2.62: y1 = 74.5
        GoSub PDF_DrawLine
        x2 = 2.8: y2 = 88.9
        GoSub PDF_DrawLine
    Else
        x2 = 2.29: y2 = 60
        GoSub PDF_DrawLine
  End If
End If

'plot the data
For i = 1 To UBound(R.Load)
    X = Val(R.VO2(i))
    Y = Val(R.VE(i))
    pdf.DrawCircle M.Left + DPI((xMin + xSize * X / xbig)), DPI((yMax - ySize * Y / ybig)), DPI(0.7)
Next i
pdf.Fill
pdf.Stroke
    
    
    
'Draw HR/SaO2 v VO2 graph ****************
xMin = 110
xMax = xMin + 50

ybig = 200
yoffset = 40
ylabel = "HR (bpm)"
ybig2 = 100
y2offset = 75
GoSub PDF_DrawGraph
'draw symbol next to HR label
pdf.DrawCircle M.Left + DPI(xMin + 5), DPI(yMin - 3.5), DPI(0.7)
pdf.Fill
pdf.Stroke
 
'draw predicted max lines
pdf.SetLineDash "[5] 0"
pdf.MoveTo M.Left + DPI(xMin), DPI(yMax - ySize * (Val(R.predMaxHR) - yoffset) / (ybig - yoffset))
pdf.DrawLineTo M.Left + DPI((0.95 * xMax)), DPI(yMax - ySize * (Val(R.predMaxHR) - yoffset) / (ybig - yoffset))
If Val(R.predMaxVO2) > 0 Then
    pdf.MoveTo M.Left + DPI(xMin + xSize * Val(R.predMaxVO2) / xbig), DPI(yMin + ySize)
    pdf.DrawLineTo M.Left + DPI(xMin + xSize * Val(R.predMaxVO2) / xbig), DPI(yMin)
    pdf.Stroke
End If

pdf.Stroke
pdf.SetLineDash "[] 0"
pdf.Stroke

'add on SaO2 labelling on RHS of graph
pdf_P pdf, "SaO2 (%)", (xMax - 5), (yMin - 5) - IDP(M.Top), Fs2, 3
For i = 80 To 100 Step 5
    Y = yMax - ySize * (i - 75) / 25
    pdf.MoveTo M.Left + DPI(xMax), DPI(Y): pdf.DrawLineTo M.Left + DPI(xMax - 1), DPI(Y)
    pdf.Stroke
    pdf_P pdf, Format(i, "###"), (xMax + 2), (Y - 1) - IDP(M.Top), Fs2, 3
    pdf.Stroke
Next i
    
  
  
'draw normal tram tracks

If R.D_Gender = "Male" Then
    If CDate(R.TestDate) < #7/12/1999# Then    'changed to fairbarn's predicted range on this date
        x1 = 0.41: y1 = 97.8: x2 = 2.89: y2 = 199.6
        GoSub PDF_DrawLine
        x1 = 0.41: y1 = 79.2: x2 = 2.89: y2 = 177.4
        GoSub PDF_DrawLine
    Else
        x1 = 0.4: y1 = 94.8: x2 = 3.31: y2 = 198   'data from Fairbarn
        GoSub PDF_DrawLine
        x1 = 0.4: y1 = 64.8: x2 = 3.99: y2 = 192.1
        GoSub PDF_DrawLine
    End If
Else
    If CDate(R.TestDate) < #7/12/1999# Then    'changed to fairbarn's predicted range on this date
        x1 = 0.26: y1 = 104.6: x2 = 1.73: y2 = 198.6
        GoSub PDF_DrawLine
        x1 = 0.41: y1 = 92.1: x2 = 2.53: y2 = 198.6
        GoSub PDF_DrawLine
    Else
        x1 = 0.4: y1 = 104: x2 = 2.27: y2 = 197.9     'data from Fairbarn
        GoSub PDF_DrawLine
        x1 = 0.4: y1 = 76: x2 = 2.83: y2 = 198
        GoSub PDF_DrawLine
    End If
End If



  'plot the data
    For i = 1 To UBound(R.Load)
        X = Val(R.VO2(i))
        Y = Val(R.HR(i)) - yoffset
        'plot the HR (should be solid circle)
        pdf.DrawCircle M.Left + DPI(xMin + xSize * X / xbig), DPI(yMax - ySize * Y / (ybig - yoffset)), DPI(0.7)
        pdf.Fill
        pdf.Stroke
        Y = Val(R.SpO2(i)) - y2offset
        'plot the SaO2 (should be open circle)
        pdf.DrawCircle M.Left + DPI(xMin + xSize * X / xbig), DPI(yMax - ySize * Y / (ybig2 - y2offset)), DPI(0.7)
        pdf.Stroke
    Next i
    'draw symbol next to SaO2 label
    pdf.DrawCircle M.Left + DPI(xMax + 10), DPI(yMin - 3.5), DPI(0.7)
    pdf.Stroke
    
Call DrawPDFReportRftReport1(pdf, CurrentPt.Get_ReportSectionInfoFromCpxUDT(R), yMax + 5)



'PRINT PAGE 2 *******************************************************************
pdf.NewPage
Heading_Line1 = Array("Time", "Load", "VE", "VT", "VO2", "VCO2", "R", "VE/", "VE/", "SpO2", "HR", "O2 Pulse", "PetO2", "PetCO2")
Heading_Line2 = Array("(min)", "(Watts)", "(L/min)", "(L)", "(L/min)", "(L/min)", "", "VO2", "VCO2", "(%)", "(bpm)", "(ml/beat)", "mmHg", "mmHg")
ReDim c(UBound(Heading_Line1) + 1)
c(0) = 5: c(1) = 15: c(2) = 30: c(3) = 40: c(4) = 50: c(5) = 60: c(6) = 70: c(7) = 80: c(8) = 90: c(9) = 100: c(10) = 110: c(11) = 120: c(12) = 135: c(13) = 150:

Call pdf_DrawReportHeader1(pdf, R.Ref_HealthServiceID, R.TestDate, D, LabType.ltResp, 2)
Y = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromCpxUDT(R), LabType.ltResp)
Y = IDP(Y + lHt)
pdf_P pdf, "CARDIO-PULMONARY EXERCISE TEST: Continued ...", 2, Y, Fs4, 3
Y = Y + lHt * 3
pdf_P pdf, "RESULTS: Gas Exchange Data Table", c(0), Y, Fs1, 3
Y = Y + lHt * 2

'Print table heading
For i = 0 To UBound(Heading_Line1)
    pdf_P pdf, Heading_Line1(i), c(i), Y, Fs2, 3
    pdf_P pdf, Heading_Line2(i), c(i), Y + lHt, Fs2, 3
Next i
Y = Y + lHt * 2
For i = 1 To UBound(R.Load)
    Y = Y + lHt
    pdf_P pdf, Format(i / 2, "#0.0"), c(0), Y, Fs2, 2
    pdf_P pdf, R.Load(i), c(1), Y, Fs2, 2
    pdf_P pdf, Format(R.VE(i), "###.0"), c(2), Y, Fs2, 2
    pdf_P pdf, Format(R.Vt(i), "0.00"), c(3), Y, Fs2, 2
    pdf_P pdf, Format(R.VO2(i), "0.00"), c(4), Y, Fs2, 2
    pdf_P pdf, Format(R.VCO2(i), "0.00"), c(5), Y, Fs2, 2
    If Val(R.VO2(i)) > 0 Then
        pdf_P pdf, Format(Val(R.VCO2(i)) / Val(R.VO2(i)), "0.00"), c(6), Y, Fs2, 2
        pdf_P pdf, Format(Val(R.VE(i)) / Val(R.VO2(i)), "##.0"), c(7), Y, Fs2, 2
    Else
        pdf_P pdf, "NA", c(6), Y, Fs2, 2
        pdf_P pdf, "NA", c(7), Y, Fs2, 2
    End If
    If Val(R.VCO2(i)) > 0 Then
        pdf_P pdf, Format(Val(R.VE(i)) / Val(R.VCO2(i)), "##.0"), c(8), Y, Fs2, 2
    Else
        pdf_P pdf, "NA", c(8), Y, Fs2, 2
    End If
    pdf_P pdf, Format(R.SpO2(i), "###"), c(9), Y, Fs2, 2
    pdf_P pdf, Format(R.HR(i), "###"), c(10), Y, Fs2, 2
    If Val(R.HR(i)) > 0 Then
        pdf_P pdf, Format(1000 * Val(R.VO2(i)) / Val(R.HR(i)), "##.0"), c(11), Y, Fs2, 2
    Else
        pdf_P pdf, "NA", c(11), Y, Fs2, 2
    End If
    pdf_P pdf, Format(R.PetO2(i), "###"), c(12), Y, Fs2, 2
    pdf_P pdf, Format(R.PetCO2(i), "##.0"), c(13), Y, Fs2, 2
Next i
Y = Y + lHt
pdf_P pdf, "Post", c(1), Y, Fs2, 2
pdf_P pdf, R.SpO2_Post, c(9), Y, Fs2, 2


'PRINT PAGE 3 *******************************************************************
pdf.NewPage
Call pdf_DrawReportHeader1(pdf, R.Ref_HealthServiceID, R.TestDate, D, ltResp, 2)
Y = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromCpxUDT(R), LabType.ltResp)
Y = IDP(Y + lHt)
pdf_P pdf, "CARDIO-PULMONARY EXERCISE TEST: Continued ...", 2, Y, Fs4, 3
Y = Y + lHt * 3
pdf_P pdf, "RESULTS: Other Gas Exchange Graphs", c(0), Y, Fs1, 3
Y = Y + lHt * 2

'*************************** Draw VO2 versus Workload graph ************
xMin = 15
xMax = xMin + 40
yMin = Y + lHt * 3
yMax = yMin + 40

    If Val(R.Load(UBound(R.Load))) < 100 Then
        xbig = 100
    ElseIf Val(R.Load(UBound(R.Load))) >= 100 And Val(R.Load(UBound(R.Load))) < 200 Then
        xbig = 200
    ElseIf Val(R.Load(UBound(R.Load))) >= 200 And Val(R.Load(UBound(R.Load))) < 300 Then
        xbig = 300
    Else: xbig = 400
    End If

    If Val(R.MaxVO2) < 2 Then
        ybig = 2
    ElseIf Val(R.MaxVO2) >= 2 And Val(R.MaxVO2) < 4 Then
       ybig = 4
    Else: ybig = 6
    End If

    yoffset = 0
    xlabel = "Load (Watts)"
    ylabel = "VO2 (L/min)"
    GoSub PDF_DrawGraph

'Draw the normal range lines
        x1 = 0: x2 = xbig
        If R.TestDate <= "14/3/2004" Then
            y1 = 0.44:  y2 = xbig * 0.0125 + 0.44
            If y2 > ybig Then
                y2 = ybig
                x2 = (ybig - 0.44) / 0.0125
            End If
        Else
            y1 = 0.62: y2 = xbig * 0.0103 + 0.62
            If y2 > ybig Then
                y2 = ybig
                x2 = (ybig - 0.62) / 0.0103
            End If
        End If
        GoSub PDF_DrawLine
        
        If R.TestDate <= "14/3/2004" Then
            y1 = 0.15:  y2 = xbig * 0.0125 + 0.15
            If y2 > ybig Then
                y2 = ybig
                x2 = (ybig - 0.15) / 0.0125
            End If
        Else
            y1 = 0.38:  y2 = xbig * 0.0103 + 0.38
            If y2 > ybig Then
                y2 = ybig
                x2 = (ybig - 0.38) / 0.0103
            End If
        End If
        GoSub PDF_DrawLine
        'draw upper limit
        If R.TestDate <= "14/3/2004" Then
            y1 = 0.73:  y2 = xbig * 0.0125 + 0.73
            If y2 > ybig Then
                y2 = ybig
                x2 = (ybig - 0.73) / 0.0125
            End If
        Else
            y1 = 0.86:  y2 = xbig * 0.0103 + 0.86
            If y2 > ybig Then
                y2 = ybig
                x2 = (ybig - 0.86) / 0.0103
            End If
        End If
        GoSub PDF_DrawLine

    'plot the data
    i = UBound(R.Load)
    Do While i > 0
        If IsNumeric(Val(R.Load(i))) Then
            X = Val(R.Load(i))
            Y = Val(R.VO2(i))
            pdf.DrawCircle M.Left + DPI(xMin + xSize * X / xbig), DPI(yMax - ySize * Y / ybig), DPI(0.7)  'plot the point
            pdf.Stroke
        End If
        i = i - 1
    Loop
   

'********************** Draw VCO2 versus VO2 graph   ***********************

xMin = xMin + (xMax - xMin) + 20
xMax = xMin + 40

    If Val(R.MaxVCO2) > Val(R.MaxVO2) Then V = Val(R.MaxVCO2) Else V = Val(R.MaxVO2)
    If V < 2 Then
        ybig = 2
    ElseIf V >= 2 And V < 4 Then
       ybig = 4
    Else: ybig = 6
    End If
    xbig = ybig
    yoffset = 0
    xlabel = "VO2 (L/min)"
    ylabel = "VCO2 (L/min)"
    GoSub PDF_DrawGraph

    
'Draw the y=x line
        x1 = 0: y1 = 0: x2 = xbig: y2 = ybig
        GoSub PDF_DrawLine

    'plot the data
    For i = 1 To UBound(R.Load)
        X = Val(R.VO2(i))
        Y = Val(R.VCO2(i))
        pdf.DrawCircle M.Left + DPI(xMin + xSize * X / xbig), DPI(yMax - ySize * Y / ybig), DPI(0.7)  'plot the point
        pdf.Stroke
    Next i


'*************************** Draw O2 pulse versus VO2 graph ************
xMin = 15
xMax = xMin + 40
yMin = yMin + (yMax - yMin) + 20
yMax = yMin + 40

    If Val(R.MaxVO2) < Val(R.predMaxVO2) Then Max = Val(R.predMaxVO2) Else Max = Val(R.MaxVO2)
    If Max < 4 Then xbig = 4 Else xbig = 6

    If Val(R.MaxO2Pulse) < Val(R.predMaxO2Pulse) Then Max = Val(R.predMaxO2Pulse) Else Max = Val(R.MaxO2Pulse)
    If Max < 20 Then
        ybig = 20
    Else
        ybig = 40
    End If

    yoffset = 0
    xlabel = "VO2 (L/min)"
    ylabel = "O2 Pulse (mL/beat)"
    GoSub PDF_DrawGraph

'Draw predicted max lines
        pdf.SetLineDash "[5] 0"
        x1 = R.predMaxVO2: y1 = 0: x2 = R.predMaxVO2: y2 = ybig
        GoSub PDF_DrawLine
        x1 = 0: y1 = R.predMaxO2Pulse: x2 = xbig: y2 = R.predMaxO2Pulse
        GoSub PDF_DrawLine
        pdf.SetLineDash "[] 0"
        pdf.Stroke

    'plot the data
    For i = 1 To UBound(R.Load)
        X = Val(R.VO2(i))
        Y = 1000 * Val(R.VO2(i)) / Val(R.HR(i))
        pdf.DrawCircle M.Left + DPI(xMin + xSize * X / xbig), DPI(yMax - ySize * Y / ybig), DPI(0.7)  'plot the point
        pdf.Stroke
    Next i
   

'*************************** draw VT versus VE graph
xMin = xMin + (xMax - xMin) + 20
xMax = xMin + 40

    If Val(R.MaxVt) < Val(R.predMaxVt) Then Max = Val(R.predMaxVt) Else Max = Val(R.MaxVt)
    If Max < 2 Then xbig = 2 Else xbig = 4

    If Val(R.MaxVE) < Val(R.predMaxVE) Then Max = Val(R.predMaxVE) Else Max = Val(R.MaxVE)
    If Max < 60 Then
        ybig = 60
    ElseIf Max < 100 Then
        ybig = 100
    ElseIf Max < 160 Then
        ybig = 160
    Else: ybig = 200
    End If

    yoffset = 0
    xlabel = "        VT (L)"
    ylabel = "VE (L/min)"
    GoSub PDF_DrawGraph

'Draw predicted max lines
        pdf.SetLineDash "[5] 0"
        x1 = Val(R.predMaxVt): y1 = 0: x2 = Val(R.predMaxVt): y2 = ybig
        If x1 > 0 And x2 > 0 Then GoSub PDF_DrawLine
        x1 = 0: y1 = Val(R.predMaxVE): x2 = xbig: y2 = Val(R.predMaxVE)
        If y1 > 0 And y2 > 0 Then GoSub PDF_DrawLine
        pdf.SetLineDash "[] 0"
        pdf.Stroke
    'plot the data
    For i = 1 To UBound(R.Load)
        X = Val(R.Vt(i))
        Y = Val(R.VE(i))
        pdf.DrawCircle M.Left + DPI(xMin + xSize * X / xbig), DPI(yMax - ySize * Y / ybig), DPI(0.7)  'plot the point
        pdf.Stroke
    Next i

Exit Function



PDF_DrawGraph:
    xSize = xMax - xMin
    ySize = yMax - yMin
    pdf_P pdf, xlabel, (xMin + 0.5 * xSize), (yMax - 11), Fs2, 3
    pdf_P pdf, ylabel, (xMin - 10), (yMin - 10), Fs2, 3
    pdf.Stroke
    pdf.DrawRectangle M.Left + DPI(xMin), DPI(yMin), DPI(xSize), DPI(ySize), 0
    
  'draw and label tick marks
    For i = xMin + xSize / 4 To xMax Step xSize / 4
        pdf.MoveTo M.Left + DPI(i), DPI(yMax): pdf.DrawLineTo M.Left + DPI(i), DPI(yMax - 1)
        pdf.Stroke
        If xlabel = "Load (Watts)" Then
            pdf_P pdf, Format(xbig * (i - xMin) / xSize, "###"), (i - 2), (yMax + 2) - IDP(M.Top), Fs2, 3
        Else
            pdf_P pdf, Format(xbig * (i - xMin) / xSize, "0.0"), (i - 2), (yMax + 2) - IDP(M.Top), Fs2, 3
        End If
    Next i
    pdf.Stroke
    For i = yMax - ySize / 4 To yMin Step -ySize / 4
        pdf.MoveTo M.Left + DPI(xMin), DPI(i): pdf.DrawLineTo M.Left + DPI(xMin + 1), DPI(i)
        pdf.Stroke
        If ylabel = "VO2 (L/min)" Or ylabel = "VCO2 (L/min)" Then
            pdf_P pdf, Format(yoffset + (ybig - yoffset) * (yMax - i) / ySize, "0.0"), (xMin - 7), (i - 1) - IDP(M.Top), Fs2, 3
        Else
            pdf_P pdf, Format(yoffset + (ybig - yoffset) * (yMax - i) / ySize, "###"), (xMin - 7), (i - 1) - IDP(M.Top), Fs2, 3
        End If
        
    Next i
    pdf.Stroke
    
Return

PDF_DrawLine:          'draws normal range lines
    cat1 = xMin + xSize * Val(x1) / xbig
    dog1 = yMax - ySize * (y1 - yoffset) / (ybig - yoffset)
    cat2 = xMin + xSize * Val(x2) / xbig
    dog2 = yMax - ySize * (y2 - yoffset) / (ybig - yoffset)
    pdf.MoveTo M.Left + DPI(cat1), DPI(dog1): pdf.DrawLineTo M.Left + DPI(cat2), DPI(dog2)
    pdf.Stroke
Return

Exit Function


DrawPDFReportCpxData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportCpxData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function DrawPDFReportTreadmillData(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, yStart As Single)

Dim Msg As String

Dim AirArray(0 To 21) As Single, O2Array(0 To 22) As Single, SpeedArray(0 To 14) As Single
Dim i As Integer
Dim j As Integer
Dim xMin As Single, xMax As Single, yMin As Single, yMax As Single, xSize As Single, ySize As Single
Dim prevX As Single, prevY As Single
Dim O2X As Single, O2Y As Single
Dim O2prevX As Single, O2prevY As Single
Dim SaO2 As Single, O2SaO2 As Single
Dim Nadir_Air, Nadir_O2, Rest_SaO2_O2
Dim Recovery_Air, Recovery_O2
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim X As Single, Y As Single, y1 As Single
Dim O2NotDone As Boolean
Dim locO2SaO2String, locAirSaO2String, locSpeedString
Dim LowSaO2, maxSpeed, minSpeed, minSaO2, speedRange, AirBorg, AirHR

Const lHt As Single = 3.6

On Error GoTo DrawPDFReportTreadmillData_Error

'Setup stuff
Y = IDP(yStart) + 2.5
c1 = 10: c2 = c1 + 2: c3 = c2 + 22: c4 = c3 + 30: c5 = c4 + 25: c6 = c5 + 28: c7 = c6 + 28
Fs1 = 9   'Test header
Fs2 = 8   'Results values, indice labels, units, normals
Fs3 = 7   'Equipment string
Fs4 = 10  'Report

locO2SaO2String = R.O2SaO2String   '("O2SaO2String")
locAirSaO2String = R.AirSaO2String '("AirSaO2String")
'locSpeedString = R.SpeedString '  Rs("SpeedString")

'Check to see if O2 dtat has been collected - first char will be a comma
If Left(locO2SaO2String, 1) = "," Then O2NotDone = True


'If Left(R.SpeedString, 1) = "," Then  ' Rs("SpeedString"), 1) = "," Then
'    flg6MWD = 1
'Else
'    flg6MWD = 0
'End If

'Print table
'If flg6MWD = 1 Then
    pdf_P pdf, "6MWD (ATS Protocol)", 2, Y, Fs1, 3
'Else
'    pdf_P pdf, "TREADMILL EXERCISE TEST", 2, Y, Fs1, 3
'End If
Y = Y + lHt * 2

'Decode the data from the strings into easier to use arrays
indexipoo = 0: j = 1
For i = 1 To Len(locAirSaO2String) 'first do air sao2
    If Mid(locAirSaO2String, i, 1) = "," Then
        AirArray(indexipoo) = Val(Mid(locAirSaO2String, j, i - j))
        indexipoo = indexipoo + 1
        j = i + 1
    End If
Next i
indexipoo = 0: j = 1
For i = 1 To Len(locO2SaO2String)  'then do O2 sao2
    If Mid(locO2SaO2String, i, 1) = "," Then
        O2Array(indexipoo) = Val(Mid(locO2SaO2String, j, i - j))
        indexipoo = indexipoo + 1
        j = i + 1
    End If
Next i
'indexipoo = 0: j = 1
'For i = 1 To Len(locSpeedString)  'finally do speed
'    If Mid(locSpeedString, i, 1) = "," Then
'        SpeedArray(indexipoo) = Val(Mid(locSpeedString, j, i - j))
'        indexipoo = indexipoo + 1
'        j = i + 1
'    End If
'Next i

'Find mins and maxes
LowSaO2 = 100   ': maxSpeed = 0: minSpeed = 100
For i = 0 To 18
    If Val(AirArray(i)) <> 0 And Val(AirArray(i)) < LowSaO2 Then LowSaO2 = Val(AirArray(i))
    If Val(O2Array(i)) <> 0 And Val(O2Array(i)) < LowSaO2 Then LowSaO2 = Val(O2Array(i))
'    If i < 12 Then
'        If Val(SpeedArray(i)) > maxSpeed Then maxSpeed = Val(SpeedArray(i))
'        If Val(SpeedArray(i)) < minSpeed And Val(SpeedArray(i)) <> 0 Then minSpeed = Val(SpeedArray(i))
'    End If
Next i

If Val(LowSaO2) < 80 Then
    minSaO2 = 60
Else
    minSaO2 = 80
End If
'If minSpeed = maxSpeed Then
'    speedRange = Format(minSpeed, "##") & " kph"
'Else
'    speedRange = Format(minSpeed, "0") & "-" & Format(maxSpeed, "##") & " kph"
'End If
'If flg6MWD = 1 Then speedRange = "self paced/30m track"

'Get other data which was tacked onto the end of the strings
AirBorg = AirArray(19)
AirHR = AirArray(20)
If O2NotDone Then O2Borg = "" Else O2Borg = O2Array(19)
If O2NotDone Then O2HR = "" Else O2HR = O2Array(20)
If O2NotDone Then O2Flow = "" Else O2Flow = O2Array(21)
'gradRange = SpeedArray(12)
'AirDistance = SpeedArray(13)
'If O2NotDone Then O2distance = "" Else O2distance = SpeedArray(14)
'If gradRange = "" Then gradRange = "0%"

'Get the nadir and recovery SpO2s
Nadir_Air = 101
Nadir_O2 = 101
For i = 1 To 12   'Skip the rest value ie start at 1
    If AirArray(i) > 0 And AirArray(i) < Nadir_Air Then Nadir_Air = AirArray(i)
    If O2Array(i) > 0 And O2Array(i) < Nadir_O2 Then Nadir_O2 = O2Array(i)
Next i
Recovery_Air = 0
Recovery_O2 = 0
For i = 13 To 18   'Skip the rest value ie start at 1
    If AirArray(i) > 0 And AirArray(i) > Recovery_Air Then Recovery_Air = AirArray(i)
    If O2Array(i) > 0 And O2Array(i) > Recovery_O2 Then Recovery_O2 = O2Array(i)
Next i
If Nadir_Air = 101 Then Nadir_Air = ""
If Nadir_O2 = 101 Then Nadir_O2 = ""
If Recovery_Air = 0 Then Recovery_Air = ""
If Recovery_O2 = 0 Then Recovery_O2 = ""


'Print numerical results
c1 = 10: c2 = c1 + 35: c3 = c2 + 22: c4 = c3 + 21: c5 = c4 + 31: c6 = c5 + 21: c7 = c6 + 21
ResultOffsett = 4
pdf_P pdf, "SpO2 (%)", c5 - 3, Y, Fs2, 3
pdf_P pdf, "SpO2 (%)", c6 - 3, Y, Fs2, 3
pdf_P pdf, "SpO2 (%)", c7 - 3, Y, Fs2, 3
Y = Y + lHt
pdf_P pdf, "Distance (m)", c2, Y, Fs2, 3
pdf_P pdf, "Dyspnoea", c3, Y, Fs2, 3
pdf_P pdf, "Max HR (bpm)", c4, Y, Fs2, 3
pdf_P pdf, "Rest", c5, Y, Fs2, 3
pdf_P pdf, "Nadir", c6, Y, Fs2, 3
pdf_P pdf, "Recovery", c7 - 3, Y, Fs2, 3
Y = Y + lHt
pdf_P pdf, "AIR:", c1, Y, Fs1, 3
pdf_P pdf, Format(AirDistance, "####"), c2 + ResultOffsett, Y, Fs2, 2
pdf_P pdf, AirBorg, c3 + ResultOffsett + 2, Y, Fs2, 2
pdf_P pdf, AirHR, c4 + ResultOffsett, Y, Fs2, 2
pdf_P pdf, AirArray(0), c5 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Nadir_Air, c6 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Recovery_Air, c7 + ResultOffsett - 2, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "OXYGEN (@" & O2Flow & "L/min):", c1, Y, Fs2, 3
pdf_P pdf, Format(O2Distance, "####"), c2 + ResultOffsett, Y, Fs2, 2
pdf_P pdf, O2Borg, c3 + ResultOffsett + 2, Y, Fs2, 2
pdf_P pdf, O2HR, c4 + ResultOffsett, Y, Fs2, 2
If O2NotDone Then Rest_SaO2_O2 = "" Else Rest_SaO2_O2 = O2Array(0)
pdf_P pdf, Rest_SaO2_O2, c5 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Nadir_O2, c6 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Recovery_O2, c7 + ResultOffsett - 2, Y, Fs2, 2

xMin = 30             'Co-ords (in mm) for exercise graph
xMax = xMin + 80
yMin = 120
yMax = yMin + 60
xSize = xMax - xMin
ySize = yMax - yMin

'DRAW EXERCISE GRAPH
pdf.SetLineWidth 0.5
pdf_P pdf, "SaO2 (%)", (xMin - 17), (yMin + 10), Fs1, 3
pdf_P pdf, "Time (mins)", (xMin + 0.4 * xSize), (yMax + 4), Fs1, 3
pdf_P pdf, "EXERCISE  (Speed: " & speedRange & ", Grade: " & gradRange & "%)", xMin + 8, yMin - 2 * lHt, Fs1, 3
'Draw horizontal lines and label y-axis
For j = minSaO2 To 100 Step 2
    Y = yMax - (j - minSaO2) * ySize / (100 - minSaO2)
    
    If j / 10 = Int(j / 10) Then
        pdf_P pdf, Format(j, "###"), (xMin - 8), (Y - 2), Fs1, 3
    End If
    pdf.MoveTo M.Left + DPI(xMin), M.Top + DPI(Y): pdf.DrawLineTo M.Left + DPI(xMax), M.Top + DPI(Y)
    pdf.Stroke
Next j

'draw vertical lines and label x-axis
For j = 0 To 6
    X = xMin + j * xSize / 6
    pdf.MoveTo M.Left + DPI(X), M.Top + DPI(yMin): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(yMax)
    pdf.Stroke
    pdf_P pdf, Format(j, "0"), (xMin - 1 + j * xSize / 6), (yMax + 1), Fs1, 3
Next j

'plot the data
prevX = xMin
prevY = yMax - (Val(AirArray(0)) - minSaO2) * ySize / (100 - minSaO2)
O2prevX = xMin
O2prevY = yMax - (Val(O2Array(0)) - minSaO2) * ySize / (100 - minSaO2)
pdf.SetLineWidth 1.5
For j = 0 To 12
    SaO2 = Val(AirArray(j))
    If SaO2 <> 0 Then
        X = xMin + (j) * xSize / 12
        Y = yMax - (SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.SetLineDash ("[] 0")
        pdf.DrawCircle M.Left + DPI(X), M.Top + DPI(Y), 0.5
        pdf.MoveTo M.Left + DPI(prevX), M.Top + DPI(prevY): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(Y)
        pdf.Stroke
        prevX = X
        prevY = Y
    End If
    O2SaO2 = Val(O2Array(j))
    If O2SaO2 <> 0 Then
        O2X = xMin + (j) * xSize / 12
        O2Y = yMax - (O2SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.DrawCircle DPI(O2X), M.Top + DPI(O2Y), DPI(0.5)
        pdf.SetLineDash "[5] 0"
        pdf.MoveTo M.Left + DPI(O2prevX), M.Top + DPI(O2prevY): pdf.DrawLineTo M.Left + DPI(O2X), M.Top + DPI(O2Y)
        pdf.Stroke
        O2prevX = O2X
        O2prevY = O2Y
    End If
Next j
pdf.SetLineDash ("[] 0")

'DRAW RECOVERY GRAPH
pdf.SetLineWidth 0.5
xMin = 120            'Co-ords (in mm) for recovery graph
xMax = xMin + 40
xSize = xMax - xMin
ySize = yMax - yMin
pdf_P pdf, "RECOVERY", (xMin + 10), yMin - 2 * lHt, Fs1, 3

For j = minSaO2 To 100 Step 2                   'draw horizontal lines
    Y = yMax - (j - minSaO2) * ySize / (100 - minSaO2)
    pdf.MoveTo M.Left + DPI(xMin), M.Top + DPI(Y): pdf.DrawLineTo M.Left + DPI(xMax), M.Top + DPI(Y)
    pdf.Stroke
Next j

'draw vertical lines and label x-axis
For j = 0 To 3
    X = xMin + j * xSize / 3
    pdf.MoveTo M.Left + DPI(X), M.Top + DPI(yMin): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(yMax)
    pdf.Stroke
    pdf_P pdf, Format(j, "0"), (xMin - 1 + j * xSize / 3), (yMax + 1), Fs1, 3
Next j

'plot the data
prevX = xMin + xSize / 6
O2prevX = xMin + xSize / 6
prevY = yMax - (Val(AirArray(13)) - minSaO2) * ySize / (100 - minSaO2)
O2prevY = yMax - (Val(O2Array(13)) - minSaO2) * ySize / (100 - minSaO2)
pdf.SetLineWidth 1.5

For j = 13 To 18
    SaO2 = Val(AirArray(j))
    If SaO2 > 0 Then
        X = xMin + ((j - 13) + 1) * xSize / 6
        Y = yMax - (SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.SetLineDash ("[] 0")
        pdf.DrawCircle M.Left + DPI(X), M.Top + DPI(Y), 0.5
        pdf.MoveTo M.Left + DPI(prevX), M.Top + DPI(prevY): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(Y)
        pdf.Stroke
        prevX = X
        prevY = Y
    End If
    O2SaO2 = Val(O2Array(j))
    If O2SaO2 > 0 Then
        O2X = xMin + ((j - 13) + 1) * xSize / 6
        O2Y = yMax - (O2SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.DrawCircle M.Left + DPI(O2X), M.Top + DPI(O2Y), DPI(0.5)
        pdf.SetLineDash "[5] 0"
        pdf.MoveTo M.Left + DPI(O2prevX), M.Top + DPI(O2prevY): pdf.DrawLineTo M.Left + DPI(O2X), M.Top + DPI(O2Y)
        pdf.Stroke
        O2prevX = O2X
        O2prevY = O2Y
    End If
Next j

'Draw legend
Y = yMin + ySize / 3
X = xMax + 5
pdf_P pdf, "Air", X, Y, Fs1, 3
pdf.SetLineDash ("[] 0")
pdf.MoveTo M.Left + DPI(X + 7), M.Top + DPI(Y + 2): pdf.DrawLineTo M.Left + DPI(X + 12), M.Top + DPI(Y + 2)
pdf.Stroke
Y = Y + lHt
pdf_P pdf, "O2", X, Y, Fs1, 3
pdf.SetLineDash ("[5] 0")
pdf.MoveTo M.Left + DPI(X + 7), M.Top + DPI(Y + 2): pdf.DrawLineTo M.Left + DPI(X + 12), M.Top + DPI(Y + 2)
pdf.Stroke
pdf.SetLineDash ("[] 0")

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), yMax + 10)
'DrawPDFReportRftReport1 Pdf, Rs, yMax + 10

Exit Function

DrawPDFReportTreadmillData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport_BronchChallengeData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function
  
Public Function DrawPDFReport6MWDData(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, yStart As Single)

Dim Msg As String
Dim AirArray(0 To 21) As Single, O2Array(0 To 22) As Single
Dim i As Integer
Dim j As Integer
Dim xMin As Single, xMax As Single, yMin As Single, yMax As Single, xSize As Single, ySize As Single
Dim prevX As Single, prevY As Single
Dim O2X As Single, O2Y As Single
Dim O2prevX As Single, O2prevY As Single
Dim SaO2 As Single, O2SaO2 As Single
Dim Nadir_Air, Nadir_O2, Rest_SaO2_O2
Dim Recovery_Air, Recovery_O2
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim X As Single, Y As Single, y1 As Single
Dim O2NotDone As Boolean
Dim locO2SaO2String, locAirSaO2String
Dim LowSaO2, minSaO2, AirBorg, AirHR, AirDistance, O2Distance, O2Flow, O2HR, O2Borg

Const lHt As Single = 3.6

On Error GoTo DrawPDFReport6MWDData_Error

'Setup stuff
Y = IDP(yStart) + 2.5
c1 = 10: c2 = c1 + 2: c3 = c2 + 22: c4 = c3 + 30: c5 = c4 + 25: c6 = c5 + 28: c7 = c6 + 28
Fs1 = 9   'Test header
Fs2 = 8   'Results values, indice labels, units, normals
Fs3 = 7   'Equipment string
Fs4 = 10  'Report

locO2SaO2String = R.O2SaO2String   '("O2SaO2String")
locAirSaO2String = R.AirSaO2String '("AirSaO2String")

'Check to see if O2 dtat has been collected - first char will be a comma
If Left(locO2SaO2String, 1) = "," Then O2NotDone = True

'Print table
pdf_P pdf, "6MWD (ATS Protocol):", 2, Y, Fs1, 3
Y = Y + lHt * 2

'Decode the data from the strings into easier to use arrays
indexipoo = 0: j = 1
For i = 1 To Len(locAirSaO2String) 'first do air sao2
    If Mid(locAirSaO2String, i, 1) = "," Then
        AirArray(indexipoo) = Val(Mid(locAirSaO2String, j, i - j))
        indexipoo = indexipoo + 1
        j = i + 1
    End If
Next i
indexipoo = 0: j = 1
For i = 1 To Len(locO2SaO2String)  'then do O2 sao2
    If Mid(locO2SaO2String, i, 1) = "," Then
        O2Array(indexipoo) = Val(Mid(locO2SaO2String, j, i - j))
        indexipoo = indexipoo + 1
        j = i + 1
    End If
Next i

'Find mins and maxes
LowSaO2 = 100   ': maxSpeed = 0: minSpeed = 100
For i = 0 To 18
    If Val(AirArray(i)) <> 0 And Val(AirArray(i)) < LowSaO2 Then LowSaO2 = Val(AirArray(i))
    If Val(O2Array(i)) <> 0 And Val(O2Array(i)) < LowSaO2 Then LowSaO2 = Val(O2Array(i))
Next i

If Val(LowSaO2) < 80 Then
    minSaO2 = 60
Else
    minSaO2 = 80
End If

'Get other data which was tacked onto the end of the strings
AirBorg = AirArray(19)
AirHR = AirArray(20)
AirDistance = AirArray(21)
If O2NotDone Then O2Borg = "" Else O2Borg = O2Array(19)
If O2NotDone Then O2HR = "" Else O2HR = O2Array(20)
If O2NotDone Then O2Flow = "" Else O2Flow = O2Array(21)
If O2NotDone Then O2Distance = "" Else O2Distance = O2Array(22)

'Get the nadir and recovery SpO2s
Nadir_Air = 101
Nadir_O2 = 101
For i = 1 To 12   'Skip the rest value ie start at 1
    If AirArray(i) > 0 And AirArray(i) < Nadir_Air Then Nadir_Air = AirArray(i)
    If O2Array(i) > 0 And O2Array(i) < Nadir_O2 Then Nadir_O2 = O2Array(i)
Next i
Recovery_Air = 0
Recovery_O2 = 0
For i = 13 To 18   'Skip the rest value ie start at 1
    If AirArray(i) > 0 And AirArray(i) > Recovery_Air Then Recovery_Air = AirArray(i)
    If O2Array(i) > 0 And O2Array(i) > Recovery_O2 Then Recovery_O2 = O2Array(i)
Next i
If Nadir_Air = 101 Then Nadir_Air = ""
If Nadir_O2 = 101 Then Nadir_O2 = ""
If Recovery_Air = 0 Then Recovery_Air = ""
If Recovery_O2 = 0 Then Recovery_O2 = ""


'Print numerical results
c1 = 10: c2 = c1 + 35: c3 = c2 + 22: c4 = c3 + 21: c5 = c4 + 31: c6 = c5 + 21: c7 = c6 + 21
ResultOffsett = 4
pdf_P pdf, "SpO2 (%)", c5 - 3, Y, Fs2, 3
pdf_P pdf, "SpO2 (%)", c6 - 3, Y, Fs2, 3
pdf_P pdf, "SpO2 (%)", c7 - 3, Y, Fs2, 3
Y = Y + lHt
pdf_P pdf, "Distance (m)", c2, Y, Fs2, 3
pdf_P pdf, "Dyspnoea", c3, Y, Fs2, 3
pdf_P pdf, "Max HR (bpm)", c4, Y, Fs2, 3
pdf_P pdf, "Rest", c5, Y, Fs2, 3
pdf_P pdf, "Nadir", c6, Y, Fs2, 3
pdf_P pdf, "Recovery", c7 - 3, Y, Fs2, 3
Y = Y + lHt
pdf_P pdf, "AIR:", c1, Y, Fs1, 3
pdf_P pdf, Format(AirDistance, "####"), c2 + ResultOffsett, Y, Fs2, 2
pdf_P pdf, AirBorg, c3 + ResultOffsett + 2, Y, Fs2, 2
pdf_P pdf, AirHR, c4 + ResultOffsett, Y, Fs2, 2
pdf_P pdf, AirArray(0), c5 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Nadir_Air, c6 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Recovery_Air, c7 + ResultOffsett - 2, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "OXYGEN (@" & O2Flow & "L/min):", c1, Y, Fs2, 3
pdf_P pdf, Format(O2Distance, "####"), c2 + ResultOffsett, Y, Fs2, 2
pdf_P pdf, O2Borg, c3 + ResultOffsett + 2, Y, Fs2, 2
pdf_P pdf, O2HR, c4 + ResultOffsett, Y, Fs2, 2
If O2NotDone Then Rest_SaO2_O2 = "" Else Rest_SaO2_O2 = O2Array(0)
pdf_P pdf, Rest_SaO2_O2, c5 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Nadir_O2, c6 + ResultOffsett - 2, Y, Fs2, 2
pdf_P pdf, Recovery_O2, c7 + ResultOffsett - 2, Y, Fs2, 2

xMin = 30             'Co-ords (in mm) for exercise graph
xMax = xMin + 80
yMin = 110
yMax = yMin + 60
xSize = xMax - xMin
ySize = yMax - yMin

'DRAW EXERCISE GRAPH
pdf.SetLineWidth 0.5
pdf_P pdf, "SaO2 (%)", (xMin - 17), (yMin + 10), Fs1, 3
pdf_P pdf, "Time (mins)", (xMin + 0.4 * xSize), (yMax + 4), Fs1, 3
pdf_P pdf, "EXERCISE", xMin + 8, yMin - 2 * lHt, Fs1, 3
'Draw horizontal lines and label y-axis
For j = minSaO2 To 100 Step 2
    Y = yMax - (j - minSaO2) * ySize / (100 - minSaO2)
    
    If j / 10 = Int(j / 10) Then
        pdf_P pdf, Format(j, "###"), (xMin - 8), (Y - 2), Fs1, 3
    End If
    pdf.MoveTo M.Left + DPI(xMin), M.Top + DPI(Y): pdf.DrawLineTo M.Left + DPI(xMax), M.Top + DPI(Y)
    pdf.Stroke
Next j

'draw vertical lines and label x-axis
For j = 0 To 6
    X = xMin + j * xSize / 6
    pdf.MoveTo M.Left + DPI(X), M.Top + DPI(yMin): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(yMax)
    pdf.Stroke
    pdf_P pdf, Format(j, "0"), (xMin - 1 + j * xSize / 6), (yMax + 1), Fs1, 3
Next j

'plot the data
prevX = xMin
prevY = yMax - (Val(AirArray(0)) - minSaO2) * ySize / (100 - minSaO2)
O2prevX = xMin
O2prevY = yMax - (Val(O2Array(0)) - minSaO2) * ySize / (100 - minSaO2)
pdf.SetLineWidth 1.5
For j = 0 To 12
    SaO2 = Val(AirArray(j))
    If SaO2 <> 0 Then
        X = xMin + (j) * xSize / 12
        Y = yMax - (SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.SetLineDash ("[] 0")
        pdf.DrawCircle M.Left + DPI(X), M.Top + DPI(Y), 0.5
        pdf.MoveTo M.Left + DPI(prevX), M.Top + DPI(prevY): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(Y)
        pdf.Stroke
        prevX = X
        prevY = Y
    End If
    O2SaO2 = Val(O2Array(j))
    If O2SaO2 <> 0 Then
        O2X = xMin + (j) * xSize / 12
        O2Y = yMax - (O2SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.DrawCircle DPI(O2X), M.Top + DPI(O2Y), DPI(0.5)
        pdf.SetLineDash "[5] 0"
        pdf.MoveTo M.Left + DPI(O2prevX), M.Top + DPI(O2prevY): pdf.DrawLineTo M.Left + DPI(O2X), M.Top + DPI(O2Y)
        pdf.Stroke
        O2prevX = O2X
        O2prevY = O2Y
    End If
Next j
pdf.SetLineDash ("[] 0")

'DRAW RECOVERY GRAPH
pdf.SetLineWidth 0.5
xMin = 120            'Co-ords (in mm) for recovery graph
xMax = xMin + 40
xSize = xMax - xMin
ySize = yMax - yMin
pdf_P pdf, "RECOVERY", (xMin + 10), yMin - 2 * lHt, Fs1, 3

For j = minSaO2 To 100 Step 2                   'draw horizontal lines
    Y = yMax - (j - minSaO2) * ySize / (100 - minSaO2)
    pdf.MoveTo M.Left + DPI(xMin), M.Top + DPI(Y): pdf.DrawLineTo M.Left + DPI(xMax), M.Top + DPI(Y)
    pdf.Stroke
Next j

'draw vertical lines and label x-axis
For j = 0 To 3
    X = xMin + j * xSize / 3
    pdf.MoveTo M.Left + DPI(X), M.Top + DPI(yMin): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(yMax)
    pdf.Stroke
    pdf_P pdf, Format(j, "0"), (xMin - 1 + j * xSize / 3), (yMax + 1), Fs1, 3
Next j

'plot the data
prevX = xMin + xSize / 6
O2prevX = xMin + xSize / 6
prevY = yMax - (Val(AirArray(13)) - minSaO2) * ySize / (100 - minSaO2)
O2prevY = yMax - (Val(O2Array(13)) - minSaO2) * ySize / (100 - minSaO2)
pdf.SetLineWidth 1.5

For j = 13 To 18
    SaO2 = Val(AirArray(j))
    If SaO2 > 0 Then
        X = xMin + ((j - 13) + 1) * xSize / 6
        Y = yMax - (SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.SetLineDash ("[] 0")
        pdf.DrawCircle M.Left + DPI(X), M.Top + DPI(Y), 0.5
        pdf.MoveTo M.Left + DPI(prevX), M.Top + DPI(prevY): pdf.DrawLineTo M.Left + DPI(X), M.Top + DPI(Y)
        pdf.Stroke
        prevX = X
        prevY = Y
    End If
    O2SaO2 = Val(O2Array(j))
    If O2SaO2 > 0 Then
        O2X = xMin + ((j - 13) + 1) * xSize / 6
        O2Y = yMax - (O2SaO2 - minSaO2) * ySize / (100 - minSaO2)
        pdf.DrawCircle M.Left + DPI(O2X), M.Top + DPI(O2Y), DPI(0.5)
        pdf.SetLineDash "[5] 0"
        pdf.MoveTo M.Left + DPI(O2prevX), M.Top + DPI(O2prevY): pdf.DrawLineTo M.Left + DPI(O2X), M.Top + DPI(O2Y)
        pdf.Stroke
        O2prevX = O2X
        O2prevY = O2Y
    End If
Next j

'Draw legend
Y = yMin + ySize / 3
X = xMax + 5
pdf_P pdf, "Air", X, Y, Fs1, 3
pdf.SetLineDash ("[] 0")
pdf.MoveTo M.Left + DPI(X + 7), M.Top + DPI(Y + 2): pdf.DrawLineTo M.Left + DPI(X + 12), M.Top + DPI(Y + 2)
pdf.Stroke
Y = Y + lHt
pdf_P pdf, "O2", X, Y, Fs1, 3
pdf.SetLineDash ("[5] 0")
pdf.MoveTo M.Left + DPI(X + 7), M.Top + DPI(Y + 2): pdf.DrawLineTo M.Left + DPI(X + 12), M.Top + DPI(Y + 2)
pdf.Stroke
pdf.SetLineDash ("[] 0")

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), yMax + 15)

Exit Function

DrawPDFReport6MWDData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport6MWDData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Private Function Suppress_ReportInPdf(ReportStatus As String, Report As String) As String
'Don't want reports with 'Reported not Entered' AKA 'results for discussion' to be viewable via Cerner ie saved in the PDF image.
'If status ='Reported not Entered' then return empty string else return the report field intact

Dim Msg As String

On Error GoTo Suppress_ReportInPdf_Error

If UCase(ReportStatus) = "REPORTED NOT ENTERED" Then
    Suppress_ReportInPdf = ""
Else
    Suppress_ReportInPdf = Report
End If

Exit Function

Suppress_ReportInPdf_Error:
    Msg = ""
    Call ErrorLog(Msg, "Suppress_ReportInPdf", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function DrawPDFReport_Rft(LaunchAcrobat As Boolean, UpdateDatabase As Boolean, ByVal ResultsID As Long, Report As ReportType, Optional ByVal FName As Variant) As Boolean

Dim z As Variant
Dim Y As Single
Dim TblName As String
Dim pdf As PDFCreatorPilotLib.PDFDocument4
Dim iReturn
Dim pdfStream As New ADODB.Stream
Dim dbLockType
Dim i
Dim Msg As String

On Error GoTo DrawPDFReport_Rft_Error

If IsMissing(FName) Then FName = Environ("tmp") & "\temp.pdf"
If UpdateDatabase Then dbLockType = adLockPessimistic Else dbLockType = adLockReadOnly

'Delete any leftover copy of the temporary file to prevent displaying the wrong patient report in event of error
If Dir(FName) <> "" Then Kill FName

'Retrieve all info required for printing report direct from tables. Do it now to ensure current data is used
Select Case Report
    Case ReportType.rtCpx
        TblName = "ExResults"
    Case Else
        TblName = "RftResults"
End Select

'Set up for the pdf
Set pdf = CreateObject("PDFCreatorPilot.PDFDocument4")
iReturn = pdf_GetLicenseData
pdf.SetLicenseData gCreatorPilot.RegistrationName, gCreatorPilot.SerialNo      ' initialize PDF Engine
Call pdf_SetFonts(pdf)
gPageResolution = pdf.PageResolution
pdf.PageSize = psA4
pdf.ScaleCoords 1, 1
pdf.InitialZoom = znFitPage
Call pdf_SetMargins(20, 8, 5, 10, 40)
        
        
'Get Demographics and results
Dim D As PtDemographics
Select Case Report
    Case ReportType.rtCpx
        Dim X As Cpx
        X = Pt.Get_CpxFromDB(ResultsID)
        X.Report = Suppress_ReportInPdf(X.ReportStatus, X.Report)
        D = Pt.Get_DemographicsFromDB(X.PtID)
        Call pdf_DrawReportHeader1(pdf, X.Ref_HealthServiceID, X.TestDate, D, LabType.ltResp, 1)
        Y = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromCpxUDT(X), LabType.ltResp)
    Case Else
        Dim R As RFTs
        R = Pt.Get_RftFromDB(ResultsID)
        R.Report = Suppress_ReportInPdf(R.ReportStatus, R.Report)
        D = Pt.Get_DemographicsFromDB(R.PtID)
        Call pdf_DrawReportHeader1(pdf, R.E_Ref_HealthServiceID, R.E_TestDate, D, LabType.ltResp, 1)
        Y = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromRftUDT(R), LabType.ltResp)
End Select


Select Case Report
    Case ReportType.rtRFT
        Call DrawPDFReportRftData1(pdf, R, D, Y)
        Call pdf_DoPageNumbering1(pdf, R.ReportStatus, rtRFT)
    Case ReportType.rtBronchChall
        Call DrawPDFReport_BronchChallengeData(pdf, R, Y)
        pdf_DoPageNumbering1 pdf, R.ReportStatus, rtBronchChall
    Case ReportType.rtHAST
        Call DrawPDFReport_HastData(pdf, R, Y)
        pdf_DoPageNumbering1 pdf, R.ReportStatus, rtHAST
    Case ReportType.rtSixMWD
        Call DrawPDFReport6MWDData(pdf, R, Y)
        pdf_DoPageNumbering1 pdf, R.ReportStatus, rtTreadmill
    Case ReportType.rtTreadmill
        Call DrawPDFReportTreadmillData(pdf, R, Y)
        pdf_DoPageNumbering1 pdf, R.ReportStatus, rtTreadmill
    Case ReportType.rtSkin
        Call DrawPDFReportSkinData(pdf, R, Y)
        pdf_DoPageNumbering1 pdf, R.ReportStatus, rtSkin
    Case ReportType.rtCpx
        Call DrawPDFReportCpxData(pdf, X, D, Y)
        pdf_DoPageNumbering1 pdf, X.ReportStatus, rtCpx
End Select

pdf.SaveToFile FName, LaunchAcrobat

If UpdateDatabase Then
    'Save as file and pointer
    Pdf_SaveAsPointer Report, FName, ResultsID
End If

Set pdf = Nothing                         '

Exit Function



DrawPDFReport_Rft_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport_Rft", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function DrawPDFRequest_Sleep(reqData As RequestFormHeaderData, LaunchAcrobat As Boolean, UpdateDatabase As Boolean) As Boolean

Dim i As Integer
Dim pdf As PDFCreatorPilotLib.PDFDocument4
Dim pdfStream As ADODB.Stream
Dim dbLockType
Dim TempFile As String



Dim Msg As String

    On Error GoTo DrawPDFRequest_Sleep_Error

If UpdateDatabase Then dbLockType = adLockPessimistic Else dbLockType = adLockReadOnly

'Delete any leftover copy of the temporary file to prevent displaying the wrong patient report in event of error
TempFile = Environ("tmp") & "\temp.pdf"
If Dir(TempFile) <> "" Then Kill TempFile

'Generate the pdf
Set pdf = CreateObject("PDFCreatorPilot.PDFDocument4")
iReturn = pdf_GetLicenseData
pdf.SetLicenseData gCreatorPilot.RegistrationName, gCreatorPilot.SerialNo      ' initialize PDF Engine

Call pdf_SetFonts(pdf)
gPageResolution = pdf.PageResolution
pdf.PageSize = psA4
pdf.InitialZoom = znFitPage
pdf.ScaleCoords 1, 1
Call pdf_SetMargins(20, 8, 5, 10, 40)
i = pdf_DrawRequestHeader(pdf, reqPsg, reqData)
i = pdf_DrawRequestSleepStudyBody(pdf, i)



pdf.SaveToFile TempFile, LaunchAcrobat

If UpdateDatabase Then
    Set pdfStream = New Stream
    pdfStream.Type = adTypeBinary
    pdfStream.Open
    pdfStream.LoadFromFile TempFile
    'Rs("pdfReport") = pdfStream.Read
    'Rs("pdf_Refresh") = "Refreshed: " & Now
    'Rs.Update
    Set pdfStream = Nothing
End If

'Tidy up
Set pdf = Nothing                         ' disconnect from library
'Rs.Close



    On Error GoTo 0
    Exit Function

DrawPDFRequest_Sleep_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFRequest_Sleep", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function DrawPDFReport_Sleep(LaunchAcrobat As Boolean, UpdateDatabase As Boolean, StudyID As Long, Optional ByVal FName) As Boolean

Dim z As Variant
Dim Y As Single
Dim pdf As PDFCreatorPilotLib.PDFDocument4
Dim iReturn
Dim SlpSections As SlpSection
Dim pdfStream As ADODB.Stream
Dim dbLockType
Dim Msg As String

On Error GoTo DrawPDFReport_Sleep_Error

If IsMissing(FName) Then FName = Environ("tmp") & "\temp.pdf"
If Dir(FName) <> "" Then Kill FName     'Delete any leftover copy of the temporary file to prevent displaying the wrong patient report in event of error

Call SetScoringRulesNotes

'Generate the pdf
Set pdf = CreateObject("PDFCreatorPilot.PDFDocument4")
iReturn = pdf_GetLicenseData
pdf.SetLicenseData gCreatorPilot.RegistrationName, gCreatorPilot.SerialNo      ' initialize PDF Engine
Call pdf_SetFonts(pdf)
gPageResolution = pdf.PageResolution
pdf.PageSize = psA4
pdf.InitialZoom = znFitPage
pdf.ScaleCoords 1, 1
Call pdf_SetMargins(20, 8, 5, 10, 40)

'Get Demographics and results
Dim D As PtDemographics
Dim s As SlpResults
s = Pt.Get_SlpFromDB(StudyID)
D = Pt.Get_DemographicsFromDB(s.PatientID)

'Check which results should be printed
If (s.R_TDT = "---" Or s.R_TDT = "" Or IsNull(s.R_TDT)) Then SlpSections.SleepStats = False Else SlpSections.SleepStats = True
If (s.R_SpO2_Time95 = "---" Or s.R_SpO2_Time95 = "" Or IsNull(s.R_SpO2_Time95)) Then SlpSections.SaO2 = False Else SlpSections.SaO2 = True
If (s.R_ArI_AllAll = "---" Or s.R_ArI_AllAll = "" Or IsNull(s.R_ArI_AllAll)) Then SlpSections.Arousals = False Else SlpSections.Arousals = True
If (s.R_ODI4_AllAll = "---" Or s.R_ODI4_AllAll = "" Or IsNull(s.R_ODI4_AllAll)) Then SlpSections.Desats = False Else SlpSections.Desats = True
If (s.R_AHI_AllAll = "---" Or s.R_AHI_AllAll = "" Or IsNull(s.R_AHI_AllAll)) Then SlpSections.Resp = False Else SlpSections.Resp = True
If (s.R_ABG_PaO21 = "---" Or s.R_ABG_PaO21 = "" Or s.R_ABG_PaCO21 = "---" Or s.R_ABG_PaCO21 = "" Or _
    IsNull(s.R_ABG_PaO21) Or IsNull(s.R_ABG_PaCO21)) Or _
    (s.R_ABG_PaO22 = "---" Or s.R_ABG_PaO22 = "" Or s.R_ABG_PaCO22 = "---" Or s.R_ABG_PaCO22 = "" Or _
    IsNull(s.R_ABG_PaO22) Or IsNull(s.R_ABG_PaCO22)) _
    Then SlpSections.Abg = False Else SlpSections.Abg = True
If s.R_Graphic Then SlpSections.Graphics = True Else SlpSections.Graphics = False

Call pdf_DrawReportHeader1(pdf, s.Request_HealthServiceID, CDate(s.StudyDate), D, LabType.ltSleep, 1)
Y = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromSlpUDT(s), LabType.ltSleep)
Y = DrawPDFReportSleepData(pdf, s, D, Y, SlpSections)
Call pdf_DoPageNumbering1(pdf, s.ReportStatus, rtSleep)

pdf.SaveToFile FName, LaunchAcrobat



If UpdateDatabase Then
    'Save as file and pointer
    Pdf_SaveAsPointer ReportType.rtSleep, FName, CLng(StudyID)
End If

'Tidy up
Set pdf = Nothing                         ' disconnect from library

Exit Function


DrawPDFReport_Sleep_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport_Sleep", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function DrawPDFReport_CpapClinic(LaunchAcrobat As Boolean, VisitID As Long, VisitDate As Date, PtID As Long) As Boolean

Dim z As Variant
Dim Y As Single
Dim pdf As PDFCreatorPilotLib.PDFDocument4
Dim iReturn
Dim pdfStream As ADODB.Stream
Dim Msg As String
Dim c As CPAPClinic
Dim Filename As String

On Error GoTo DrawPDFReport_CpapClinic_Error

If UpdateDatabase Then dbLockType = adLockPessimistic Else dbLockType = adLockReadOnly

'Delete any leftover copy of the temporary file to prevent displaying the wrong patient report in event of error
Filename = Environ("tmp") & "\temp.pdf"
If Dir(Filename) <> "" Then Kill Filename

'Generate the pdf
Set pdf = CreateObject("PDFCreatorPilot.PDFDocument4")
iReturn = pdf_GetLicenseData
pdf.SetLicenseData gCreatorPilot.RegistrationName, gCreatorPilot.SerialNo      ' initialize PDF Engine
Call pdf_SetFonts(pdf)
gPageResolution = pdf.PageResolution
pdf.PageSize = psA4
pdf.InitialZoom = znFitPage
pdf.ScaleCoords 1, 1
Call pdf_SetMargins(20, 8, 5, 10, 40)

Call pdf_DrawReportHeader1(pdf, HSID.AustinHealth, VisitDate, Pt.Get_DemographicsFromDB(PtID), LabType.ltCpapClinic, 1)
c = Pt.Get_CpapFromDB(VisitID)
Y = pdf_DrawReportInfo1(pdf, Pt.Get_ReportInfoFromCpapUDT(c), LabType.ltCpapClinic)
Y = DrawPDFReportCPAPClinicData(pdf, c, Y)
Call pdf_DoPageNumbering1(pdf, "", rtCpapClinic)

Call pdf.SaveToFile(Filename, LaunchAcrobat)

Set pdf = Nothing                         ' disconnect from library

Exit Function

DrawPDFReport_CpapClinic_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport_CpapClinic", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function pdf_DrawReportInfo1(pdf As PDFCreatorPilotLib.PDFDocument4, R As ReportInfoSection, Lab As LabType) As Single
'Draws info section for sleep and resp lab reports in pdf format
'Returns the number of lines of clinical note printed

Dim i As Integer
Dim temp As String
Dim X As Single, Y As Single        'Current print co-ords
Dim x1 As Single, x2 As Single, x3 As Single
Dim y1 As Single, y2 As Single, y3 As Single, y4 As Single, y5 As Single
Dim LFs As Integer      'Number of linefeeds
Dim Hb As String
Dim Smoke As String
Dim ReqBy As String
Dim ReqAddress As String
Dim TestDate As String
Dim TestTime As String
Dim BD As String
Dim CopyTo As String
Dim techNote As String
Dim clinNote As String
Dim Location As String
Dim BMI
Dim Rs As New ADODB.Recordset
Dim Msg As String
Dim Ht As Single, Wt As Single

On Error GoTo pdf_DrawReportInfo1_Error
    
'Lab specific stuff
Select Case Lab
    Case LabType.ltResp
        If Val(R.Resp_Hb) > 0 Then
            If R.Resp_Hb_Info & "" = "" Then
                Hb = "Hb: " & R.Resp_Hb & " g/dL"
            Else
                Hb = "Hb: " & R.Resp_Hb & " g/dL" & " (" & R.Resp_Hb_Info & ")"
            End If
        Else
            Hb = "Hb: N/A"
        End If
        If R.Resp_Packyrs & "" = "" Then
            Smoke = "Smoking Hx:  " & R.Resp_Smoke
        Else
            Smoke = "Smoking Hx:  " & R.Resp_Smoke & " (" & R.Resp_Packyrs & " pack yrs)"
        End If
        LFs = 0
        ReqBy = "Requested by: " & R.RequestingMO
        ReqAddress = "Report to: " & R.Resp_RefLocation
        TestDate = "Test Date:   " & Format(R.TestDate, "d/mm/yyyy")
        TestTime = "Test Time:   " & Format(R.TestTime, "hh:mm") & " hrs"
        BD = "Last BD:     " & R.Resp_LastBd
        CopyTo = "Copy to:   " & R.Resp_ReportCopyTo
        clinNote = R.ClinicalNote
        Location = R.Lab
        Ht = R.ThisVisit_height
        Wt = R.ThisVisit_weight
    Case LabType.ltSleep
        Hb = ""
        Smoke = "Smoking Hx:"
        LFs = 0
        ReqBy = "Requested by: " & R.RequestingMO
        ReqAddress = "Report to: " & R.Slp_ReportDest_1
        TestDate = "Test Date:   " & Format(R.TestDate, "d/mm/yyyy")
        TestTime = "Study Type:   " & R.Slp_StudyType
        If R.Slp_TreatmentMode = "" Then BD = "Treatment Mode: Nil" Else BD = "Treatment Mode:   " & R.Slp_TreatmentMode
        CopyTo = "Copy to:   " & R.Slp_ReportDest_2
        clinNote = R.ClinicalNote
        Location = R.Lab
        Ht = R.ThisVisit_height
        Wt = R.ThisVisit_weight
     Case LabType.ltCpapClinic
        Hb = ""
        Smoke = "Referral mode: " & R.Cpap_ReferralMode
        LFs = 0
        ReqBy = "Referred by: " & R.RequestingMO
        ReqAddress = "Report to: " & R.Cpap_ReportTo
        TestDate = "Date:   " & Format(R.TestDate, "d/mm/yyyy")
        TestTime = ""
        BD = ""
        CopyTo = ""
        clinNote = R.Cpap_ReasonForReferral
        Location = R.Cpap_VisitLocation
        Ht = 0
        Wt = 0
End Select

y1 = M.Top + DPI(41)                            'top of main rectangle
y2 = DPI(25)                                    'height of main rectangle
y3 = DPI(4.5)                                   'Line height
y4 = M.Top + DPI(41) + y2                       'top of clin note rectangle
y5 = DPI(15)                                    'height of clin note rectangle

x1 = M.Left                                     'left coord of main rectangle
x2 = pdf.PageWidth - M.RIGHT - M.Left           'width of main rectangle
x3 = DPI(135)                                   'left coord of testdate rectangle

pdf.DrawRectangle x1, y1, x2, y2, 0                     'Main
pdf.DrawRectangle x1, y1, x3, y2, 0                 'Test date etc
pdf.Stroke

pdf.UseFont gPdfFonts(6).ID, 10
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2), "Height:" & Ht & " cm     Weight:" & Wt & " kg"
pdf.ShowTextAt x1 + DPI(52), y1 + DPI(2), "BMI: " & Calculate_BMI(CSng(Ht), CSng(Wt)) & " kg/m2"

x3 = x3 + M.Left
pdf.ShowTextAt x1 + DPI(90), y1 + DPI(2), Hb
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3, Smoke
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3 * (2 - LFs), ReqBy
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3 * (3 - LFs), ReqAddress
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3 * (4 - LFs), CopyTo
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2), TestDate
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2) + y3, TestTime
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2) + y3 * 2, BD
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2) + y3 * 3, "Location: " & Location

'Clincial notes section
pdf.UseFont gPdfFonts(3).ID, 9
pdf.ShowTextAt x1 + DPI(2), y4 + DPI(2), "CLINICAL NOTES / REASON FOR REFERRAL"
pdf.UseFont gPdfFonts(2).ID, 9
i = pdf.ShowTextLines(x1 + DPI(2), y4 + DPI(6), pdf.PageWidth - M.RIGHT - DPI(2), y4 + DPI(24), 10, taLeft, vaTop, clinNote)

'pdf.DrawRectangle x1, y4, x2, i, 0           'Clin note
'pdf.DrawRectangle x1, y4, x2, pdf.GetCurrentTextY - y4 + pdf.GetTextHeight(clinNote) + DPI(1), 0           'Clin note
pdf.Stroke
pdf_DrawReportInfo1 = pdf.GetCurrentTextY

Exit Function


pdf_DrawReportInfo1_Error:
    Msg = "Error drawing report header"
    Call ErrorLog(Msg, "pdf_DrawReportInfo1", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function pdf_DrawReportInfoxxx(pdf As PDFCreatorPilotLib.PDFDocument4, Rs As ADODB.Recordset, Lab As LabType) As Single
'Draws info section for sleep and resp lab reports in pdf format
'Returns the number of lines of clinical note printed

Dim i As Integer
Dim temp As String
Dim X As Single, Y As Single        'Current print co-ords
Dim x1 As Single, x2 As Single, x3 As Single
Dim y1 As Single, y2 As Single, y3 As Single, y4 As Single, y5 As Single
Dim LFs As Integer      'Number of linefeeds
Dim Hb As String
Dim Smoke As String
Dim ReqBy As String
Dim ReqAddress As String
Dim TestDate As String
Dim TestTime As String
Dim BD As String
Dim CopyTo As String
Dim techNote As String
Dim clinNote As String
Dim Location As String

Dim Msg As String

    On Error GoTo pdf_DrawReportInfo_Error


'Lab specific stuff
Select Case Lab
    Case LabType.ltResp
        If Not IsNull(Rs("Hb")) And Rs("Hb") > 0 Then
            If Rs("HbInfo") & "" = "" Then
                Hb = "Hb: " & Rs("Hb") & " g/dL"
            Else
                Hb = "Hb: " & Rs("Hb") & " g/dL" & " (" & Rs("HbInfo") & ")"
            End If
        Else
            Hb = "Hb: N/A"
        End If
        If Rs("Packyrs") & "" = "" Then
            Smoke = "Smoking Hx:  " & Rs("Smoke")
        Else
            Smoke = "Smoking Hx:  " & Rs("Smoke") & " (" & Rs("Packyrs") & " pack yrs)"
        End If
        LFs = 0
        ReqBy = "Requested by: " & Rs("MO") & ""
        ReqAddress = "Report to: " & Rs("RefLocation") & ""
        TestDate = "Test Date:   " & Format(Rs("TestDate"), "d/mm/yyyy")
        TestTime = "Test Time:   " & Format(Rs("TestTime"), "hh:mm") & " hrs"
        BD = "Last BD:     " & Rs("LastBd")
        CopyTo = "Copy to:   " & Rs("Report_CopyTo") & ""
        clinNote = Rs("ClinicalNote") & ""
        Location = Rs("Campus")
        
    Case LabType.ltSleep
        Hb = ""
        Smoke = "Smoking Hx:"
        LFs = 0
        ReqBy = "Requested by: " & Rs("RequestingMO") & ""
        ReqAddress = "Report to: " & Rs("ReportDest_1") & ""
        TestDate = "Test Date:   " & Format(Rs("StudyDate"), "d/mm/yyyy")
        TestTime = "Study Type:   " & Rs("StudyType")
        If Rs("TreatmentMode") & "" = "" Then BD = "Treatment Mode: Nil" Else BD = "Treatment Mode:   " & Rs("TreatmentMode")
        CopyTo = "Copy to:   " & Rs("ReportDest_2")
        clinNote = Rs("ReasonForTest") & ""
        Location = Rs("Campus")
        
     Case LabType.ltCpapClinic
        Hb = ""
        Smoke = "Referral mode: " & Rs("Clinic_ReferralMode")
        LFs = 0
        ReqBy = "Referred by: " & Rs("Clinic_ReferredBy") & ""
        ReqAddress = "Report to: " & Rs("Clinic_ReportTo") & ""
        TestDate = "Date:   " & Format(Rs("VisitDate"), "d/mm/yyyy")
        TestTime = ""
        BD = ""
        CopyTo = ""
        clinNote = Rs("Clinic_ReferralReason") & ""
        Location = Rs("VisitLocation")
     
End Select


y1 = M.Top + DPI(41)                            'top of main rectangle
y2 = DPI(25)                                    'height of main rectangle
y3 = DPI(4.5)                                   'Line height
y4 = M.Top + DPI(41) + y2                       'top of clin note rectangle
y5 = DPI(15)                                    'height of clin note rectangle

x1 = M.Left                                     'left coord of main rectangle
x2 = pdf.PageWidth - M.RIGHT - M.Left           'width of main rectangle
x3 = DPI(135)                                   'left coord of testdate rectangle

pdf.DrawRectangle x1, y1, x2, y2, 0                     'Main
pdf.DrawRectangle x1, y1, x3, y2, 0                 'Test date etc
pdf.Stroke

pdf.UseFont gPdfFonts(6).ID, 10
'*** If this is a RFT report use ht at test else print demographic ht
If Lab = LabType.ltResp Then
    pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2), "Height:" & Rs("Height_rft") & " cm     Weight:" & Rs("Weight") & " kg"
Else
    pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2), "Height:" & Rs("Height") & " cm     Weight:" & Rs("Weight") & " kg"
End If

'*** If there is a Height at rft test then use this for BMI else
'*** use generic height since this section is shared across Resp, Sleep & VRSS
If Lab = LabType.ltResp Then
    If Rs("Height_rft") > 0 Then
        If Val(Rs("weight") & "") > 0 Then
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: " & Format(10000 * Rs("weight") / Rs("Height_rft") / Rs("Height_rft"), "##.0 kg/m/m")
        Else
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: N/A"
        End If
    Else
        If Val(Rs("weight") & "") > 0 And Rs("Height") > 0 Then
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: " & Format(10000 * Rs("weight") / Rs("Height") / Rs("Height"), "##.0 kg/m/m")
        Else
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: N/A"
        End If
    End If
Else
    If Rs("Height") > 0 Then
        If Val(Rs("weight") & "") > 0 Then
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: " & Format(10000 * Rs("weight") / Rs("Height") / Rs("Height"), "##.0 kg/m/m")
        Else
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: N/A"
        End If
    Else
        If Val(Rs("weight") & "") > 0 And Rs("Height") > 0 Then
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: " & Format(10000 * Rs("weight") / Rs("Height") / Rs("Height"), "##.0 kg/m/m")
        Else
            pdf.ShowTextAt x1 + DPI(50), y1 + DPI(2), "BMI: N/A"
        End If
    End If
End If

x3 = x3 + M.Left
pdf.ShowTextAt x1 + DPI(90), y1 + DPI(2), Hb
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3, Smoke
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3 * (2 - LFs), ReqBy
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3 * (3 - LFs), ReqAddress
pdf.ShowTextAt x1 + DPI(2), y1 + DPI(2) + y3 * (4 - LFs), CopyTo
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2), TestDate
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2) + y3, TestTime
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2) + y3 * 2, BD
pdf.ShowTextAt x3 + DPI(2), y1 + DPI(2) + y3 * 3, "Location: " & Location

'Clincial notes section
pdf.UseFont gPdfFonts(3).ID, 9
pdf.ShowTextAt x1 + DPI(2), y4 + DPI(2), "CLINICAL NOTES / REASON FOR REFERRAL"
pdf.UseFont gPdfFonts(2).ID, 9
i = pdf.ShowTextLines(x1 + DPI(2), y4 + DPI(6), pdf.PageWidth - M.RIGHT, y4 + DPI(14), 10, taLeft, vaTop, clinNote)

pdf.DrawRectangle x1, y4, x2, pdf.GetCurrentTextY - y4 + pdf.GetTextHeight(clinNote) + DPI(1), 0           'Clin note
pdf.Stroke
pdf_DrawReportInfo = pdf.GetCurrentTextY

'Exit Function



'Error_pdf_DrawReportInfo:
'MsgBox "Error drawing report header (pdf_DrawReportInfo)" & vbCrLf & Err.Description
'Exit Function
'Resume

    On Error GoTo 0
    Exit Function

pdf_DrawReportInfo_Error:
    Msg = "Error drawing report header"
    Call ErrorLog(Msg, "pdf_DrawReportInfo", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function pdf_DrawReportHeaderxxx(pdf As PDFCreatorPilotLib.PDFDocument4, Rs As ADODB.Recordset, Lab As LabType, PageNum As Integer)
'Draws a standardised pdf format pageheader for sleep and resp lab reports

Dim i As Integer
Dim temp As String
Dim X As Single, Y As Single        'Current print co-ords
Dim x1 As Single, x2 As Single, x3 As Single
Dim flagNataLogo As Boolean
Dim flagTsanzLogo As Boolean
Dim flagRespLabForm As Boolean
Dim flagSleepLabForm As Boolean
Dim flagCpapClinicForm As Boolean
Dim ImageList(1 To 7) As Integer
Dim LabName As String
Dim labPhone As String
Dim labReportName As String
Dim Page As String
Dim Age As Single
Dim URforBarcode As String

Dim Msg As String

    On Error GoTo pdf_DrawReportHeader_Error



'Lab specific stuff
Select Case Lab
    Case LabType.ltResp
        flagNataLogo = True
        flagTsanzLogo = True
        LabName = "Respiratory Laboratory"
        labPhone = "Ph:(03)9496 5739  Fax:(03)9496 3723"
        labReportName = "RESPIRATORY FUNCTION REPORT"
        Page = ""
        '*B*A*R*
        '*B*A*R* Set flag for RespLab Barcode
        '*B*A*R*
        flagRespLabForm = True
        Age = GetAge(Rs("DOB"), Rs("TestDate"), 0)
    Case LabType.ltSleep
        flagNataLogo = True
        flagTsanzLogo = False
        LabName = "Sleep Disorders Laboratory"
        labPhone = "Ph:(03)9496 3508  Fax:(03)9496 5124"
        labReportName = "SLEEP STUDY REPORT"
        Page = ""
        '*B*A*R*
        '*B*A*R* Set flag for SleepLab Barcode
        '*B*A*R*
        flagSleepLabForm = True
        Age = GetAge(Rs("DOB") & "", Rs("StudyDate"), 0)
    Case LabType.ltCpapClinic
        flagNataLogo = True
        flagTsanzLogo = False
        LabName = "Sleep Disorders Laboratory"
        labPhone = "Ph:(03)9496 3508  Fax:(03)9496 5124"
        labReportName = "CPAP CLINIC FOLLOWUP"
        Page = ""
        '*B*A*R*
        '*B*A*R* Set flag for SleepLab Barcode
        '*B*A*R*
        flagCpapClinicForm = True
        Age = GetAge(Rs("DOB") & "", Rs("VisitDate"), 0)
End Select


'Draw graphic images
ImageList(1) = pdf.AddImageFromFile(pathAustinLogo)
ImageList(2) = pdf.AddImageFromFile(pathTsanzLogo)
ImageList(3) = pdf.AddImageFromFile(pathNataLogo)
ImageList(4) = pdf.AddImageFromFile(pathRcpaLogo)
ImageList(5) = pdf.AddImageFromFile(pathRespLabForm) '*B*A*R code for RespLab D22 form
ImageList(6) = pdf.AddImageFromFile(pathSleepLabForm) '*B*A*R code for SleepLab form

'*B*A*R* code load URtemp Form for barCode but only for Austin UR's
If Rs("request_HealthServiceID") = HSID.AustinHealth Then
    URforBarcode = Get_PtString(Rs("PatientID"), Rs("request_HealthServiceID"), CurrentPatientItems.URforBarcode)
    frmBarCodeTemp.txtUR.Text = URforBarcode
    '*B*A*R* Set licence info for BarCode Active X
    frmBarCodeTemp.BarcodeC1.SetLicense "BC21C5-C718C3-D8430D-0F1214-27EB1F-4608"
    '*B*A*R*
    '*B*A*R* call routine to print UR Bar code to right of UR number
    '*B*A*R*
    frmBarCodeTemp.BarcodeC1.Orientation = BC_Angle_0
    
    '*B*A*R* 'Draw UR Barcode in Demographic box area
    frmBarCodeTemp.BarcodeC1.AutoCheck = True
    frmBarCodeTemp.BarcodeC1.BarcodeType = BC_CODE39
    frmBarCodeTemp.BarcodeC1.WideToNarrowRatio = 3
    frmBarCodeTemp.BarcodeC1.BarHeight = 0.5
    frmBarCodeTemp.BarcodeC1.ShowText = False
    frmBarCodeTemp.BarcodeC1.ShowCheckDigit = False
    frmBarCodeTemp.BarcodeC1.AutoCheck = False
    '*B*A*R* when no UR data entered ditch Barcode
    frmBarCodeTemp.BarcodeC1.BarcodeData = URforBarcode
     
    '*B*A*R* set local path to environ
    Filename = Environ("tmp") & "\BarCodeURTemp.jpg"
    frmBarCodeTemp.BarcodeC1.SaveAs Filename
    '*B*A*R* set printer stylus back to where it was

    ImageList(7) = pdf.AddImageFromFile(Filename)  '*B*A*R* code for UR saved locally (environ)
End If

pdf.DrawImage ImageList(1), M.Left - DPI(2), M.Top - DPI(1), DPI(45), DPI(11), 0

' *B*A*R* set correct bar code for Lab.
If flagRespLabForm = True Then
    pdf.DrawImage ImageList(5), M.Left - DPI(17), M.Top + DPI(30), DPI(15), DPI(42), 0
ElseIf flagSleepLabForm Then
    pdf.DrawImage ImageList(6), M.Left - DPI(17), M.Top + DPI(30), DPI(15), DPI(42), 0
ElseIf flagCpapClinicForm Then

End If
' *B*A*R* *C*O*D*E* end

pdf.UseFont gPdfFonts(2).ID, 5
If flagTsanzLogo Then
    pdf.DrawImage ImageList(2), M.Left + DPI(58), M.Top + DPI(4), DPI(10), DPI(7), 0
    pdf.ShowTextAt M.Left + DPI(68), M.Top + DPI(7), "TSANZ Accredited"
End If
If flagNataLogo Then
    pdf.DrawImage ImageList(3), M.Left + DPI(58), M.Top + DPI(12), DPI(10), DPI(10), 0
    pdf.DrawImage ImageList(4), M.Left + DPI(72), M.Top + DPI(14), DPI(20), DPI(6), 0
    pdf.ShowTextAt M.Left + DPI(69), M.Top + DPI(21), "NATA Accreditation No 15760"
    pdf.ShowTextAt M.Left + DPI(60), M.Top + DPI(23), "Accredited for compliance with AS 4633 (ISO 15189)"
End If

'Draw lines and rectangles
pdf.SetLineWidth 0.5
pdf.DrawRectangle M.Left + DPI(107), M.Top + DPI(10), pdf.PageWidth - M.RIGHT - M.Left - DPI(107), DPI(31), 0      'patient label
pdf.DrawRectangle M.Left, M.Top + DPI(41), pdf.PageWidth - M.RIGHT - M.Left, pdf.PageHeight - M.Bottom - M.Top - DPI(41), 0 'Main
pdf.Stroke

pdf.UseFont gPdfFonts(3).ID, 12
X = M.Left + DPI(110): Y = M.Top + DPI(4)
pdf.ShowTextAt X, Y, labReportName
pdf.UseFont gPdfFonts(2).ID, 9
X = M.Left + DPI(165): Y = M.Top + DPI(4)
pdf.ShowTextAt X, Y, Page
pdf.UseFont gPdfFonts(7).ID, 11
X = M.Left: Y = Y + DPI(8)
pdf.ShowTextAt X, Y, LabName
pdf.UseFont gPdfFonts(7).ID, 10
X = M.Left: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Respiratory and Sleep Medicine"
Y = Y + DPI(4)
pdf.UseFont gPdfFonts(7).ID, 7
pdf.ShowTextAt X, Y, "Studley Rd, Heidelberg, Vic 3084."
Y = Y + DPI(3.5)
pdf.ShowTextAt X, Y, labPhone
pdf.UseFont gPdfFonts(4).ID, 7
X = M.Left: x1 = X + DPI(25): x2 = X + DPI(50): x3 = X + DPI(75): Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Prof CF McDonald"
pdf.ShowTextAt x1, Y, "Dr N Antoniades"
pdf.ShowTextAt x2, Y, "Dr M Barnes"
pdf.ShowTextAt x3, Y, "Dr M Caldecott"
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr CJ Worsnop"
pdf.ShowTextAt x1, Y, "Dr KA Detering"
pdf.ShowTextAt x2, Y, "Dr N Goh"
pdf.ShowTextAt x3, Y, "Dr S Haque"
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr M Howard"
pdf.ShowTextAt x1, Y, "Dr M Ibrahim"
pdf.ShowTextAt x2, Y, "Dr C Lanteri"
pdf.ShowTextAt x3, Y, "Dr M McMahon"
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr F O'Donoghue"
pdf.ShowTextAt x1, Y, "Dr M Sutherland"
pdf.ShowTextAt x2, Y, ""
pdf.ShowTextAt x3, Y, ""

'Print patient demographics
X = M.Left + DPI(110): Y = M.Top + DPI(11)
pdf.UseFont gPdfFonts(7).ID, 11
pdf.ShowTextAt X, Y, "Name: " & Rs("FirstName") & " " & Replace(Rs("Surname"), "~", "'")
Y = Y + DPI(6)
pdf.ShowTextAt X, Y, "UR:    " & Get_URForVAED(Rs("PatientID"), Rs("Request_HealthServiceID"))

'*B*A*R* If the UR is not Austin, or is blank then don't print barcode image as it is not valid
If Rs("Request_HealthServiceID") = HSID.AustinHealth Then
    pdf.DrawImage ImageList(7), X + 25, Y + DPI(4), DPI(35), DPI(8), 0
End If

pdf.UseFont gPdfFonts(6).ID, 10
Y = Y + DPI(10)
pdf.ShowTextAt X, Y, "DOB: " & Rs("DOB") & " (" & Age & "yrs)" & "      Gender: " & Rs("Sex")
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Address: " & Rs("Address")
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, Space(15) & Rs("Suburb") & ", " & Rs("PostCode") & "."


'Exit Function



'Error_pdf_PrintPageHeader:
'MsgBox "Error drawing report header (pdf_PrintPageHeader)" & vbCrLf & Err.Description
'Exit Function

'Resume

    On Error GoTo 0
    Exit Function

pdf_DrawReportHeader_Error:
    Msg = "Error drawing report header."
    Call ErrorLog(Msg, "pdf_DrawReportHeader", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function pdf_DrawReportHeader1(pdf As PDFCreatorPilotLib.PDFDocument4, HealthServiceID As HSID, TestDate As Date, D As PtDemographics, Lab As LabType, PageNum As Integer)
'Draws a standardised pdf format pageheader for sleep and resp lab reports

Dim i As Integer
Dim temp As String
Dim X As Single, Y As Single        'Current print co-ords
Dim x1 As Single, x2 As Single, x3 As Single, x4 As Single
Dim flagNataLogo As Boolean
Dim flagTsanzLogo As Boolean
Dim flagRespLabForm As Boolean
Dim flagSleepLabForm As Boolean
Dim flagCpapClinicForm As Boolean
Dim ImageList(1 To 7) As Integer
Dim LabName As String
Dim labPhone As String
Dim labReportName As String
Dim Page As String
Dim Age As Single
Dim URforBarcode As String
Dim Msg As String

On Error GoTo pdf_DrawReportHeader1_Error


'Lab specific stuff
Select Case Lab
    Case LabType.ltResp
        flagNataLogo = True
        flagTsanzLogo = True
        LabName = "Respiratory Laboratory"
        labPhone = "Ph:(03)9496 5739  Fax:(03)9496 3723"
        labReportName = "RESPIRATORY FUNCTION REPORT"
        Page = ""
        '*B*A*R*
        '*B*A*R* Set flag for RespLab Barcode
        '*B*A*R*
        flagRespLabForm = True
        Age = GetAge(D.DOB, TestDate, 0)
    Case LabType.ltSleep
        flagNataLogo = True
        flagTsanzLogo = False
        LabName = "Sleep Disorders Laboratory"
        labPhone = "Ph:(03)9496 3508  Fax:(03)9496 5124"
        labReportName = "SLEEP STUDY REPORT"
        Page = ""
        '*B*A*R*
        '*B*A*R* Set flag for SleepLab Barcode
        '*B*A*R*
        flagSleepLabForm = True
        Age = GetAge(D.DOB, TestDate, 0)
    Case LabType.ltCpapClinic
        flagNataLogo = True
        flagTsanzLogo = False
        LabName = "Sleep Disorders Laboratory"
        labPhone = "Ph:(03)9496 3508  Fax:(03)9496 5124"
        labReportName = "CPAP CLINIC FOLLOWUP"
        Page = ""
        '*B*A*R*
        '*B*A*R* Set flag for SleepLab Barcode
        '*B*A*R*
        flagCpapClinicForm = True
        Age = GetAge(D.DOB & "", TestDate, 0)
End Select


'Draw graphic images
ImageList(1) = pdf.AddImageFromFile(pathAustinLogo)
ImageList(2) = pdf.AddImageFromFile(pathTsanzLogo)
'changed code to flip Nata and Rcpa images as per Nata visit 09/2013 - WRR
ImageList(4) = pdf.AddImageFromFile(pathNataLogo)
ImageList(3) = pdf.AddImageFromFile(pathRcpaLogo)
ImageList(5) = pdf.AddImageFromFile(pathRespLabForm) '*B*A*R code for RespLab D22 form
ImageList(6) = pdf.AddImageFromFile(pathSleepLabForm) '*B*A*R code for SleepLab form

'*B*A*R* code load URtemp Form for barCode but only for Austin UR's

URforBarcode = Get_PtString(D.ID, HealthServiceID, CurrentPatientItems.URforBarcode)
If URforBarcode <> "" Then
    frmBarCodeTemp.txtUR.Text = URforBarcode
    '*B*A*R* Set licence info for BarCode Active X
    frmBarCodeTemp.BarcodeC1.SetLicense "BC21C5-C718C3-D8430D-0F1214-27EB1F-4608"
    '*B*A*R*
    '*B*A*R* call routine to print UR Bar code to right of UR number
    '*B*A*R*
    frmBarCodeTemp.BarcodeC1.Orientation = BC_Angle_0
    
    '*B*A*R* 'Draw UR Barcode in Demographic box area
    frmBarCodeTemp.BarcodeC1.AutoCheck = True
    frmBarCodeTemp.BarcodeC1.BarcodeType = BC_CODE39
    frmBarCodeTemp.BarcodeC1.WideToNarrowRatio = 3
    frmBarCodeTemp.BarcodeC1.BarHeight = 0.5
    frmBarCodeTemp.BarcodeC1.ShowText = False
    frmBarCodeTemp.BarcodeC1.ShowCheckDigit = False
    frmBarCodeTemp.BarcodeC1.AutoCheck = False
    '*B*A*R* when no UR data entered ditch Barcode
    frmBarCodeTemp.BarcodeC1.BarcodeData = URforBarcode
     
    '*B*A*R* set local path to environ
    Filename = Environ("tmp") & "\BarCodeURTemp.jpg"
    frmBarCodeTemp.BarcodeC1.SaveAs Filename
    '*B*A*R* set printer stylus back to where it was

    ImageList(7) = pdf.AddImageFromFile(Filename)  '*B*A*R* code for UR saved locally (environ)

End If
pdf.DrawImage ImageList(1), M.Left - DPI(2), M.Top - DPI(1), DPI(45), DPI(11), 0

' *B*A*R* set correct bar code for Lab.
If flagRespLabForm = True Then
    pdf.DrawImage ImageList(5), M.Left - DPI(17), M.Top + DPI(30), DPI(15), DPI(42), 0
ElseIf flagSleepLabForm Then
    pdf.DrawImage ImageList(6), M.Left - DPI(17), M.Top + DPI(30), DPI(15), DPI(42), 0
ElseIf flagCpapClinicForm Then

End If
' *B*A*R* *C*O*D*E* end

pdf.UseFont gPdfFonts(2).ID, 5
If flagTsanzLogo Then
    pdf.DrawImage ImageList(2), M.Left + DPI(58), M.Top + DPI(4), DPI(10), DPI(7), 0
    pdf.ShowTextAt M.Left + DPI(68), M.Top + DPI(7), "TSANZ Accredited"
End If
If flagNataLogo Then
    'changed code to flip Nata and Rcpa images as per Nata visit 09/2013 - WRR
    'pdf.DrawImage ImageList(3), M.Left + DPI(58), M.Top + DPI(12), DPI(10), DPI(10), 0
    pdf.DrawImage ImageList(3), M.Left + DPI(58), M.Top + DPI(14), DPI(20), DPI(6), 0
    'pdf.DrawImage ImageList(4), M.Left + DPI(72), M.Top + DPI(14), DPI(20), DPI(6), 0
    pdf.DrawImage ImageList(4), M.Left + DPI(82), M.Top + DPI(12), DPI(10), DPI(10), 0
    pdf.ShowTextAt M.Left + DPI(59), M.Top + DPI(21), "NATA Accreditation No 15760"
    pdf.ShowTextAt M.Left + DPI(59), M.Top + DPI(23), "Accredited for compliance with AS 4633 (ISO 15189)"
End If

'Draw lines and rectangles
pdf.SetLineWidth 0.5
pdf.DrawRectangle M.Left + DPI(107), M.Top + DPI(10), pdf.PageWidth - M.RIGHT - M.Left - DPI(107), DPI(31), 0      'patient label
pdf.DrawRectangle M.Left, M.Top + DPI(41), pdf.PageWidth - M.RIGHT - M.Left, pdf.PageHeight - M.Bottom - M.Top - DPI(41), 0 'Main
pdf.Stroke

pdf.UseFont gPdfFonts(3).ID, 12
X = M.Left + DPI(110): Y = M.Top + DPI(4)
pdf.ShowTextAt X, Y, labReportName
pdf.UseFont gPdfFonts(2).ID, 9
X = M.Left + DPI(165): Y = M.Top + DPI(4)
pdf.ShowTextAt X, Y, Page
pdf.UseFont gPdfFonts(7).ID, 11
X = M.Left: Y = Y + DPI(8)
pdf.ShowTextAt X, Y, LabName
pdf.UseFont gPdfFonts(7).ID, 10
X = M.Left: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Respiratory and Sleep Medicine"
Y = Y + DPI(4)
pdf.UseFont gPdfFonts(7).ID, 7
pdf.ShowTextAt X, Y, "Studley Rd, Heidelberg, Vic 3084."
Y = Y + DPI(3.5)
pdf.ShowTextAt X, Y, labPhone
pdf.UseFont gPdfFonts(4).ID, 7
X = M.Left: x1 = X + DPI(24): x2 = X + DPI(46): x3 = X + DPI(63): x4 = X + DPI(83): Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Prof CF McDonald"
pdf.ShowTextAt x1, Y, "Dr N Antoniades"
pdf.ShowTextAt x2, Y, "Dr M Barnes"
pdf.ShowTextAt x3, Y, "Dr M Caldecott"
pdf.ShowTextAt x4, Y, ""
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr CJ Worsnop"
pdf.ShowTextAt x1, Y, "Dr S Haque"
pdf.ShowTextAt x2, Y, "Dr N Goh"
pdf.ShowTextAt x3, Y, "Dr B Jennings"
pdf.ShowTextAt x4, Y, ""
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr M Howard"
pdf.ShowTextAt x1, Y, "Dr M Ibrahim"
pdf.ShowTextAt x2, Y, "Dr C Lanteri"
pdf.ShowTextAt x3, Y, "Dr M McMahon"
pdf.ShowTextAt x4, Y, ""
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr F O'Donoghue"
pdf.ShowTextAt x1, Y, "Dr M Sutherland"
pdf.ShowTextAt x2, Y, "Dr N Atkins"
pdf.ShowTextAt x3, Y, "Dr J Ward"
pdf.ShowTextAt x4, Y, ""

'Print patient demographics
X = M.Left + DPI(110): Y = M.Top + DPI(11)
pdf.UseFont gPdfFonts(7).ID, 11
pdf.ShowTextAt X, Y, "Name: " & D.Firstname & " " & D.Surname
Y = Y + DPI(5)
If HealthServiceID = AustinHealth Then
    pdf.ShowTextAt X, Y, "UR:    " & D.UR_PrimaryAustin
Else
    pdf.ShowTextAt X, Y, "UR:    " & Get_URForVAED(D.ID, HealthServiceID)
End If

'*B*A*R* If the UR is not Austin, or is blank then don't print barcode image as it is not valid
If URforBarcode <> "" Then
    pdf.DrawImage ImageList(7), X + 25, Y + DPI(4), DPI(40), DPI(8), 0
End If

pdf.UseFont gPdfFonts(6).ID, 10
Y = Y + DPI(10.5)
pdf.ShowTextAt X, Y, "DOB: " & D.DOB & " (" & Age & "yrs)" & "      Gender: " & D.Gender
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Address: " & D.Address1
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, Space(15) & D.Suburb & ", " & D.Postcode & "."

Exit Function


pdf_DrawReportHeader1_Error:
    Msg = "Error drawing report header."
    Call ErrorLog(Msg, "pdf_DrawReportHeader1", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function pdf_DrawRequestHeader(pdf As PDFCreatorPilotLib.PDFDocument4, Req As RequestType, reqData As RequestFormHeaderData) As Integer
'Draws a standardised pdf format pageheader for sleep and resp lab request forms
'Returns the y position of the lowest part of the header

Dim i As Integer
Dim temp As String
Dim X As Single, Y As Single        'Current print co-ords
Dim x1 As Single, x2 As Single, x3 As Single
Dim flagNataLogo As Boolean
Dim flagTsanzLogo As Boolean
Dim ImageList(1 To 4) As Integer
Dim LabName As String
Dim labPhone As String
Dim labReportName As String
Dim Page As String
Dim Age As Single

Dim Msg As String

    On Error GoTo pdf_DrawRequestHeader_Error

'Request type specific stuff
Select Case Req
    Case RequestType.reqRft
        flagNataLogo = True
        flagTsanzLogo = True
        LabName = "Respiratory Laboratory"
        labPhone = "Ph:(03)9496 5739  Fax:(03)9496 3723"
        labReportName = "LUNG FUNCTION TEST REQUEST FORM"
    Case RequestType.reqPsg
        flagNataLogo = True
        flagTsanzLogo = False
        LabName = "Sleep Disorders Laboratory"
        labPhone = "Ph:(03)9496 3508  Fax:(03)9496 5124"
        labReportName = "SLEEP STUDY REQUEST FORM"
End Select


'Draw graphic images
ImageList(1) = pdf.AddImageFromFile(pathAustinLogo)
ImageList(2) = pdf.AddImageFromFile(pathTsanzLogo)
ImageList(3) = pdf.AddImageFromFile(pathNataLogo)
ImageList(4) = pdf.AddImageFromFile(pathRcpaLogo)
pdf.DrawImage ImageList(1), M.Left - DPI(2), M.Top - DPI(1), DPI(45), DPI(11), 0
pdf.UseFont gPdfFonts(2).ID, 5
If flagTsanzLogo Then
    pdf.DrawImage ImageList(2), M.Left + DPI(58), M.Top + DPI(4), DPI(10), DPI(7), 0
    pdf.ShowTextAt M.Left + DPI(68), M.Top + DPI(7), "TSANZ Accredited"
End If
If flagNataLogo Then
    pdf.DrawImage ImageList(3), M.Left + DPI(58), M.Top + DPI(12), DPI(10), DPI(10), 0
    pdf.DrawImage ImageList(4), M.Left + DPI(72), M.Top + DPI(14), DPI(20), DPI(6), 0
    pdf.ShowTextAt M.Left + DPI(69), M.Top + DPI(21), "NATA Accreditation No 15760"
    pdf.ShowTextAt M.Left + DPI(60), M.Top + DPI(23), "Accredited for compliance with AS 4633 (ISO 15189)"
End If

'Draw lines and rectangles
pdf.SetLineWidth 0.5
pdf.DrawRectangle M.Left + DPI(107), M.Top + DPI(10), pdf.PageWidth - M.RIGHT - M.Left - DPI(107), M.Top + DPI(36), 0   'patient label
'PDF.DrawRectangle m.Left, m.Top + DPI(55), PDF.PageWidth - m.Right - m.Left, PDF.PageHeight - m.Bottom - m.Top - DPI(41), 0 'Main
pdf.Stroke

pdf.UseFont gPdfFonts(3).ID, 12
X = M.Left + DPI(110): Y = M.Top + DPI(4)
pdf.ShowTextAt X, Y, labReportName
pdf.UseFont gPdfFonts(2).ID, 9
X = M.Left + DPI(165): Y = M.Top + DPI(4)
pdf.ShowTextAt X, Y, Page
pdf.UseFont gPdfFonts(7).ID, 11
X = M.Left: Y = Y + DPI(8)
pdf.ShowTextAt X, Y, LabName
pdf.UseFont gPdfFonts(7).ID, 10
X = M.Left: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Respiratory and Sleep Medicine"
Y = Y + DPI(4)
pdf.UseFont gPdfFonts(7).ID, 7
pdf.ShowTextAt X, Y, "Studley Rd, Heidelberg, Vic 3084."
Y = Y + DPI(3.5)
pdf.ShowTextAt X, Y, labPhone
pdf.UseFont gPdfFonts(4).ID, 7
X = M.Left: x1 = X + DPI(25): x2 = X + DPI(50): x3 = X + DPI(75): Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Prof CF McDonald"
pdf.ShowTextAt x1, Y, "Dr N Antoniades"
pdf.ShowTextAt x2, Y, "Dr M Barnes"
pdf.ShowTextAt x3, Y, "Dr M Caldecott"
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr P Canty"
pdf.ShowTextAt x1, Y, "Dr KA Detering"
pdf.ShowTextAt x2, Y, "Dr N Goh"
pdf.ShowTextAt x3, Y, "Dr S Haque"
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr M Howard"
pdf.ShowTextAt x1, Y, "Dr M Ibrahim"
pdf.ShowTextAt x2, Y, "Dr C Lanteri"
pdf.ShowTextAt x3, Y, "Dr M McMahon"
Y = Y + DPI(3)
pdf.ShowTextAt X, Y, "Dr F O'Donoghue"
pdf.ShowTextAt x1, Y, "Dr M Sutherland"
pdf.ShowTextAt x2, Y, "Dr CJ Worsnop"
pdf.ShowTextAt x3, Y, ""

'Print patient demographics
X = M.Left + DPI(110): Y = M.Top + DPI(12)
pdf.UseFont gPdfFonts(7).ID, 11
pdf.ShowTextAt X, Y, "Name: " & reqData.FName & " " & reqData.Surname
Y = Y + DPI(5)
pdf.ShowTextAt X, Y, "UR:    " & reqData.UR
pdf.UseFont gPdfFonts(6).ID, 10
Y = Y + DPI(5)
pdf.ShowTextAt X, Y, "DOB: " & reqData.DOB & " (" & Age & "yrs)" & "      Gender: " & reqData.Gender
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Address: " & reqData.Address
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, Space(15) & reqData.Suburb & ", " & reqData.PCode & "."
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Home phone: " & reqData.Ph_Home
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Work phone: " & reqData.Ph_Work
Y = Y + DPI(4.5)
pdf.ShowTextAt X, Y, "Mobile phone: " & reqData.Ph_Mobile

pdf_DrawRequestHeader = M.Top + DPI(36)

'Exit Function



'Error_pdf_DrawRequestHeader:
'MsgBox "Error drawing request form header (pdf_DrawRequestHeader)" & vbCrLf & Err.Description
'Exit Function

'Resume

    On Error GoTo 0
    Exit Function

pdf_DrawRequestHeader_Error:
    Msg = "Error drawing request form header."
    Call ErrorLog(Msg, "pdf_DrawRequestHeader", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function pdf_DrawRequestSleepStudyBody(pdf As PDFCreatorPilotLib.PDFDocument4, yStart As Integer) As Integer
'Draws non-header part of PSG request form
'Returns y position of lowest pos drawn in dpi

Dim i As Integer
Dim temp, temp1
Dim X As Single, Y As Single        'Current print co-ords
Dim y1 As Single, y2 As Single, y3 As Single
Dim YN()


Dim Msg As String

    On Error GoTo pdf_DrawRequestSleepStudyBody_Error

Y = yStart + DPI(20)
y1 = Y
YN = Array("No", "Yes")

'Draw lines and rectangles
'PDF.SetLineWidth 0.5
'PDF.DrawRectangle m.Left, m.Top + DPI(55), PDF.PageWidth - m.Right - m.Left, PDF.PageHeight - m.Bottom - m.Top - DPI(41), 0 'Main
'PDF.Stroke


'Print study details
X = M.Left
pdf.UseFont gPdfFonts(3).ID, 10
pdf.ShowTextAt X, Y, "Study Details------------------------------": Y = Y + DPI(6): temp = Y
pdf.UseFont gPdfFonts(3).ID, 8
pdf.ShowTextAt X, Y, "Study Type:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Treatment Mode(s):": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Location:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Special monitoring:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Study required by:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Patient review date:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Booked study date:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Fee type:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Reason for test, other clinical notes:":  y2 = Y + DPI(6)

pdf.UseFont gPdfFonts(2).ID, 8
X = M.Left + DPI(30): Y = temp
pdf.ShowTextAt X, Y, GetListboxSelectedItem(frmRequestWizard.lstStudyType): Y = Y + DPI(4)
temp1 = GetListboxSelectedItem(frmRequestWizard.lstTreatment): If temp1 = "" Then temp1 = "Nil"
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, GetListboxSelectedItem(frmRequestWizard.lstLocation): Y = Y + DPI(4)
temp1 = GetListboxSelectedItem(frmRequestWizard.lstSpecial): If temp1 = "" Then temp1 = "Nil"
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtRequiredByDate: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtReviewDate: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtBookedDate: Y = Y + DPI(4)

temp1 = ""
If frmRequestWizard.optFee(0) Then temp1 = frmRequestWizard.optFee(0).Caption
If frmRequestWizard.optFee(1) Then temp1 = frmRequestWizard.optFee(1).Caption
If frmRequestWizard.optFee(2) Then
    temp1 = frmRequestWizard.optFee(2).Caption
    If frmRequestWizard.txtPensionHCCNo <> "" Then temp1 = temp1 & " (HCC/Pension: " & frmRequestWizard.txtPensionHCCNo & ")"
End If
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)


If frmRequestWizard.txtReasonForTest <> "" Then
    i = pdf.ShowTextLines(X + DPI(22), Y, X + DPI(80), Y + DPI(100), 10, taLeft, vaTop, frmRequestWizard.txtReasonForTest)
    y2 = pdf.GetCurrentTextY + DPI(8)
Else
    y2 = pdf.GetCurrentTextY + DPI(12)
End If

'Print referral data
X = M.Left: Y = y2
pdf.UseFont gPdfFonts(3).ID, 10
pdf.ShowTextAt X, Y, "Referral Details-----------------------------------": Y = Y + DPI(6): temp = Y
pdf.UseFont gPdfFonts(3).ID, 8
pdf.ShowTextAt X, Y, "Referring Doctor:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Provider Number:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Address:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Referral Date:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Referral Duration:": Y = Y + DPI(4)
pdf.UseFont gPdfFonts(2).ID, 8
X = X + DPI(30): Y = temp
pdf.ShowTextAt X, Y, frmRequestWizard.cmboRef_Name: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.cmboRef_Provider: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.cmboRef_Address: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtRef_RefDate: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtRef_RefDuration: Y = Y + DPI(8): y3 = Y


'Print other patient info
X = M.Left + DPI(90): Y = y1
pdf.UseFont gPdfFonts(3).ID, 10
pdf.ShowTextAt X, Y, "Other Patient Information ------------------------------------": Y = Y + DPI(6): temp = Y
pdf.UseFont gPdfFonts(3).ID, 8
pdf.ShowTextAt X, Y, "Nursing care required:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Supplemenatl O2 during the study:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Mobility Assistance:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Interpreter booked:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "VRSS sleep study:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Patient's estimated weight (kg):": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Existing Diseases:": Y = Y + DPI(4)

pdf.UseFont gPdfFonts(2).ID, 8
X = X + DPI(50): Y = temp
pdf.ShowTextAt X, Y, YN(Abs(CInt(frmRequestWizard.chkRequires(0)))): Y = Y + DPI(4)
pdf.ShowTextAt X, Y, YN(Abs(CInt(frmRequestWizard.chkRequires(1)))): Y = Y + DPI(4)
temp1 = YN(Abs(CInt(frmRequestWizard.chkRequires(2))))
If frmRequestWizard.chkRequires(2) Then
    temp1 = temp1 & " ("
    If frmRequestWizard.optMobility(0) Then temp1 = temp1 & frmRequestWizard.optMobility(0).Caption & ", "
    If frmRequestWizard.optMobility(1) Then temp1 = temp1 & frmRequestWizard.optMobility(1).Caption & ", "
    If frmRequestWizard.optMobility(2) Then temp1 = temp1 & frmRequestWizard.optMobility(2).Caption & ", "
    'PDF.ShowTextAt x, Y, "": Y = Y + DPI(4) 'Turns
    'PDF.ShowTextAt x, Y, "": Y = Y + DPI(4) 'Other
    temp1 = temp1 & ")"
End If
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)

temp1 = YN(Abs(CInt(frmRequestWizard.chkRequires(2))))
If frmRequestWizard.cmboInterpreter_language <> "" Then temp1 = temp1 & " (" & frmRequestWizard.cmboInterpreter_language & ")"
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)

If frmRequestWizard.optVRSS(0) Then temp1 = "Yes" Else temp1 = "No"
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)

pdf.ShowTextAt X, Y, frmRequestWizard.txtWeight_Estimated: Y = Y + DPI(4)

temp1 = ""
If frmRequestWizard.chkDiseases(0) Then temp1 = temp1 & frmRequestWizard.chkDiseases(0).Caption & ", "
If frmRequestWizard.chkDiseases(1) Then temp1 = temp1 & frmRequestWizard.chkDiseases(1).Caption & ", "
If frmRequestWizard.chkDiseases(2) Then temp1 = temp1 & frmRequestWizard.chkDiseases(2).Caption & ", "
If frmRequestWizard.chkDiseases(3) Then temp1 = temp1 & frmRequestWizard.chkDiseases(3).Caption & ", "
If frmRequestWizard.chkDiseases(4) Then temp1 = temp1 & frmRequestWizard.txtOtherDisease
If temp1 = "" Then
    temp1 = "No"
Else
    temp1 = "Yes (" & temp1 & ")"
End If
pdf.ShowTextAt X, Y, temp1: Y = Y + DPI(4)


'Print requesting doctor info
X = M.Left: Y = y3
pdf.UseFont gPdfFonts(3).ID, 10
pdf.ShowTextAt X, Y, "Requesting Doctor Information ------------------------------------": Y = Y + DPI(6): temp = Y
pdf.UseFont gPdfFonts(3).ID, 8
pdf.ShowTextAt X, Y, "Requested by:": y2 = Y: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Request date:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Request time:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Report to:": Y = Y + DPI(4)
pdf.ShowTextAt X, Y, "Report copy to:": Y = Y + DPI(4)

pdf.UseFont gPdfFonts(2).ID, 8
X = X + DPI(30): Y = y2
pdf.ShowTextAt X, Y, frmRequestWizard.cmboRequestingMO_Name: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtRequestDate: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.txtRequestTime: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.cmboReportTo: Y = Y + DPI(4)
pdf.ShowTextAt X, Y, frmRequestWizard.cmboCopyTo: Y = Y + DPI(4)







'Exit Function



'Error_pdf_DrawRequestSleepStudy:
'MsgBox "Error drawing body of request form  (pdf_DrawRequestSleepStudy)" & vbCrLf & Err.Description
'Exit Function

'Resume

    On Error GoTo 0
    Exit Function

pdf_DrawRequestSleepStudyBody_Error:
    Msg = "Error drawing body of request form."
    Call ErrorLog(Msg, "pdf_DrawRequestSleepStudyBody", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function
Public Sub pdf_SetFonts(pdf As PDFCreatorPilotLib.PDFDocument4)
'  Adds available fonts to the requested pdf object

gPdfFonts(1).Description = "Arial, bold, italic"
gPdfFonts(2).Description = "Arial"
gPdfFonts(3).Description = "Arial, bold"
gPdfFonts(4).Description = "Arial, italic"
gPdfFonts(5).Description = "Times New Roman, bold, italic"
gPdfFonts(6).Description = "Times New Roman"
gPdfFonts(7).Description = "Times New Roman, bold, "
gPdfFonts(8).Description = "Times New Roman, italic"

gPdfFonts(1).ID = pdf.AddFont("Arial", True, True, False, False, 0)    'bold, italic
gPdfFonts(2).ID = pdf.AddFont("Arial", False, False, False, False, 0)  'normal
gPdfFonts(3).ID = pdf.AddFont("Arial", True, False, False, False, 0)  'bold
gPdfFonts(4).ID = pdf.AddFont("Arial", False, True, False, False, 0)  'italics
gPdfFonts(5).ID = pdf.AddFont("Times New Roman", True, True, False, False, 0)    'bold, italic
gPdfFonts(6).ID = pdf.AddFont("Times New Roman", False, False, False, False, 0)  'normal
gPdfFonts(7).ID = pdf.AddFont("Times New Roman", True, False, False, False, 0)  'bold
gPdfFonts(8).ID = pdf.AddFont("Times New Roman", False, True, False, False, 0)  'italics


End Sub

Public Function DPI(mm As Single) As Single
'  Applies conversion factor to convert from mm to dpi at the current page resolution

DPI = mm * gPageResolution / 25.4


End Function
Public Function IDP(dots As Single) As Single
'  Applies conversion factor to convert from dpi to mm at the current page resolution

IDP = 25.4 * dots / gPageResolution

End Function

Public Function pdf_GetLicenseData()

gCreatorPilot.LicenseType = "Small Business License"
gCreatorPilot.RegistrationName = "neil.cook@austin.org.au"
gCreatorPilot.SerialNo = "JQP4Z-9QSPB-UPR29-DCSXT"

End Function

Public Sub pdf_SetMargins(L As Single, R As Single, T As Single, B As Single, HH As Single)
' Margins passed in mm and calculated in dpi

Dim iLeftOffsett, iTopOffsett, iRightOffsett, iBottomOffsett

'Start a print job to get a valid printer.hdc and get offsetts to printable area from the default printer
Printer.Print ""
Printer.ScaleMode = vbTwips   '(56.7 twips = 1 mm)
iLeftOffsett = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
iTopOffsett = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
iRightOffsett = iLeftOffsett
iBottomOffsett = iTopOffsett
Printer.KillDoc

'Calculate margins (dpi), adjusting for the non-printing margins, assume A4 paper
M.Left = DPI(L)
M.RIGHT = DPI(R)
M.Top = DPI(T)
M.Bottom = DPI(B)
M.HeaderHeight = DPI(HH)


End Sub
Public Function pdf_P(pdf As PDFCreatorPilotLib.PDFDocument4, txt, X As Single, Y As Single, Size As Single, Fnt As Integer, Optional SuperScript As Boolean)

'Prints text to the passed pdf object at X,Y
'Text = text to be printed; x,y are co-ords in mm
'Size = fontsize
'Fnt= font style
   
Dim Msg As String
Dim Ht As Single

On Error GoTo pdf_P_Error

If IsMissing(SuperScript) Then
    pdf.UseFont gPdfFonts(Fnt).ID, Size
    pdf.ShowTextAt M.Left + DPI(X), M.Top + DPI(Y), txt & ""
Else
    Select Case SuperScript
        Case True
            pdf.UseFont gPdfFonts(Fnt).ID, Size
            Ht = pdf.GetTextHeight(1) * 0.75
            pdf.SetTextRise Ht
            pdf.ShowTextAt M.Left + DPI(X), M.Top + DPI(Y), txt & ""
            pdf.SetTextRise -Ht
        Case False
            pdf.UseFont gPdfFonts(Fnt).ID, Size
            pdf.ShowTextAt M.Left + DPI(X), M.Top + DPI(Y), txt & ""
    End Select
End If

Exit Function

pdf_P_Error:
    Msg = ""
    Call ErrorLog(Msg, "pdf_P", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function DrawPDFReportRftReportxxx(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, yStart As Single)

Dim Tmp
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim Y As Single
Const lHt As Single = 3.6

Dim Msg As String

    On Error GoTo DrawPDFReportRftReport_Error


'setups stuff
c1 = 2: c2 = c1 + 2: c3 = c2 + 22: c4 = c3 + 30: c5 = c4 + 30: c6 = c5 + 28: c7 = c6 + 28
Fs1 = 9   'Test header
Fs2 = 8    'Results values, indice labels, units, normals
Fs3 = 6    'Equipment string
Fs4 = 10    'Report


Y = yStart

'Print NATA disclaimer - accreditation requirement from Oct 2007
pdf_P pdf, NataDisclaimer, c4 + 5, Y - 1, 7, 4
Y = Y + 8
pdf.SetLineWidth 0.5
pdf.MoveTo M.Left, DPI(Y)
pdf.DrawLineTo pdf.PageWidth - M.RIGHT, DPI(Y)
pdf.Stroke

'Technical notes section
  'print ethnicity note in bottom margin thingy
'If UCase(Rs("Race")) <> "CAUCASIAN" And UCase(Rs("Race")) <> "" And UCase(Rs("Race")) <> "UNKNOWN" Then
'    Tmp = " Predicted values corrected for " & Rs("Race") & " ethnicity."
'Else
'    Tmp = ""
'End If

Y = Y - lHt
pdf_P pdf, "TECHNICAL COMMENTS:", c1, Y, Fs4, 7
'pdf_P Pdf, "Scientist: " & Rs("Scientist"), c1 + 50, Y, Fs4, 6
pdf.UseFont gPdfFonts(6).ID, 10
'i = Pdf.ShowTextLines(m.Left + DPI(c1), DPI(Y + lHt * 2.5), Pdf.PageWidth - m.RIGHT, DPI(Y + lHt * 10), -1, taLeft, vaTop, Trim(Rs("Comments")) & Tmp)
i = IDP(pdf.GetCurrentTextY)

'Report section
  'Check if report has been verified and print message if necessary
Y = i + lHt * 2
If ReportStatus = "Verified NOT printed" Or ReportStatus = "Verified AND printed" Then
    pdf_P pdf, "REPORT:", c1, Y, Fs4, 7
Else
    pdf_P pdf, "REPORT:                     ** Unverified Report, Do Not File **", c1, Y, Fs4, 7
End If
Y = Y + lHt * 2.5
pdf.UseFont gPdfFonts(6).ID, 10
'Tmp = Trim(Rs("Report")) & vbCrLf & vbCrLf & "    Reported by:   " & Rs("Reporter") & "   " & Rs("ReportDate")
i = pdf.ShowTextLines(M.Left + DPI(c1), DPI(Y), pdf.PageWidth - M.RIGHT, pdf.PageHeight - M.Bottom, -1, taLeft, vaTop, Tmp)




    On Error GoTo 0
    Exit Function

DrawPDFReportRftReport_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportRftReport", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function


Public Function DrawPDFReportRftReport1(pdf As PDFCreatorPilotLib.PDFDocument4, R As ReportSection, yStart As Single, Optional predRef)

Dim Tmp
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim Y As Single
Const lHt As Single = 3.6
Dim Msg As String

On Error GoTo DrawPDFReportRftReport1_Error


'setup stuff
c1 = 2: c2 = c1 + 2: c3 = c2 + 22: c4 = c3 + 30: c5 = c4 + 30: c6 = c5 + 28: c7 = c6 + 28
Fs1 = 9   'Test header
Fs2 = 8    'Results values, indice labels, units, normals
Fs3 = 6    'Equipment string
Fs4 = 10    'Report

Y = yStart

'Print NATA disclaimer - accreditation requirement from Oct 2007
If IsMissing(predRef) Then
    pdf_P pdf, NataDisclaimer, c4 + 5, Y - 5, 7, 4
Else
    pdf_P pdf, "Normal values from: " & predRef & Space(15) & NataDisclaimer, c1, Y - 5, 7, 4
End If

Y = Y + 4
pdf.SetLineWidth 0.5
pdf.MoveTo M.Left, DPI(Y)
pdf.DrawLineTo pdf.PageWidth - M.RIGHT, DPI(Y)
pdf.Stroke

'Technical notes section
  'print ethnicity note in bottom margin thingy
If UCase(R.Race_Rft) <> "CAUCASIAN" And UCase(R.Race_Rft) <> "" And UCase(R.Race_Rft) <> "UNKNOWN" Then
    Tmp = " Predicted values corrected for " & R.Race_Rft & " ethnicity."
Else
    Tmp = ""
End If

Y = Y - lHt
pdf_P pdf, "TECHNICAL COMMENTS:", c1, Y, Fs4, 7
pdf_P pdf, "Scientist: " & R.Scientist, c1 + 50, Y, Fs4, 6
pdf.UseFont gPdfFonts(6).ID, 10
i = pdf.ShowTextLines(M.Left + DPI(c1), DPI(Y + lHt * 2.5), pdf.PageWidth - M.RIGHT, DPI(Y + lHt * 10), -1, taLeft, vaTop, Trim(R.Comments) & Tmp)
i = IDP(pdf.GetCurrentTextY)

'Report section, check if report has been verified and print message if necessary
Y = i + lHt * 2
If R.ReportStatus = "Verified NOT printed" Or R.ReportStatus = "Verified AND printed" Then
    pdf_P pdf, "REPORT:", c1, Y, Fs4, 7
Else
    pdf_P pdf, "REPORT:                     ** Unverified Report, Do Not File **", c1, Y, Fs4, 7
End If
Y = Y + lHt * 2.5
pdf.UseFont gPdfFonts(6).ID, 10
Tmp = Trim(R.Report) & vbCrLf & vbCrLf & "    Reported by:   " & R.Reporter & "   " & R.ReportDate
i = pdf.ShowTextLines(M.Left + DPI(c1), DPI(Y), pdf.PageWidth - M.RIGHT, pdf.PageHeight - M.Bottom, -1, taLeft, vaTop, Tmp)

Exit Function


DrawPDFReportRftReport1_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportRftReport1", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function



Public Function DrawPDFReportRftData1(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, D As PtDemographics, Y As Single)

'Declare local variables to hold demographics needed to calc predicteds
Dim locHb As Variant
Dim locDOB As Variant
Dim locTestDate As Variant
Dim locRace As String, locSex As String
Dim locHt As Single, locHt_rft As Single, locWt As Variant
Dim ycownt As Single
Dim LineHeight As Single
Dim SectionTop As Single
Dim SectionBottom As Single
Dim Tmp
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim FVLoop As FlowVolImage
Dim Filename As String
Dim Rs As New ADODB.Recordset
Dim pRefSpiro As String, pRefLV As String, pRefTLCO As String, pRefMRP As String, pRefRaw As String, predRef As String
Dim RefCount As Integer

'Declare local variables to hold mean predicted values (p???) and normal limits (norm???)
Dim normFEV1 As Variant, normFVC As Variant, normVC As Variant, normFER As Variant, normFEF As Variant
Dim normTLCO As Variant, normKCO As Variant, normVA As Variant, normFRC As Variant, normTLC As Variant
Dim normRV As Variant, normRVTLC As Variant, normShunt As Variant, normMEPs, normMIPs, normSNIP As Variant
Dim normSgaw As Variant, normRaw As Variant, normpH As Variant, normpCO2 As Variant, normpO2 As Variant
Dim pFEV1, pFVC, pVC, pFER, pFEF, pTLCO, pKCO, pVA, pFRC, pTLC, pRV, pMEPs, pMIPs

Dim Msg As String

On Error GoTo DrawPDFReportRftData1_Error

Screen.MousePointer = 11
    
'Set local demographic variables from globals
If IsDate(D.DOB) Then locDOB = CVDate(D.DOB)
If IsDate(R.E_TestDate) Then locTestDate = CVDate(R.E_TestDate)
locRace = D.Race_Rft
locSex = D.Gender
locHt_rft = Val(R.E_ThisVisit_height)
locWt = Val(R.E_ThisVisit_weight)
    
'get predicted reference
pRefSpiro = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FEV1", "MPV").Type
pRefLV = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "TLC", "MPV").Type
pRefTLCO = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "TLCO", "MPV").Type
pRefMRP = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "MEP", "MPV").Type
pRefRaw = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "Raw", "Custom").Type

'Get required predicteds
pFEV1 = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FEV1", "MPV").Value
pFVC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FVC", "MPV").Value
pVC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "VC", "MPV").Value
pFER = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FER", "MPV").Value
pFEF = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FEF", "MPV").Value
pTLCO = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "TLCO", "MPV").Value
pKCO = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "KCO", "MPV").Value
pVA = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "VA", "MPV").Value
pFRC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FRC", "MPV").Value
pTLC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "TLC", "MPV").Value
pRV = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "RV", "MPV").Value
pMEPs = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "MEP", "MPV").Value
pMIPs = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "MIP", "MPV").Value
pSNIP = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "SNIP", "MPV").Value
normFEV1 = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FEV1", "LLN").Value
normFVC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FVC", "LLN").Value
normVC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "VC", "LLN").Value
normFER = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FER", "LLN", 0).Value
normFEF = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FEF", "LLN", 1).Value
normTLCO = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "TLCO", "LLN", 1).Value
normKCO = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "KCO", "Range", 1).Value
normVA = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "VA", "LLN", 1).Value
normFRC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "FRC", "Range", 1).Value
normTLC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "TLC", "Range", 1).Value
normRV = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "RV", "Range", 1).Value
normRVTLC = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "RVTLC", "ULN", 0).Value
normSgaw = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "Sgaw", "Custom").Value
normRaw = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "Raw", "Custom").Value
normpH = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "pH", "Custom").Value
normpCO2 = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "PaCO2", "Custom").Value
normpO2 = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "PaO2", "LLN", 0).Value
normShunt = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "Shunt", "ULN", 1).Value
normBE = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "BE", "Custom").Value
normHCO3 = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "HCO3", "Custom").Value
normSaO2 = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "SaO2", "Custom").Value
normMEPs = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "MEP", "LLN", 0).Value
normMIPs = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "MIP", "LLN", 0).Value
normSNIP = GetRFTPredicted(locRace, locDOB, locTestDate, locSex, locHt_rft, locWt, "SNIP", "LLN", 0).Value

'Get results needing formatting
preFEV1 = Format(R.preFEV1, "0.00")
preFVC = Format(R.preFVC, "0.00")
preVC = Format(R.preVC, "0.00")
post1FEV1 = Format(R.postFEV1_1, "0.00")
post1FVC = Format(R.postFVC_1, "0.00")
post1VC = Format(R.postVC_1, "0.00")
post2FEV1 = Format(R.postFEV1_2, "0.00")
post2FVC = Format(R.postFVC_2, "0.00")
post2VC = Format(R.postVC_2, "0.00")
preFEF = Format(R.preFEF, "#0.0")
post1FEF = Format(R.postFEF_1, "#0.0")
post2FEF = Format(R.postFEF_2, "#0.0")
VA = Format(R.VA, "0.0")
TLCO = Format(R.TLCO, "##.0")
KCO = Format(R.KCO, "#0.0")
FRC = Format(R.FRC, "#0.00")
TLC = Format(R.TLC, "#0.00")
VC = Format(R.VC, "0.00")
RV = Format(R.RV, "0.00")
Raw = Format(R.Raw, "#0.0")
Sgaw = Format(R.Sgaw, "0.00")
Raw2 = Format(R.Raw_2, "#0.0")
Sgaw2 = Format(R.Sgaw_2, "0.00")
ph1 = Format(R.pH_1, "0.00")
pco21 = Format(R.pCO2_1, "###")
po21 = Format(R.pO2_1, "###")
Shunt1 = Format(R.Shunt_1, "#0.0")
BE1 = Format(R.BE_1, "#0.0")
HCO31 = Format(R.HCO3_1, "#0.0")
SaO21 = Format(R.SaO2_1, "###")
pH2 = Format(R.pH_2, "0.00")
pCO22 = Format(R.pCO2_2, "###")
po22 = Format(R.pO2_2, "###")
Shunt2 = Format(R.Shunt_2, "#0.0")
BE2 = Format(R.BE_2, "#0.0")
HCO32 = Format(R.HCO3_2, "#0.0")
SaO22 = Format(R.SaO2_2, "###")
mips = Format(R.MIP, "###")
meps = Format(R.MEP, "###")
SNIP = Format(R.SNIP, "###")

'Determine which tests were done
pre_done = Not (preFEV1 = "" And preFVC = "" And preVC = "")
post1_done = Not (post1FEV1 = "" And post1FVC = "" And post1VC = "")
post2_done = Not (post2FEV1 = "" And post2FVC = "" And post2VC = "")
anyspir_done = pre_done Or post1_done Or post2_done
tlco_done = Not (TLCO = "")
If IsNull(R.Hb) Then locHb = 0 Else locHb = R.Hb
hb_done = locHb > 0
vols_done = Not (FRC = "" And Raw = "" And Sgaw = "" And Raw2 = "" And Sgaw2 = "")
abgs_done = Not (ph1 = "" And pco21 = "" And po21 = "")
mips_done = Not (mips = "" And meps = "" And SNIP = "")
comments_done = Not (Commens = "")
FVCflag = preFVC <> "" Or post1FVC <> ""
VCflag = preVC <> "" Or post1VC <> ""

'setups stuff
LineHeight = 3.6
SectionTop = IDP(Y) + 2.5
FVLoop.Height = DPI(50)
FVLoop.Width = DPI(45)
FVLoop.Top = DPI(SectionTop + 40 + LineHeight * Lines)
FVLoop.Left = DPI(150)
ycownt = SectionTop
c1 = 2: c2 = c1 + 2: c3 = c2 + 22: c4 = c3 + 30: c5 = c4 + 30: c6 = c5 + 28: c7 = c6 + 28
Fs1 = 9   'Test header
Fs2 = 8    'Results values, indice labels, units, normals
Fs3 = 6    'Equipment string
Fs4 = 10    'Report

pdf_P pdf, "RESULTS:", c1, ycownt, Fs4, 7
ycownt = ycownt + LineHeight
pdf_P pdf, "Units", c3, ycownt, Fs2, 4
pdf_P pdf, "Normal", c4, ycownt, Fs2, 4
pdf_P pdf, "Range", c4, ycownt + LineHeight, Fs2, 4
ycownt = ycownt + 2 * LineHeight

If anyspir_done Then                       'print headings and normal values
    RefCount = RefCount + 1
    pdf_P pdf, "SPIROMETRY", c1, ycownt, Fs1, 3
    pdf_P pdf, "(" & RefCount & ")", 24, ycownt, Fs3, 3
    predRef = predRef & RefCount & ". " & pRefSpiro & Space(5)

    pdf_P pdf, "FEV1", c2, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, "(L BTPS)", c3, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, normFEV1, c4, (ycownt + LineHeight), Fs2, 4
    If FVCflag Then
        pdf_P pdf, "FVC", c2, (ycownt + 2 * LineHeight), Fs2, 2
        pdf_P pdf, "(L BTPS)", c3, (ycownt + 2 * LineHeight), Fs2, 2
        pdf_P pdf, normFVC, c4, (ycownt + 2 * LineHeight), Fs2, 4
    End If
    If FVCflag And VCflag Then      'ie both FVC and VC data have been entered
        Offset = LineHeight                  '   - need to add one line feed
    Else
        Offset = 0
    End If
    If VCflag Then
        pdf_P pdf, "VC", c2, (ycownt + Offset + 2 * LineHeight), Fs2, 2
        pdf_P pdf, "(L BTPS)", c3, (ycownt + Offset + 2 * LineHeight), Fs2, 2
        pdf_P pdf, normVC, c4, (ycownt + Offset + 2 * LineHeight), Fs2, 4
    End If
   
    pdf_P pdf, "FER", c2, (ycownt + Offset + 3 * LineHeight), Fs2, 2
    pdf_P pdf, "(%)", c3, (ycownt + Offset + 3 * LineHeight), Fs2, 2
    pdf_P pdf, normFER, c4, (ycownt + Offset + 3 * LineHeight), Fs2, 4
    
    If (preFEF <> "" And Val(preFEF) <> 0) Or (post1FEF <> "" And Val(post1FEF) <> 0) Then
        pdf_P pdf, "FEF25-75", c2, (ycownt + Offset + 4 * LineHeight), Fs2, 2
        pdf_P pdf, "(L/s)", c3, (ycownt + Offset + 4 * LineHeight), Fs2, 2   'print L/s a bit further on since FEF25-75 is a long word
        pdf_P pdf, normFEF, c4, (ycownt + Offset + 4 * LineHeight), Fs2, 4
    End If
End If
    
pdf_P pdf, "Baseline", c5, (ycownt - 2 * LineHeight), Fs2, 3    'Print this heading irrespective of which
pdf_P pdf, "(%mean pred)", c5, (ycownt - LineHeight), Fs2, 3      ' tests have been done

If pre_done Then
    pdf_P pdf, prepare(preFEV1, pFEV1), c5, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, prepare(preFVC, pFVC), c5, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, prepare(preVC, pVC), c5, (ycownt + Offset + 2 * LineHeight), Fs2, 2
    If (preVC = "" And preFVC = "") Or preFEV1 = "" Then
        FER = ""
    ElseIf preVC < preFVC Then
            FER = Format(100 * preFEV1 / preFVC, "###")
        Else
            FER = Format(100 * preFEV1 / preVC, "###")
    End If
    FER = prepare(FER, pFER)
    pdf_P pdf, FER, c5, (ycownt + Offset + 3 * LineHeight), Fs2, 2
    pdf_P pdf, prepare(preFEF, pFEF), c5, (ycownt + Offset + 4 * LineHeight), Fs2, 2
End If
        
temp1 = ""
Temp2 = ""
Temp3 = ""
temp4 = ""
temp5 = ""
If post1_done Or Raw2 <> "" Or Sgaw2 <> "" Then
    pdf_P pdf, R.BD_1, c6, (ycownt - 2 * LineHeight), Fs2, 3
End If

If post1_done Then
    If pre_done Then     'if baseline done then give % change
        
        pdf_P pdf, "(%change)", c6, (ycownt - LineHeight), Fs2, 3
        
        If post1FVC <> "" Then Temp2 = delta(preFVC, post1FVC)
        If post1VC <> "" Then temp5 = delta(preVC, post1VC)
        If post1FEV1 <> "" Then
            temp1 = delta(preFEV1, post1FEV1)
            If post1FVC <> "" Or post1VC <> "" Then
                If post1FVC > post1VC Then
                    Temp3 = Format(100 * post1FEV1 / post1FVC, "###")
                Else
                    Temp3 = Format(100 * post1FEV1 / post1VC, "###")
                End If
            End If
        End If
        If post1FEF <> "" And Val(post1FEF) <> 0 Then
            temp4 = delta(preFEF, post1FEF)
        End If
    Else              'if no baseline, give % predicted
        pdf_P pdf, "(%mean pred)", c6, (ycownt - LineHeight), Fs2, 3
        If post1FVC <> "" Then Temp2 = prepare(post1FVC, pFVC)
        If post1VC <> "" Then temp5 = prepare(post1VC, pVC)
        
        If post1FEV1 <> "" Then
            temp1 = prepare(post1FEV1, pFEV1)
            If post1FVC <> "" Or post1VC <> "" Then
                If post1FVC > post1VC Then
                    Temp3 = Format(100 * post1FEV1 / post1FVC, "###")
                Else
                    Temp3 = Format(100 * post1FEV1 / post1VC, "###")
                End If
            End If
            
        End If
        Temp3 = prepare(Temp3, pFER)
        temp4 = prepare(post1FEF, pFEF)
        
    End If
    pdf_P pdf, temp1, c6, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, Temp2, c6, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, temp5, c6, (ycownt + Offset + 2 * LineHeight), Fs2, 2
    pdf_P pdf, Temp3, c6, (ycownt + Offset + 3 * LineHeight), Fs2, 2
    pdf_P pdf, temp4, c6, (ycownt + Offset + 4 * LineHeight), Fs2, 2
End If
temp1 = ""
Temp2 = ""
Temp3 = ""
temp4 = ""
temp5 = ""
If post2_done Then
    pdf_P pdf, R.BD_2, c7, (ycownt - 2 * LineHeight), Fs2, 3
    pdf_P pdf, "(%change)", c7, (ycownt - LineHeight), Fs2, 3

    If post2FVC <> "" Then Temp2 = delta(post1FVC, post2FVC)
    If post2VC <> "" Then temp5 = delta(post1VC, post2VC)
    If post2FEV1 <> "" Then
        temp1 = delta(post1FEV1, post2FEV1)
        If post2FVC <> "" Or post2VC <> "" Then
            If post2FVC > post2VC Then
                Temp3 = Format(100 * post2FEV1 / post2FVC, "###")
            Else
                Temp3 = Format(100 * post2FEV1 / post2VC, "###")
            End If
        End If
    End If
    If post2FEF <> "" Then
        temp4 = delta(post1FEF, post2FEF)
    End If
    
    pdf_P pdf, temp1, c7, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, Temp2, c7, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, temp5, c7, (ycownt + Offset + 2 * LineHeight), Fs2, 2
    pdf_P pdf, Temp3, c7, (ycownt + Offset + 3 * LineHeight), Fs2, 2
    pdf_P pdf, temp4, c7, (ycownt + Offset + 4 * LineHeight), Fs2, 2
End If

If anyspir_done Then ycownt = ycownt + Offset + 4 * LineHeight        ' increment ycownt
If preFEF <> "" Or post1FEF <> "" Then ycownt = ycownt + LineHeight

ycownt = ycownt + LineHeight * 0.3

If tlco_done Then                               'Start TLCO print
    RefCount = RefCount + 1
    pdf_P pdf, "CO TRANSFER FACTOR", c1, ycownt, Fs1, 3
    pdf_P pdf, "(" & RefCount & ")", 40, ycownt, Fs3, 3
    predRef = predRef & RefCount & ". " & pRefTLCO & Space(5)
    
    pdf_P pdf, "TLCO", c2, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, "(ml/min/mmHg)", c3, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, normTLCO, c4, (ycownt + LineHeight), Fs2, 4
    If pTLCO <> "" Then
        temp$ = prepare(TLCO, pTLCO)
    Else
        temp$ = TLCO
    End If
    pdf_P pdf, temp$, c5, (ycownt + LineHeight), Fs2, 2
    
    If hb_done Then
        corrTLCO = Format(TLCO * HbFac(locHb, locDOB, locSex, locTestDate), "##.0")
        pdf_P pdf, "TLCO (Hb corr)", c2, (ycownt + 2 * LineHeight), Fs2, 2
        If pTLCO <> "" Then
            temp$ = prepare(corrTLCO, pTLCO)
        Else
            temp$ = corrTLCO
        End If
        pdf_P pdf, temp$, c5, (ycownt + 2 * LineHeight), Fs2, 2
        ycownt = ycownt + LineHeight
    End If
    
    pdf_P pdf, "KCO", c2, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, "(ml/min/mmHg/L)", c3, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, normKCO, c4, (ycownt + 2 * LineHeight), Fs2, 4
    If pKCO <> "" Then
        temp$ = prepare(KCO, pKCO)
    Else
        temp$ = KCO
    End If
    pdf_P pdf, temp$, c5, (ycownt + 2 * LineHeight), Fs2, 2
    
    If hb_done Then
        corrKCO = Format(corrTLCO / VA, "#0.0")
        pdf_P pdf, "KCO (Hb corr)", c2, (ycownt + 3 * LineHeight), Fs2, 2
        If pKCO <> "" Then
            temp$ = prepare(corrKCO, pKCO)
        Else
            temp$ = corrKCO
        End If
        pdf_P pdf, temp$, c5, (ycownt + 3 * LineHeight), Fs2, 2
        ycownt = ycownt + LineHeight
    End If

    pdf_P pdf, "VA", c2, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, "(L BTPS)", c3, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, normVA, c4, (ycownt + 3 * LineHeight), Fs2, 4
    If pVA <> "" Then
        temp$ = prepare(VA, pVA)
    Else
        temp$ = VA
    End If
    pdf_P pdf, temp$, c5, (ycownt + 3 * LineHeight), Fs2, 2
    ycownt = ycownt + 4 * LineHeight      'Increment ycownt
End If          'End TLCO print
ycownt = ycownt + LineHeight * 0.3

If vols_done Then                               'Start lung vols print
    RefCount = RefCount + 1
    pdf_P pdf, "LUNG VOLUMES", c1, ycownt, Fs1, 3
    pdf_P pdf, "(" & RefCount & ")", 28, ycownt, Fs3, 3
    predRef = predRef & RefCount & ". " & pRefLV & Space(5)

    pdf_P pdf, "FRC", c2, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, "(L BTPS)", c3, ycownt + LineHeight, Fs2, 2
    pdf_P pdf, normFRC, c4, ycownt + LineHeight, Fs2, 4
    pdf_P pdf, prepare(FRC, pFRC), c5, ycownt + LineHeight, Fs2, 2
    
    pdf_P pdf, "TLC", c2, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, "(L BTPS)", c3, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, normTLC, c4, (ycownt + 2 * LineHeight), Fs2, 4
    pdf_P pdf, prepare(TLC, pTLC), c5, (ycownt + 2 * LineHeight), Fs2, 2
    
    pdf_P pdf, "VC", c2, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, "(L BTPS)", c3, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, normVC, c4, (ycownt + 3 * LineHeight), Fs2, 4
    pdf_P pdf, prepare(VC, pVC), c5, (ycownt + 3 * LineHeight), Fs2, 2
    
    pdf_P pdf, "RV", c2, (ycownt + 4 * LineHeight), Fs2, 2
    pdf_P pdf, "(L BTPS)", c3, (ycownt + 4 * LineHeight), Fs2, 2
    pdf_P pdf, normRV, c4, (ycownt + 4 * LineHeight), Fs2, 4
    pdf_P pdf, prepare(RV, pRV), c5, (ycownt + 4 * LineHeight), Fs2, 2
    
    pdf_P pdf, "RV/TLC", c2, (ycownt + 5 * LineHeight), Fs2, 2
    pdf_P pdf, "(%)", c3, (ycownt + 5 * LineHeight), Fs2, 2
    pdf_P pdf, normRVTLC, c4, (ycownt + 5 * LineHeight), Fs2, 4
    If Val(TLC) <> 0 And Val(RV) <> 0 Then pdf_P pdf, Format(100 * RV / TLC, "##"), c5, ycownt + 5 * LineHeight, Fs2, 2
                                              
    If Raw <> "" Or Sgaw <> "" Or Raw2 <> "" Or Sgaw2 <> "" Then
        pdf_P pdf, "Raw", c2, (ycownt + 6 * LineHeight), Fs2, 2
        pdf_P pdf, "(cmH2O/L/s)", c3, (ycownt + 6 * LineHeight), Fs2, 2
        pdf_P pdf, normRaw, c4, (ycownt + 6 * LineHeight), Fs2, 4
        pdf_P pdf, Raw, c5, (ycownt + 6 * LineHeight), Fs2, 2
        pdf_P pdf, Raw2, c6, (ycownt + 6 * LineHeight), Fs2, 2

        pdf_P pdf, "Sgaw", c2, (ycownt + 7 * LineHeight), Fs2, 2
        pdf_P pdf, "(1/cmH2O/L/s/L)", c3, (ycownt + 7 * LineHeight), Fs2, 2
        pdf_P pdf, normSgaw, c4, (ycownt + 7 * LineHeight), Fs2, 4
        pdf_P pdf, Sgaw, c5, (ycownt + 7 * LineHeight), Fs2, 2
        pdf_P pdf, Sgaw2, c6, (ycownt + 7 * LineHeight), Fs2, 2
        ycownt = ycownt + 2 * LineHeight
    End If
        
    ycownt = ycownt + 6 * LineHeight      'Increment ycownt
End If          'End lung vols print

ycownt = ycownt + LineHeight * 0.3

If abgs_done Then                               'Start blood gas print
    pdf_P pdf, "ARTERIAL BLOOD GASES:", c1, ycownt, Fs1, 3
    pdf_P pdf, "FiO2", c2, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, R.FiO2_1, c5, (ycownt + LineHeight), Fs2, 2
    pdf_P pdf, "pH", c2, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, normpH, c4, (ycownt + 2 * LineHeight), Fs2, 4
    pdf_P pdf, ph1, c5, (ycownt + 2 * LineHeight), Fs2, 2
    pdf_P pdf, "PaCO2", c2, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, "(mmHg)", c3, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, normpCO2, c4, (ycownt + 3 * LineHeight), Fs2, 4
    pdf_P pdf, pco21, c5, (ycownt + 3 * LineHeight), Fs2, 2
    pdf_P pdf, "PaO2", c2, (ycownt + 4 * LineHeight), Fs2, 2
    pdf_P pdf, "(mmHg)", c3, (ycownt + 4 * LineHeight), Fs2, 2
    If UCase(R.FiO2_1) = "ROOM AIR" Or R.FiO2_1 = "R/A" Or R.FiO2_1 = "21%" Or R.FiO2_1 = "0.21" Then pdf_P pdf, normpO2, c4, (ycownt + 4 * LineHeight), Fs2, 4
    pdf_P pdf, po21, c5, (ycownt + 4 * LineHeight), Fs2, 2
    If BE1 <> "" Or BE2 <> "" Then
        goog = LineHeight
        pdf_P pdf, "BE", c2, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, "(mmol/L)", c3, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, normBE, c4, (ycownt + 4 * LineHeight + goog), Fs2, 4
        If BE1 <> "" Then pdf_P pdf, BE1, c5, (ycownt + 4 * LineHeight + goog), Fs2, 2
        If BE2 <> "" Then pdf_P pdf, BE2, c6, (ycownt + 4 * LineHeight + goog), Fs2, 2
    End If
    If HCO31 <> "" Or HCO32 <> "" Then
        goog = goog + LineHeight
        pdf_P pdf, "HCO3", c2, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, "(mmol/L)", c3, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, normHCO3, c4, (ycownt + 4 * LineHeight + goog), Fs2, 4
        If HCO31 <> "" Then pdf_P pdf, HCO31, c5, (ycownt + 4 * LineHeight + goog), Fs2, 2
        If HCO32 <> "" Then pdf_P pdf, HCO32, c6, (ycownt + 4 * LineHeight + goog), Fs2, 2
    End If
    If SaO21 <> "" Or SaO22 <> "" Then
        goog = goog + LineHeight
        pdf_P pdf, "SaO2", c2, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, "(%)", c3, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, normSaO2, c4, (ycownt + 4 * LineHeight + goog), Fs2, 4
        If SaO21 <> "" Then pdf_P pdf, SaO21, c5, (ycownt + 4 * LineHeight + goog), Fs2, 2
        If SaO22 <> "" Then pdf_P pdf, SaO22, c6, (ycownt + 4 * LineHeight + goog), Fs2, 2
    End If
    
    If R.FiO2_1 = "100%" Or R.FiO2_2 = "100%" Then
        goog = goog + LineHeight
        pdf_P pdf, "Shunt", c2, (ycownt + 4 * LineHeight + goog), Fs2, 2: pdf_P pdf, "(%)", c3, (ycownt + 4 * LineHeight + goog), Fs2, 2
        pdf_P pdf, normShunt, c4, (ycownt + 4 * LineHeight + goog), Fs2, 4
        If Shunt1 <> "" Then pdf_P pdf, Shunt1, c5, (ycownt + 4 * LineHeight + goog), Fs2, 2
        If Shunt2 <> "" Then pdf_P pdf, Shunt2, c6, (ycownt + 4 * LineHeight + goog), Fs2, 2
    End If

    If R.FiO2_2 <> "" Then
        pdf_P pdf, R.FiO2_2, c6, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, pH2, c6, (ycownt + 2 * LineHeight), Fs2, 2
        pdf_P pdf, pCO22, c6, (ycownt + 3 * LineHeight), Fs2, 2
        pdf_P pdf, po22, c6, (ycownt + 4 * LineHeight), Fs2, 2
    End If
    ycownt = ycownt + 5 * LineHeight + goog    'Increment ycownt
End If          'End ABGs print
    
ycownt = ycownt + LineHeight * 0.3
    
If mips_done Then                               'Start max respiratory pressures print
    RefCount = RefCount + 1
    pdf_P pdf, "MAXIMAL RESPIRATORY PRESSURES", c1, ycownt, Fs1, 3
    pdf_P pdf, "(" & RefCount & ")", 62, ycownt, Fs3, 3
    predRef = predRef & RefCount & ". " & pRefMRP & Space(5)
    If meps <> "" Then
        pdf_P pdf, "MEP @ TLC", c2, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, "(cmH2O)", c3, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, normMEPs, c4, (ycownt + LineHeight), Fs2, 4
        pdf_P pdf, prepare(meps, pMEPs), c5, (ycownt + LineHeight), Fs2, 2
        ycownt = ycownt + LineHeight
    End If
    If mips <> "" Then
        pdf_P pdf, prepare(mips, pMIPs), c5, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, "MIP @ FRC", c2, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, "(cmH2O)", c3, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, normMIPs, c4, (ycownt + LineHeight), Fs2, 4
        ycownt = ycownt + LineHeight
    End If
    If SNIP <> "" Then
        pdf_P pdf, prepare(SNIP, pSNIP), c5, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, "SNIP @ FRC", c2, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, "(cmH2O)", c3, (ycownt + LineHeight), Fs2, 2
        pdf_P pdf, normSNIP, c4, (ycownt + LineHeight), Fs2, 4
        ycownt = ycownt + LineHeight
    End If
    ycownt = ycownt + LineHeight
End If          'End MIPs print


'Get flow vol image from database and add to pdf
Filename = Environ("tmp") & "\Temp.jpg"
If Dir(Filename) <> "" Then Kill Filename
If Pt.FlowVol_GetFromDBToFile(Filename, R.RecordID) Then
    If DPI(ycownt) > FVLoop.Top + FVLoop.Height Then SectionBottom = DPI(ycownt) + DPI(9) Else SectionBottom = FVLoop.Top + FVLoop.Height + DPI(9)
    f = pdf.AddImageFromFile(Filename)
    pdf.DrawImage f, FVLoop.Left, FVLoop.Top, FVLoop.Height, FVLoop.Width, 0
    pdf.DrawRectangle FVLoop.Left, FVLoop.Top, FVLoop.Height, FVLoop.Width, 0
    pdf.Stroke
Else
    SectionBottom = DPI(ycownt) + DPI(9)
End If

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), IDP(SectionBottom), predRef)

Exit Function

DrawPDFReportRftData1_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportRftData1", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function DrawPDFReportSleepData(pdf As PDFCreatorPilotLib.PDFDocument4, s As SlpResults, D As PtDemographics, yS As Single, Section As SlpSection)

Dim crazy
Dim Pages As Integer
Dim Special As String
Dim dFont As Integer
Dim dFSize As Integer
Dim temp As Single
Dim temp1 As Single
Dim Temp2 As Variant
Dim Temp3
Dim sTemp As String
Dim lineWidth As Single
Dim yStart As Single
Dim Tst As Single, TSTtemp As Single, Nrem12 As Single, Nrem34 As Single
Dim NREM12_Label As String, NREM34_Label As String
Dim c0 As Single, c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single, c8 As Single, c9 As Single, c10 As Single, c11 As Single
Dim y0 As Single, y1 As Single, y2 As Single, y3 As Single, y4 As Single, y5 As Single, y6 As Single, y7 As Single, y8 As Single, y9 As Single, y10 As Single, y11 As Single
Dim X As Variant
Dim Y As Variant
Dim z As String
Dim i As Single, j As Single
Dim DOB As Variant
Dim TestDate As Variant
Dim Sex As String
Dim Ht As Single
Dim Wt As Single
Dim Race As String
Dim Report As String
Dim pdffont()
Dim yPos As Single
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim SectionTop As Single
Dim NumChars_All As Integer, NumChars_Page1 As Integer
Dim s_Page1 As String, s_Page2 As String
Const lHt As Single = 3.5     'Lineheight
Const FontMain As Integer = 1
Dim Msg As String

On Error GoTo DrawPDFReportSleepData_Error

'Setup stuff
SectionTop = IDP(yS) + 2.5
yPos = SectionTop
c1 = 2: c2 = c1 + 23: c3 = c1 + 42: y1 = 54: y2 = y1 + 4: y3 = y2
Fs1 = 9    'Test header
Fs2 = 9    'Results values, indice labels, normals
Fs3 = 6    'Equipment string
Fs4 = 10    'Report

'Special procedures section
pdf_P pdf, "SPECIAL PROCEDURES:", c1, yPos, Fs1, 3
If s.Special_TCO2 Then Special = "PtcCO2, "
If s.Special_Poes Then Special = Special & "Poes, "
If s.Special_EMGd Then Special = Special & "EMGdi, "
If s.Special_Temp Then Special = Special & "Rectal Temp, "
If s.Special_OtherProcedure <> "" Then Special = Special & s.Special_OtherProcedure & ", "
If Len(Special) > 2 Then Special = Mid(Special, 1, Len(Special) - 2) & "."       'remove the trailing comma and add a full stop
If Special = "" Then Special = "Nil."
pdf_P pdf, Special, c3, yPos, Fs2, 2

yPos = yPos + lHt * 2

If Section.SleepStats Then
    c0 = 2: c1 = 4: c2 = c1 + 27: c3 = c2 + 25: c4 = c3 + 20: c5 = c4 + 27: c6 = c5 + 32: c7 = c6 + 15: c8 = c7 + 8
    Nrem12 = Val(s.R_NREM1) + Val(s.R_NREM2)
    Nrem34 = Val(s.R_NREM3) + Val(s.R_NREM4)
    
    'Set NREM labelling
    Select Case GetScoringRuleSet(CDate(s.StudyDate), label)
        Case "Before_20-5-2002", "From_20-5-2002"
            NREM12_Label = "NREM1/2 (min):"
            NREM34_Label = "NREM3/4 (min):"
        Case "From_1-6-2011"
            NREM12_Label = "N1/N2 (min):"
            NREM34_Label = "N3 (min):"
    End Select
    
    Tst = Nrem12 + Nrem34 + Val(s.R_REMTotal)
    If Tst = 0 Then TSTtemp = 0.00001 Else TSTtemp = Tst    'Avoid divide by zero errors when tst=0
    y0 = yPos
    pdf_P pdf, "SLEEP STATISTICS:", c0, yPos, Fs1, 3
    y1 = yPos + lHt
    y2 = y1 + lHt
    y3 = y2 + lHt
    y4 = y3 + lHt
    pdf_P pdf, "N1 (min):", c3, y0, Fs2, 2
    pdf_P pdf, s.R_NREM1 & " (" & Format(100 * Val(s.R_NREM1) / TSTtemp, "0") & "%)", c4, y0, Fs2, 2
    pdf_P pdf, "Lights off/on (hrs):", c1, y1, Fs1, 2
    pdf_P pdf, Format(s.R_LOut, "hh:mm") & " / " & Format(s.R_LOn, "hh:mm"), c2, y1, Fs2, 2
    pdf_P pdf, "N2 (min): ", c3, y1, Fs2, 2
    pdf_P pdf, s.R_NREM2 & " (" & Format(100 * Val(s.R_NREM2) / TSTtemp, "0") & "%)", c4, y1, Fs2, 2
    pdf_P pdf, "Sleep Latency (min):", c5, y1, Fs2, 2
    pdf_P pdf, Format(s.R_SLat, "##0.0"), c6, y1, Fs2, 2
    pdf_P pdf, "ESS:", c7, y1, Fs2, 2
    pdf_P pdf, s.ESS, c8, y1, Fs2, 2
    
    
    pdf_P pdf, "Total dark (min):", c1, y2, Fs2, 2
    pdf_P pdf, Format(s.R_TDT, "##0.0"), c2, y2, Fs2, 2
    pdf_P pdf, NREM34_Label, c3, y2, Fs2, 2
    pdf_P pdf, Format(Nrem34, "##0.0") & " (" & Format(100 * Nrem34 / TSTtemp, "0") & "%)", c4, y2, Fs2, 2
    pdf_P pdf, "REM Latency (min):", c5, y2, Fs2, 2
    pdf_P pdf, Format(s.R_REMLat, "##0.0"), c6, y2, Fs2, 2
    pdf_P pdf, "Total sleep (min):", c1, y3, Fs2, 2
    pdf_P pdf, Format(Tst, "##0.0"), c2, y3, Fs2, 2
    pdf_P pdf, "REM (min):", c3, y3, Fs2, 2
    pdf_P pdf, Format(s.R_REMTotal, "##0.0") & " (" & Format(100 * s.R_REMTotal / TSTtemp, "0") & "%)", c4, y3, Fs2, 2
    pdf_P pdf, "Sleep Efficiency (%): ", c5, y3, Fs2, 2
    pdf_P pdf, Format(s.R_SlpEff, "##0.0"), c6, y3, Fs2, 2
    y1 = y4 + lHt
End If

If Section.Resp Or Section.Arousals Or Section.Desats Then
    y1 = y1 + lHt * 0.5
    i = 12
    c0 = 2: c1 = 4: c2 = c1 + 44: c3 = c2 + i: c4 = c3 + i: c5 = c4 + i + 8: c6 = c5 + i: c7 = c6 + i: c8 = c7 + i + 8: c9 = c8 + i: c10 = c9 + i
    pdf_P pdf, "EVENTS:", c0, y1, Fs1, 3
    pdf_P pdf, "NREM", c3, y1, Fs2, 2
    pdf_P pdf, "REM", c6, y1, Fs2, 2
    pdf_P pdf, "ALL SLEEP", c8 + 5, y1, Fs2, 2
    y1 = y1 + 4
    pdf_P pdf, "Supine", c2, y1, Fs2, 2
    pdf_P pdf, "Other", c3, y1, Fs2, 2
    pdf_P pdf, "All", c4, y1, Fs2, 2
    pdf_P pdf, "Supine", c5, y1, Fs2, 2
    pdf_P pdf, "Other", c6, y1, Fs2, 2
    pdf_P pdf, "All", c7, y1, Fs2, 2
    pdf_P pdf, "Supine", c8, y1, Fs2, 2
    pdf_P pdf, "Other", c9, y1, Fs2, 2
    pdf_P pdf, "All", c10, y1, Fs2, 2

    'Do time spent indices
    y1 = y1 + lHt
    pdf_P pdf, "Time spent (min):", c1, y1, Fs2, 2
    X = Format(s.R_TST_SupineNREM, "##0")
    pdf_P pdf, X, c2, y1, Fs2, 2
    X = Format(s.R_TST_NonSupineNREM, "##0")
    pdf_P pdf, X, c3, y1, Fs2, 2
    'X = s.R_TST_SupineNREM & ""
    'Y = s.R_TST_NonSupineNREM & ""
    'If X <> "" And Y <> "" Then
        X = Format(s.R_NREMTotal, "##0")
        pdf_P pdf, X, c4, y1, Fs2, 2
    'End If
    
    X = Format(s.R_TST_SupineREM, "##0")
    pdf_P pdf, X, c5, y1, Fs2, 2
    X = Format(s.R_TST_NonSupineREM, "##0")
    pdf_P pdf, X, c6, y1, Fs2, 2
    'X = s.R_TST_SupineREM & ""
    'Y = s.R_TST_NonSupineREM & ""
    'If X <> "" And Y <> "" Then
        X = Format(s.R_REMTotal, "##0")
        pdf_P pdf, X, c7, y1, Fs2, 2
    'End If
        
    X = Format(s.R_TST_SupineTotal, "##0")
    pdf_P pdf, X, c8, y1, Fs2, 2
    X = Format(s.R_TST_NonSupineTotal, "##0")
    pdf_P pdf, X, c9, y1, Fs2, 2
    'X = s.R_TST_SupineTotal & ""
    'Y = s.R_TST_NonSupineTotal & ""
    'If X <> "" And Y <> "" Then
        X = Format(s.R_TST, "##0")
        pdf_P pdf, X, c10, y1, Fs2, 2
    'End If
    
    If Section.Arousals Then
        y1 = y1 + lHt
        pdf_P pdf, "Arousals (/hr):", c1, y1, Fs2, FontMain
        X = Format(s.R_ArI_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, FontMain
        X = Format(s.R_ArI_NonSupNrem & "", "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, FontMain
        X = Format(s.R_ArI_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, FontMain
        X = Format(s.R_ArI_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, FontMain
        X = Format(s.R_ArI_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, FontMain
        X = Format(s.R_ArI_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, FontMain
        X = Format(s.R_ArI_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, FontMain
        X = Format(s.R_ArI_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, FontMain
        X = Format(s.R_ArI_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, FontMain
        
    End If
    
    If Section.Desats Then
    
        y1 = y1 + lHt
        pdf_P pdf, "ODI4% (/hr): ", c1, y1, Fs2, FontMain
        X = Format(s.R_ODI4_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, FontMain
        X = Format(s.R_ODI4_NonSupNrem & "", "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, FontMain
        X = Format(s.R_ODI4_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, FontMain
        X = Format(s.R_ODI4_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, FontMain
        X = Format(s.R_ODI4_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, FontMain
        X = Format(s.R_ODI4_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, FontMain
        X = Format(s.R_ODI4_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, FontMain
        X = Format(s.R_ODI4_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, FontMain
        X = Format(s.R_ODI4_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, FontMain
    
    End If
    
    If Section.Resp Then
    
        y1 = y1 + lHt
        pdf_P pdf, "Hypopnoeas (/hr):", c1, y1, Fs2, 2
        X = Format(s.R_HI_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, 2
        X = Format(s.R_HI_NonSupNrem, "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, 2
        X = Format(s.R_HI_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, 2
        X = Format(s.R_HI_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, 2
        X = Format(s.R_HI_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, 2
        X = Format(s.R_HI_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, 2
        X = Format(s.R_HI_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, 2
        X = Format(s.R_HI_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, 2
        X = Format(s.R_HI_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, 2

        y1 = y1 + lHt
        pdf_P pdf, "Obs Apnoeas (/hr):", c1, y1, Fs2, 2
        X = Format(s.R_OAI_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, 2
        X = Format(s.R_OAI_NonSupNrem, "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, 2
        X = Format(s.R_OAI_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, 2
        X = Format(s.R_OAI_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, 2
        X = Format(s.R_OAI_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, 2
        X = Format(s.R_OAI_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, 2
        X = Format(s.R_OAI_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, 2
        X = Format(s.R_OAI_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, 2
        X = Format(s.R_OAI_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, 2
        
        y1 = y1 + lHt
        pdf_P pdf, "Mixed Apnoeas (/hr):", c1, y1, Fs2, 2
        X = Format(s.R_MAI_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, 2
        X = Format(s.R_MAI_NonSupNrem, "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, 2
        X = Format(s.R_MAI_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, 2
        X = Format(s.R_MAI_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, 2
        X = Format(s.R_MAI_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, 2
        X = Format(s.R_MAI_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, 2
        X = Format(s.R_MAI_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, 2
        X = Format(s.R_MAI_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, 2
        X = Format(s.R_MAI_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, 2
        
        y1 = y1 + lHt
        pdf_P pdf, "Central Apnoeas (/hr):", c1, y1, Fs2, 2
        X = Format(s.R_CAI_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, 2
        X = Format(s.R_CAI_NonSupNrem, "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, 2
        X = Format(s.R_CAI_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, 2
        X = Format(s.R_CAI_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, 2
        X = Format(s.R_CAI_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, 2
        X = Format(s.R_CAI_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, 2
        X = Format(s.R_CAI_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, 2
        X = Format(s.R_CAI_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, 2
        X = Format(s.R_CAI_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, 2
        
        y1 = y1 + lHt
        pdf_P pdf, "AHI (/hr):", c1, y1, Fs2, FontMain
        X = Format(s.R_AHI_SupNrem, "##0.0")
        pdf_P pdf, X, c2, y1, Fs2, FontMain
        X = Format(s.R_AHI_NonSupNrem, "##0.0")
        pdf_P pdf, X, c3, y1, Fs2, FontMain
        X = Format(s.R_AHI_AllNrem, "##0.0")
        pdf_P pdf, X, c4, y1, Fs2, FontMain
        X = Format(s.R_AHI_SupRem, "##0.0")
        pdf_P pdf, X, c5, y1, Fs2, FontMain
        X = Format(s.R_AHI_NonSupRem, "##0.0")
        pdf_P pdf, X, c6, y1, Fs2, FontMain
        X = Format(s.R_AHI_AllRem, "##0.0")
        pdf_P pdf, X, c7, y1, Fs2, FontMain
        X = Format(s.R_AHI_SupAll, "##0.0")
        pdf_P pdf, X, c8, y1, Fs2, FontMain
        X = Format(s.R_AHI_NonSupAll, "##0.0")
        pdf_P pdf, X, c9, y1, Fs2, FontMain
        X = Format(s.R_AHI_AllAll, "##0.0")
        pdf_P pdf, X, c10, y1, Fs2, FontMain
        
        If GetScoringRuleSet(CDate(s.StudyDate), label) = "From_1-6-2011" Then

            y1 = y1 + lHt
            pdf_P pdf, "RERAs (/hr):", c1, y1, Fs2, 2
            X = Format(s.R_RERA_SupNrem, "##0.0")
            pdf_P pdf, X, c2, y1, Fs2, 2
            X = Format(s.R_RERA_NonSupNrem, "##0.0")
            pdf_P pdf, X, c3, y1, Fs2, 2
            X = Format(s.R_RERA_AllNrem, "##0.0")
            pdf_P pdf, X, c4, y1, Fs2, 2
            X = Format(s.R_RERA_SupRem, "##0.0")
            pdf_P pdf, X, c5, y1, Fs2, 2
            X = Format(s.R_RERA_NonSupRem, "##0.0")
            pdf_P pdf, X, c6, y1, Fs2, 2
            X = Format(s.R_RERA_AllRem, "##0.0")
            pdf_P pdf, X, c7, y1, Fs2, 2
            X = Format(s.R_RERA_SupAll, "##0.0")
            pdf_P pdf, X, c8, y1, Fs2, 2
            X = Format(s.R_RERA_NonSupAll, "##0.0")
            pdf_P pdf, X, c9, y1, Fs2, 2
            X = Format(s.R_RERA_AllAll, "##0.0")
            pdf_P pdf, X, c10, y1, Fs2, 2
           
            y1 = y1 + lHt
            pdf_P pdf, "RDI (/hr):", c1, y1, Fs2, FontMain
            X = Format(s.R_RDI_SupNrem, "##0.0")
            pdf_P pdf, X, c2, y1, Fs2, FontMain
            X = Format(s.R_RDI_NonSupNrem, "##0.0")
            pdf_P pdf, X, c3, y1, Fs2, FontMain
            X = Format(s.R_RDI_AllNrem, "##0.0")
            pdf_P pdf, X, c4, y1, Fs2, FontMain
            X = Format(s.R_RDI_SupRem, "##0.0")
            pdf_P pdf, X, c5, y1, Fs2, FontMain
            X = Format(s.R_RDI_NonSupRem, "##0.0")
            pdf_P pdf, X, c6, y1, Fs2, FontMain
            X = Format(s.R_RDI_AllRem, "##0.0")
            pdf_P pdf, X, c7, y1, Fs2, FontMain
            X = Format(s.R_RDI_SupAll, "##0.0")
            pdf_P pdf, X, c8, y1, Fs2, FontMain
            X = Format(s.R_RDI_NonSupAll, "##0.0")
            pdf_P pdf, X, c9, y1, Fs2, FontMain
            X = Format(s.R_RDI_AllAll, "##0.0")
            pdf_P pdf, X, c10, y1, Fs2, FontMain
        End If
    End If
    y1 = y1 + lHt
End If

If Section.SaO2 Then
    y1 = y1 + lHt * 0.5
    c1 = 4: c2 = 31: c3 = 52: c4 = 71: c5 = 92: c6 = 111: c7 = 132: c8 = 151
    pdf_P pdf, "SpO2 STATISTICS:", 2, y1, Fs1, 3
    y1 = y1 + lHt
    y2 = y1 + lHt
    pdf_P pdf, "Baseline Wake:", c1, y1, Fs2, 2
    pdf_P pdf, s.R_SpO2_BLAwake & "%", c2, y1, Fs2, 2
    pdf_P pdf, "Min. REM:   ", c3, y1, Fs2, 2
    pdf_P pdf, s.R_SpO2_MinREM & "%", c4, y1, Fs2, 2
    pdf_P pdf, "%TST<95:", c5, y1, Fs2, 2
    pdf_P pdf, s.R_SpO2_Time95 & "%", c6, y1, Fs2, 2
    pdf_P pdf, "%TST<88:", c7, y1, Fs2, 2
    pdf_P pdf, s.R_SpO2_Time88 & "%", c8, y1, Fs2, 2
    pdf_P pdf, "Baseline Asleep:", c1, y2, Fs2, 2
    pdf_P pdf, s.R_SpO2_BLAsleep & "%", c2, y2, Fs2, 2
    pdf_P pdf, "Min. NREM:   ", c3, y2, Fs2, 2
    pdf_P pdf, s.R_SpO2_MinNREM & "%", c4, y2, Fs2, 2
    pdf_P pdf, "%TST<90:", c5, y2, Fs2, 2
    pdf_P pdf, s.R_SpO2_Time90 & "%", c6, y2, Fs2, 2
    pdf_P pdf, "%TST<85:", c7, y2, Fs2, 2
    pdf_P pdf, s.R_SpO2_Time85 & "%", c8, y2, Fs2, 2
    y1 = y2
End If

If Section.Abg Then
    'Print normal values
    If IsDate(s.D_DOB) Then DOB = s.D_DOB
    If IsDate(s.StudyDate) Then TestDate = s.StudyDate
    Sex = s.D_Gender
    Race = ""
    Ht = Val(s.Height_slp)
    Wt = Val(s.Weight)
    pH = GetRFTPredicted(Race, DOB, TestDate, Sex, Ht, Wt, "pH", "Custom").Value
    PaCO2 = GetRFTPredicted(Race, DOB, TestDate, Sex, Ht, Wt, "PaCO2", "Custom").Value
    PaO2 = GetRFTPredicted(Race, DOB, TestDate, Sex, Ht, Wt, "PaO2", "LLN", "0").Value
    HCO3 = GetRFTPredicted(Race, DOB, TestDate, Sex, Ht, Wt, "HCO3", "Custom").Value
    BE = GetRFTPredicted(Race, DOB, TestDate, Sex, Ht, Wt, "BE", "Custom").Value
    SaO2 = GetRFTPredicted(Race, DOB, TestDate, Sex, Ht, Wt, "SaO2", "Custom").Value
    
    y1 = y1 + lHt * 1.5
    i = 14
    c1 = 10: c2 = c1 + i: c3 = c2 + i + 4: c4 = c3 + i: c5 = c4 + i: c6 = c5 + i: c7 = c6 + i: c8 = c7 + i: c9 = c8 + i: c10 = c9 + i: c11 = c10 + i
    pdf_P pdf, "ARTERIAL BLOOD GASES:", 2, y1, Fs1, 3
    y1 = y1 + lHt
    y2 = y1 + lHt
    y3 = y2 + lHt
    y4 = y3 + lHt
    y5 = y4 + lHt
    'First row
    pdf_P pdf, "Time", c1, y1, Fs2, 2
    pdf_P pdf, "FiO2", c2, y1, Fs2, 2
    pdf_P pdf, "pH", c3, y1, Fs2, 2
    pdf_P pdf, "PaCO2", c4 - 3, y1, Fs2, 2
    pdf_P pdf, "PaO2", c5 - 2, y1, Fs2, 2
    pdf_P pdf, "SaO2", c6 - 3, y1, Fs2, 2
    pdf_P pdf, "HCO3", c7 - 3, y1, Fs2, 2
    pdf_P pdf, "BE", c8 - 0, y1, Fs2, 2
    pdf_P pdf, "PtcCO2", c9 - 3, y1, Fs2, 2
    pdf_P pdf, "tc-aPCO2", c10 - 3, y1, Fs2, 2
    'Second row
    pdf_P pdf, "(mmHg)", c4 - 3, y2, Fs2, 2
    pdf_P pdf, "(mmHg)", c5 - 3, y2, Fs2, 2
    pdf_P pdf, "(%)", c6 - 1, y2, Fs2, 2
    pdf_P pdf, "(mmol/L)", c7 - 4, y2, Fs2, 2
    pdf_P pdf, "(mmol/L)", c8 - 3, y2, Fs2, 2
    pdf_P pdf, "(mmHg)", c9 - 3, y2, Fs2, 2
    pdf_P pdf, "(mmHg)", c10 - 2, y2, Fs2, 2
    'Third row
    pdf_P pdf, s.R_ABG_Time1, c1, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_FiO21, c2, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_pH1, c3, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_PaCO21, c4, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_PaO21, c5, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_SaO21, c6, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_HCO31, c7, y3, Fs2, 2
    pdf_P pdf, Format(s.R_ABG_BE1, "+0.0;-0.0;0.0"), c8, y3, Fs2, 2
    pdf_P pdf, s.R_ABG_PtcCO21, c9, y3, Fs2, 2
    If (s.R_ABG_PtcCO21 <> "" And s.R_ABG_PaCO21 <> "") Then
        pdf_P pdf, Format(Val(s.R_ABG_PtcCO21) - Val(s.R_ABG_PaCO21), "+0.0;-0.0;0.0"), c10, y3, Fs2, 2
    End If
    'Fourth row
    pdf_P pdf, s.R_ABG_Time2, c1, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_FiO22, c2, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_pH2, c3, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_PaCO22, c4, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_PaO22, c5, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_SaO22, c6, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_HCO32, c7, y4, Fs2, 2
    pdf_P pdf, Format(s.R_ABG_BE2, "+0.0;-0.0;0.0"), c8, y4, Fs2, 2
    pdf_P pdf, s.R_ABG_PtcCO22, c9, y4, Fs2, 2
    If (s.R_ABG_PtcCO22 <> "" And s.R_ABG_PaCO22 <> "") Then
        pdf_P pdf, Format(Val(s.R_ABG_PtcCO22) - Val(s.R_ABG_PaCO22), "+0.0;-0.0;0.0"), c10, y4, Fs2, 2
    End If
    'Fifth row
    pdf_P pdf, "Normal Range", c1, y5, Fs2, 4
    pdf_P pdf, pH, c3 - 5, y5, Fs2, 4
    pdf_P pdf, PaCO2, c4 - 3, y5, Fs2, 4
    pdf_P pdf, PaO2, c5 - 2, y5, Fs2, 4
    pdf_P pdf, SaO2, c6 - 3, y5, Fs2, 4
    pdf_P pdf, HCO3, c7 - 3, y5, Fs2, 4
    pdf_P pdf, BE, c8 - 3, y5, Fs2, 4
    y1 = y5
End If

'Print NATA disclaimer - accreditation requirement from Oct 2007
y1 = y1 + lHt * 1.5
If y1 < yPos Then y1 = yPos
pdf_P pdf, NataDisclaimer, 10, y1, 7, 4
'Add scoring rules note
y1 = y1 + lHt
pdf_P pdf, "Scoring Criteria: " & GetScoringRuleSet(CDate(s.StudyDate), Description), 10, y1, 7, 4

'Draw divider line
pdf.MoveTo M.Left, DPI(y1 + lHt * 2.5)
pdf.DrawLineTo pdf.PageWidth - M.RIGHT, DPI(y1 + lHt * 2.5)
pdf.Stroke

'Print report
y1 = y1 + lHt * 2
If s.ReportStatus = "Verified NOT printed" Or s.ReportStatus = "Verified AND printed" Then
    pdf_P pdf, "REPORT:", 2, y1, Fs2, 3
Else
    pdf_P pdf, "REPORT:                     ** Unverified Report, Do Not File **", 2, y1, Fs2, 3
End If

'Build report string
'  Remove any trailing crlf's from tech note
z = "STUDY NOTES:  " & s.TechnicalNote
Do While RIGHT(z, 2) = vbCrLf
    z = Left(z, Len(z) - 2)
Loop
If s.ScoredBy <> "" Then Temp2 = "   Scored by: " & s.ScoredBy
z = z & Temp2 & vbCrLf & vbCrLf

'Add report
'  Remove any trailing crlf's from report
Temp2 = s.Report
Do While RIGHT(Temp2, 2) = vbCrLf
    Temp2 = Left(Temp2, Len(Temp2) - 2)
Loop
z = z & Temp2 & vbCrLf & vbCrLf
'Add reporter
z = z & Space(20) & "Reported by: " & s.Reporter_MO & ",  " & Format(s.ReportDate, "d/mm/yyyy") & "." & vbCrLf

'Check for report running onto page 2
y1 = y1 + lHt * 3
pdf.UseFont gPdfFonts(6).ID, 10

'Print chars available on page 1
NumChars_Page1 = pdf.ShowTextLines(M.Left + DPI(2), DPI(y1), pdf.PageWidth - M.RIGHT, pdf.PageHeight - M.Bottom, -1, taLeft, vaTop, z)

If Len(z) > NumChars_Page1 Then    'Report carries over
    pdf_P pdf, "Continued over ...", 150, IDP(pdf.PageHeight - M.Bottom) - 4, Fs2, 4
    
    'Do page 2 bit
    pdf.NewPage
    Call pdf_DrawReportHeader1(pdf, s.Request_HealthServiceID, CDate(s.StudyDate), D, LabType.ltSleep, 2)
    i = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromSlpUDT(s), LabType.ltSleep)
    y1 = IDP(i) + lHt
    If s.ReportStatus = "Verified NOT printed" Or s.ReportStatus = "Verified AND printed" Then
        pdf_P pdf, "REPORT: Continued ...", 2, y1, Fs2, 3
    Else
        pdf_P pdf, "REPORT: Continued ...                     ** Unverified Report, Do Not File **", 2, y1, Fs2, 3
    End If
    y1 = y1 + lHt * 3

    'Print last bit of report (remove leading spaces and carriage returns)
    s_Page2 = Trim((Mid(z, NumChars_Page1 + 1)))
    Do While Left(s_Page2, 2) = vbCrLf
         s_Page2 = RIGHT(s_Page2, Len(s_Page2) - 2)
    Loop
    pdf.UseFont gPdfFonts(6).ID, 10
    i = pdf.ShowTextLines(M.Left + DPI(2), DPI(y1), pdf.PageWidth - M.RIGHT, pdf.PageHeight - M.Bottom, -1, taLeft, vaTop, s_Page2)
    pdf_P pdf, "Continued over ...", 150, IDP(pdf.PageHeight - M.Bottom) - 4, Fs2, 4
    y1 = y1 + i * lHt
    'Draw divider line
    pdf.MoveTo M.Left, DPI(y1 + lHt * 1.5)
    pdf.DrawLineTo pdf.PageWidth - M.RIGHT, DPI(y1 + lHt * 1.5)
    pdf.Stroke
Else
    i = pdf.ShowTextLines(M.Left + DPI(2), DPI(y1), pdf.PageWidth - M.RIGHT, pdf.PageHeight - M.Bottom, -1, taLeft, vaTop, z)
End If

'Do graphics page
pdf.NewPage
Call pdf_DrawReportHeader1(pdf, s.Request_HealthServiceID, CDate(s.StudyDate), D, LabType.ltSleep, 2)
i = pdf_DrawReportInfo1(pdf, CurrentPt.Get_ReportInfoFromSlpUDT(s), LabType.ltSleep)
y1 = SectionTop

pdf_P pdf, "GRAPHICAL SUMMARY:", 2, y1, Fs2, 3



'Get the graphic summary
If s.R_Graphic Then
    Dim TempDir As String
    
    TempDir = Environ("Tmp")
    
    
    Fname_emf = "temp.emf"  'nOTE: BUG IN pRINT2EdOC - FILENAME CREATED IS tEMP1.EMF
    Fname_rtf = CurrentPt.Get_SlpGraphicFromDB(s.RecordID)
 
    If Dir$(TempDir & "\" & Fname_emf) <> "" Then Kill TempDir & "\" & Fname_emf
    If Dir$(TempDir & "\temp.emf") <> "" Then Kill TempDir & "\temp.emf"
    If Dir$(TempDir & "\temp1.emf") <> "" Then Kill TempDir & "\temp1.emf"
    
    Dim PrinterObj As New gtPrint2eDoc
    PrinterObj.ActivateLicense "6A536F5477486943764570473473384C616D2F4773737257"
    PrinterObj.SetAsDefaultPrinter
   
    PrinterObj.KeepOfficeApplicationsOpen = False
    PrinterObj.UseCustomDocumentName = True
    PrinterObj.ApplyOfficeAddinSettings = False
    PrinterObj.DefaultOutputDirectory = TempDir
    PrinterObj.OutputDocumentName = Fname_emf
    PrinterObj.OutputDocumentFormat = TxgtDocumentFormat.EMF
    PrinterObj.ShowPreview = False
    PrinterObj.ShowSaveDialog = False
    PrinterObj.ShowSettingsDialog = False
    PrinterObj.ViewGeneratedDocuments = False
    PrinterObj.PrintDocument (Fname_rtf)
    PrinterObj.RestoreDefaultPrinter
    
    '***Need to delay here until emf file generation finished before
    i = Timer
    Do
        DoEvents
    Loop Until Timer > i + 3
    If Dir(TempDir & "\temp1.emf") <> "" Then
        pdf.PlayMetaFile TempDir & "\temp1.emf", DPI(10), DPI(y1 - lHt * 2), 0.8, 0.7
    Else
        MsgBox "Graphics file not found" & vbCrLf & Environ("tmp") & "\temp1.emf" & vbCrLf & "Login: " & GetLogonID, vbOKOnly, "Problem generating Pdf" & vbCrLf
    End If
End If

Set PrinterObj = Nothing

Screen.MousePointer = 0

Exit Function


z = pdf.AddImageFromFile("C:\Documents and Settings\rochpd\Desktop\SleepGraphic.bmp")
pdf.ShowImage z, 1, 1
pdf.SaveToFile Environ("tmp") & "\pdr.pdf", True

DrawPDFReportSleepData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportSleepData", "mPdf", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Sub LoadGlobalsChallenge(Rs As ADODB.Recordset)

Dim Msg As String

    On Error GoTo LoadGlobalsChallenge_Error

'This loads provocation data only from the database

Lab = Rs("Campus")
Source = Rs("Source")
AdmissionStatus = Rs("AdmissionStatus")
BillingItems = Rs("BillingItems")
BillingMO = Rs("BillingMO")
BillingStatus = Rs("BillingStatus")
ReferringMO = Rs("mo")
RefLocation = Rs("RefLocation")
ReportStatus = Rs("ReportStatus")
TestType = Rs("TestType")
TestDate = Rs("TestDate")
TestTime = Rs("TestTime")
Scientist = Rs("scientist")
Commens = Rs("comments")
ClinicalNote = Rs("clinicalnote")
Smoke = Rs("Smoke")
Hb = Rs("Hb")
HbInfo = Rs("HbInfo")
Packyrs = Rs("Packyrs")
Weight = Rs("Weight")
'*** added for Ht at test
hight_RFT = Rs("Height_rft")

spir = Rs("spir")
LastBd = Rs("LastBd")
bdtype = Rs("PostProvBDType")
If bdtype = "" Then bdtype = "BD"
provFEV1 = Format(Rs("prefev1"), "0.00")
provFVC = Format(Rs("prefvc"), "0.00")
provVC = Format(Rs("prevc"), "0.00")

Report = Rs("Report")
Reporter = Rs("Reporter")
ReportDate = Rs("ReportDate")

    'Load Challenge data
tempProvDrug(1) = "Baseline"
tempProvDrug(2) = Rs("control")
tempProvDrug(3) = Rs("Dose1")        'Challenge drug doses
tempProvDrug(4) = Rs("Dose2")
tempProvDrug(5) = Rs("Dose3")
tempProvDrug(6) = Rs("Dose4")
tempProvDrug(7) = Rs("Dose5")
tempProvDrug(8) = Rs("Dose6")
tempProvDrug(9) = Rs("Dose7")
tempProvDrug(10) = Rs("Dose8")
tempProvDrug(11) = Rs("Dose9")
tempProvDrug(12) = Rs("post")

chall(1) = Format(provFEV1, "0.00")
chall(2) = Format(Rs("FEV1-Control"), "0.00")
chall(3) = Format(Rs("FEV1-1"), "0.00")
chall(4) = Format(Rs("FEV1-2"), "0.00")
chall(5) = Format(Rs("FEV1-3"), "0.00")
chall(6) = Format(Rs("FEV1-4"), "0.00")
chall(7) = Format(Rs("FEV1-5"), "0.00")
chall(8) = Format(Rs("FEV1-6"), "0.00")
chall(9) = Format(Rs("FEV1-7"), "0.00")
chall(10) = Format(Rs("FEV1-8"), "0.00")
chall(11) = Format(Rs("FEV1-9"), "0.00")
chall(12) = Format(Rs("FEV1-post"), "0.00")



    On Error GoTo 0
    Exit Sub

LoadGlobalsChallenge_Error:
    Msg = ""
    Call ErrorLog(Msg, "LoadGlobalsChallenge", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Sub
    Resume

End Sub
Public Function DrawPDFReport_BronchChallengeData(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, Y As Single) As Boolean

'Declare local variables to hold demographics needed to calc predicteds
Dim SectionTop As Single
Dim SectionBottom As Single
Dim Tmp
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim X As Single, prevX As Single, prevY As Single
Dim xMin As Single, xMax As Single, yMin As Single, yMax As Single, xSize As Single, ySize As Single
Dim startpos As Single
Dim blFEV1 As Single, provFEV1 As Single, bdFEV1 As Single
Dim blFVC As Single
Dim blVC As Single
Dim blFER As Single
Dim LastDose As String
Dim LastFEV1 As Single
Dim Drug As ProvocationTestType
Dim DrugString As String
Dim FER As Single
Dim Fev1ToPlot As Single
Dim FEV1(1 To 9) As Single
Dim FEV1pre As Single
Dim FEV1_Control As Single
Dim FEV1_Post As Single
Dim Dose(1 To 9) As String
Dim Dose_pre As String
Dim Dose_control As String
Dim Dose_post As String

Const lHt As Single = 3.6   'Line height in mm

'Declare local variables to hold mean predicted values (p???) and normal limits (norm???)
Dim normFEV1 As Variant, normFVC As Variant, normVC As Variant, normFER As Variant
Dim pFEV1 As Variant, pFVC As Variant, pVC As Variant, pFER As Variant
Dim Msg As String

On Error GoTo DrawPDFReport_BronchChallengeData_Error

Screen.MousePointer = 11

'Get required predicteds
pFEV1 = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "FEV1", "MPV").Value
pFVC = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "FVC", "MPV").Value
pVC = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "VC", "MPV").Value
pFER = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "FER", "MPV").Value
normFEV1 = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "FEV1", "LLN").Value
normFVC = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "FVC", "LLN").Value
normVC = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "VC", "LLN").Value
normFER = GetRFTPredicted(R.D_Race_Rft, R.D_DOB, R.E_TestDate, R.D_Gender, R.E_ThisVisit_height, R.E_ThisVisit_weight, "FER", "LLN", 0).Value

'setup stuff
Y = IDP(Y) + 2.5
c1 = 10: c2 = c1 + 2: c3 = c2 + 22: c4 = c3 + 30: c5 = c4 + 25: c6 = c5 + 28: c7 = c6 + 28
Fs1 = 9   'Test header
Fs2 = 8    'Results values, indice labels, units, normals
Fs3 = 7    'Equipment string
Fs4 = 10    'Report

'Fill array with data
FEV1pre = Val(R.prov_FEV1pre)
FEV1_Control = Val(R.prov_FEV1_Control)
FEV1(1) = Val(R.prov_FEV1_Dose1)
FEV1(2) = Val(R.prov_FEV1_Dose2)
FEV1(3) = Val(R.prov_FEV1_Dose3)
FEV1(4) = Val(R.prov_FEV1_Dose4)
FEV1(5) = Val(R.prov_FEV1_Dose5)
FEV1(6) = Val(R.prov_FEV1_Dose6)
FEV1(7) = Val(R.prov_FEV1_Dose7)
FEV1(8) = Val(R.prov_FEV1_Dose8)
FEV1(9) = Val(R.prov_FEV1_Dose9)
FEV1_Post = Val(R.prov_FEV1_Post)
Dose_pre = "B/L"
Dose_control = R.prov_Control
Dose(1) = R.prov_Dose1
Dose(2) = R.prov_Dose2
Dose(3) = R.prov_Dose3
Dose(4) = R.prov_Dose4
Dose(5) = R.prov_Dose5
Dose(6) = R.prov_Dose6
Dose(7) = R.prov_Dose7
Dose(8) = R.prov_Dose8
Dose(9) = R.prov_Dose9
Dose_post = R.prov_Post

Drug = Get_ProvocationDrug(R.TestType)

'Get last dose and last fev1
For i = 9 To 1 Step -1
    If FEV1(i) <> 0 Then
        LastFEV1 = FEV1(i)
        LastDose = Dose(i)
        Exit For
    End If
Next i

If Drug = Mannitol Then
    pdf_P pdf, "MANNITOL BRONCHO-PROVOCATION CHALLENGE (Aridol method)", 2, Y, Fs1, 3
    DrugString = "Mannitol"
ElseIf Drug = Methacholine Then
    pdf_P pdf, "METHACHOLINE BRONCHO-PROVOCATION CHALLENGE (Yan method)", 2, Y, Fs1, 3
    DrugString = "Methacholine"
ElseIf Drug = Histamine Then
    pdf_P pdf, "HISTAMINE BRONCHO-PROVOCATION CHALLENGE (Yan method)", 2, Y, Fs1, 3
    DrugString = "Histamine"
End If

Y = Y + lHt * 2
pdf_P pdf, "Units", c3, Y, Fs2, 4
pdf_P pdf, "Normal", c4, Y, Fs2, 4
pdf_P pdf, "Range", c4, Y + lHt, Fs2, 4
pdf_P pdf, "Baseline", c5, Y, Fs2, 3
pdf_P pdf, "(%mean pred)", c5, (Y + lHt), Fs2, 3
pdf_P pdf, "Post", c6, Y, Fs2, 3
pdf_P pdf, DrugString, c6, (Y + lHt), Fs2, 3
pdf_P pdf, "Post", c7, Y, Fs2, 3
pdf_P pdf, Mid(R.prov_PostBDType, 6), c7, (Y + lHt), Fs2, 3

Y = Y + 2 * lHt
pdf_P pdf, "SPIROMETRY:", c1, Y, Fs2, 3
pdf_P pdf, R.spir, c1 + 25, Y + lHt * 0.1, Fs3, 2
Y = Y + lHt
pdf_P pdf, "FEV1", c2, Y, Fs2, 2
pdf_P pdf, "(L BTPS)", c3, Y, Fs2, 2
pdf_P pdf, normFEV1, c4, Y, Fs2, 4
pdf_P pdf, prepare(R.prov_FEV1pre, pFEV1), c5, Y, Fs2, 2
pdf_P pdf, LastFEV1, c6, Y, Fs2, 2
pdf_P pdf, R.prov_FEV1_Post, c7, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "FVC", c2, Y, Fs2, 2
pdf_P pdf, "(L BTPS)", c3, Y, Fs2, 2
pdf_P pdf, normFVC, c4, Y, Fs2, 4
pdf_P pdf, prepare(R.prov_FVCpre, pFVC), c5, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "VC", c2, Y, Fs2, 2
pdf_P pdf, "(L BTPS)", c3, Y, Fs2, 2
pdf_P pdf, normVC, c4, Y, Fs2, 4
pdf_P pdf, prepare(R.prov_VCpre, pVC), c5, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "FER", c2, Y, Fs2, 2
pdf_P pdf, "(%)", c3, Y, Fs2, 2
pdf_P pdf, normFER, c4, Y, Fs2, 4
If (Val(R.prov_VCpre) = 0 And Val(R.prov_FVCpre) = 0) Or Val(R.prov_FEV1pre) = 0 Then
    FER = 0
ElseIf Val(R.prov_VCpre) < Val(R.prov_FVCpre) Then
    FER = Format(100 * Val(R.prov_FEV1pre) / Val(R.prov_FVCpre), "###")
Else
    FER = Format(100 * Val(R.prov_FEV1pre) / Val(R.prov_VCpre), "###")
End If
pdf_P pdf, prepare(FER, pFER), c5, Y, Fs2, 2
  
'Print Provocation dose
Y = Y + lHt * 2
If Drug = Mannitol Then
    pdf_P pdf, "PD15:", c1, Y, Fs2, 3
    pdf_P pdf, "mg", c3, Y, Fs2, 2
    pdf_P pdf, Format(Calculate_PDx(R, 15), "#"), c6, Y, Fs2, 2
ElseIf (Drug = Methacholine Or Drug = Histamine) Then
    pdf_P pdf, "PD20:", c1, Y, Fs2, 3
    pdf_P pdf, "umol", c3, Y, Fs2, 2
    pdf_P pdf, Format(Calculate_PDx(R, 20), "#.#"), c6, Y, Fs2, 2
End If
     
'Dose response curve
Y = Y + lHt * 2
pdf_P pdf, "DOSE RESPONSE GRAPH:", c1, Y, Fs2, 3

'Co-ords (in mm) for graph
xMin = IDP(M.Left) + 30
xMax = xMin + 110
yMin = Int(Y) + 20
yMax = yMin + 70
xSize = xMax - xMin
ySize = yMax - yMin

'Draw graph
pdf_P pdf, "FEV1 (% of control)", xMin - 5 - IDP(M.Left), yMin - 10, Fs2, 3
If Drug = Mannitol Then
    pdf_P pdf, "Log Cumulative Mannitol Dose (mg)", xMin + 0.25 * xSize - IDP(M.Left), yMax + 2, Fs2, 3
ElseIf Drug = Methacholine Then
    pdf_P pdf, "Log Cumulative Methacholine Dose (umol)", xMin + 0.25 * xSize - IDP(M.Left), yMax + 2, Fs2, 3
ElseIf Drug = Histamine Then
    pdf_P pdf, "Log Cumulative Histamine Dose (umol)", xMin + 0.25 * xSize - IDP(M.Left), yMax + 2, Fs2, 3
End If

'Draw y-axis
pdf.SetLineWidth 0.5
pdf.DrawRectangle DPI(xMin), DPI(yMax), DPI(xSize), DPI(ySize), 0
i = 0
For j = yMin To yMax Step ySize / 7
    Y = j
    pdf_P pdf, Format(120 - i, "#0"), xMin + 4 - IDP(M.Left), Y - 7, Fs2, 3
    'draw little bip marks for baseline
    pdf.MoveTo DPI(xMin), DPI(Y): pdf.DrawLineTo DPI(xMin + 1), DPI(Y)
    'draw little bip marks for post
    pdf.MoveTo DPI(xMax - 1), DPI(Y): pdf.DrawLineTo DPI(xMax), DPI(Y)
    i = i + 10
    pdf.Stroke
Next j

'Draw x-axis
For j = 0 To UBound(Dose) + 2
    X = xMin + j * xSize / (UBound(Dose) + 2)
    pdf.MoveTo DPI(X), DPI(yMin): pdf.DrawLineTo DPI(X), DPI(yMax)
    pdf.Stroke
    Select Case j
        Case Is = 0:            Tmp = Dose_pre
        Case 1:                 Tmp = Dose_control
        Case UBound(Dose) + 2:  Tmp = R.prov_PostBDType
        Case Else:              Tmp = Dose(j - 1)
    End Select
    pdf.ShowTextAt DPI(X - 2), DPI(yMax + 1), Tmp
Next j

'Draw the gridlines
pdf.SetLineDash "[2] 2"
For j = yMin To yMax Step ySize / 7
    Y = j
    pdf.MoveTo DPI(xMin + xSize / 11), DPI(Y): pdf.DrawLineTo DPI(xMax - xSize / 11), DPI(Y)
    pdf.Stroke
Next j
For j = xMin To xMax Step xSize / 11
    X = j
    pdf.MoveTo DPI(X), DPI(yMin): pdf.DrawLineTo DPI(X), DPI(yMax)
    pdf.Stroke
Next j

'Plot the data
pdf.SetLineDash "[] 0"
pdf.SetLineWidth 1
prevX = xMin
prevY = yMin - 1.2 * ySize + 2 * ySize * Val(FEV1pre) / Val(FEV1_Control)
For j = 0 To UBound(Dose) + 2
    X = xMin + j * xSize / (UBound(Dose) + 2)
    Select Case j
        Case 0:                 Fev1ToPlot = FEV1pre
        Case 1:                 Fev1ToPlot = FEV1_Control
        Case UBound(Dose) + 2:  Fev1ToPlot = FEV1_Post
        Case Else:              Fev1ToPlot = FEV1(j - 1)
    End Select
    If Fev1ToPlot > 0 Then
        Y = yMax - ((100 * Val(Fev1ToPlot) / Val(FEV1_Control) - 50) / 70) * ySize
        'Y = yMin - 1.2 * ySize + 2 * ySize * Val(Fev1ToPlot) / Val(FEV1_Control)
        pdf.DrawCircle DPI(X), DPI(Y), DPI(0.5): pdf.Stroke: pdf.SetLineWidth 1
        pdf.MoveTo DPI(prevX), DPI(prevY): pdf.DrawLineTo DPI(X), DPI(Y)
        prevX = X: prevY = Y
    End If
    
'    If j = 0 Then
'        Y = yMin - 1.2 * ySize + 2 * ySize * Val(FEV1pre) / Val(FEV1_Control)
'        pdf.DrawCircle DPI(X), DPI(Y), DPI(0.5): pdf.Stroke: pdf.SetLineWidth 1
'        pdf.MoveTo DPI(prevX), DPI(prevY): pdf.DrawLineTo DPI(X), DPI(Y)
'        prevX = X: prevY = Y
'    ElseIf j = 1 Then
'        Y = yMin - 1.2 * ySize + 2 * ySize * Val(FEV1_Control) / Val(FEV1_Control)
'        pdf.DrawCircle DPI(X), DPI(Y), DPI(0.5): pdf.Stroke: pdf.SetLineWidth 1
'        pdf.MoveTo DPI(prevX), DPI(prevY): pdf.DrawLineTo DPI(X), DPI(Y)
'        prevX = X: prevY = Y
'    ElseIf j = UBound(Dose) + 2 Then
'        If Val(FEV1_Post) <> 0 Then
'            Y = yMin - 1.2 * ySize + 2 * ySize * Val(FEV1_Post) / Val(FEV1_Control)
'            pdf.DrawCircle DPI(X), DPI(Y), DPI(0.5): pdf.Stroke: pdf.SetLineWidth 1
'            pdf.MoveTo DPI(prevX), DPI(prevY): pdf.DrawLineTo DPI(X), DPI(Y)
'            prevX = X: prevY = Y
'        End If
'    Else
'        If FEV1(j - 1) <> 0 Then
'            Y = yMin - 1.2 * ySize + 2 * ySize * Val(FEV1(j - 1)) / Val(FEV1_Control)
'            pdf.DrawCircle DPI(X), DPI(Y), DPI(0.5): pdf.Stroke: pdf.SetLineWidth 1
'            pdf.MoveTo DPI(prevX), DPI(prevY): pdf.DrawLineTo DPI(X), DPI(Y)
'            prevX = X: prevY = Y
'        End If
'    End If
Next j
pdf.Stroke

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), yMax + 15)

Exit Function


DrawPDFReport_BronchChallengeData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport_BronchChallengeData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Sub PDF_DrawDashedLine(PDFname, x1, y1, x2, y2)

Dim X, Y, deltX, deltY, R As Single
Dim DashLen As Integer
Dim TestX, TestY As Boolean

    DashLen = 3
    deltX = x2 - x1: deltY = y2 - y1
    R = (deltX ^ 2 + deltY ^ 2) ^ 0.5
    'need to do this since can get rounding errors so that r is not exactly equal to zero
    If R < 0.0001 Then Exit Sub
    X = x1: Y = y1
    TestX = True: TestY = True
    PDFname.MoveTo x1, y1
    Do While TestX And TestY
        PDFname.DrawLineTo X + deltX * DashLen / R - DashLen, Y + deltY * DashLen / R
        PDFname.MoveTo X + 2 * deltX * DashLen / R, Y + 2 * deltY * DashLen / R
        X = X + 2 * deltX * DashLen / R: Y = Y + 2 * deltY * DashLen / R
        If deltX > 0 Then
            If X > x2 Then TestX = False Else TestX = True
        Else
            If X < x2 Then TestX = False Else TestX = True
        End If
        If deltY > 0 Then
            If Y > y2 Then TestY = False Else TestY = True
        Else
            If Y < y2 Then TestY = False Else TestY = True
        End If
    Loop
    PDFname.Stroke
End Sub

Function pdf_ViewOrPrintSilent(repType As ReportType, RecordID As Long, prnt As Boolean, Optional pHDC As Long, Optional JobNum, Optional Copies) As Boolean
    'Displays pdf in pdf viewer or prints it
    'Retrieves pdf report from the server and makes a local copy to work on - this avoids locking issues on the server copy
    'prnt=true: print it   =false: display it
    '  if printing, need to reconstruct image, add timestamp to bottom and print but dont resave pdf image to the database
    'pHDC=printer hdc
    'Copies=number of copies to print
    
    
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim pdfStream As New ADODB.Stream
Dim pdfFound As Boolean
Dim Fname_local As String
Dim Fname_server As String
Dim Tbl As String
Dim pointerField_name As String
Dim idField_name As String
Dim pdf As New PDFCreatorPilotLib.PDFDocument4
Dim i, j
Dim PageCounter()
Dim Msg As String

On Error GoTo pdf_ViewOrPrintSilent_Error

Screen.MousePointer = 11

If IsMissing(Copies) Then Copies = 1
If IsMissing(JobNum) Then
    Fname_local = Environ("Tmp") & "\Temp.pdf"
Else
    Fname_local = Environ("Tmp") & "\Temp" & JobNum & ".pdf"
End If

'Delete any leftover copy of the temporary file to rpevent displaying the wrong patient report in event of error
If Dir(Fname_local) <> "" Then Kill Fname_local

Select Case repType
    Case ReportType.rtBronchChall:
        FilenamePrefix = "RFT_BronchChall"
        Tbl = "RftResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
    Case ReportType.rtCpapClinic:
        FilenamePrefix = "SLEEP_CpapClinic"
        Tbl = ""
        Fld = ""
        idField_name = ""
    Case ReportType.rtCpx:
        FilenamePrefix = "RFT_Cpx"
        Tbl = "ExResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
    Case ReportType.rtHAST:
        FilenamePrefix = "RFT_HAST"
        Tbl = "RftResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
    Case ReportType.rtRFT:
        FilenamePrefix = "RFT_Rft"
        Tbl = "RftResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
    Case ReportType.rtSkin:
        FilenamePrefix = "RFT_Skin"
        Tbl = "RftResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
    Case ReportType.rtSleep:
        FilenamePrefix = "SLEEP_Psg"
        Tbl = "SlpResults"
        pointerField_name = "TypedBy"
        idField_name = "StudyID"
    Case ReportType.rtTreadmill:
        FilenamePrefix = "RFT_Treadmill"
        Tbl = "RftResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
    Case ReportType.rtSixMWD:
        FilenamePrefix = "RFT_6MWD"
        Tbl = "RftResults"
        pointerField_name = "TypedBy"
        idField_name = "ResultsID"
End Select

'Get the pointer to the pdf file.
sql = "SELECT * FROM FilePointers WHERE ReportType=" & repType & " AND RecordID=" & RecordID & " AND Tablename='" & Tbl & "'"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
        Screen.MousePointer = 0
        MsgBox "PDF image unavailable. Re-save report and try again.", vbOKOnly, "PDF Viewer"
        pdfFound = False
Else
    'Pointer found, check that file exists
    Fname_server = gServerPdfLocation & Rs("PointerValue")
    If Dir(Fname_server) = "" Then
        MsgBox "PDF image unavailable. Re-save report and try again.", vbOKOnly, "PDF Viewer"
        pdfFound = False
    Else
        pdfFound = True
        'Make a copy on the local hard drive
        FileCopy Fname_server, Fname_local
    End If
End If
Rs.Close

'Call print routines if print selected Then
If pdfFound Then
    Select Case prnt
        Case False
            'Display it
            iReturn = ShellExecute(0, "open", Fname_local, 0, 0, SW_SHOW)
        Case True
            'Append a time stamp
            iReturn = pdf_GetLicenseData
            pdf.SetLicenseData gCreatorPilot.RegistrationName, gCreatorPilot.SerialNo      ' initialize PDF Engine
            pdf.Open Fname_local, ""
            ReDim PageCounter(1 To pdf.GetPageCount)
            For i = 1 To pdf.GetPageCount
                pdf.CurrentPage = i - 1
                PageCounter(i) = pdf.AddEditBox(200, 815, 400, 825, "MyField" & PageCounter(i))
                pdf.CurrentAnnotation = PageCounter(i)
                pdf.SetAnnotText "Printed: " & Format(Now, "hh:mm:ss  dd/mm/yyyy"), 1
                pdf.ControlShowBorder = False
                pdf.ControlTextColor = 0 'Black
                pdf.ControlBackColor = &HFFFFFF 'White
                pdf.ControlFontSize = 7
                pdf.AnnotPrint = True
            Next i
            pdf.SaveToFile Fname_local, False
            'Print it
            iReturn = OpenPrinter("", pHDC, 0)
            For i = 1 To Copies
                iReturn = ShellExecute(0, "print", Fname_local, 0, 0, SW_HIDE)
            Next i
            iReturn = ClosePrinter(pHDC)
    End Select
End If
Screen.MousePointer = 0

Exit Function


pdf_ViewOrPrintSilent_Error:
    Select Case Err.Number
        Case 70
            Screen.MousePointer = 0
            Msg = "Error loading pdf - close currently open pdf files and retry"
            Call ErrorLog(Msg, "pdf_ViewOrPrintSilent", "mPdfRoutines", Err.Number, Err.Description)
        Case Else
            Msg = ""
            Call ErrorLog(Msg, "pdf_ViewOrPrintSilent", "mPdfRoutines", Err.Number, Err.Description)
    End Select
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function


Function DrawPDFReport_HastData(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, yStart As Single)

Dim ASTarray(1 To 7, 1 To 8) As String
Dim AST As String
Dim RowNum As Integer
Dim ColumnNum As Integer
Dim Age As Single
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single, c8 As Single, c9 As Single 'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim Y As Single, y1 As Single
Dim Msg As String

Const lHt As Single = 3.6

On Error GoTo DrawPDFReport_HastData_Error

'Setup stuff
Y = IDP(yStart) + 2.5
c1 = 15: c2 = c1 + 15: c3 = c2 + 13: c4 = c3 + 22: c5 = c4 + 29: c6 = c5 + 14: c7 = c6 + 11: c8 = c7 + 17: c9 = c8 + 15
Fs1 = 9   'Test header
Fs2 = 8.5    'Results values, indice labels, units, normals
Fs3 = 7    'Equipment string
Fs4 = 10    'Report

'Print table
pdf_P pdf, "ALTITUDE SIMULATION TESTING", 2, Y, Fs1, 3
Y = Y + lHt * 3

pdf_P pdf, "Stage", c1, Y, Fs2, 2
pdf_P pdf, "FiO2", c2, Y, Fs2, 2
pdf_P pdf, "Simulated", c3, Y, Fs2, 2
pdf_P pdf, "Supplemental", c4, Y, Fs2, 2
pdf_P pdf, "SpO2", c5, Y, Fs2, 2
pdf_P pdf, "pH", c6, Y, Fs2, 2
pdf_P pdf, "PaCO2", c7, Y, Fs2, 2
pdf_P pdf, "PaO2", c8, Y, Fs2, 2
pdf_P pdf, "SaO2", c9, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "(%)", c2, Y, Fs2, 2
pdf_P pdf, "Altitude", c3, Y, Fs2, 2
pdf_P pdf, "Intranasal O2", c4, Y, Fs2, 2
pdf_P pdf, "(%)", c5, Y, Fs2, 2
pdf_P pdf, "(mmHg)", c7, Y, Fs2, 2
pdf_P pdf, "(mmHg)", c8, Y, Fs2, 2
pdf_P pdf, "(%)", c9, Y, Fs2, 2
Y = Y + lHt
pdf_P pdf, "(feet)", c3, Y, Fs2, 2
pdf_P pdf, "(L/min)", c4, Y, Fs2, 2
Y = Y + lHt * 2.7

pdf.MoveTo M.Left + DPI(c1), DPI(Y)
pdf.DrawLineTo pdf.PageWidth - M.RIGHT - DPI(20), DPI(Y)
pdf.Stroke


'Want to extract data from string and build array - easier to manipulate when printing
'decompress the string to the text boxes
AST = R.AST_TestDataString
j = 1: RowNum = 1: ColumnNum = 1
For i = 1 To Len(AST)
    If Mid(AST, i, 1) = "$" Then
       ASTarray(RowNum, ColumnNum) = Mid(AST, j, i - j)
        ColumnNum = ColumnNum + 1
        If ColumnNum = 9 Then
            ColumnNum = 1
            RowNum = RowNum + 1
        End If
        j = i + 1
    End If
Next i

For i = 1 To 7   'step through the rows
    If ASTarray(i, 1) = "" Then Exit For
    pdf_P pdf, i, c1, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 1), c2, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 2), c3, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 3), c4, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 4), c5, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 5), c6, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 6), c7, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 7), c8, Y, Fs2, 2
    pdf_P pdf, ASTarray(i, 8), c9, Y, Fs2, 2
    Y = Y + lHt
Next i

Y = Y + lHt * 2

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), Y + 35)

Exit Function

DrawPDFReport_HastData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReport_HastData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function DrawPDFReportSkinData(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, yStart As Single)

Dim SkinArray(0 To 57)
Dim skinString As String
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim Y As Single, y1 As Single, yEnd As Single
Dim TestSite As String
Dim Drugs, Drug, AllergenLabel, Weal, Meds
Dim OtherDrugs(0 To 3)
Dim Msg As String
Dim GrpName As String

Dim SptData As SptAllData

Const lHt As Single = 3.6

On Error GoTo DrawPDFReportSkinData_Error

Drugs = Array("Antihistamines", "Theophylline", "Corticosteroids", "Beta blockers")

'Setup stuff
Y = IDP(yStart) + 2.5
c1 = 10: c2 = c1 + 5: c3 = c2 + 35: c4 = c3 + 35: c5 = 130
Fs1 = 9   'Test header
Fs2 = 8.5    'Results values, indice labels, units, normals
Fs3 = 7    'Equipment string
Fs4 = 10    'Report

'Print table
pdf_P pdf, "SKIN PRICK TESTING", 2, Y, Fs1, 3
Y = Y + lHt * 2
pdf_P pdf, "ALLERGEN GROUP", c2, Y, Fs2, 2
pdf_P pdf, "ALLERGEN", c3, Y, Fs2, 2
pdf_P pdf, "WHEAL (mm)", c4 - 5, Y, Fs2, 2

'Get spt data
SptData = gDBFunctions.Decode_SptString(R.SkinTestDataString)
For i = 1 To UBound(SptData.Allergens)
    If i = 1 Then
        Y = Y + lHt * 2.5
        pdf.MoveTo M.Left + DPI(c1), DPI(Y)
        pdf.DrawLineTo M.Left + DPI(c4) + DPI(20), DPI(Y)
        pdf.Stroke
        Y = Y - lHt
        pdf_P pdf, SptData.Allergens(i).GroupName, c2, Y, Fs2, 2
        pdf_P pdf, SptData.Allergens(i).AllergenName, c3, Y, Fs2, 2
        pdf_P pdf, SptData.Allergens(i).Wheal_mm, c4, Y, Fs2, 2
    Else
        If SptData.Allergens(i).GroupName = SptData.Allergens(i - 1).GroupName Then
            GrpName = ""
            Y = Y + lHt
        Else
            Y = Y + lHt * 1.5
            GrpName = SptData.Allergens(i).GroupName
        End If
        pdf_P pdf, GrpName, c2, Y, Fs2, 2
        pdf_P pdf, SptData.Allergens(i).AllergenName, c3, Y, Fs2, 2
        pdf_P pdf, SptData.Allergens(i).Wheal_mm, c4, Y, Fs2, 2
    End If
Next i
y1 = Y + lHt * 2

Y = 80
pdf_P pdf, "TEST SITE:  " & SptData.Site, c5, Y, Fs1, 2
Y = Y + lHt * 1.5
pdf_P pdf, "MEDICATIONS (last 48 hrs):", c5, Y, Fs1, 2
Y = Y + lHt
If SptData.Antihistamine Then Drug = "Yes" Else Drug = "No"
pdf_P pdf, "Antihistamines: " & Drug, c5 + 5, Y, Fs1, 2
If SptData.BetaBlocker Then Drug = "Yes" Else Drug = "No"
pdf_P pdf, "Beta-blockers: " & Drug, c5 + 5, Y + lHt, Fs1, 2

If Y < y1 Then Y = y1

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), Y + 5)

Exit Function

DrawPDFReportSkinData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportSkinData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function DrawPDFReportSkinData_OLD(pdf As PDFCreatorPilotLib.PDFDocument4, R As RFTs, yStart As Single)

Dim SkinArray(0 To 57)
Dim skinString As String
Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim Y As Single, y1 As Single, yEnd As Single
Dim TestSite As String
Dim Drugs, Drug, AllergenLabel, Weal, Meds
Dim OtherDrugs(0 To 3)

Const lHt As Single = 3.6
Dim Msg As String

On Error GoTo DrawPDFReportSkinData_Error

Drugs = Array("Antihistamines", "Theophylline", "Corticosteroids", "Beta blockers")

'Setup stuff
Y = IDP(yStart) + 2.5
c1 = 10: c2 = c1 + 5: c3 = c2 + 35: c4 = c3 + 35: c5 = 130
Fs1 = 9   'Test header
Fs2 = 8.5    'Results values, indice labels, units, normals
Fs3 = 7    'Equipment string
Fs4 = 10    'Report

'Print table
pdf_P pdf, "SKIN PRICK TESTING", 2, Y, Fs1, 3
Y = Y + lHt * 2
pdf_P pdf, "ALLERGEN GROUP", c2, Y, Fs2, 2
pdf_P pdf, "ALLERGEN", c3, Y, Fs2, 2
pdf_P pdf, "WHEAL (mm)", c4 - 5, Y, Fs2, 2
Y = Y + 4

'Want to extract data from string and build array - easier to manipulate when printing in allergenGroup order
skinString = Mid$(R.SkinTestDataString, 6)  'excise the first 5 characters which contain the testsite and other medications info
indexipoo = 0: j = 1
For i = 1 To Len(skinString)
    If Mid(skinString, i, 1) = "|" Then
        SkinArray(indexipoo) = Mid(skinString, j, i - j)
        indexipoo = indexipoo + 1
        j = i + 1
    End If
Next i

For i = 1 To 7   'need to step through the groups in order (1=control, 2=pollen, 3=dust,4 = mould, 5=dander, 6=food, 7=other
    Select Case i
        Case 1
            guppy = "Control"
        Case 2
            guppy = "Pollens"
        Case 3
            guppy = "Dust Mite"
        Case 4
            guppy = "Moulds"
        Case 5
            guppy = "Danders"
        Case 6
            guppy = "Foods"
        Case 7
            guppy = "Other"
    End Select
    notDoneLabel = True
    For j = 0 To indexipoo - 1
        If Mid(SkinArray(j), 4, 2) = Left(guppy, 2) Then
            firstOne = 0: secondOne = 0: thirdOne = 0
            
            For k = 1 To Len(SkinArray(j)) 'find the 3 slashes which separate the data
                If Mid(SkinArray(j), k, 1) = "/" Then
                    If firstOne = 0 Then
                        firstOne = k
                    Else
                        If secondOne = 0 Then secondOne = k Else thirdOne = k
                    End If
                End If
            Next k
            
            AllergenLabel = Mid(SkinArray(j), secondOne + 1, thirdOne - secondOne - 1)
            Weal = Mid(SkinArray(j), thirdOne + 1)
            If Weal <> "" Then
                If notDoneLabel Then
                    Y = Y + 3
                    pdf_P pdf, guppy, c2, Y, Fs2, 2
                    pdf.MoveTo M.Left + DPI(c1), DPI(Y + 2)
                    pdf.DrawLineTo M.Left + DPI(c4) + DPI(20), DPI(Y + 2)
                    pdf.Stroke
                    notDoneLabel = False
                End If
                pdf_P pdf, AllergenLabel, c3, Y, Fs2, 2
                pdf_P pdf, Weal, c4, Y, Fs2, 2
                Y = Y + lHt
            End If
        End If
    Next j
Next i
y1 = Y + lHt

'Get site
Select Case Left$(R.SkinTestDataString, 1)
    Case "L": TestSite = "Left Forearm"
    Case "R": TestSite = "Right Forearm"
    Case "B": TestSite = "Both Forearms"
    Case "K": TestSite = "Back"
End Select

'Get drugs
Meds = Mid$(R.SkinTestDataString, 2, 4)      'testsite and other medications info
If InStr(Meds, "A") Then OtherDrugs(0) = True Else OtherDrugs(0) = False
If InStr(Meds, "T") Then OtherDrugs(1) = True Else OtherDrugs(1) = False
If InStr(Meds, "C") Then OtherDrugs(2) = True Else OtherDrugs(2) = False
If InStr(Meds, "B") Then OtherDrugs(3) = True Else OtherDrugs(3) = False
Y = 80
pdf_P pdf, "TEST SITE:  " & TestSite, c5, Y, Fs1, 2
Y = Y + lHt * 1.5
pdf_P pdf, "MEDICATIONS (last 48 hrs):", c5, Y, Fs1, 2
Y = Y + lHt
For i = 0 To 3
    If OtherDrugs(i) Then Drug = "Yes" Else Drug = "No"
    pdf_P pdf, Drugs(i) & ":", c5 + 5, Y + lHt * i, Fs1, 2
    pdf_P pdf, Drug, c5 + 30, Y + lHt * i, Fs1, 2
Next i
Y = Y + lHt * i
If Y < y1 Then Y = y1

Call DrawPDFReportRftReport1(pdf, Pt.Get_ReportSectionInfoFromRftUDT(R), Y)

Exit Function

DrawPDFReportSkinData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportSkinData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function DrawPDFReportCPAPClinicData(pdf As PDFCreatorPilotLib.PDFDocument4, c As CPAPClinic, yStart As Single)

Dim c1 As Single, c2 As Single, c3 As Single, c4 As Single, c5 As Single, c6 As Single, c7 As Single  'Column spacings
Dim Fs1 As Single, Fs2 As Single, Fs3 As Single, Fs4 As Single 'Font sizes
Dim Y As Single, y1 As Single
Dim yTemp As Single
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim MachineFound As Boolean
Dim Msg As String

Const lHt As Single = 3.6

On Error GoTo DrawPDFReportCPAPClinicData_Error

'Setup stuff
Y = IDP(yStart) + 2.5
c1 = 5: c2 = c1 + 5: c3 = c2 + 25: c4 = c3 + 30: c5 = c4 + 25: c6 = c5 + 25: c7 = c6 + 25
Fs1 = 9         'Test header
Fs2 = 8.5       'Results values, indice labels, units, normals
Fs3 = 7         'Equipment string
Fs4 = 10        'Report

'Get pump details
If c.MachineID > 0 Then
    sql = "SELECT VentilatorSettings.Manufacturer, VentilatorSettings.Model, CPAP_VentEquip.SerialNo "
    sql = sql & "FROM CPAP_VentEquip INNER JOIN VentilatorSettings ON CPAP_VentEquip.ModelID = VentilatorSettings.ModelID "
    sql = sql & "WHERE CPAP_VentEquip.MachineID=" & c.MachineID & ";"
    Rs.Open sql, ConnectString, adOpenStatic   'nb read only
    MachineFound = True
End If

'Print CPAP table
pdf_P pdf, "CPAP THERAPY DETAILS", 2, Y, Fs1, 3: Y = Y + lHt * 2
yTemp = Y
pdf_P pdf, "CPAP PUMP:", c2, Y, Fs2, 3: Y = Y + lHt
yTemp1 = Y
pdf_P pdf, "Manufacturer: ", c2, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Model: ", c2, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Serial#: ", c2, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Hourmeter: ", c2, Y, Fs2, 2: Y = Y + lHt
If MachineFound Then
    Y = yTemp1
    pdf_P pdf, Rs("Manufacturer"), c3, Y, Fs2, 2: Y = Y + lHt
    pdf_P pdf, Rs("Model"), c3, Y, Fs2, 2: Y = Y + lHt
    pdf_P pdf, Rs("SerialNo"), c3, Y, Fs2, 2: Y = Y + lHt
    pdf_P pdf, c.HourMeter, c3, Y, Fs2, 2: Y = Y + lHt
End If

Y = yTemp
pdf_P pdf, "SETTINGS:", c4, Y, Fs2, 3: Y = Y + lHt
pdf_P pdf, "CPAP (cmH2O): ", c4, Y, Fs2, 2: Y = Y + lHt
Y = yTemp1
pdf_P pdf, c.CPAP_Set, c5, Y, Fs2, 2: Y = Y + lHt

Y = yTemp
pdf_P pdf, "MASK:", c6, Y, Fs2, 3: Y = Y + lHt
pdf_P pdf, "Manufacturer: ", c6, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Model: ", c6, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Frame style: ", c6, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Mask size: ", c6, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Headgear: ", c6, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, "Headgear size: ", c6, Y, Fs2, 2: Y = Y + lHt
Y = yTemp1
pdf_P pdf, c.Mask_Manufacturer, c7, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, c.Mask_Model, c7, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, c.Mask_Frame, c7, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, c.Mask_CushionSize, c7, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, c.Mask_HeadgearType, c7, Y, Fs2, 2: Y = Y + lHt
pdf_P pdf, c.Mask_HeadgearSize, c7, Y, Fs2, 2: Y = Y + lHt * 3

'pdf.MoveTo M.Left, DPI(Y)
'pdf.DrawLineTo pdf.PageWidth - M.RIGHT, DPI(Y)
'pdf.Stroke

If Rs.State = adStateOpen Then Rs.Close

'Report section
pdf_P pdf, "FOLLOWUP NOTES:", 2, Y - lHt * 0.5, Fs1, 3
Y = Y + lHt * 2.5

pdf.UseFont gPdfFonts(6).ID, 10
If InStr(c.Visit_Clinician, ",") Then
    Tmp = Trim(c.Visit_Note) & vbCrLf & vbCrLf & "    Clinician:   " & Mid(c.Visit_Clinician, InStr(c.Visit_Clinician, ",") + 1) & " " & Mid(c.Visit_Clinician, 1, InStr(c.Visit_Clinician, ",") - 1)
Else
    Tmp = Trim(c.Visit_Note) & vbCrLf & vbCrLf & "    Clinician:   " & c.Visit_Clinician
End If
i = pdf.ShowTextLines(M.Left + DPI(c1), DPI(Y), pdf.PageWidth - M.RIGHT - DPI(2), pdf.PageHeight - M.Bottom, -1, taLeft, vaTop, Tmp)

Exit Function


DrawPDFReportCPAPClinicData_Error:
    Msg = ""
    Call ErrorLog(Msg, "DrawPDFReportCPAPClinicData", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function pdf_DoPageNumberingxxx(pdf As PDFCreatorPilotLib.PDFDocument4, ReportStatus As String, Report_Type As ReportType) As Boolean

Dim i As Integer
Dim Msg As String

On Error GoTo pdf_DoPageNumbering_Error

Select Case Report_Type
    Case ReportType.rtSleep
        X = M.Left + DPI(165): Y = M.Top + DPI(4)
        For i = 0 To pdf.GetPageCount - 1
            pdf.CurrentPage = i
            pdf.UseFont gPdfFonts(2).ID, 9
            pdf.ShowTextAt X, Y, "(Page " & i + 1 & "/" & pdf.GetPageCount & ")"
            pdf_P pdf, "Report Status: " & ReportStatus, 2, IDP(pdf.PageHeight) - 14, 7, gPdfFonts(2).ID
        Next i
    Case ReportType.rtCpapClinic
        X = M.Left + DPI(165): Y = M.Top + DPI(4)
        For i = 0 To pdf.GetPageCount - 1
            pdf.CurrentPage = i
            pdf.UseFont gPdfFonts(2).ID, 9
            pdf.ShowTextAt X, Y, "(Page " & i + 1 & "/" & pdf.GetPageCount & ")"
        Next i
    Case Else
        pdf_P pdf, "Report Status: " & ReportStatus, 2, IDP(pdf.PageHeight) - 14, 7, gPdfFonts(2).ID
End Select

Exit Function

pdf_DoPageNumbering_Error:
    Msg = ""
    Call ErrorLog(Msg, "pdf_DoPageNumbering", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function pdf_DoPageNumbering1(pdf As PDFCreatorPilotLib.PDFDocument4, ReportStatus As String, Report_Type As ReportType) As Boolean

Dim i As Integer


Dim Msg As String

    On Error GoTo pdf_DoPageNumbering1_Error

Select Case Report_Type
    Case ReportType.rtSleep
        X = M.Left + DPI(165): Y = M.Top + DPI(4)
        For i = 0 To pdf.GetPageCount - 1
            pdf.CurrentPage = i
            pdf.UseFont gPdfFonts(2).ID, 9
            pdf.ShowTextAt X, Y, "(Page " & i + 1 & "/" & pdf.GetPageCount & ")"
            pdf_P pdf, "Report Status: " & ReportStatus, 2, IDP(pdf.PageHeight) - 14, 7, gPdfFonts(2).ID
        Next i
    Case ReportType.rtCpapClinic
        X = M.Left + DPI(165): Y = M.Top + DPI(4)
        For i = 0 To pdf.GetPageCount - 1
            pdf.CurrentPage = i
            pdf.UseFont gPdfFonts(2).ID, 9
            pdf.ShowTextAt X, Y, "(Page " & i + 1 & "/" & pdf.GetPageCount & ")"
        Next i
    Case Else
        'added page numbers to respiratory reports as per Nata vist 09/2013 - WRR
        X = M.Left + DPI(165): Y = pdf.PageHeight - DPI(10) 'M.Top + DPI(1)
        For i = 0 To pdf.GetPageCount - 1
            pdf.CurrentPage = i
            pdf.UseFont gPdfFonts(2).ID, 9
            pdf.ShowTextAt X, Y, "(Page " & i + 1 & "/" & pdf.GetPageCount & ")"
            pdf_P pdf, "Report Status: " & ReportStatus, 2, IDP(pdf.PageHeight) - 14, 7, gPdfFonts(2).ID
        Next i        '
        'original code below WRR
        'pdf_P pdf, "Report Status: " & ReportStatus, 2, IDP(pdf.PageHeight) - 14, 7, gPdfFonts(2).ID
End Select


    On Error GoTo 0
    Exit Function

pdf_DoPageNumbering1_Error:
    Msg = ""
    Call ErrorLog(Msg, "pdf_DoPageNumbering1", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function Pdf_SaveAsPointer(repType As ReportType, SourceFile As Variant, RecordID As Long) As String
'Pdf files are stored in a directory on the server. A pointer to the file is saved in the database.
'Input parameters:
'   Sourcefile - full path and filename of the file to be saved
'   RecordID - record number of the table where the pointer is stored
'Server filename format eg - RFT_Cpx_PointerID


Dim FilenamePrefix As String
Dim Tbl As String
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Pdf_SaveAsPointer_Error

Pdf_SaveAsPointer = ""

Select Case repType
    Case ReportType.rtBronchChall:
        FilenamePrefix = "RFT_BronchChall"
        Tbl = "RftResults"
    Case ReportType.rtCpapClinic:
        FilenamePrefix = "SLEEP_CpapClinic"
        Tbl = "Visit_CPAP_ImpReview"
    Case ReportType.rtCpx:
        FilenamePrefix = "RFT_Cpx"
        Tbl = "ExResults"
    Case ReportType.rtHAST:
        FilenamePrefix = "RFT_HAST"
        Tbl = "RftResults"
    Case ReportType.rtRFT:
        FilenamePrefix = "RFT_Rft"
        Tbl = "RftResults"
    Case ReportType.rtSkin:
        FilenamePrefix = "RFT_Skin"
        Tbl = "RftResults"
    Case ReportType.rtSleep:
        FilenamePrefix = "SLEEP_Psg"
        Tbl = "SlpResults"
    Case ReportType.rtTreadmill:
        FilenamePrefix = "RFT_Treadmill"
        Tbl = "RftResults"
    Case ReportType.rtSixMWD:
        FilenamePrefix = "RFT_6MWD"
        Tbl = "RftResults"
End Select

'Find record and save pointer - if no existing record, create one
sql = "SELECT * FROM FilePointers WHERE ReportType=" & repType & " AND TableName='" & Tbl & "' AND RecordID=" & RecordID
Rs.CursorLocation = adUseClient
Rs.Open sql, ConnectString, adOpenStatic, adLockOptimistic
If Rs.EOF Then
    Rs.AddNew
        Rs("ReportType") = repType
        Rs("PointerValue") = 0
        Rs("RecordID") = RecordID
        Rs("TableName") = Tbl
        Rs("LastUpdated") = Now
    Rs.Update
    Rs.MoveFirst   'Get the new auto increment record ID
    'Construct unique filename
    Filename = FilenamePrefix & "_" & Rs("PointerID") & ".pdf"
    Rs("PointerValue") = Filename
    Rs.Update
    'Save file
    FileCopy SourceFile, gServerPdfLocation & Filename
    'Return the new filename
    Pdf_SaveAsPointer = gServerPdfLocation & Filename
Else
    'Should only ever be a single record
    If Rs.RecordCount = 1 Then
        'Construct unique filename
        Filename = FilenamePrefix & "_" & Rs("PointerID") & ".pdf"
        Rs("PointerValue") = Filename
        Rs("LastUpdated") = Now
        Rs.Update
        'Save file
        FileCopy SourceFile, gServerPdfLocation & Filename
        'Return the new filename
        Pdf_SaveAsPointer = gServerPdfLocation & Filename
    Else
        'Warning and dont save the file
        MsgBox "Warning: Multiple file pointers for this file" & vbCrLf & "File not saved" & vbCrLf & "Note this message and contact Peter Rochford x3673"
    End If
End If
Rs.Close



    On Error GoTo 0
    Exit Function

Pdf_SaveAsPointer_Error:
    Msg = ""
    Call ErrorLog(Msg, "Pdf_SaveAsPointer", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function


Public Function GetScoringRuleSet(StudyDate As Date, Item As ScoringRulesItem) As String

Dim i As Integer
Dim Msg As String

On Error GoTo GetScoringRuleSet_Error

If Not IsDate(StudyDate) Then
    GetScoringRuleSet = ""
    Exit Function
End If

For i = LBound(ScoringRule) To UBound(ScoringRule)
    If StudyDate >= ScoringRule(i).StartDate Then
        Select Case Item
            Case ScoringRulesItem.Description:  GetScoringRuleSet = ScoringRule(i).Description
            Case ScoringRulesItem.label:        GetScoringRuleSet = ScoringRule(i).label
            Case ScoringRulesItem.StartDate:    GetScoringRuleSet = ScoringRule(i).StartDate
            Case ScoringRulesItem.EndDate:      GetScoringRuleSet = ScoringRule(i).EndDate
        End Select
        Exit For
    End If
Next i

Exit Function

GetScoringRuleSet_Error:
    Msg = ""
    Call ErrorLog(Msg, "GetScoringRuleSet", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Sub SetScoringRulesNotes()

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim i As Integer

Dim Msg As String

    On Error GoTo SetScoringRulesNotes_Error


sql = "SELECT * FROM Config_PsgScoring ORDER BY RuleSet_StartDate DESC;"
Rs.CursorLocation = adUseClient
Rs.Open sql, ConnectString, adOpenStatic, adLockOptimistic
If Not Rs.EOF Then
    i = 0
    Do While Not Rs.EOF
        ReDim Preserve ScoringRule(i)
        ScoringRule(i).Description = Rs("RuleSet_Description")
        ScoringRule(i).label = Rs("RuleSet_Label")
        ScoringRule(i).StartDate = Rs("RuleSet_StartDate")
        If IsNull(Rs("RuleSet_EndDate")) Then
            ScoringRule(i).EndDate = Now()
        Else
            ScoringRule(i).EndDate = Rs("RuleSet_EndDate")
        End If
        i = i + 1
        Rs.MoveNext
    Loop
End If
Rs.Close

Exit Sub


ScoringRulesNotes.Before_20_5_2002.Description = "Respiratory events - local rules, Sleep staging - R&K, Arousals - ASDA 1992."
ScoringRulesNotes.Before_20_5_2002.StartDate = #1/1/1980#
ScoringRulesNotes.Before_20_5_2002.EndDate = #5/19/2002#

ScoringRulesNotes.From_20_5_2002.Description = "Respiratory events - Chicago 1999, Sleep staging - R&K, Arousals - ASDA 1992."
ScoringRulesNotes.From_20_5_2002.StartDate = #5/20/2002#
ScoringRulesNotes.From_20_5_2002.EndDate = #5/31/2011#

ScoringRulesNotes.From_01_06_2011.Description = "Respiratory events, Sleep staging, Arousals - AASM2007 (incorporating ASTA/ASA2011 guidelines)."
ScoringRulesNotes.From_01_06_2011.StartDate = #6/1/2011#
ScoringRulesNotes.From_01_06_2011.EndDate = Now()



    On Error GoTo 0
    Exit Sub

SetScoringRulesNotes_Error:
    Msg = ""
    Call ErrorLog(Msg, "SetScoringRulesNotes", "mPdfRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Sub
    Resume

End Sub
