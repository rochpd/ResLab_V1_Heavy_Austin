VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum RecordMode
    newRecord
    editRecord
    viewRecord
End Enum
Private Enum MedicareNoFormat
    Formatted
    Unformatted
End Enum
Public Enum DRSMContactType
    All
    RFTs
    Cpx
    RftAndCpx
    CPAP
    Psgs
    VRSS_Assess
    VRSS_Review
    VRSS_Episode
    VRSS_ClinicalContact
End Enum

Private mPtID As Long
Private mRftID As RftRecord
Private mSurname As String
Private mFirstName As String
Private mNumberOfAlerts As Integer
Private mDisplayPtPanel As Panel

Public Property Let DisplayPtControl(pnl As Panel)
    Set mDisplayPtPanel = pnl
End Property

Friend Property Let CurrentRft(R As RftRecord)
    mRftID.ID = R.ID
    mRftID.SourceTable = R.SourceTable
End Property

Friend Property Get CurrentRft() As RftRecord
    CurrentRft.ID = mRftID.ID
    CurrentRft.SourceTable = mRftID.SourceTable
End Property



Public Property Let PtID(ByVal PtID As Long)
    
    Dim Rs As New ADODB.Recordset
    Dim sql As String
    
    mPtID = PtID
    mSurname = Get_PtString(mPtID, cHS.CurrentVAED, CurrentPatientItems.Lastname)
    mFirstName = Get_PtString(mPtID, cHS.CurrentVAED, CurrentPatientItems.Firstname)

    'Display the current pt's details
    ' Set mDisplayPtControl equal to the actual panel control that will be used by the calling app to display pt's details
    mDisplayPtPanel.Text = Get_PtString(mPtID, cHS.CurrentVAED, FullName_DisplayUR)
    
    'Check for alerts
    If InStr(ConnectString, "server303s") Then
        sql = "SELECT id FROM tbl_HL7_AllergyAlert_Cerner WHERE patient_ID=" & PtID
        sql = sql & " UNION "
        sql = sql & "SELECT id FROM tbl_HL7_AllergyAlert_Trak WHERE patient_ID=" & PtID
        Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        If Not Rs.EOF Then
            mNumberOfAlerts = Rs.RecordCount
            frmMainMenu.txtAlert.Enabled = True
        Else
            mNumberOfAlerts = 0
            frmMainMenu.txtAlert.Enabled = False
        End If
        Rs.Close
    End If
    
End Property

Public Property Get Surname() As String
    Surname = mSurname
End Property
Public Property Get GivenName() As String
    GivenName = mFirstName
End Property

Public Property Get PtID() As Long
    PtID = mPtID
End Property

Public Property Get PatientExists() As Boolean
    PatientExists = CBool(mPtID <> 0)
End Property

Public Property Get HasAlerts() As Boolean
    HasAlerts = CBool(mNumberOfAlerts > 0)
End Property

Public Function Clear_RftRecordID()

    mRftID.ID = 0
    mRftID.SourceTable = DbTables.Not_Specified

End Function

Friend Function Get_DRSMContacts_V2(PatientID As Long, ContactType As DRSMContactType, Optional VrssEpisodeID) As DRSMContacts
'Returns lists of record IDs for previous contacts in DRSM -
'Version for VRSS episodes

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Contacts As DRSMContacts
Dim Msg As String
Dim tempDate As Date

ReDim Contacts.RFTs(0)
ReDim Contacts.Cpx(0)
ReDim Contacts.RftsOrCpx(0)
ReDim Contacts.CPAP(0)
ReDim Contacts.Psgs(0)
ReDim Contacts.Vrss_A(0)
ReDim Contacts.Vrss_R(0)
ReDim Contacts.Vrss_Contacts(0)

On Error GoTo Get_DRSMContacts_Error

Select Case ContactType
    Case DRSMContactType.All
        sql = "      (SELECT ResultsID as ID, 'RftResults'              AS Tbl, 'RFT'                  AS Type, TestDate as xDate,  TestType AS Description  FROM RftResults              WHERE PatientID=" & PatientID & ") UNION "
        sql = sql & "(SELECT ResultsID as ID, 'ExResults'               AS Tbl, 'CPX'                  AS Type, TestDate as xDate,  TestType AS Description  FROM ExResults               WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT StudyID as ID,   'SlpResults'              AS Tbl, 'SLP'                  AS Type, StudyDate as xDate, StudyType AS Description FROM SlpResults              WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT EpisodeID as ID, 'VRSS_Episodes'           AS Tbl, 'VRSS_EPISODE'         AS Type, Ep_StartDate as xDate, Ep_EndDate as EndDate, 'EPISODE' AS Description FROM VRSS_Episodes WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT VisitID as ID,   'Visit_CPAP_ImpReview'    AS Tbl, 'CPAP'                 AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_CPAP_ImpReview    WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT ContactID as ID, 'VRSS_ClinicalContacts'   AS Tbl, 'VRSS_CLINICALCONTACT' AS Type, Contact_Date as xDate, Contact_Type AS Description FROM VRSS_ClinicalContacts WHERE PatientID=" & PatientID & ")"
        sql = sql & "(SELECT VisitID as ID,   'Visit_NonCPAP_ImpReview' AS Tbl, 'VRSS_REVIEW'          AS Type, VisitDate AS xDate, VisitType AS Description FROM Visit_NonCPAP_ImpReview WHERE PatientID=" & PatientID & ")"
    Case DRSMContactType.CPAP
        sql = sql & "SELECT VisitID as ID, 'Visit_CPAP_ImpReview' AS Tbl, 'CPAP' AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_CPAP_ImpReview WHERE PatientID=" & PatientID
    Case DRSMContactType.Cpx
        sql = sql & "SELECT ResultsID as ID, 'ExResults' AS Tbl, 'CPX' AS Type, TestDate as xDate,  TestType AS Description  FROM ExResults WHERE PatientID=" & PatientID
    Case DRSMContactType.Psgs
        sql = sql & "SELECT StudyID as ID, 'SlpResults' AS Tbl, 'SLP'         AS Type, StudyDate as xDate, StudyType AS Description FROM SlpResults WHERE PatientID=" & PatientID
    Case DRSMContactType.RFTs
        sql = sql & "SELECT ResultsID as ID,  'RftResults' AS Tbl,  'RFT'         AS Type, TestDate as xDate,  TestType AS Description  FROM RftResults WHERE PatientID=" & PatientID
    Case DRSMContactType.RftAndCpx
        sql = "      (SELECT ResultsID as ID, " & DbTables.RftResults & " AS TblEnum, 'RftResults' AS Tbl, 'RFTandCPX' AS Type, TestDate as xDate, TestType AS Description FROM RftResults WHERE PatientID=" & PatientID & ") "
        sql = sql & " UNION "
        sql = sql & "(SELECT ResultsID as ID, " & DbTables.ExResults & " AS TblEnum, 'ExResults'  AS Tbl, 'RFTandCPX' AS Type, TestDate as xDate, TestType AS Description FROM ExResults  WHERE PatientID=" & PatientID & ")"
    Case DRSMContactType.VRSS_Episode
        sql = sql & "SELECT EpisodeID as ID, 'VRSS_Episodes'           AS Tbl,     'VRSS_EPISODE' AS Type, Ref_RecDate as xDate, Ep_EndDate_VRSS as EndDate, 'EPISODE' AS Description FROM VRSS_Episodes WHERE PatientID=" & PatientID
    Case DRSMContactType.VRSS_ClinicalContact
        sql = sql & "SELECT ContactID as ID, 'VRSS_ClinicalContacts'   AS Tbl,    'VRSS_CLINICALCONTACT' AS Type, Contact_Date as xDate, Contact_Type AS Description FROM VRSS_ClinicalContacts "
        sql = sql & "WHERE PatientID=" & PatientID
        If Not IsMissing(VrssEpisodeID) Then sql = sql & " AND EpisodeID=" & VrssEpisodeID
    Case DRSMContactType.VRSS_Review
        sql = sql & "SELECT VisitID as ID, 'Visit_NonCPAP_ImpReview' AS Tbl,     'VRSS_REVIEW' AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_NonCPAP_ImpReview "
        sql = sql & "WHERE PatientID=" & PatientID
        If Not IsMissing(VrssEpisodeID) Then sql = sql & " AND EpisodeID=" & VrssEpisodeID
End Select
sql = sql & " ORDER BY xDate DESC"

Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    If IsDate(Rs("xDate")) Then tempDate = Rs("xDate") Else tempDate = Empty
    Select Case Rs("Type")
        Case "RFT"
            ReDim Preserve Contacts.RFTs(UBound(Contacts.RFTs) + 1)
            Contacts.RFTs(UBound(Contacts.RFTs)).ID = Rs("ID")
            Contacts.RFTs(UBound(Contacts.RFTs)).ContactDate = tempDate
            Contacts.RFTs(UBound(Contacts.RFTs)).Description = Rs("Description") & ""
            Contacts.RFTs(UBound(Contacts.RFTs)).SourceTbl = Rs("Tbl") & ""
            Contacts.RFTs(UBound(Contacts.RFTs)).SourceTblEnum = DbTables.RftResults
        Case "CPX"
            ReDim Preserve Contacts.Cpx(UBound(Contacts.Cpx) + 1)
            Contacts.Cpx(UBound(Contacts.Cpx)).ID = Rs("ID")
            Contacts.Cpx(UBound(Contacts.Cpx)).ContactDate = tempDate
            Contacts.Cpx(UBound(Contacts.Cpx)).Description = Rs("Description") & ""
            Contacts.Cpx(UBound(Contacts.Cpx)).SourceTbl = Rs("Tbl") & ""
            Contacts.Cpx(UBound(Contacts.Cpx)).SourceTblEnum = DbTables.ExResults
         Case "RFTandCPX"
            ReDim Preserve Contacts.RftsOrCpx(UBound(Contacts.RftsOrCpx) + 1)
            Contacts.RftsOrCpx(UBound(Contacts.RftsOrCpx)).ID = Rs("ID")
            Contacts.RftsOrCpx(UBound(Contacts.RftsOrCpx)).ContactDate = tempDate
            Contacts.RftsOrCpx(UBound(Contacts.RftsOrCpx)).Description = Rs("Description") & ""
            Contacts.RftsOrCpx(UBound(Contacts.RftsOrCpx)).SourceTbl = Rs("Tbl") & ""
            Contacts.RftsOrCpx(UBound(Contacts.RftsOrCpx)).SourceTblEnum = Rs("TblEnum")
       Case "SLP"
            ReDim Preserve Contacts.Psgs(UBound(Contacts.Psgs) + 1)
            Contacts.Psgs(UBound(Contacts.Psgs)).ID = Rs("ID")
            Contacts.Psgs(UBound(Contacts.Psgs)).ContactDate = tempDate
            Contacts.Psgs(UBound(Contacts.Psgs)).Description = Rs("Description") & ""
            Contacts.Psgs(UBound(Contacts.Psgs)).SourceTbl = Rs("Tbl") & ""
            Contacts.Psgs(UBound(Contacts.Psgs)).SourceTblEnum = DbTables.SlpResults
        Case "VRSS_EPISODE"
            ReDim Preserve Contacts.Vrss_A(UBound(Contacts.Vrss_A) + 1)
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).ID = Rs("ID")
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).ContactDate = tempDate
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).Description = Rs("Description") & ""
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).SourceTbl = Rs("Tbl") & ""
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).SourceTblEnum = DbTables.VRSS_Episodes
           If IsUDTDate(Rs("EndDate")) Then
                Contacts.Vrss_A(UBound(Contacts.Vrss_A)).DateClosed = Rs("EndDate")
                Contacts.Vrss_A(UBound(Contacts.Vrss_A)).IsClosed = True
            Else
                Contacts.Vrss_A(UBound(Contacts.Vrss_A)).DateClosed = Empty
                Contacts.Vrss_A(UBound(Contacts.Vrss_A)).IsClosed = False
            End If
        Case "CPAP"
            ReDim Preserve Contacts.CPAP(UBound(Contacts.CPAP) + 1)
            Contacts.CPAP(UBound(Contacts.CPAP)).ID = Rs("ID")
            Contacts.CPAP(UBound(Contacts.CPAP)).ContactDate = tempDate
            Contacts.CPAP(UBound(Contacts.CPAP)).Description = Rs("Description") & ""
            Contacts.CPAP(UBound(Contacts.CPAP)).SourceTbl = Rs("Tbl") & ""
            Contacts.CPAP(UBound(Contacts.CPAP)).SourceTblEnum = DbTables.Visit_CPAP_ImpReview
        Case "VRSS_CLINICALCONTACT"
            ReDim Preserve Contacts.Vrss_Contacts(UBound(Contacts.Vrss_Contacts) + 1)
            Contacts.Vrss_Contacts(UBound(Contacts.Vrss_Contacts)).ID = Rs("ID")
            Contacts.Vrss_Contacts(UBound(Contacts.Vrss_Contacts)).ContactDate = tempDate
            Contacts.Vrss_Contacts(UBound(Contacts.Vrss_Contacts)).Description = Rs("Description") & ""
            Contacts.Vrss_Contacts(UBound(Contacts.Vrss_Contacts)).SourceTbl = Rs("Tbl") & ""
            Contacts.Vrss_Contacts(UBound(Contacts.Vrss_Contacts)).SourceTblEnum = DbTables.VRSS_ClinicalContacts
        Case "VRSS_REVIEW"
            ReDim Preserve Contacts.Vrss_R(UBound(Contacts.Vrss_R) + 1)
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).ID = Rs("ID")
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).ContactDate = tempDate
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).Description = Rs("Description") & ""
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).SourceTbl = Rs("Tbl") & ""
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).SourceTblEnum = DbTables.Visit_NonCPAP_ImpReview
    End Select
    Rs.MoveNext
Loop
Rs.Close

Get_DRSMContacts_V2 = Contacts

Exit Function


Get_DRSMContacts_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_DRSMContacts", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_DRSMContacts_old(PatientID As Long, ContactType As DRSMContactType) As DRSMContacts
'Returns lists of record IDs for previous contacts in DRSM

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Contacts As DRSMContacts
Dim Msg As String
Dim tempDate As Date

ReDim Contacts.RFTs(0)
ReDim Contacts.Cpx(0)
ReDim Contacts.CPAP(0)
ReDim Contacts.Psgs(0)
ReDim Contacts.Vrss_A(0)
ReDim Contacts.Vrss_R(0)

On Error GoTo Get_DRSMContacts_Error

Select Case ContactType
    Case DRSMContactType.All
        sql = "      (SELECT ResultsID as ID,   'RFT'         AS Type, TestDate as xDate,  TestType AS Description  FROM RftResults              WHERE PatientID=" & PatientID & ") UNION "
        sql = sql & "(SELECT ResultsID as ID,   'CPX'         AS Type, TestDate as xDate,  TestType AS Description  FROM ExResults               WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT StudyID as ID,     'SLP'         AS Type, StudyDate as xDate, StudyType AS Description FROM SlpResults              WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT VisitID as ID,     'VRSS_ASSESS' AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_NonCPAP_Assess    WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT VisitID as ID,     'CPAP'        AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_CPAP_ImpReview    WHERE PatientID=" & PatientID & ") UNION"
        sql = sql & "(SELECT VisitID as ID,     'VRSS_REVIEW' AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_NonCPAP_ImpReview WHERE PatientID=" & PatientID & ")"
    Case DRSMContactType.CPAP
        sql = sql & "SELECT VisitID as ID,     'CPAP'        AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_CPAP_ImpReview WHERE PatientID=" & PatientID
    Case DRSMContactType.Cpx
        sql = sql & "SELECT ResultsID as ID,   'CPX'         AS Type, TestDate as xDate,  TestType AS Description  FROM ExResults WHERE PatientID=" & PatientID
    Case DRSMContactType.Psgs
        sql = sql & "SELECT StudyID as ID,     'SLP'         AS Type, StudyDate as xDate, StudyType AS Description FROM SlpResults WHERE PatientID=" & PatientID
    Case DRSMContactType.RFTs
        sql = sql & "SELECT ResultsID as ID,   'RFT'         AS Type, TestDate as xDate,  TestType AS Description  FROM RftResults WHERE PatientID=" & PatientID
    Case DRSMContactType.VRSS_Assess
        sql = sql & "SELECT VisitID as ID,     'VRSS_ASSESS' AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_NonCPAP_Assess WHERE PatientID=" & PatientID
    Case DRSMContactType.VRSS_Review
        sql = sql & "SELECT VisitID as ID,     'VRSS_REVIEW' AS Type, VisitDate as xDate, VisitType AS Description FROM Visit_NonCPAP_ImpReview WHERE PatientID=" & PatientID
End Select
sql = sql & " ORDER BY xDate DESC"

Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    If IsDate(Rs("xDate")) Then tempDate = Rs("xDate") Else tempDate = Empty
    Select Case Rs("Type")
        Case "RFT"
            ReDim Preserve Contacts.RFTs(UBound(Contacts.RFTs) + 1)
            Contacts.RFTs(UBound(Contacts.RFTs)).ID = Rs("ID")
            Contacts.RFTs(UBound(Contacts.RFTs)).ContactDate = tempDate
            Contacts.RFTs(UBound(Contacts.RFTs)).Description = Rs("Description") & ""
        Case "CPX"
            ReDim Preserve Contacts.Cpx(UBound(Contacts.Cpx) + 1)
            Contacts.Cpx(UBound(Contacts.Cpx)).ID = Rs("ID")
            Contacts.Cpx(UBound(Contacts.Cpx)).ContactDate = tempDate
            Contacts.Cpx(UBound(Contacts.Cpx)).Description = Rs("Description") & ""
        Case "SLP"
            ReDim Preserve Contacts.Psgs(UBound(Contacts.Psgs) + 1)
            Contacts.Psgs(UBound(Contacts.Psgs)).ID = Rs("ID")
            Contacts.Psgs(UBound(Contacts.Psgs)).ContactDate = tempDate
            Contacts.Psgs(UBound(Contacts.Psgs)).Description = Rs("Description") & ""
        Case "VRSS_ASSESS"
            ReDim Preserve Contacts.Vrss_A(UBound(Contacts.Vrss_A) + 1)
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).ID = Rs("ID")
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).ContactDate = tempDate
            Contacts.Vrss_A(UBound(Contacts.Vrss_A)).Description = Rs("Description") & ""
        Case "CPAP"
            ReDim Preserve Contacts.CPAP(UBound(Contacts.CPAP) + 1)
            Contacts.CPAP(UBound(Contacts.CPAP)).ID = Rs("ID")
            Contacts.CPAP(UBound(Contacts.CPAP)).ContactDate = tempDate
            Contacts.CPAP(UBound(Contacts.CPAP)).Description = Rs("Description") & ""
        Case "VRSS_REVIEW"
            ReDim Preserve Contacts.Vrss_R(UBound(Contacts.Vrss_R) + 1)
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).ID = Rs("ID")
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).ContactDate = tempDate
            Contacts.Vrss_R(UBound(Contacts.Vrss_R)).Description = Rs("Description") & ""
    End Select
    Rs.MoveNext
Loop
Rs.Close

Get_DRSMContacts_old = Contacts

Exit Function


Get_DRSMContacts_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_DRSMContacts", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function


Friend Function Get_ABGsAll(PatientID As Long) As ABGResult()
'Returns from RFTs, Sleep studies and Dom O2 clinic

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim ABGs() As ABGResult
Dim Msg As String

On Error GoTo Get_ABGsAll_Error

'From RFTs
ReDim ABGs(0)
sql = "SELECT * FROM RFTResults WHERE PatientID = " & PatientID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    If InStr(Rs("TestType"), "ABG1") Then
        ReDim Preserve ABGs(UBound(ABGs) + 1)
        ABGs(UBound(ABGs)).Source = ABGResultSource.RespLab
        If IsDate(Rs("TestDate")) Then ABGs(UBound(ABGs)).TestDate = Rs("TestDate") Else ABGs(UBound(ABGs)).TestDate = Empty
        If IsDate(Rs("TestTime")) Then ABGs(UBound(ABGs)).TestTime = Rs("TestTime") Else ABGs(UBound(ABGs)).TestTime = Empty
        ABGs(UBound(ABGs)).FiO2 = Rs("FiO21")
        ABGs(UBound(ABGs)).Condition = ""
        ABGs(UBound(ABGs)).pH = Rs("PH1")
        ABGs(UBound(ABGs)).pCO2 = Rs("PCO21")
        ABGs(UBound(ABGs)).pO2 = Rs("PO21")
        ABGs(UBound(ABGs)).HCO3 = ""
        ABGs(UBound(ABGs)).BE = ""
        ABGs(UBound(ABGs)).SaO2 = ""
        ABGs(UBound(ABGs)).Shunt = Rs("Shunt1") & ""
    End If
    If InStr(Rs("TestType"), "ABG12") Then
        ReDim Preserve ABGs(UBound(ABGs) + 1)
        ABGs(UBound(ABGs)).Source = ABGResultSource.RespLab
        If IsDate(Rs("TestDate")) Then ABGs(UBound(ABGs)).TestDate = Rs("TestDate") Else ABGs(UBound(ABGs)).TestDate = Empty
        If IsDate(Rs("TestTime")) Then ABGs(UBound(ABGs)).TestTime = Rs("TestTime") Else ABGs(UBound(ABGs)).TestTime = Empty
        ABGs(UBound(ABGs)).FiO2 = Rs("FiO22")
        ABGs(UBound(ABGs)).Condition = ""
        ABGs(UBound(ABGs)).pH = Rs("PH2")
        ABGs(UBound(ABGs)).pCO2 = Rs("PCO22")
        ABGs(UBound(ABGs)).pO2 = Rs("PO22")
        ABGs(UBound(ABGs)).HCO3 = ""
        ABGs(UBound(ABGs)).BE = ""
        ABGs(UBound(ABGs)).SaO2 = ""
        ABGs(UBound(ABGs)).Shunt = Rs("Shunt2") & ""
    End If
    Rs.MoveNext
Loop
Rs.Close

'From Sleep studies
sql = "SELECT * FROM SlpResults WHERE PatientID = " & PatientID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    If Rs("R_ABG_Time1") & "" <> "" Then
        ReDim Preserve ABGs(UBound(ABGs) + 1)
        ABGs(UBound(ABGs)).Source = ABGResultSource.SlpLab
        If IsDate(Rs("StudyDate")) Then ABGs(UBound(ABGs)).TestDate = Rs("StudyDate") Else ABGs(UBound(ABGs)).TestDate = Empty
        If IsDate(Rs("R_ABG_Time1")) Then ABGs(UBound(ABGs)).TestTime = Rs("R_ABG_Time1") Else ABGs(UBound(ABGs)).TestTime = Empty
        ABGs(UBound(ABGs)).FiO2 = Rs("R_ABG_FiO21")
        ABGs(UBound(ABGs)).Condition = Rs("R_ABG_Conditions1")
        ABGs(UBound(ABGs)).pH = Rs("R_ABG_pH1")
        ABGs(UBound(ABGs)).pCO2 = Rs("R_ABG_PaCO21")
        ABGs(UBound(ABGs)).pO2 = Rs("R_ABG_PaO21")
        ABGs(UBound(ABGs)).HCO3 = Rs("R_ABG_HCO31")
        ABGs(UBound(ABGs)).BE = Rs("R_ABG_BE1")
        ABGs(UBound(ABGs)).SaO2 = Rs("R_ABG_SaO21")
        ABGs(UBound(ABGs)).Shunt = ""
    End If
    If Rs("R_ABG_Time2") & "" <> "" Then
        ReDim Preserve ABGs(UBound(ABGs) + 1)
        ABGs(UBound(ABGs)).Source = ABGResultSource.SlpLab
        If IsDate(Rs("StudyDate")) Then ABGs(UBound(ABGs)).TestDate = Rs("StudyDate") Else ABGs(UBound(ABGs)).TestDate = Empty
        If IsDate(Rs("R_ABG_Time2")) Then ABGs(UBound(ABGs)).TestTime = Rs("R_ABG_Time2") Else ABGs(UBound(ABGs)).TestTime = Empty
        ABGs(UBound(ABGs)).FiO2 = Rs("R_ABG_FiO22")
        ABGs(UBound(ABGs)).Condition = Rs("R_ABG_Conditions2")
        ABGs(UBound(ABGs)).pH = Rs("R_ABG_pH2")
        ABGs(UBound(ABGs)).pCO2 = Rs("R_ABG_PaCO22")
        ABGs(UBound(ABGs)).pO2 = Rs("R_ABG_PaO22")
        ABGs(UBound(ABGs)).HCO3 = Rs("R_ABG_HCO32")
        ABGs(UBound(ABGs)).BE = Rs("R_ABG_BE2")
        ABGs(UBound(ABGs)).SaO2 = Rs("R_ABG_SaO22")
        ABGs(UBound(ABGs)).Shunt = ""
    End If
    Rs.MoveNext
Loop
Rs.Close

'From Dom O2. ADO/SQL code
'sql = "SELECT * FROM [BLOOD GASES] WHERE HOSP_ID = '" & Get_AustinURFromPtID(PatientID) & "'"
'Rs.Open sql, gDomOxyConnectFull, adOpenStatic, adLockReadOnly
'Do While Not Rs.EOF
'    If Rs("ABG PO2") & "" <> "" And Rs("ABG CO2") & "" <> "" Then
'        ReDim Preserve ABGs(UBound(ABGs) + 1)
'        ABGs(UBound(ABGs)).Source = ABGResultSource.DomO2Clinic
'        If IsDate(Rs("Testdate")) Then ABGs(UBound(ABGs)).TestDate = Rs("Testdate") Else ABGs(UBound(ABGs)).TestDate = Empty
'        ABGs(UBound(ABGs)).TestTime = Empty
'        ABGs(UBound(ABGs)).FiO2 = Rs("ABG FiO2")
'        ABGs(UBound(ABGs)).Condition = Rs("COMMENTS") & ""
'        ABGs(UBound(ABGs)).pH = Rs("ABG PH")
'        ABGs(UBound(ABGs)).pCO2 = Rs("ABG CO2")
'        ABGs(UBound(ABGs)).pO2 = Rs("ABG PO2")
'        ABGs(UBound(ABGs)).HCO3 = ""
'        ABGs(UBound(ABGs)).BE = ""
'        ABGs(UBound(ABGs)).SaO2 = "ABG SAO2"
'        ABGs(UBound(ABGs)).Shunt = ""
'    End If
'    Rs.MoveNext
'Loop
'Rs.Close

Get_ABGsAll = ABGs

Exit Function


Get_ABGsAll_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_ABGsAll", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Public Function GetMedicareNo_formatted(PtID As Long) As String

GetMedicareNo_formatted = GetMedicareNo(PtID, Formatted)

End Function

Public Function GetMedicareNo_unformatted(PtID As Long) As String

GetMedicareNo_unformatted = GetMedicareNo(PtID, Unformatted)

End Function

Private Function GetMedicareNo(PtID As Long, mFormat As MedicareNoFormat) As String

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo GetMedicareNo_Error

'Get medicare num
sql = "SELECT TOP 1 tbl_PatientIdentifier.* FROM tbl_PatientIdentifier "
sql = sql & "INNER JOIN tbl_patient ON tbl_PatientIdentifier.PatientID = tbl_patient.ID "
sql = sql & "WHERE (tbl_PatientIdentifier.PatientID = '" & PtID & "') AND (tbl_patient.Type = 'P') AND (tbl_PatientIdentifier.idTypeCode = 'MC') "
sql = sql & "ORDER BY EffectiveDate DESC,PatientIdentifierID DESC;"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Select Case mFormat
        Case MedicareNoFormat.Formatted:        GetMedicareNo = Format(Rs("PatientIdentifier"), "####-#####-# #")
        Case MedicareNoFormat.Unformatted:      GetMedicareNo = Rs("PatientIdentifier")
    End Select
Else
    GetMedicareNo = ""
End If
Rs.Close

Exit Function

GetMedicareNo_Error:
    Msg = ""
   Call ErrorLog(Msg, "GetMedicareNo", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_GPDetailsFromTrak(PatientID As Long) As GPDetails

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim gP As GPDetails
Dim dummyGP As GPDetails

On Error GoTo Get_GPDetailsFromTrak_Error

gP = dummyGP

'Get GP details
sql = "SELECT tbl_GP.*, tbl_ClinicLocation.* FROM tbl_ClinicLocation INNER JOIN tbl_patient ON "
sql = sql & "tbl_ClinicLocation.site_id = dbo.tbl_patient.ClinicLocationId INNER JOIN tbl_GP ON tbl_patient.GPId = tbl_GP.practitioner_id "
sql = sql & "WHERE tbl_patient.ID =" & PatientID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    If Rs.RecordCount > 1 Then
        MsgBox "There are " & Rs.RecordCount & " sets of GP info in Trak." & vbCrLf & "Only one is diplayed.", vbOKOnly, "GP details"
    End If
    gP.Practitioner_id = Rs("Practitioner_id") & ""
    gP.Surname = Rs("Surname") & ""
    gP.GivenName = Rs("GivenName") & ""
    gP.Title = Rs("prefix") & ""
    gP.PracticeName = Rs("address_line_1") & ""
    gP.Address = Rs("address_line_2") & ""
    gP.Suburb = Rs("Suburb") & ""
    gP.State = Rs("State") & ""
    gP.Postcode = Rs("Postcode") & ""
    gP.Phone = Rs("number") & ""
    gP.Email = Rs("email_address") & ""
End If
Rs.Close

Get_GPDetailsFromTrak = gP

Exit Function

Get_GPDetailsFromTrak_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_GPDetailsFromTrak", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Get_Alerts(PatientID As Long) As Alert()

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim A() As Alert
Dim dummyA As Alert

On Error GoTo Get_Alerts_Error

ReDim A(0)
A(0) = dummyA

'Get alert details
sql = "SELECT NULL AS TypeCode, NULL AS TypeDescription, ZAM_5_2_AllergyReactionDescription AS Reaction, ZAM_20_unk AS CreateDate, AL1_6_IdentificationDate AS LastUpdatedDate, AL1_4_1_AllergySeverityCode as Severity, "
sql = sql & "AL1_3_1_AllergyAlertCode AS AlertCode, AL1_3_2_AllergyAlertDescription AS AlertDescription, NULL AS Comments, "
sql = sql & AlertSourceTable.Cerner & " AS SourceTable "
sql = sql & "FROM tbl_HL7_AllergyAlert_Cerner WHERE patient_ID=" & PatientID
sql = sql & " UNION "
sql = sql & "SELECT IAM_2_1_AllergenTypeCode AS TypeCode, IAM_2_2_AllergenTypeCodeDescription AS TypeDescription, NULL as Reaction, IAM_11_OnsetDate AS CreateDate, NULL AS LastUpdatedDate, NULL as Severity,  "
sql = sql & "IAM_3_1_AllergenCodeMnominicCode AS AlertCode, IAM_3_2_AllergenCodeMnomincDescription AS AlertDescription, IAM_10_AllergenGroupCode AS Comments, "
sql = sql & AlertSourceTable.MedTrak & " AS SourceTable "
sql = sql & "FROM tbl_HL7_AllergyAlert_Trak WHERE patient_ID=" & PatientID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve A(UBound(A) + 1)
    Select Case Rs("SourceTable")
        Case AlertSourceTable.Cerner:   A(UBound(A)).SourceTable = "Cerner"
        Case AlertSourceTable.MedTrak:  A(UBound(A)).SourceTable = "Trak"
    End Select
    A(UBound(A)).AlertGroup = Empty
    A(UBound(A)).AlertType_Code = Rs("TypeCode") & ""
    A(UBound(A)).AlertType_Description = Rs("TypeDescription") & ""
    A(UBound(A)).Alert_Code = Rs("AlertCode") & ""
    A(UBound(A)).Alert_Description = Rs("AlertDescription") & ""
    A(UBound(A)).Alert_Comment = Rs("Comments") & ""
    A(UBound(A)).Alert_severity = Rs("severity") & ""
    A(UBound(A)).Alert_reactions = Rs("Reaction") & ""
    Select Case Rs("SourceTable")
        Case AlertSourceTable.Cerner
            If Len(Rs("LastUpdatedDate")) = 8 Then
                A(UBound(A)).Date_LastUpdated = Format(RIGHT(Rs("LastUpdatedDate"), 2) & "/" & Mid(Rs("LastUpdatedDate"), 5, 2) & "/" & Left(Rs("LastUpdatedDate"), 4), "dd/mm/yyyy")
            Else
                A(UBound(A)).Date_LastUpdated = Empty
            End If
        Case AlertSourceTable.MedTrak
            If IsUDTDate(Rs("CreateDate")) Then
                A(UBound(A)).Date_Create = Format(Rs("CreateDate"), "dd/mm/yyyy")
            Else
                A(UBound(A)).Date_Create = Empty
            End If
    End Select
    Rs.MoveNext
Loop
Rs.Close

Get_Alerts = A

Exit Function

Get_Alerts_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_Alerts", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_GPDetailsFromOldVRSSAssess(PatientID As Long) As GPDetails

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim gP As GPDetails
Dim dummyGP As GPDetails

On Error GoTo Get_GPDetailsFromOldVRSSAssess_Error

gP = dummyGP

'Get GP details
sql = "SELECT VRSS_Episodes.* FROM VRSS_Episodes WHERE patientID=" & PatientID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    gP.Surname = Rs("GP_Name") & ""
    gP.GivenName = ""
    gP.Title = ""
    gP.PracticeName = Rs("GP_PracticeName") & ""
    gP.Address = Rs("GP_Address") & ""
    gP.Suburb = Rs("GP_Suburb") & ""
    gP.State = ""
    gP.Postcode = Rs("GP_PCode") & ""
    gP.Phone = ""
    gP.Email = ""
End If
Rs.Close

Get_GPDetailsFromOldVRSSAssess = gP

Exit Function

Get_GPDetailsFromOldVRSSAssess_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_GPDetailsFromOldVRSSAssess", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_LastRFTEpisode_ID(ThisTestDate As Date, PtID As Long) As Long
' Returns the RFT/Ex recordID for the most recent previous visit to allow retrieval of last episode data

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_LastRFTEpisode_ID_Error

sql = "SELECT Height_rft, TestDate FROM ExResults WHERE PatientID = " & PtID & " AND Testdate<>" & ThisTestDate
sql = sql & " UNION "
sql = sql & "SELECT Height_rft, TestDate FROM RftResults WHERE PatientID = " & PtID & " AND Testdate<>" & ThisTestDate
sql = sql & " ORDER BY TestDate DESC;"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.EOF
    Case True: Get_LastRFTEpisode_ID = ""
    Case False: Get_LastRFTEpisode_ID = Rs("Height_rft") & ""
End Select
Rs.Close

Exit Function


Get_LastRFTEpisode_ID_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_LastRFTEpisode_ID", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Public Function Get_LastRFT_Ht(ThisTestDate As Date, PtID As Long) As String

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_LastRFT_Ht_Error

sql = "SELECT Height_rft, TestDate FROM ExResults WHERE PatientID = " & PtID & " AND Testdate<>" & ThisTestDate
sql = sql & " UNION ALL "
sql = sql & "SELECT Height_rft, TestDate FROM RftResults WHERE PatientID = " & PtID & " AND Testdate<>" & ThisTestDate
sql = sql & " ORDER BY TestDate DESC;"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.EOF
    Case True: Get_LastRFT_Ht = 0
    Case False: Get_LastRFT_Ht = CLng(Rs("Height_rft") & "")
End Select
Rs.Close

Exit Function


Get_LastRFT_Ht_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_LastRFT_Ht", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Public Function Get_LastRFT_Wt(ThisTestDate As Date, PtID As Long) As String

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_LastRFT_Wt_Error

sql = "SELECT Weight, TestDate FROM ExResults WHERE PatientID = " & PtID & " AND Testdate<>" & ThisTestDate
sql = sql & " UNION ALL "
sql = sql & "SELECT Weight, TestDate FROM RftResults WHERE PatientID = " & PtID & " AND Testdate<>" & ThisTestDate
sql = sql & " ORDER BY TestDate DESC;"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.EOF
    Case True: Get_LastRFT_Wt = 0
    Case False: Get_LastRFT_Wt = CLng(Rs("Weight") & "")
End Select
Rs.Close

Exit Function


Get_LastRFT_Wt_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_LastRFT_Wt", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Public Function Get_RaceRft(PtID As Long) As String

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_RaceRft_Error
   
sql = "SELECT List_RacesForRFTs.Race_Description FROM tbl_RM_Demographics INNER JOIN "
sql = sql & "List_RacesForRFTs ON tbl_RM_Demographics.Race_RFTID = List_RacesForRFTs.RaceID "
sql = sql & "WHERE (tbl_RM_Demographics.PatientID = " & PtID & ")"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Get_RaceRft = Rs("Race_Description")
Else
    Get_RaceRft = ""
End If
Rs.Close

Exit Function


Get_RaceRft_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_RaceRft", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_IndigenousStatus(PtID As Long) As String

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_IndigenousStatus_Error
   
sql = "SELECT Race FROM tbl_patient WHERE ID = " & PtID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Get_IndigenousStatus = Map_RefTbl(Rs("Race") & "", Ethnicity, HL7Code_RMText)
Else
    Get_IndigenousStatus = ""
End If
Rs.Close

Exit Function


Get_IndigenousStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_IndigenousStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_PreferredLanguage(PtID As Long) As String
'Returns the Cerner language code (compliant with Vinah)

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_PreferredLanguage_Error
   
sql = "SELECT Language FROM tbl_patient WHERE ID = " & PtID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Get_PreferredLanguage = Map_RefTbl(Rs("Language") & "", Language, MapMode.HL7Code_RMText)
Else
    Get_PreferredLanguage = ""
End If
Rs.Close

Exit Function


Get_PreferredLanguage_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_PreferredLanguage", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSCurrentVentDependentStatus(PatientID As Long, Optional EpisodeID, Optional CensorDate) As VRSS_VentDependentStatus
'If episodeID is supplied, return the latest status limited to that episode, otherwise return the ststus for the latest episode.
'If Censordate is supplied, ignore any records after this date

Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim sql As String
Dim sqlEpisode As String
Dim sqlCensorDate As String
Dim Msg As String
Dim D As VRSS_VentDependentStatus

On Error GoTo Get_VRSSCurrentVentDependentStatus_Error

If IsMissing(EpisodeID) Or CLng(EpisodeID) <= 0 Then
    sqlEpisode = ""
Else
    sqlEpisode = " AND EpisodeID=" & EpisodeID
End If

If IsMissing(CensorDate) Then
    sqlCensorDate = ""
Else
    If IsUDTDate(CDate(CensorDate)) Then
        sqlCensorDate = " AND StatusDate <= " & gDBFunctions.Convert_ToSQLQueryDate(CDate(CensorDate))
    Else
        sqlCensorDate = ""
    End If
End If


sql = "SELECT TOP 1 * FROM VRSS_VentDependentStatus WHERE PatientID =" & PatientID
sql = sql & sqlEpisode & sqlCensorDate
sql = sql & " ORDER BY StatusDate DESC"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.EOF
    Case False
        D.EnteredBy = Rs("EnteredBy")
        D.EpisodeID = Rs("EpisodeID")
        D.PatientID = Rs("PatientID")
        D.Status_code = Rs("Status_code")
        D.Status_text = Rs("Status_text")
        If IsUDTDate(Rs("StatusDate")) Then D.StatusDate = Rs("StatusDate") Else D.StatusDate = Empty
        D.StatusTime = Rs("StatusTime") & ""
        D.StatusID = Rs("StatusID")
    Case True
        D.EnteredBy = ""
        D.EpisodeID = 0
        D.PatientID = PatientID
        D.Status_code = 6
        D.Status_text = "Not set"
        D.StatusDate = Empty
        D.StatusTime = ""
        D.StatusID = 0
End Select
Rs.Close

Get_VRSSCurrentVentDependentStatus = D

Exit Function


Get_VRSSCurrentVentDependentStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSCurrentVentDependentStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function



Friend Function Get_VRSSStatus(PatientID As Long) As VRSS_PtStatus

Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim S As VRSS_PtStatus

On Error GoTo Get_VRSSStatus_Error

'Definitions
' IsVRSS = has one or more episode/contact records (although should never be a contact without an episode
' HasOpenEpisode = one or more open episodes ie no close date

sql = "SELECT TOP 1 Ref_RecDate, Ep_EndDate_VRSS FROM VRSS_Episodes WHERE PatientID =" & PatientID & " ORDER BY Ref_RecDate DESC"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.EOF
    Case True
        'Check if any orphaned contacts
        'sql = "SELECT ContactID FROM VRSS_ClinicalContacts WHERE PatientID =" & PatientID
        sql = "SELECT VisitID FROM Visit_NonCPAP_ImpReview WHERE PatientID =" & PatientID
        Rs1.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        If Rs1.EOF Then
            S.IsVRSS = False
            S.HasOpenEpisode = False
            S.LatestEpisode_ID = 0
            S.LatestEpisode_StartDate = Empty
            S.LatestEpisode_EndDate = Empty
        Else
            S.IsVRSS = True
            S.HasOpenEpisode = False
            S.LatestEpisode_ID = 0
            S.LatestEpisode_StartDate = Empty
            S.LatestEpisode_EndDate = Empty
        End If
        Rs1.Close
    Case False
        S.IsVRSS = True
            S.LatestEpisode_ID = 0
            If IsUDTDate(Rs("Ref_RecDate")) Then S.LatestEpisode_StartDate = Rs("Ref_RecDate") Else S.LatestEpisode_StartDate = Empty
        If IsUDTDate(Rs("Ep_EndDate_VRSS")) Then
            S.HasOpenEpisode = False
            S.LatestEpisode_EndDate = Rs("Ep_EndDate_VRSS")
        Else
            S.HasOpenEpisode = True
            S.LatestEpisode_EndDate = Empty
        End If
End Select
Rs.Close

Get_VRSSStatus = S

Exit Function


Get_VRSSStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_EquipIDFromAllocationID(AllocationID As Long) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_EquipIDFromAllocationID_Error

sql = "SELECT EquipID FROM VRSS_EquipAllocations WHERE AllocationID =" & AllocationID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.EOF
    Case True
        Get_EquipIDFromAllocationID = 0
    Case False
        Rs.MoveFirst
        Get_EquipIDFromAllocationID = Rs("EquipID")
End Select
Rs.Close

Exit Function


Get_EquipIDFromAllocationID_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_EquipIDFromAllocationID", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_VRSSLastContactDate(EpisodeID As Long, CensorDate As Date) As Date
'Ignore contacts after CensorDate

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim D As Date

On Error GoTo Get_VRSSLastContactDate_Error

If Not IsUDTDate(CensorDate) Or EpisodeID = 0 Then
    Get_VRSSLastContactDate = Empty
    Exit Function
End If

sql = "SELECT TOP (1) VisitDate FROM Visit_NonCPAP_ImpReview WHERE EpisodeID=" & EpisodeID
sql = sql & " AND VisitDate <= " & gDBFunctions.Convert_ToSQLQueryDate(CensorDate) & " ORDER BY VisitDate DESC"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    D = Empty
Else
    If IsUDTDate(Rs("VisitDate")) Then D = Rs("VisitDate") Else D = Empty
End If

Get_VRSSLastContactDate = D

Exit Function

Get_VRSSLastContactDate_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSLastContactDate", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_VRSSNumberOfContacts(EpisodeID As Long, CensorDate As Date) As Integer
'Ignore contacts after CensorDate

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim D As Date

On Error GoTo Get_VRSSNumberOfContacts_Error

If Not IsUDTDate(CensorDate) Or EpisodeID = 0 Then
    Get_VRSSNumberOfContacts = -1
    Exit Function
End If

sql = "SELECT count(VisitID) AS Num FROM Visit_NonCPAP_ImpReview WHERE EpisodeID=" & EpisodeID & " AND VisitDate <= " & gDBFunctions.Convert_ToSQLQueryDate(CensorDate)
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    Get_VRSSNumberOfContacts = 0
Else
    Get_VRSSNumberOfContacts = Rs("Num")
End If

Exit Function

Get_VRSSNumberOfContacts_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSNumberOfContacts", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function



Friend Function Get_VRSSLatestPtData(PatientID As Long, Optional EpisodeID As Long) As Scripting.Dictionary

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim Dic As New Scripting.Dictionary
Dim D As PtDemographics
Dim dummyD As PtDemographics
Dim E As VRSS_Episode
Dim dummyE As VRSS_Episode
Dim locEpisodeID As Long

On Error GoTo Get_VRSSLatestPtData_Error

If IsMissing(EpisodeID) Or EpisodeID = 0 Then locEpisodeID = Pt.Get_VRSSCurrentEpisodeID(PatientID) Else locEpisodeID = EpisodeID

Set Dic = Pt.Get_DicOfLatestVrssPtDataItems()

'Get demographics
D = dummyD
D = Pt.Get_DemographicsFromDB(PatientID)
Dic("PatientID") = D.ID
Dic("Title") = D.Title
Dic("Surname") = D.Surname
Dic("Firstname") = D.Firstname
Dic("URforDisplay") = D.URforDisplay
Dic("Gender") = D.Gender
Dic("DOB") = D.DOB
Dic("Address1") = D.Address1
Dic("Address2") = D.Address2
Dic("Suburb") = D.Suburb
Dic("Postcode") = D.Postcode
Dic("PhoneHome") = D.PhoneHome
Dic("PhoneMobile") = D.PhoneMobile
Dic("PhoneWork") = D.PhoneWork
Dic("Email") = D.Email

'Latest episode data
E = dummyE
E = Pt.Get_VRSSEpisodeFromDB(locEpisodeID)
Dic("EpisodeID") = locEpisodeID
Dic("EpisodeStartDate") = E.EpisodeStartDate
If IsUDTDate(E.EpisodeEndDate) Then Dic("EpisodeEndDate") = E.EpisodeEndDate Else Dic("EpisodeEndDate") = ""
Dic("EpisodeEndReason") = E.Vinah_EpisodeEndReason
Dic("PrimaryDiagnosis") = E.Diagnosis_Pri
Dic("SecondaryDiagnosis") = E.Diagnosis_Sec
Dic("VrssRegion") = E.VRSS_Region
Dic("MelwayMapRef") = E.Melways
Dic("OtherClinicalInfo") = E.OtherClinicalInfo




Set Get_VRSSLatestPtData = Dic

Exit Function


Get_VRSSLatestPtData_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSLatestPtData", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_DicOfLatestVrssPtDataItems() As Scripting.Dictionary

Dim Msg As String
Dim D As New Scripting.Dictionary

On Error GoTo Get_ListOfLatestVrssPtDataItems_Error

D.Add "PatientID", ""
D.Add "Title", ""
D.Add "Surname", ""
D.Add "Firstname", ""
D.Add "URforDisplay", ""
D.Add "Gender", ""
D.Add "DOB", ""
D.Add "Address1", ""
D.Add "Address2", ""
D.Add "Suburb", ""
D.Add "Postcode", ""
D.Add "PhoneHome", ""
D.Add "PhoneMobile", ""
D.Add "PhoneWork", ""
D.Add "Email", ""

D.Add "EpisodeStartDate", ""
D.Add "EpisodeEndDate", ""
D.Add "EpisodeEndReason", ""
D.Add "PrimaryDiagnosis", ""
D.Add "SecondaryDiagnosis", ""
D.Add "VrssRegion", ""
D.Add "MelwayMapRef", ""
D.Add "OtherClinicalInfo", ""
D.Add "NumberOfContacts", ""
D.Add "LastContactDate", ""
D.Add "LastContactMonthsAgo", ""
D.Add "No of machines", ""
D.Add "Machine details", ""
D.Add "Vent dependency", ""
D.Add "Latest vent prescription(s)", ""

GoTo Here

D.Add "VisitDate_V", ""
D.Add "VisitTime", ""
D.Add "VisitType", ""
D.Add "Clinician", ""
D.Add "Model", ""
D.Add "SerialNo", ""
D.Add "MedPhysicsNumber", ""
D.Add "PurchaseDate", ""
D.Add "VentilationMode", ""
D.Add "PrescribedUsage", ""
D.Add "Hourmeter", ""
D.Add "Humidifier", ""
D.Add "Sett_RR", ""
D.Add "Sett_TidalVolume", ""
D.Add "Sett_IPAP", ""
D.Add "Sett_EPAP", ""

D.Add "VisitDate_M", ""
D.Add "MaskManufacturer", ""
D.Add "MaskModel", ""
D.Add "FrameSize", ""
D.Add "CushionSize", ""
D.Add "CushionType", ""
D.Add "HeadgearSize", ""
D.Add "HeadgearType", ""
D.Add "Chinstrap", ""

Here:

Set Get_DicOfLatestVrssPtDataItems = D

Exit Function

Get_ListOfLatestVrssPtDataItems_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_ListOfLatestVrssPtDataItems", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSEquipAllocations(PatientID As Long, MachineFilter As pdrFlags, Optional CensorDate) As MachineAllocation()

'MachineStatus options: pdrAllocated, pdrNotAllocated, pdrAll

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim MachineID() As Long
Dim All() As MachineAllocation
Dim Keep() As MachineAllocation
Dim sqlCensorOnDate As String
Dim Include As Boolean
Dim Regions() As String

On Error GoTo Get_VRSSEquipAllocations_Error

ReDim All(0)
ReDim MachineID(0)
ReDim Regions(0)
ReDim Keep(0)

If IsMissing(CensorDate) Then
    sqlCensorOnDate = ""
Else
    If IsUDTDate(CensorDate) Then
        sqlCensorOnDate = " AND VisitDate <=" & gDBFunctions.Convert_ToSQLQueryDate(CDate(CensorDate))
    Else
        MsgBox "Error", vbOKOnly, "Problem retrieving machine allocations for patient ID " & PatientID & vbCrLf & "Invalid censor date"
        Get_VRSSEquipAllocations = All
        Exit Function
    End If
End If

'First get all VRSS machines ever allocated to this patient
sql = "SELECT DISTINCT MachineID FROM Visit_NonCPAP_ImpReview WHERE PatientID = " & PatientID & sqlCensorOnDate
sql = sql & " AND MachineID>0 GROUP BY MachineID"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve MachineID(UBound(MachineID) + 1)
    MachineID(UBound(MachineID)) = Rs("MachineID")
    Rs.MoveNext
Loop
Rs.Close

'Now get the details for each machine
All = gReg.Get_MachineAllocation_Latest(MachineID, Regions, CensorDate)
For i = 1 To UBound(All)
    Select Case MachineFilter
        Case pdrFlags.pdrAll
            Include = True
        Case pdrFlags.pdrAllocated
            If All(i).D_PatientID = PatientID Then Include = True Else Include = False
        Case pdrFlags.pdrNotAllocated
            If All(i).D_PatientID <> PatientID Then Include = True Else Include = False
        Case Else
            Include = True
    End Select
    If Include Then
        ReDim Preserve Keep(UBound(Keep) + 1)
        Keep(UBound(Keep)) = All(i)
    End If
Next i

Get_VRSSEquipAllocations = Keep

Exit Function


Get_VRSSEquipAllocations_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSEquipAllocations", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSEquipAllocations1(PatientID As Long) As Long()
'Returns ModelIDs

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim ModelIDs() As Long

On Error GoTo Get_VRSSEquipAllocations1_Error

ReDim ModelIDs(0)

'Get all VRSS models ever allocated to this patient
sql = "SELECT CPAP_VentEquip.ModelID FROM Visit_NonCPAP_ImpReview INNER JOIN CPAP_VentEquip ON "
sql = sql & "Visit_NonCPAP_ImpReview.MachineID = CPAP_VentEquip.MachineID "
sql = sql & "WHERE (Visit_NonCPAP_ImpReview.PatientID = " & PatientID & ") "
sql = sql & "AND (CPAP_VentEquip.MachineID > 0) GROUP BY CPAP_VentEquip.ModelID"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve ModelIDs(UBound(ModelIDs) + 1)
    If IsNull(Rs("ModelID")) Then
        ModelIDs(UBound(ModelIDs)) = 0
    Else
        ModelIDs(UBound(ModelIDs)) = Rs("ModelID")
    End If
    Rs.MoveNext
Loop
Rs.Close

Get_VRSSEquipAllocations1 = ModelIDs

Exit Function


Get_VRSSEquipAllocations1_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSEquipAllocations1", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Get_DeathInfo(PatientID As Long) As DeathInfo
'Trak [Date of Death] = valid date PLUS Trak [DeathIndicator] = Y   --> Accurate deathdate
'Trak [Date of Death] = empty      PLUS Trak [DeathIndicator] = Y   --> estimated death date (text not available to ResMed)

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim Info As DeathInfo

On Error GoTo Get_DeathDate_Error

sql = "SELECT DateOfDeath, DeathIndicator FROM tbl_Patient WHERE ID =" & PatientID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Select Case IsUDTDate(Rs("DateOfDeath"))
        Case True
            Info.DeathDate = Rs("DateOfDeath")
            Info.DeathStatusString = Map_RefTbl(Rs("DeathIndicator"), DeathIndicator, HL7Code_RMText)
        Case False
            If Rs("DeathIndicator") = "Y" Then
                Info.DeathStatusString = Map_RefTbl(Rs("DeathIndicator"), DeathIndicator, HL7Code_RMText)
                Info.DeathDate = "Estimated death date"
            Else
                Info.DeathStatusString = Map_RefTbl(Rs("DeathIndicator"), DeathIndicator, HL7Code_RMText)
                Info.DeathDate = ""
            End If
    End Select
Else
    Info.DeathDate = ""
    Info.DeathStatusString = ""
End If
Rs.Close

Get_DeathInfo = Info

Exit Function


Get_DeathDate_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_DeathDate", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Public Function IsVRSSEpisodeOpen(EpisodeID As Long) As Boolean

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo IsVRSSEpisodeOpen_Error

sql = "SELECT Ep_EndDate_VRSS FROM VRSS_Episodes WHERE EpisodeID =" & EpisodeID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    If IsUDTDate(Rs("Ep_EndDate_VRSS")) Then
        IsVRSSEpisodeOpen = False
    Else
        IsVRSSEpisodeOpen = True
    End If
End If
Rs.Close

Exit Function


IsVRSSEpisodeOpen_Error:
    Msg = ""
    Call ErrorLog(Msg, "IsVRSSEpisodeOpen", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_ReportInfoFromCpxUDT(X As Cpx) As ReportInfoSection

Dim H As ReportInfoSection
Dim Msg As String
   
On Error GoTo Get_ReportInfoFromCpxUDT_Error

H.ClinicalNote = X.ClinicalNote
H.Resp_Hb = X.Hb
H.Resp_Hb_Info = X.Hb_Info
H.Lab = X.Lab
H.Resp_LastBd = X.LastBd
H.Resp_Packyrs = X.Packyrs
H.Resp_RefLocation = X.RefLocation
H.Resp_ReportCopyTo = X.ReportCopyTo
H.RequestingMO = X.RequestingMO
H.Resp_Smoke = X.Smoke
H.Resp_Source = X.Source
H.TestDate = X.TestDate
H.TestTime = X.TestTime
H.Resp_TestType = X.TestType
H.ThisVisit_height = X.ThisVisit_height
H.ThisVisit_weight = X.ThisVisit_weight

Get_ReportInfoFromCpxUDT = H

Exit Function

Get_ReportInfoFromCpxUDT_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_ReportInfoFromCpxUDT", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_ReportInfoFromCpapUDT(X As CPAPClinic) As ReportInfoSection

Dim H As ReportInfoSection
Dim Msg As String
   
On Error GoTo Get_ReportInfoFromCpapUDT_Error

H.TestDate = X.Visit_Date
H.ClinicalNote = X.Visit_Note
H.Cpap_ReasonForReferral = X.Clinic_ReferralReason
H.Cpap_VisitLocation = X.Visit_Location
H.Cpap_ReferralMode = X.Clinic_ReferralMode
H.RequestingMO = X.Clinic_ReferredBy
H.Cpap_ReportTo = X.Clinic_ReportTo
H.ThisVisit_weight = X.Weight

Get_ReportInfoFromCpapUDT = H

Exit Function

Get_ReportInfoFromCpapUDT_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_ReportInfoFromCpapUDT", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_ReportInfoFromRftUDT(R As RFTs) As ReportInfoSection

Dim H As ReportInfoSection
Dim Msg As String

On Error GoTo Get_ReportInfoFromRftUDT_Error

H.ClinicalNote = R.E_ClinicalNote
H.Resp_Hb = R.Hb
H.Resp_Hb_Info = R.HbInfo
H.Lab = R.E_Lab
H.Resp_LastBd = R.LastBd
H.Resp_Packyrs = R.E_Packyrs
H.Resp_RefLocation = R.E_RefLocation
H.Resp_ReportCopyTo = R.E_ReportCopyTo
H.RequestingMO = R.E_RequestingMO
H.Resp_Smoke = R.E_Smoke
'H.Resp_Source = R.Source
H.TestDate = R.E_TestDate
H.TestTime = R.E_TestTime
H.Resp_TestType = R.TestType
H.ThisVisit_height = R.E_ThisVisit_height
H.ThisVisit_weight = R.E_ThisVisit_weight

Get_ReportInfoFromRftUDT = H

Exit Function

Get_ReportInfoFromRftUDT_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_ReportInfoFromRftUDT", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Friend Function Get_ReportInfoFromSlpUDT(S As SlpResults) As ReportInfoSection

Dim H As ReportInfoSection
Dim Msg As String

On Error GoTo Get_ReportInfoFromSlpUDT_Error

H.ClinicalNote = S.ReasonForTest
H.Lab = S.Lab
H.Slp_ReportDest_1 = S.ReportDest_1
H.Slp_ReportDest_2 = S.ReportDest_2
H.RequestingMO = S.RequestingMO
H.TestDate = S.StudyDate
H.TestTime = S.StudyTime
H.Slp_StudyType = S.StudyType
H.ThisVisit_height = S.Height_slp
H.ThisVisit_weight = S.Weight
H.Slp_TreatmentMode = S.TreatmentMode

Get_ReportInfoFromSlpUDT = H

Exit Function

Get_ReportInfoFromSlpUDT_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_ReportInfoFromSlpUDT", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function

Friend Function Get_ReportSectionInfoFromCpxUDT(X As Cpx) As ReportSection

Dim Rs As ReportSection
Dim Msg As String

On Error GoTo Get_ReportSectionInfoFromCpxUDT_Error

Rs.Comments = X.Comments
Rs.Race_Rft = X.D_Race_Rft
Rs.Report = X.Report
Rs.ReportDate = X.ReportDate
Rs.Reporter = X.Reporter
Rs.Scientist = X.Scientist
Rs.ReportStatus = X.ReportStatus

Get_ReportSectionInfoFromCpxUDT = Rs

Exit Function

Get_ReportSectionInfoFromCpxUDT_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_ReportSectionInfoFromCpxUDT", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_ReportSectionInfoFromSlpUDT(S As SlpResults) As ReportSection

Dim Rs As ReportSection
Dim Msg As String
   
On Error GoTo Get_ReportSectionInfoFromSlpUDT_Error

Rs.Comments = S.TechnicalNote
Rs.Race_Rft = S.D_Race_Rft
Rs.Report = S.Report
Rs.ReportDate = S.ReportDate
Rs.Reporter = S.Reporter_MO
Rs.Scientist = S.ScoredBy
Rs.ReportStatus = S.ReportStatus

Get_ReportSectionInfoFromSlpUDT = Rs

Exit Function


Get_ReportSectionInfoFromSlpUDT_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_ReportSectionInfoFromSlpUDT", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume

End Function


Friend Function Get_ReportSectionInfoFromRftUDT(R As RFTs) As ReportSection

Dim Rs As ReportSection
Dim Msg As String

On Error GoTo Get_ReportSectionInfoFromRftUDT_Error

Rs.Comments = R.TechComments
Rs.Race_Rft = R.D_Race_Rft
Rs.Report = R.Report
Rs.ReportDate = R.ReportDate
Rs.Reporter = R.Reporter
Rs.Scientist = R.Scientist
Rs.ReportStatus = R.ReportStatus

Get_ReportSectionInfoFromRftUDT = Rs

Exit Function

Get_ReportSectionInfoFromRftUDT_Error:
    Msg = ""

   Call ErrorLog(Msg, "Get_ReportSectionInfoFromRftUDT", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_CpxFromDB(CpxID As Long) As Cpx

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim X As Cpx
Dim emptyUDT As Cpx
Dim Msg As String

On Error GoTo Get_CpxFromDB_Error

'Clear UDT
X = emptyUDT

'Retrieve Cpx data
sql = "SELECT tbl_Name.*, tbl_RM_Demographics.*, ExResults.*, "
sql = sql & "tbl_patient.ID, tbl_patient.URNumber, tbl_patient.Type AS Expr1, tbl_patient.DOB, tbl_patient.Gender, "
sql = sql & "tbl_patient.MaritalStatus, tbl_patient.Language, tbl_patient.Religion, tbl_patient.Race, "
sql = sql & "tbl_patient.Nationality, tbl_patient.Citizenship, tbl_patient.MothersID, tbl_patient.DateofDeath, "
sql = sql & "tbl_patient.DeathIndicator, tbl_patient.IdentityUnknown, tbl_patient.PrimaryFacility, "
sql = sql & "tbl_patient.LastModifiedBy, tbl_patient.LastModifiedDate, tbl_patient.MergedUR, "
sql = sql & "tbl_Address.Type AS Expr2, tbl_Address.TypeID AS Expr3, tbl_Address.AddressID, tbl_Address.Address1, tbl_Address.Address2, "
sql = sql & "tbl_Address.Suburb, tbl_Address.State, tbl_Address.Postcode, tbl_Address.Country , tbl_Address.AddressType "
sql = sql & "FROM tbl_patient INNER JOIN "
sql = sql & "tbl_Name ON tbl_patient.ID = tbl_Name.TypeID INNER JOIN ExResults ON "
sql = sql & "tbl_patient.ID = ExResults.PatientID LEFT OUTER JOIN tbl_Address ON tbl_patient.ID = "
sql = sql & "tbl_Address.TypeID LEFT OUTER JOIN tbl_RM_Demographics ON tbl_patient.ID = tbl_RM_Demographics.PatientID "
sql = sql & "WHERE (tbl_Name.Type = 'P') AND (tbl_Address.AddressType = '4' OR tbl_Address.AddressType IS NULL) AND (tbl_Address.Type = 'P') AND "
sql = sql & "(tbl_Name.Type = 'P') AND (tbl_patient.Type = 'P') AND (ExResults.ResultsID = " & CpxID & ")"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving CPX results - aborted"
    Exit Function
End If
    
X.D_DOB = Format(Rs("DOB"), "dd/mm/yyyy")
X.D_Gender = Map_RefTbl(Rs("Gender"), Gender, HL7Code_RMText)
X.D_Race_Rft = CurrentPt.Get_RaceRft(Rs("PatientID"))
X.D_Address1 = Rs("Address1") & ""
X.D_Address2 = Rs("Address2") & ""
X.D_FirstName = Rs("FirstName") & ""
X.D_Middlename = Rs("Middlename") & ""
X.D_Postcode = Rs("Postcode") & ""
X.D_State = Rs("State") & ""
X.D_Suburb = Rs("Suburb") & ""
X.D_Surname = Rs("Surname")
X.D_UR_PrimaryAustin = Rs("URNumber")
X.D_URforDisplay = Get_PtString(X.PtID, cHS.CurrentVAED, DisplayUR)
    
X.RecordID = Rs("ResultsID")
X.PtID = Rs("PatientID")
X.Lab = Rs("Campus") & ""
X.RequestingMO = Rs("MO") & ""
X.RequestingMO_Pn = Rs("RequestingMO_Pn") & ""
X.RefLocation = Rs("RefLocation") & ""
Select Case IsNull(Rs("Request_HealthServiceID"))
    Case True
        X.Ref_HealthService = ""
        X.Ref_HealthServiceID = 0
    Case False
        X.Ref_HealthService = cHS.Get_HSFromVAED(Rs("Request_HealthServiceID"))
        X.Ref_HealthServiceID = Rs("Request_HealthServiceID") & ""
End Select
X.ReportCopyTo = Rs("Report_CopyTo") & ""
X.TestType = Rs("TestType")
X.TestDate = Format(Rs("TestDate"), "dd/mm/yyyy")
If IsDate(Rs("RequestDate")) Then X.RequestDate = Format(Rs("RequestDate"), "dd/mm/yyyy") Else X.RequestDate = Empty
If IsDate(Rs("TestTime")) Then X.TestTime = Format(Rs("TestTime"), "hh:mm") Else X.TestTime = Empty
X.Scientist = Rs("scientist") & ""
X.Comments = Rs("comments") & ""
X.LastBd = Rs("LastBD") & ""
X.Smoke = Rs("Smoke") & ""
X.Packyrs = Rs("Packyrs") & ""
X.ThisVisit_weight = Rs("Weight") & ""
X.ThisVisit_height = Rs("Height_rft") & ""
X.ClinicalNote = Rs("ClinicalNote") & ""
X.Source = Rs("Source") & ""
X.AdmissionStatus = Rs("AdmissionStatus") & ""
X.BillingMO = Rs("BillingMO") & ""
X.BillingItems = Rs("BillingItems") & ""
X.BillingStatus = Rs("BillingStatus") & ""
X.BillingMO_ProviderNumber = Rs("BillingMO_ProviderNumber") & ""
X.BillingSessionID = Rs("BillingSessionID") & ""

If Rs("BE") = 0 Then X.ABG_BE = "" Else X.ABG_BE = Rs("BE") & ""
If Rs("HCO3") = 0 Then X.ABG_HCO3 = "" Else X.ABG_HCO3 = Rs("HCO3") & ""
If Rs("pCO2") = 0 Then X.ABG_PaCO2 = "" Else X.ABG_PaCO2 = Rs("pCO2") & ""
If Rs("pO2") = 0 Then X.ABG_PaO2 = "" Else X.ABG_PaO2 = Rs("pO2") & ""
If Rs("BE") = 0 Then X.ABG_pH = "" Else X.ABG_pH = Rs("BE") & ""
If Rs("exSaO2") = 0 Then X.ABG_SaO2 = "" Else X.ABG_SaO2 = Rs("exSaO2") & ""

X.Borg_Dyspnoea = Rs("Dysp") & ""
X.Borg_Legs = Rs("Legs") & ""
X.Borg_Other = Rs("Other") & ""
X.Borg_OtherSymptom = Rs("OtherSympt") & ""

X.BP_Post = Rs("PostBP") & ""
X.BP_Rest = Rs("RestBP") & ""

X.EndLoad = Rs("EndLoad") & ""
X.FEV1 = Rs("FEV1") & ""
X.FVC = Rs("FVC") & ""
X.SpO2_Post = Rs("PostSaO2") & ""
X.Stages = Rs("Stages") & ""

X.string_HR = Rs("HRstring") & ""
X.string_Load = Rs("Loadstring") & ""
X.string_PetCO2 = Rs("PetCO2string") & ""
X.string_PetO2 = Rs("PetO2string") & ""
X.string_SpO2 = Rs("SaO2string") & ""
X.string_VCO2 = Rs("VCO2string") & ""
X.string_VE = Rs("VEstring") & ""
X.string_VO2 = Rs("VO2string") & ""
X.string_Vt = Rs("Vtstring") & ""

X.HR = Decode_CpxString(Rs("HRstring") & "")
X.Load = Decode_CpxString(Rs("Loadstring") & "")
X.PetCO2 = Decode_CpxString(Rs("PetCO2string") & "")
X.PetO2 = Decode_CpxString(Rs("PetO2string") & "")
X.SpO2 = Decode_CpxString(Rs("SaO2string") & "")
X.VCO2 = Decode_CpxString(Rs("VCO2string") & "")
X.VO2 = Decode_CpxString(Rs("VO2string") & "")
X.Vt = Decode_CpxString(Rs("Vtstring") & "")
X.VE = Decode_CpxString(Rs("VEstring") & "")

'Work out maximum values
X.MaxVt = 0
X.MaxVE = 0
X.MaxVO2 = 0
X.MaxHR = 0
X.MaxVCO2 = 0
X.MaxVeVO2 = 0
X.MaxO2Pulse = 0
For i = 1 To UBound(X.Load)
    If Val(X.Vt(i)) > Val(X.MaxVt) Then X.MaxVt = X.Vt(i)
    If Val(X.VE(i)) > Val(X.MaxVE) Then X.MaxVE = X.VE(i)
    If Val(X.VO2(i)) > Val(X.MaxVO2) Then X.MaxVO2 = X.VO2(i)
    If Val(X.VCO2(i)) > Val(X.MaxVCO2) Then X.MaxVCO2 = X.VCO2(i)
    If Val(X.HR(i)) > Val(X.MaxHR) Then X.MaxHR = X.HR(i)
    If Val(X.VO2(i)) > 0 Then
        If Val(X.VE(i)) / Val(X.VO2(i)) > Val(X.MaxVeVO2) Then X.MaxVeVO2 = Format(Val(X.VE(i)) / Val(X.VO2(i)), "0.0")
    End If
    If 1000 * Val(X.VO2(i)) / Val(X.HR(i)) > Val(X.MaxO2Pulse) Then X.MaxO2Pulse = Format(1000 * Val(X.VO2(i)) / Val(X.HR(i)), "0.0")
Next i
X.MaxVO2 = Format(X.MaxVO2, "0.00")
X.MaxVE = Format(X.MaxVE, "###")
If Val(X.ThisVisit_weight) > 0 Then X.MaxVO2ml_kg = Format(1000 * X.MaxVO2 / Val(X.ThisVisit_weight), "##.0") Else X.MaxVO2ml_kg = ""

X.Report = Rs("Report") & ""
X.Reporter = Rs("Reporter") & ""
If IsDate(Rs("ReportDate")) Then X.ReportDate = Format(Rs("ReportDate"), "dd/mm/yyyy") Else X.ReportDate = Empty
X.ReportStatus = Rs("ReportStatus")
X.VerifiedBy = Rs("VerifiedBy") & ""
If IsDate(Rs("VerifiedDate")) Then X.VerifiedDate = Format(Rs("VerifiedDate"), "dd/mm/yyyy") Else X.VerifiedDate = Empty

Rs.Close

'Get normal values
X.predMaxVE = Format(GetExPredicted(X.D_DOB, X.TestDate, X.D_Gender, X.ThisVisit_height, X.ThisVisit_weight, "VEMAX", "MPV", X.FEV1, X.FVC, 0), "###")
X.predMaxHR = Format(GetExPredicted(X.D_DOB, X.TestDate, X.D_Gender, X.ThisVisit_height, X.ThisVisit_weight, "HRMAX", "MPV"), "###")
X.predMaxVO2 = Format(GetExPredicted(X.D_DOB, X.TestDate, X.D_Gender, X.ThisVisit_height, X.ThisVisit_weight, "VO2MAX", "MPV"), "#.##")
X.predMaxW = Format(GetExPredicted(X.D_DOB, X.TestDate, X.D_Gender, X.ThisVisit_height, X.ThisVisit_weight, "WMAX", "MPV"), "###")
If Val(X.predMaxHR) > 0 Then X.predMaxO2Pulse = Format(Val(1000 * Val(X.predMaxVO2)) / Val(X.predMaxHR), "0.0") Else X.predMaxO2Pulse = ""
X.predMaxVt = Format(GetExPredicted(X.D_DOB, X.TestDate, X.D_Gender, X.ThisVisit_height, X.ThisVisit_weight, "VTMAX", "MPV", FEV1, FVC, 0), "##.#")

Get_CpxFromDB = X

Exit Function

Get_CpxFromDB_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_CpxFromDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Private Function Decode_CpxString(S As String) As String()
'comma delimited

Dim sData() As String
Dim i As Integer
Dim j As Integer
Dim Item As String
Dim Msg As String
Const Delim As String = ","

On Error GoTo Decode_CpxString_Error

ReDim sData(0)

'Add a terminal comma to make decoding easier
S = S & Delim
i = 1
Do While i < Len(S)
    j = InStr(i, S, Delim)
    Item = Mid(S, i, j - i)
    ReDim Preserve sData(UBound(sData) + 1)
    sData(UBound(sData)) = Item
    i = j + 1
Loop
Decode_CpxString = sData

Exit Function


Decode_CpxString_Error:
    Msg = ""
    Call ErrorLog(Msg, "Decode_CpxString", "DisplayCycle", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_SlpGraphicFromDB(RecordID As Long) As String
'Creates a temp binary file of the graphic and returns filepath

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim S As ADODB.Stream
Dim Msg As String
Dim filepath As String
Dim Tbl As String
Dim Fname_local As String
Dim Filetype As Integer

On Error GoTo Get_SlpGraphicFromDB_Error

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'code added by WRR 28/5/2013 to obtain graphic file from server (if exists) rather than graphics field in SlpResults table
'first step is to check the filepointers table for a record that matches (i) the record ID, (ii) 'slpResults' in the tablename field (iii) 'Sleep graphics' in the description field
'if a match  and the file is found then copy the file to a temporary directory and pass the filename back to the calling function
Tbl = "SlpResults"
Fname_local = Environ("Tmp") & "\graphic.rtf"
Filetype = FileTypes.rtf_text

If Dir(Fname_local) <> "" Then Kill Fname_local
'Get the pointer to the rtf file.
sql = "SELECT * FROM FilePointers WHERE DocDescription='Sleep Graphics'" & " AND TableName='" & Tbl & "' AND FileType=" & Filetype & " AND RecordID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
        'continue on and get graphics from graphics field in slpResults table below
Else
    'Pointer found, check that file exists
    Fname_server = gServerPdfLocation & Rs("PointerValue")
    'Fname_server = Environ("tmp") & "\" & Rs("PointerValue") 'this is temporary path; revert to above after concept checks
    If Dir(Fname_server) = "" Then
        MsgBox "Cant find record - graphic not returned"
        Get_SlpGraphicFromDB = ""
        'continue on to get graphic from Graphics db field in SlpResults table.
    Else
        'Make a copy on the local hard drive
        FileCopy Fname_server, Fname_local
        'pass back the file name
        Get_SlpGraphicFromDB = Fname_local
        Exit Function
    End If
End If
Rs.Close
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'original code continues only if no pointer or can't find the graphic file the pointer refers to.

sql = "SELECT Graphic FROM SlpResults WHERE StudyID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields("Graphic").Value) Then
        filepath = Environ("tmp") & "\graphic.rtf"
        Set S = New ADODB.Stream
        S.Type = adTypeBinary
        S.Open
        S.Write Rs.Fields("Graphic").Value
        S.SaveToFile filepath, adSaveCreateOverWrite
        S.Close
        Get_SlpGraphicFromDB = filepath
    Else
        Get_SlpGraphicFromDB = ""
    End If
Else
    MsgBox "Cant find record - graphic not returned"
    Get_SlpGraphicFromDB = ""
End If
Rs.Close

Exit Function


Get_SlpGraphicFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_SlpGraphicFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Save_SlpGraphicToDB(RecordID As Long, tx As TXTextControl) As Boolean

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim S As ADODB.Stream
Dim Msg As String
Dim filepath As String

On Error GoTo Save_SlpGraphicToDB_Error

filepath = Environ("tmp") & "graphic.dat"

sql = "SELECT Graphic FROM SlpResults WHERE StudyID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
        Set S = New ADODB.Stream
        S.Type = adTypeBinary
        S.Open
        tx.Save filepath, 0, 5
        S.LoadFromFile filepath
        Rs("Graphic") = S.Read
        Rs("LastEdit") = Now
        Rs.Update
        Rs.Close
        S.Close
        Save_SlpGraphicToDB = True
Else
    MsgBox "Cant find record - graphic not saved"
    Save_SlpGraphicToDB = False
    Rs.Close
End If

Exit Function


Save_SlpGraphicToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_SlpGraphicToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function



Friend Function Get_SlpRequestsFromDB(PatientID As Long) As SlpRequest()

Dim Rs As New ADODB.Recordset
Dim sql
Dim S() As SlpRequest
Dim Msg As String
   
On Error GoTo Get_SlpRequestsFromDB_Error

ReDim S(0)

sql = "SELECT * FROM SlpResults WHERE PatientID=" & PatientID
sql = sql & " AND (StudyDate IS NULL OR StudyDate >= " & " CONVERT(DATE,'" & Format(Now, "yyyy/mm/dd") & "',102))"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve S(UBound(S) + 1)
    S(UBound(S)).PatientID = Rs("PatientID")
    S(UBound(S)).StudyID = Rs("StudyID")
    If IsUDTDate(Rs("RequestDate")) Then S(UBound(S)).RequestDate = Rs("RequestDate") Else S(UBound(S)).RequestDate = Empty
    If IsUDTDate(Rs("StudyDate")) Then S(UBound(S)).StudyDate = Rs("StudyDate") Else S(UBound(S)).StudyDate = Empty
    If IsNull(Rs("Urgent_scoring")) Then S(UBound(S)).Urgent_scoring = False Else S(UBound(S)).Urgent_scoring = Rs("Urgent_scoring")
    If IsNull(Rs("Urgent_PSG")) Then S(UBound(S)).Urgent_PSG = False Else S(UBound(S)).Urgent_PSG = Rs("Urgent_PSG")

    S(UBound(S)).StudyType = Rs("StudyType")
    S(UBound(S)).TreatmentMode = Rs("TreatmentMode")
    '''''''''''''''
    'added by WR to get around null bug in slpAppointments
    If Not IsNull(Rs("RequestedLocation")) Then
        S(UBound(S)).RequestedLocation = Rs("RequestedLocation")
    Else
        S(UBound(S)).RequestedLocation = ""
    End If
    ''''''''''''''
    S(UBound(S)).Special_TCO2 = Rs("Special_TCO2")
    S(UBound(S)).Special_Poes = Rs("Special_Poes")
    S(UBound(S)).Special_EMGd = Rs("Special_EMGd")
    S(UBound(S)).Special_Temp = Rs("Special_Temp")
    S(UBound(S)).Special_Other = Rs("Special_Other")
    If Rs("TreatmentOxygen") = "Yes" Then
        S(UBound(S)).TreatmentOxygen = True
    Else
        S(UBound(S)).TreatmentOxygen = False
    End If
    S(UBound(S)).RequestingMO = Rs("RequestingMO")
    Rs.MoveNext
Loop
Rs.Close

Get_SlpRequestsFromDB = S

Exit Function


Get_SlpRequestsFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_SlpRequestsFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_SlpRequestFromDB(StudyID As Long) As SlpRequest

Dim Rs As New ADODB.Recordset
Dim sql
Dim S As SlpRequest
Dim Msg As String
   
On Error GoTo Get_SlpRequestFromDB_Error

'Clear UDT
Dim emptyUDT As SlpRequest
S = emptyUDT

'Retrieve Slp data
sql = "SELECT  * FROM SlpResults WHERE StudyID = " & StudyID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    'MsgBox "Problem retrieving request record " & vbCrLf & "Record not found.", vbOKOnly, "Sleep study"
    Exit Function
End If

S.StudyID = Rs("StudyID")
S.PatientID = Rs("PatientID")

If IsDate(Rs("Studydate")) Then S.StudyDate = Format(Rs("Studydate"), "dd/mm/yyyy") Else S.StudyDate = Empty
If IsUDTDate(Rs("ApptTime")) Then S.ArrivalTime = Format(Rs("ApptTime"), "hh:mm") Else S.ArrivalTime = Empty
If IsNull(Rs("Urgent_scoring")) Then S.Urgent_scoring = False Else S.Urgent_scoring = Rs("Urgent_scoring")
If IsNull(Rs("Urgent_PSG")) Then S.Urgent_PSG = False Else S.Urgent_PSG = Rs("Urgent_PSG")
S.BookedLocation = Rs("Campus") & ""
S.BookedBy = Rs("Appt_BookedBy") & ""
If Rs("Admission") = "Yes" Then S.Admission = True Else S.Admission = False
S.Note = Rs("Note") & ""
S.ApptNote = Rs("ApptNotes") & ""

If IsDate(Rs("RequestDate")) Then S.RequestDate = Format(Rs("RequestDate"), "dd/mm/yyyy") Else S.RequestDate = Empty
S.RequestingMO = Rs("RequestingMO") & ""
S.RequestingMO_ProviderNum = Rs("RequestingMO_ProviderNum") & ""
If IsNull(Rs("Request_HealthServiceID")) Then S.Request_HealthServiceID = Empty Else S.Request_HealthServiceID = Rs("Request_HealthServiceID") & ""
S.RequestingMO_ReviewLocation = Rs("RequestingMO_ReviewLocation") & ""
S.ReportDest_1 = Rs("ReportDest_1") & ""
S.ReportDest_2 = Rs("ReportDest_2") & ""
S.BilledTo = Rs("BillingStatus") & ""
S.RequestStatus = Rs("RequestStatus") & ""
S.Request_HardCopyRec = Rs("Request_HardCopyRec")
If IsNull(Rs("Need_Interpreter_Booked")) Or Not Rs("Need_Interpreter_Booked") Then S.Interpreter_Booked = False Else S.Interpreter_Booked = True
S.Interpreter_Language = Rs("Need_Interpreter_Language") & ""
If IsNull(Rs("TransportNeeded")) Or Rs("TransportNeeded") = False Then S.Transport_Required = False Else S.Transport_Required = True
If IsNull(Rs("TransportBooked")) Or Rs("TransportBooked") = False Then S.Transport_Booked = False Else S.Transport_Booked = True
If IsNull(Rs("Emergency")) Or Rs("Emergency") = False Then S.EmergencyList = False Else S.EmergencyList = True
If IsNull(Rs("InfoSent")) Or Rs("InfoSent") = False Then S.InfoSent = False Else S.InfoSent = True

S.StudyType = Rs("Studytype") & ""
S.TreatmentMode = Rs("TreatmentMode") & ""
S.RequestedLocation = Rs("RequestedLocation") & ""
If IsNull(Rs("Special_TCO2")) Or Rs("Special_TCO2") = False Then S.Special_TCO2 = False Else S.Special_TCO2 = True
If IsNull(Rs("Special_EMGd")) Or Rs("Special_EMGd") = False Then S.Special_EMGd = False Else S.Special_EMGd = True
If IsNull(Rs("Special_Poes")) Or Rs("Special_Poes") = False Then S.Special_Poes = False Else S.Special_Poes = True
If IsNull(Rs("Special_Temp")) Or Rs("Special_Temp") = False Then S.Special_Temp = False Else S.Special_Temp = True
S.Special_Other = Rs("Special_Other") & ""
If IsNull(Rs("TreatmentOxygen")) Or Rs("TreatmentOxygen") = "No" Or Rs("TreatmentOxygen") = "0" Then S.TreatmentOxygen = False Else S.TreatmentOxygen = True
S.ClinicalNotes = Rs("ReasonForTest") & ""

S.Ref_MO = Rs("Ref_MO") & ""
S.Ref_Address = Rs("Ref_Address") & ""
If IsDate(Rs("Ref_Date")) Then S.Ref_Date = Format(Rs("Ref_Date"), "dd/mm/yyyy") Else S.Ref_Date = Empty
If IsDate(Rs("Ref_ActivationDate")) Then S.Ref_ActivateDate = Format(Rs("Ref_ActivationDate"), "dd/mm/yyyy") Else S.Ref_ActivateDate = Empty

S.Ref_Duration = Rs("Ref_Duration") & ""
S.Ref_ProviderNum = Rs("Ref_ProviderNum") & ""

If IsNull(Rs("ApptConfirmed")) Or Rs("ApptConfirmed") = False Then S.ApptConfirmed = False Else S.ApptConfirmed = True
If IsDate(Rs("ApptConfirmedDate")) Then S.ApptConfirmedDate = Rs("ApptConfirmedDate") Else S.ApptConfirmedDate = Empty

S.BedNumber_AtPSG = Rs("R_Bed") & ""
Rs.Close

Get_SlpRequestFromDB = S

Exit Function


Get_SlpRequestFromDB_Error:
   Msg = ""
   Call ErrorLog(Msg, "Get_SlpRequestFromDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_SlpFromDB(SlpID As Long) As SlpResults

Dim Rs As New ADODB.Recordset
Dim sql
Dim S As SlpResults
Dim Msg As String
   
On Error GoTo Get_SlpFromDB_Error

'Clear UDT
Dim emptyUDT As SlpResults
S = emptyUDT

'Retrieve Slp data
sql = "SELECT tbl_Name.*, tbl_RM_Demographics.*, SlpResults.*, "
sql = sql & "tbl_patient.ID, tbl_patient.URNumber, tbl_patient.Type AS Expr1, tbl_patient.DOB, tbl_patient.Gender, "
sql = sql & "tbl_patient.MaritalStatus, tbl_patient.Language, tbl_patient.Religion, tbl_patient.Race, "
sql = sql & "tbl_patient.Nationality, tbl_patient.Citizenship, tbl_patient.MothersID, tbl_patient.DateofDeath, "
sql = sql & "tbl_patient.DeathIndicator, tbl_patient.IdentityUnknown, tbl_patient.PrimaryFacility, "
sql = sql & "tbl_patient.LastModifiedBy, tbl_patient.LastModifiedDate, tbl_patient.MergedUR, "
sql = sql & "tbl_Address.Type AS Expr2, tbl_Address.TypeID AS Expr3, tbl_Address.AddressID, tbl_Address.Address1, tbl_Address.Address2, "
sql = sql & "tbl_Address.Suburb, tbl_Address.State, tbl_Address.Postcode, tbl_Address.Country , tbl_Address.AddressType "
sql = sql & "FROM tbl_patient INNER JOIN "
sql = sql & "tbl_Name ON tbl_patient.ID = tbl_Name.TypeID INNER JOIN SlpResults ON "
sql = sql & "tbl_patient.ID = SlpResults.PatientID LEFT OUTER JOIN tbl_Address ON tbl_patient.ID = "
sql = sql & "tbl_Address.TypeID LEFT OUTER JOIN tbl_RM_Demographics ON tbl_patient.ID = tbl_RM_Demographics.PatientID "
sql = sql & "WHERE (tbl_Name.Type = 'P') AND (tbl_Address.AddressType = '4' OR tbl_Address.AddressType IS NULL) AND (tbl_Address.Type = 'P') AND "
sql = sql & "(tbl_Name.Type = 'P') AND (tbl_patient.Type = 'P') AND (SlpResults.StudyID = " & SlpID & ")"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    'MsgBox "Problem retrieving PSG results" & vbCrLf & "Record not found.", vbOKOnly, "Sleep study"
    Exit Function
End If

'Get demo data
S.RecordID = Rs("StudyID")
S.PatientID = Rs("PatientID")
S.D_DOB = Format(Rs("DOB"), "dd/mm/yyyy")
S.D_Gender = Map_RefTbl(Rs("Gender"), Gender, HL7Code_RMText)
S.D_Race_Rft = CurrentPt.Get_RaceRft(Rs("PatientID"))
S.D_Address1 = Rs("Address1") & ""
S.D_Address2 = Rs("Address2") & ""
S.D_FirstName = Rs("FirstName") & ""
S.D_Middlename = Rs("Middlename") & ""
S.D_Postcode = Rs("Postcode") & ""
S.D_State = Rs("State") & ""
S.D_Suburb = Rs("Suburb") & ""
S.D_Surname = Rs("Surname")
S.D_UR_PrimaryAustin = Rs("URNumber")
S.D_URforDisplay = Get_PtString(S.PatientID, cHS.CurrentVAED, DisplayUR)
If IsNull(Rs("Urgent_scoring")) Then S.Urgent_scoring = False Else S.Urgent_scoring = Rs("Urgent_scoring")
If IsNull(Rs("Urgent_PSG")) Then S.Urgent_PSG = False Else S.Urgent_PSG = Rs("Urgent_PSG")

'Study data (request)
S.Lab = Rs("Campus") & ""
S.StudyType = Rs("Studytype")
If IsDate(Rs("Studydate")) Then S.StudyDate = Format(Rs("Studydate"), "dd/mm/yyyy") Else S.StudyDate = Empty
If IsDate(Rs("RequestDate")) Then S.Request_Date = Format(Rs("RequestDate"), "dd/mm/yyyy") Else S.Request_Date = Empty
If IsDate(Rs("RequestTime")) Then S.Request_Time = Format(Rs("RequestTime"), "hh:mm") Else S.Request_Time = Empty
S.StudyBy = Rs("StudyBy") & ""
S.ReasonForTest = Rs("ReasonForTest") & ""
If Rs("Admission") = "Yes" Then S.Admission = "Yes" Else S.Admission = "No"
S.TreatmentMode = Rs("TreatmentMode") & ""
S.RequestedLocation = Rs("RequestedLocation") & ""
S.Note = Rs("Note") & ""
S.ApptNotes = Rs("ApptNotes") & ""

'Request status tab (request)
S.Request_Status = Rs("RequestStatus") & ""
S.Request_HardCopyRec = Rs("Request_HardCopyRec")
S.RequestMode = Rs("RequestMode") & ""
If IsNull(Rs("Request_Acknowledged")) Then S.Request_Acknowledged = False Else S.Request_Acknowledged = Rs("Request_Acknowledged")
S.Request_Acknowledged_By = Rs("Request_Acknowledged_By") & ""

'Load special procedure info
If IsNull(Rs("Special_TCO2")) Or Rs("Special_TCO2") = False Then S.Special_TCO2 = False Else S.Special_TCO2 = True
If IsNull(Rs("Special_EMGd")) Or Rs("Special_EMGd") = False Then S.Special_EMGd = False Else S.Special_EMGd = True
If IsNull(Rs("Special_Poes")) Or Rs("Special_Poes") = False Then S.Special_Poes = False Else S.Special_Poes = True
If IsNull(Rs("Special_Temp")) Or Rs("Special_Temp") = False Then S.Special_Temp = False Else S.Special_Temp = True
S.Special_OtherProcedure = Rs("Special_Other") & ""
If IsNull(Rs("TreatmentOxygen")) Or Rs("TreatmentOxygen") = "No" Or Rs("TreatmentOxygen") = 0 Then S.TreatmentOxygen = False Else S.TreatmentOxygen = True
If IsNull(Rs("Request_Acknowledged")) Or Not Rs("Request_Acknowledged") Then S.Request_Acknowledged = False Else S.Request_Acknowledged = True
If IsNull(Rs("Request_HardCopyRec")) Or Not Rs("Request_HardCopyRec") Then S.Request_HardCopyRec = False Else S.Request_HardCopyRec = True
If IsNull(Rs("Need_NursingCare")) Or Not Rs("Need_NursingCare") Then S.Need_NursingCare = False Else S.Need_NursingCare = True
If IsNull(Rs("Need_Mobility")) Or Not Rs("Need_Mobility") Then S.Need_Mobility = False Else S.Need_Mobility = True
S.Need_Mobility_Type = Rs("Need_Mobility_Type") & ""
If IsNull(Rs("Need_Interpreter")) Or Not Rs("Need_Interpreter") Then S.Need_Interpreter = False Else S.Need_Interpreter = True
If IsNull(Rs("Need_Interpreter_Booked")) Or Not Rs("Need_Interpreter_Booked") Then S.Need_Interpreter_Booked = False Else S.Need_Interpreter_Booked = True
S.Need_Interpreter_Language = Rs("Need_Interpreter_Language") & ""
S.ExistingDiseases = Rs("ExistingDiseases") & ""
If IsNull(Rs("VRSSPatient")) Or Not Rs("VRSSPatient") Then S.VRSSPatient = False Else S.VRSSPatient = True
S.FeeType = Rs("FeeType") & ""
S.HccPensionNumber = Rs("HccPensionNumber") & ""
S.ApproximateWeight = Rs("ApproximateWeight") & ""
S.Request_HealthServiceID = AustinHealth

'RequestingMO
S.RequestingMO = Rs("RequestingMO") & ""
S.RequestingMO_ProviderNum = Rs("RequestingMO_ProviderNum") & ""
S.ReportDest_2 = Rs("ReportDest_2") & ""
If IsDate(Rs("RequestingMO_ReviewDate")) Then S.RequestingMO_ReviewDate = Format(Rs("RequestingMO_ReviewDate"), "dd/mm/yyyy") Else S.RequestingMO_ReviewDate = Empty
If IsDate(Rs("Ref_Date")) Then S.Ref_Date = Format(Rs("Ref_Date"), "dd/mm/yyyy") Else S.Ref_Date = Empty
S.RequestingMO_ReviewLocation = Rs("RequestingMO_ReviewLocation") & ""
S.ReportDest_1 = Rs("ReportDest_1") & ""
S.ReportDest_2 = Rs("ReportDest_2") & ""

'ReferringMO
S.Ref_MO = Rs("Ref_MO") & ""
S.Ref_Address = Rs("Ref_Address") & ""
S.Ref_Duration = Rs("Ref_Duration") & ""
S.Ref_ProviderNum = Rs("Ref_ProviderNum") & ""
If IsDate(Rs("Ref_Date")) Then S.Ref_Date = Format(Rs("Ref_Date"), "dd/mm/yyyy") Else S.Ref_Date = Empty
If IsDate(Rs("Ref_ActivationDate")) Then S.Ref_ActivateDate = Format(Rs("Ref_ActivationDate"), "dd/mm/yyyy") Else S.Ref_ActivateDate = Empty

'Study data (results)
If IsNumeric(Rs("Weight")) Then S.Weight = Rs("Weight") Else S.Weight = Empty
If IsNumeric(Rs("Height_slp")) Then S.Height_slp = Rs("Height_slp") Else S.Height_slp = Empty
S.Bed = Rs("R_Bed") & ""
S.OpticalDisk = Rs("OpticalDisk") & ""
S.R_ABG_PtcCO2_Unit = Rs("R_TcCO2_Unit") & ""
S.R_ABG_PtcCO2_Electrode = Rs("R_TcCO2_Electrode") & ""
S.PortableDevice = Rs("PortableDevice") & ""

'Reporting
S.ReportStatus = Rs("ReportStatus") & ""
S.ScoredBy = Rs("ScoredBy") & ""
If IsDate(Rs("ScoredDate")) Then S.ScoredDate = Format(Rs("ScoredDate"), "dd/mm/yyyy") Else S.ScoredDate = Empty
S.Reporter_MO = Rs("Reporter_MO") & ""
If IsDate(Rs("ReportDate")) Then S.ReportDate = Format(Rs("ReportDate"), "dd/mm/yyyy") Else S.ReportDate = Empty
S.VerifiedBy = Rs("VerifiedBy") & ""
If IsDate(Rs("VerifiedDate")) Then S.VerifiedDate = Format(Rs("VerifiedDate"), "dd/mm/yyyy") Else S.VerifiedDate = Empty
S.TechnicalNote = Rs("TecnicalNote") & ""
S.Report = Rs("Report") & ""

'Billing data (request)
S.BillingMO = Rs("BillingMO") & ""
S.BillingItems = Rs("BillingItems") & ""
S.BillingStatus = Rs("BillingStatus") & ""
S.BillingMO_ProviderNumber = Rs("BillingMO_ProviderNumber") & ""
If IsDate(Rs("Payment_ReceivedDate_Medicare")) Then S.Payment_ReceivedDate_Medicare = Format(Rs("Payment_ReceivedDate_Medicare"), "dd/mm/yyyy") Else S.Payment_ReceivedDate_Medicare = Empty
If IsDate(Rs("Payment_ReceivedDate_Patient")) Then S.Payment_ReceivedDate_Patient = Format(Rs("Payment_ReceivedDate_Patient"), "dd/mm/yyyy") Else S.Payment_ReceivedDate_Patient = Empty
S.Payment_AccountNote = Rs("Payment_AccountNote") & ""
S.Payment_AccountReminderDate = Rs("Payment_AccountReminderDate")
S.Payment_Location = Rs("Payment_Location") & ""
S.Payment_Method = Rs("Payment_Method") & ""
S.Payment_OwingAmount_Medicare = Rs("Payment_OwingAmount_Medicare") & ""
S.Payment_OwingAmount_Patient = Rs("Payment_OwingAmount_Patient") & ""
S.Payment_ReceiptIssued = Rs("Payment_ReceiptIssued") & ""
S.Payment_ReceivedAmount_Medicare = Rs("Payment_ReceivedAmount_Medicare") & ""
S.Payment_ReceivedAmount_Patient = Rs("Payment_ReceivedAmount_Patient") & ""
S.Payment_TransactedBy_Medicare = Rs("Payment_TransactedBy_Medicare") & ""
S.Payment_TransactedBy_Patient = Rs("Payment_TransactedBy_Patient") & ""

'Results
If IsDate(Rs("Est_LO")) Then S.Est_LOff = Format(Rs("Est_LO"), "hh:mm") Else S.Est_LOff = Empty
If IsDate(Rs("Est_WakeUpTime")) Then S.Est_WakeUpTime = Format(Rs("Est_WakeUpTime"), "hh:mm") Else S.Est_WakeUpTime = Empty
S.Est_TimeToSleep = Rs("Est_TimeToSleep") & ""
S.Est_SleepDuration = Rs("Est_SleepDuration") & ""
S.ESS = Rs("R_ESS") & ""
S.R_Recording_Start = Rs("R_Recording_Start") & ""
S.R_Recording_End = Rs("R_Recording_End") & ""
S.R_Report_Start = Rs("R_Report_Start") & ""
S.R_Report_End = Rs("R_Report_End") & ""
S.R_LOut = Rs("R_LOut") & ""
S.R_LOn = Rs("R_LOn") & ""
S.R_ArI_RespNrem = Rs("R_AI_RespNrem") & ""
S.R_ArI_RespRem = Rs("R_AI_RespRem") & ""
S.R_ArI_RespAll = Rs("R_AI_RespAll") & ""
S.R_ArI_PlmNrem = Rs("R_AI_PlmNrem") & ""
S.R_ArI_PlmRem = Rs("R_AI_PlmRem") & ""
S.R_ArI_PlmAll = Rs("R_AI_PlmAll") & ""
S.R_ArI_OtherNrem = Rs("R_AI_OtherNrem") & ""
S.R_ArI_OtherRem = Rs("R_AI_OtherRem") & ""
S.R_ArI_OtherAll = Rs("R_AI_OtherAll") & ""
S.R_ArI_AllNrem = Rs("R_AI_AllNrem") & ""
S.R_ArI_AllRem = Rs("R_AI_AllRem") & ""
S.R_ArI_AllAll = Rs("R_AI_AllAll") & ""
S.R_ArI_SupRem = Rs("R_AI_SupRem") & ""
S.R_ArI_SupNrem = Rs("R_AI_SupNrem") & ""
S.R_ArI_SupAll = Rs("R_AI_SupAll") & ""
S.R_ArI_NonSupRem = Rs("R_AI_NonSupRem") & ""
S.R_ArI_NonSupNrem = Rs("R_AI_NonSupNrem") & ""
S.R_ArI_NonSupAll = Rs("R_AI_NonSupAll") & ""
S.R_AHI_SupNrem = Rs("R_AHI_SupNrem") & ""
S.R_AHI_SupRem = Rs("R_AHI_SupRem") & ""
S.R_AHI_SupAll = Rs("R_AHI_SupAll") & ""
S.R_AHI_NonSupNrem = Rs("R_AHI_NonSupNrem") & ""
S.R_AHI_NonSupRem = Rs("R_AHI_NonSupRem") & ""
S.R_AHI_NonSupAll = Rs("R_AHI_NonSupAll") & ""
S.R_AHI_AllNrem = Rs("R_AHI_AllNrem") & ""
S.R_AHI_AllRem = Rs("R_AHI_AllRem") & ""
S.R_AHI_AllAll = Rs("R_AHI_AllAll") & ""
S.R_HI_SupNrem = Rs("R_HI_SupNrem") & ""
S.R_HI_SupRem = Rs("R_HI_SupRem") & ""
S.R_HI_SupAll = Rs("R_HI_SupAll") & ""
S.R_HI_NonSupNrem = Rs("R_HI_NonSupNrem") & ""
S.R_HI_NonSupRem = Rs("R_HI_NonSupRem") & ""
S.R_HI_NonSupAll = Rs("R_HI_NonSupAll") & ""
S.R_HI_AllNrem = Rs("R_HI_AllNrem") & ""
S.R_HI_AllRem = Rs("R_HI_AllRem") & ""
S.R_HI_AllAll = Rs("R_HI_AllAll") & ""
S.R_OAI_SupNrem = Rs("R_OAI_SupNrem") & ""
S.R_OAI_SupRem = Rs("R_OAI_SupRem") & ""
S.R_OAI_SupAll = Rs("R_OAI_SupAll") & ""
S.R_OAI_NonSupNrem = Rs("R_OAI_NonSupNrem") & ""
S.R_OAI_NonSupRem = Rs("R_OAI_NonSupRem") & ""
S.R_OAI_NonSupAll = Rs("R_OAI_NonSupAll") & ""
S.R_OAI_AllNrem = Rs("R_OAI_AllNrem") & ""
S.R_OAI_AllRem = Rs("R_OAI_AllRem") & ""
S.R_OAI_AllAll = Rs("R_OAI_AllAll") & ""
S.R_CAI_SupNrem = Rs("R_CAI_SupNrem") & ""
S.R_CAI_SupRem = Rs("R_CAI_SupRem") & ""
S.R_CAI_SupAll = Rs("R_CAI_SupAll") & ""
S.R_CAI_NonSupNrem = Rs("R_CAI_NonSupNrem") & ""
S.R_CAI_NonSupRem = Rs("R_CAI_NonSupRem") & ""
S.R_CAI_NonSupAll = Rs("R_CAI_NonSupAll") & ""
S.R_CAI_AllNrem = Rs("R_CAI_AllNrem") & ""
S.R_CAI_AllRem = Rs("R_CAI_AllRem") & ""
S.R_CAI_AllAll = Rs("R_CAI_AllAll") & ""
S.R_MAI_SupNrem = Rs("R_MAI_SupNrem") & ""
S.R_MAI_SupRem = Rs("R_MAI_SupRem") & ""
S.R_MAI_SupAll = Rs("R_MAI_SupAll") & ""
S.R_MAI_NonSupNrem = Rs("R_MAI_NonSupNrem") & ""
S.R_MAI_NonSupRem = Rs("R_MAI_NonSupRem") & ""
S.R_MAI_NonSupAll = Rs("R_MAI_NonSupAll") & ""
S.R_MAI_AllNrem = Rs("R_MAI_AllNrem") & ""
S.R_MAI_AllRem = Rs("R_MAI_AllRem") & ""
S.R_MAI_AllAll = Rs("R_MAI_AllAll") & ""
S.R_PLM_Nrem = Rs("R_PLM_Nrem") & ""
S.R_PLM_Rem = Rs("R_PLM_Rem") & ""
S.R_PLM_All = Rs("R_PLM_All") & ""
S.R_PLM_Wake = Rs("R_PLM_Wake") & ""
S.R_PLM_Ep_All = Rs("R_PLM_Ep_All") & ""
S.R_PLM_Ep_NREM = Rs("R_PLM_Ep_NREM") & ""
S.R_PLM_Ep_REM = Rs("R_PLM_Ep_REM") & ""
S.R_PLM_Ep_Wake = Rs("R_PLM_Ep_Wake") & ""
S.R_NREM1 = Rs("R_NREM1") & ""
S.R_NREM2 = Rs("R_NREM2") & ""
S.R_NREM3 = Rs("R_NREM3") & ""
S.R_NREM4 = Rs("R_NREM4") & ""
S.R_NREM12 = Rs("R_NREM12") & ""
S.R_NREM34 = Rs("R_NREM34") & ""
S.R_NREMTotal = Rs("R_NREMTotal") & ""
S.R_REMTotal = Rs("R_REMTotal") & ""
S.R_TST = Rs("R_TST") & ""
S.R_TST_SupineNREM = Rs("R_TST_SupineNREM") & ""
S.R_TST_NonSupineNREM = Rs("R_TST_NonSupineNREM") & ""
S.R_TST_SupineREM = Rs("R_TST_SupineREM") & ""
S.R_TST_NonSupineREM = Rs("R_TST_NonSupineREM") & ""
S.R_TST_SupineTotal = Rs("R_TST_SupineTotal") & ""
S.R_TST_NonSupineTotal = Rs("R_TST_NonSupineTotal") & ""
S.R_SLat = Rs("R_SLat") & ""
S.R_REMLat = Rs("R_REMLat") & ""
S.R_TDT = Rs("R_TDT") & ""
S.R_SlpEff = Rs("R_SlpEff") & ""
S.R_Awakenings = Rs("R_Awakenings") & ""
S.R_WakeTimeDuringSleep = Rs("R_WakeTimeDuringSleep") & ""
S.R_SpO2_BLAwake = Rs("R_SaO2_BLAwake") & ""
S.R_SpO2_BLAsleep = Rs("R_SaO2_BLAsleep") & ""
S.R_SpO2_MinREM = Rs("R_SaO2_MinREM") & ""
S.R_SpO2_MinNREM = Rs("R_SaO2_MinNREM") & ""
S.R_SpO2_Time95 = Rs("R_SaO2_Time95") & ""
S.R_SpO2_Time90 = Rs("R_SaO2_Time90") & ""
S.R_SpO2_Time88 = Rs("R_SaO2_Time88") & ""
S.R_SpO2_Time85 = Rs("R_SaO2_Time85") & ""
S.R_ABG_Conditions1 = Rs("R_ABG_Conditions1") & ""
S.R_ABG_Conditions2 = Rs("R_ABG_Conditions2") & ""
S.R_ABG_pH1 = Rs("R_ABG_pH1") & ""
S.R_ABG_pH2 = Rs("R_ABG_pH2") & ""
S.R_ABG_PaO21 = Rs("R_ABG_PaO21") & ""
S.R_ABG_PaO22 = Rs("R_ABG_PaO22") & ""
S.R_ABG_PaCO21 = Rs("R_ABG_PaCO21") & ""
S.R_ABG_PaCO22 = Rs("R_ABG_PaCO22") & ""
S.R_ABG_HCO31 = Rs("R_ABG_HCO31") & ""
S.R_ABG_HCO32 = Rs("R_ABG_HCO32") & ""
S.R_ABG_BE1 = Rs("R_ABG_BE1") & ""
S.R_ABG_BE2 = Rs("R_ABG_BE2") & ""
S.R_ABG_SaO21 = Rs("R_ABG_SaO21") & ""
S.R_ABG_SaO22 = Rs("R_ABG_SaO22") & ""
S.R_ABG_FiO21 = Rs("R_ABG_FiO21") & ""
S.R_ABG_FiO22 = Rs("R_ABG_FiO22") & ""
S.R_ABG_Time1 = Rs("R_ABG_Time1") & ""
S.R_ABG_Time2 = Rs("R_ABG_Time2") & ""
S.R_ABG_PtcCO21 = Rs("R_PtcCO21") & ""
S.R_ABG_PtcCO22 = Rs("R_PtcCO22") & ""
's.R_ABG_PCO2Diff1 = Rs("R_ABG_PCO2Diff1")
's.R_ABG_PCO2Diff2 = Rs("R_ABG_PCO2Diff2")
S.R_ABG_PtcCO2_Unit = Rs("R_tcCO2_Unit") & ""
S.R_ABG_PtcCO2_Electrode = Rs("R_tcCO2_Electrode") & ""
's.R_ABG_PtcCO2_Drift = Rs("R_ABG_PtcCO2_Drift")
S.R_ABG_SpO21 = Rs("R_SpO21") & ""
S.R_ABG_SpO22 = Rs("R_SpO22") & ""
's.R_ABG_SO2Diff1 = Rs("R_ABG_SO2Diff1")
's.R_ABG_SO2Diff2 = Rs("R_ABG_SO2Diff2")
S.R_RERA_AllAll = Rs("R_RERA_AllAll") & ""
S.R_RERA_SupAll = Rs("R_RERA_SupAll") & ""
S.R_RERA_NonSupAll = Rs("R_RERA_NonSupAll") & ""
S.R_RERA_AllRem = Rs("R_RERA_AllRem") & ""
S.R_RERA_AllNrem = Rs("R_RERA_AllNrem") & ""
S.R_RERA_SupRem = Rs("R_RERA_SupRem") & ""
S.R_RERA_SupNrem = Rs("R_RERA_SupNrem") & ""
S.R_RERA_NonSupRem = Rs("R_RERA_NonSupRem") & ""
S.R_RERA_NonSupNrem = Rs("R_RERA_NonSupNrem") & ""
S.R_RDI_AllAll = Rs("R_RDI_AllAll") & ""
S.R_RDI_SupAll = Rs("R_RDI_SupAll") & ""
S.R_RDI_NonSupAll = Rs("R_RDI_NonSupAll") & ""
S.R_RDI_AllRem = Rs("R_RDI_AllRem") & ""
S.R_RDI_AllNrem = Rs("R_RDI_AllNrem") & ""
S.R_RDI_SupRem = Rs("R_RDI_SupRem") & ""
S.R_RDI_SupNrem = Rs("R_RDI_SupNrem") & ""
S.R_RDI_NonSupRem = Rs("R_RDI_NonSupRem") & ""
S.R_RDI_NonSupNrem = Rs("R_RDI_NonSupNrem") & ""
S.R_ODI4_AllAll = Rs("R_Desat4_AllAll") & ""
S.R_ODI4_AllRem = Rs("R_Desat4_AllRem") & ""
S.R_ODI4_AllNrem = Rs("R_Desat4_AllNrem") & ""
S.R_ODI4_NonSupAll = Rs("R_Desat4_NonSupAll") & ""
S.R_ODI4_NonSupRem = Rs("R_Desat4_NonSupRem") & ""
S.R_ODI4_NonSupNrem = Rs("R_Desat4_NonSupNrem") & ""
S.R_ODI4_SupAll = Rs("R_Desat4_SupAll") & ""
S.R_ODI4_SupRem = Rs("R_Desat4_SupRem") & ""
S.R_ODI4_SupNrem = Rs("R_Desat4_SupNrem") & ""
If Len(Rs("Graphic")) > 0 Then S.R_Graphic = True Else S.R_Graphic = False

'Overnight notes (log)
S.StudyLog = Rs("StudyLog") & ""
S.StudyPerformedBy = Rs("StudyPerformedBy") & ""
S.StudyLogNote = Rs("StudyLogNote") & ""

Get_SlpFromDB = S

Exit Function


Get_SlpFromDB_Error:
   Msg = ""
   Call ErrorLog(Msg, "Get_SlpFromDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Public Function Get_Document(DocID As Long) As String
'Get doc from DB, save to temp file and return path to temp file

Dim myStream As ADODB.Stream
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim temp As String
Dim sFullPathToFile  As String
Const SW_SHOWDEFAULT As Long = 10
Dim Msg As String

On Error GoTo Get_Document_Error

'Get the document from the database
sql = "SELECT * FROM Documents WHERE DocumentID=" & DocID
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
If Not Rs.EOF Then
    'Build temp filename
    Select Case Rs("FileType")
        Case 1
            sFullPathToFile = Environ("tmp") & "\Temp.xls"
        Case 2
            sFullPathToFile = Environ("tmp") & "\Temp.pdf"
        Case 3
            sFullPathToFile = Environ("tmp") & "\Temp.doc"
        Case 4
            sFullPathToFile = Environ("tmp") & "\Temp.docx"
    End Select
    
    'stick in stream object
    Set myStream = New Stream
    myStream.Type = adTypeBinary
    myStream.Open
    myStream.Write Rs.Fields("DocData").Value
    myStream.SaveToFile sFullPathToFile, adSaveCreateOverWrite
    Get_Document = sFullPathToFile
Else
    Get_Document = ""
End If
Rs.Close

Exit Function

Get_Document_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_Document", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_DocumentPath(RecordID As Long, TblName As String) As String
'Get doc path and filename from DB

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_DocumentPath_Error

sql = "SELECT PointerValue FROM FilePointers WHERE RecordID=" & RecordID & " AND Tablename='" & TblName & "'"
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
If Not Rs.EOF Then
    Get_DocumentPath = gServerPdfLocation & Rs("PointerValue")
Else
    Get_DocumentPath = ""
End If
Rs.Close

Exit Function

Get_DocumentPath_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_DocumentPath", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function View_Document(DocID As Long) As String

Dim myStream As ADODB.Stream
Dim Rs As New ADODB.Recordset
Dim sql As String
Dim temp As String
Dim sFullPathToFile  As String
Const SW_SHOWDEFAULT As Long = 10
Dim Msg As String

On Error GoTo View_Document_Error

'Get the document from the database
sql = "SELECT * FROM Documents WHERE DocumentID=" & DocID
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic

'Build temp filename
Select Case Rs("FileType")
    Case 1
        sFullPathToFile = Environ("tmp") & "\Temp.xls"
    Case 2
        sFullPathToFile = Environ("tmp") & "\Temp.pdf"
    Case 3
        sFullPathToFile = Environ("tmp") & "\Temp.doc"
    Case 4
        sFullPathToFile = Environ("tmp") & "\Temp.docx"
End Select

'stick in stream object
Set myStream = New Stream
myStream.Type = adTypeBinary
myStream.Open
myStream.Write Rs.Fields("DocData").Value
Rs.Close
myStream.SaveToFile sFullPathToFile, adSaveCreateOverWrite
Call ShellExecute(0, "Open", sFullPathToFile, 0&, 0&, SW_SHOWDEFAULT)   'Execute app

Exit Function

View_Document_Error:
    Msg = "Ensure that only one document is open at a time."
    Call ErrorLog(Msg, "View_Document", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Get_VRSSDocuments(RecordID As Long, LinkTable As String, grd As MSFlexGrid)

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Stuff As String
Dim Msg As String

On Error GoTo Get_VRSSDocuments_Error

'Get list of documents from the database for this review
sql = "SELECT * FROM Documents WHERE LinkID=" & RecordID & " AND LinkTable='" & LinkTable & "'"
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
grd.Rows = 1
Do While Not Rs.EOF
    Stuff = Rs("DocumentID")
    Stuff = Stuff & vbTab & grd.Rows
    Stuff = Stuff & vbTab & ""
    Stuff = Stuff & vbTab & Format(Rs("docLoadDate"), "dd/mm/yyyy  hh:mm")
    Stuff = Stuff & vbTab & Rs("docDescription")
    Stuff = Stuff & vbTab & Rs("docAuthor")
    Stuff = Stuff & vbTab & Rs("OriginalFileName")
    Stuff = Stuff & vbTab & Rs("OriginalFilePath")
    Stuff = Stuff & vbTab & Docs(Rs("FileType")).Filetype
    grd.AddItem Stuff
       
    grd.Row = grd.Rows - 1
    grd.Col = 2
    Set grd.CellPicture = LoadPicture(Docs(Rs("FileType")).BitmapFile, vbLPSmall)
    Rs.MoveNext
Loop
Rs.Close

Exit Function


Get_VRSSDocuments_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSDocuments", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSS_NIVStatus(PtID As Long, Q_Params As VRSS_QueryParams) As VRSS_NIVStatus

Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim NIVStatus As VRSS_NIVStatus
Dim dummy As VRSS_NIVStatus
Dim AtThisDate As String
Dim CeaseDate() As Date
Dim i As Integer
Dim temp As Date

On Error GoTo Get_VRSS_NIVStatus_Error

NIVStatus = dummy   'Clear UDT

'Exit if not a relevant call
'If Q_Params.VentPrescribed = vbIgnore Then
'    NIVStatus.HasNIV = vbIgnore
'    NIVStatus.VentDependent = vbIgnore
'    Get_VRSS_NIVStatus = NIVStatus
'    Exit Function
'End If
                                                                                                                                                    
'Format date to sql
If IsUDTDate(Q_Params.VentPrescribed) Then
    AtThisDate = gDBFunctions.Convert_ToSQLQueryDate(Q_Params.VentPrescribed_AtDate)
Else
    AtThisDate = gDBFunctions.Convert_ToSQLQueryDate(Q_Params.ProgramStatus_AtDate)
End If

ReDim CeaseDate(0)
ReDim NIVStatus.Machines(0)

'Get all machines (non-cpap) ever used by pt
sql = "SELECT Visit_NonCPAP_ImpReview.MachineID FROM CPAP_VentEquip INNER JOIN VentilatorSettings ON CPAP_VentEquip.ModelID = VentilatorSettings.ModelID "
sql = sql & "INNER JOIN Visit_NonCPAP_ImpReview ON CPAP_VentEquip.MachineID = Visit_NonCPAP_ImpReview.MachineID "
sql = sql & "WHERE (VentilatorSettings.CPAPOnlymachine = 0) AND (PatientID = " & PtID & ") AND "
sql = sql & "(NOT (Visit_NonCPAP_ImpReview.MachineID IS NULL OR Visit_NonCPAP_ImpReview.MachineID = 0)) AND (VisitDate <= " & AtThisDate & ") "
sql = sql & "GROUP BY Visit_NonCPAP_ImpReview.MachineID"
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
NIVStatus.HasNIV = vbNo
Do While Not Rs.EOF
    'Get history for each machine and check who had it last - if pt was last for >=1 machine then NIV-using = true
    sql = "SELECT PatientID, VisitDate AS ThisDate, PrescribedUsage FROM Visit_NonCPAP_ImpReview "
    sql = sql & "WHERE MachineID = " & Rs("MachineID") & " AND VisitDate <= " & AtThisDate
    sql = sql & " UNION ALL "
    sql = sql & "SELECT 0 AS PatientID, LocationDate AS ThisDate, NULL AS PrescribedUsage FROM CPAP_VentEquip_NonPt "
    sql = sql & "WHERE MachineID = " & Rs("MachineID") & " AND LocationDate <= " & AtThisDate
    sql = sql & " ORDER BY ThisDate DESC"
    Rs1.CursorLocation = adUseClient
    Rs1.Open sql, ConnectString, adOpenDynamic, adLockPessimistic   'The first record is the latest
        ReDim Preserve NIVStatus.Machines(UBound(NIVStatus.Machines) + 1)
        NIVStatus.Machines(UBound(NIVStatus.Machines)).Location_PatientID = PtID
        NIVStatus.Machines(UBound(NIVStatus.Machines)).MachineID = Rs("MachineID")
        If Rs1("PatientID") = PtID Then
            NIVStatus.Machines(UBound(NIVStatus.Machines)).Location_Date = Rs1("ThisDate")
            NIVStatus.Machines(UBound(NIVStatus.Machines)).MachineReturnedDate = Empty
        Else
            'Find when the patient had this machine deallocated
            Do While Not Rs1.EOF
                If Rs1("PatientID") = PtID Then
                    NIVStatus.Machines(UBound(NIVStatus.Machines)).Location_Date = Rs1("ThisDate")
                    Rs1.MovePrevious
                    NIVStatus.Machines(UBound(NIVStatus.Machines)).MachineReturnedDate = Rs1("ThisDate")
                    Exit Do
                End If
                Rs1.MoveNext
            Loop
        End If
        
        'FInd latest prescription for each machine
        Rs1.MoveFirst
        Do While Not Rs1.EOF
            If (Rs1("PatientID") = PtID) And (Rs1("PrescribedUsage") <> "") Then
                NIVStatus.Machines(UBound(NIVStatus.Machines)).Prescription_Latest = Rs1("PrescribedUsage")
                Exit Do
            Else
                NIVStatus.Machines(UBound(NIVStatus.Machines)).Prescription_Latest = ""
            End If
            Rs1.MoveNext
        Loop
    Rs1.Close
    Rs.MoveNext
Loop
Rs.Close

'Now scan the end dates for this patient for each machine, if any=now then currently on NIV else no longer on NIV (use most recent for end date)
temp = "1/1/1900"
For i = 1 To UBound(NIVStatus.Machines)
    If Not IsUDTDate(NIVStatus.Machines(i).MachineReturnedDate) Then
        NIVStatus.HasNIV = vbYes
        NIVStatus.CeasedNIVDate = Empty
        Exit For
    Else
        If NIVStatus.Machines(i).MachineReturnedDate > temp Then
            NIVStatus.HasNIV = vbNo
            NIVStatus.CeasedNIVDate = NIVStatus.Machines(i).MachineReturnedDate
            temp = NIVStatus.Machines(i).MachineReturnedDate
        End If
    End If
Next i

'Get vent dependency if selected and if has NIV - find latest record with prescription info and use that
If Q_Params.VentDependent = vbIgnore Then
    NIVStatus.VentDependent = vbIgnore
Else
    If NIVStatus.HasNIV = vbYes Then
        sql = "SELECT TOP (1) PrescribedUsage FROM Visit_NonCPAP_ImpReview "
        sql = sql & "WHERE (PatientID = " & PtID & ") AND (PrescribedUsage <> N'') AND (VisitDate <= " & AtThisDate & ") "
        sql = sql & "ORDER BY VisitDate DESC"
        Rs.Open sql, ConnectString, adOpenKeyset, adLockPessimistic
        If Rs.EOF Then
            NIVStatus.VentDependent = VRSS_VentDependent.Unknown    'No prescription info
        Else
            Select Case Pt.VRSS_IsNIVdependent(Rs("PrescribedUsage"))
                Case True:  NIVStatus.VentDependent = vbYes
                Case False: NIVStatus.VentDependent = vbNo
            End Select
        End If
    Else
        NIVStatus.VentDependent = vbIgnore
    End If
End If

Get_VRSS_NIVStatus = NIVStatus

Exit Function

Get_VRSS_NIVStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSS_NIVStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSDocumentList(RecordID As Long, SourceTable As DbTables) As VRSS_DocInfo()

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim D() As VRSS_DocInfo
Dim D_dummy As VRSS_DocInfo
Dim i As Integer
Dim Tbl As String

On Error GoTo Get_VRSSDocumentList_Error

ReDim D(0)
D(0) = D_dummy

Select Case SourceTable
    Case DbTables.VRSS_Episodes:            Tbl = "VRSS_Episodes"
    Case DbTables.VRSS_ClinicalContacts:    Tbl = "VRSS_ClinicalContacts"
End Select

'Get list of documents from the database for this review
sql = "SELECT * FROM Documents WHERE LinkID=" & RecordID & " AND LinkTable='" & Tbl & "'"
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
Do While Not Rs.EOF
    ReDim Preserve D(UBound(D) + 1)
    D(UBound(D)).DocumentID = Rs("DocumentID")
    If IsDate(Rs("docLoadDate")) Then D(UBound(D)).LoadDate = Rs("docLoadDate") Else D(UBound(D)).LoadDate = Empty
    D(UBound(D)).Description = Rs("docDescription") & ""
    D(UBound(D)).Author = Rs("docAuthor") & ""
    D(UBound(D)).OriginalFileName = Rs("OriginalFileName") & ""
    D(UBound(D)).OriginalFilePath = Rs("OriginalFilePath") & ""
    D(UBound(D)).Filetype = Docs(Rs("FileType")).Filetype & ""
    Rs.MoveNext
Loop
Rs.Close
Get_VRSSDocumentList = D

Exit Function


Get_VRSSDocumentList_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSDocumentList", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_VRSSDocumentToDB(RecordID As Long, SourceTable As String, Description As String, Author As String, filepath As String, Filetype As String) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim S As ADODB.Stream
Dim Tbl As String

On Error GoTo Save_VRSSDocumentToDB_Error

'Save the document to stream object
Set S = New Stream
S.Type = adTypeBinary
S.Open
S.LoadFromFile filepath

'Save stream and other doc data to a new database record
sql = "SELECT * FROM Documents WHERE DocumentID=0;"
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
Rs.AddNew
Rs("LinkTable") = SourceTable
Rs("LinkID") = RecordID
Rs("Docdata") = S.Read
Rs("DocDescription") = Description
Rs("DocAuthor") = Author
Rs("docLoadDate") = Format(Now, "dd/mm/yyyy")
Rs("FileType") = Filetype
Rs("OriginalFilepath") = filepath
Rs("OriginalFileName") = Fname

Rs.Update
Rs.MoveFirst
Save_VRSSDocumentToDB = Rs("DocumentID")
Rs.Close
S.Close

Exit Function


Save_VRSSDocumentToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSDocumentToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_VRSSVentDependentStatus(D As VRSS_VentDependentStatus, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_VRSSVentDependentStatus_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM VRSS_VentDependentStatus WHERE StatusID =" & D.StatusID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving status data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM VRSS_VentDependentStatus WHERE StatusID = 0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = D.PatientID
        Rs("EpisodeID") = D.EpisodeID
End Select

Rs("Status_code") = D.Status_code
Rs("Status_text") = D.Status_text
Rs("StatusDate") = D.StatusDate
Rs("StatusTime") = D.StatusTime
Rs("EnteredBy") = D.EnteredBy

Rs("Updated") = Now
Rs.Update
Rs.MoveFirst
Save_VRSSVentDependentStatus = Rs("StatusID")
Rs.Close

Exit Function

Save_VRSSVentDependentStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSVentDependentStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_VRSSProgramToDB(P As Vent_Program, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_VRSSProgramToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM VRSS_NIVPrograms WHERE ProgramID =" & P.ProgramID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving ventilator program data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM VRSS_NIVPrograms WHERE ProgramID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
End Select

Rs("AllocationID") = P.AllocationID
Rs("Description") = P.Description
If IsUDTDate(P.Date_Entered) Then Rs("Date_Entered") = CDate(P.Date_Entered) Else Rs("Date_Entered") = Null
Rs("EnteredBy") = P.EnteredBy

Rs.Update
Rs.MoveFirst
Save_VRSSProgramToDB = Rs("ProgramID")
Rs.Close

Exit Function


Save_VRSSProgramToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSProgramToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_VRSSVentSettingsToDB(S As Vent_Settings, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_VRSSVentSettingsToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM VRSS_NIVSettings WHERE SettingsID =" & S.SettingsID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving ventilator program data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM VRSS_NIVSettings WHERE SettingsID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
End Select

Rs("ProgramID") = S.ProgramID
Rs("Mode_VentilationMode") = S.Mode
If IsUDTDate(S.Set_StartDate) Then Rs("Date_Entered") = S.Set_StartDate Else Rs("Date_Entered") = Null
If IsUDTDate(S.Set_EndDate) Then Rs("Date_Closed") = S.Set_EndDate Else Rs("Date_Closed") = Null
For i = 1 To UBound(S.Settings())
    Rs("Sett_" & S.Settings(i).Name) = S.Settings(i).Value
Next i

Rs("EnteredBy") = S.EnteredBy

Rs.Update
Rs.MoveFirst
Save_VRSSVentSettingsToDB = Rs("SettingsID")
Rs.Close

Exit Function


Save_VRSSVentSettingsToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSVentSettingsToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Save_VRSSEpisodeToDB(V As VRSS_Episode, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_VRSSEpisodeToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM VRSS_Episodes WHERE EpisodeID =" & V.EpisodeID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving VRSS episode data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM VRSS_Episodes WHERE EpisodeID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = V.PatientID
End Select

Rs("EntryCreatedBy") = V.EnteredBy & ""
Rs("EntryCreateDate") = V.EnteredByDate

Rs("VRSS_Region") = V.VRSS_Region & ""
Rs("VRSS_Melways") = V.Melways & ""

Rs("RefMO_Name") = V.RefMO_Name & ""
Rs("RefMO_Profession") = V.RefMO_Profession & ""
Rs("RefMO_Institution") = V.RefMO_Institution & ""
Rs("RefMO_Address") = V.RefMO_Address & ""
Rs("RefMO_Suburb") = V.RefMO_Suburb & ""
Rs("RefMO_PCode") = V.RefMO_PCode & ""
If IsUDTDate(V.RefDate) Then Rs("Ref_Date") = V.RefDate Else Rs("Ref_Date") = Null
If IsUDTDate(V.Ref_ReceivedDate) Then Rs("Ref_RecDate") = V.Ref_ReceivedDate Else Rs("Ref_RecDate") = Null
If IsUDTDate(V.Ref_AcknowledgeDate) Then Rs("Ref_AckDate") = V.Ref_AcknowledgeDate Else Rs("Ref_AckDate") = Null
Rs("GP_Name") = V.RefGP_Name & ""
Rs("GP_PracticeName") = V.RefGP_PracticeName & ""
Rs("GP_Address") = V.RefGP_Address & ""
Rs("GP_Suburb") = V.RefGP_Suburb & ""
Rs("GP_PCode") = V.RefGP_PCode

Rs("CarerName") = V.CarerName
Rs("CarerRelationship") = V.CarerRelationship_VRSS
Rs("CarerAddress") = V.CarerAddress
Rs("CarerSuburb") = V.CarerSuburb
Rs("CarerPostcode") = V.CarerPostcode
Rs("CarerPhone") = V.CarerPhone

Rs("Diagnosis_Pri") = V.Diagnosis_Pri
Rs("Diagnosis_Sec") = V.Diagnosis_Sec
Rs("OtherClinicalInfo") = V.OtherClinicalInfo
If IsUDTDate(V.VentilationStartDate) Then Rs("VentilationStartDate") = V.VentilationStartDate Else Rs("VentilationStartDate") = Null

Rs("Vinah_ClaimNo") = V.Vinah_ClaimNo & ""
Rs("Vinah_AccountClass") = V.VINAH_AccountClass & ""
Rs("Vinah_AccountClass_Code") = V.VINAH_AccountClass_Code & ""
Rs("Vinah_HealthCondition_Desc") = V.Vinah_HealthCondition_Desc & ""
Rs("Vinah_HealthCondition_Code") = V.Vinah_HealthCondition_Code & ""
Rs("Vinah_Interpreter") = V.Vinah_Interpreter & ""
Rs("Vinah_InterpreterCode") = V.Vinah_InterpreterCode & ""
Rs("Vinah_PtLivingArrangements") = V.Vinah_PtLivingArrangements & ""
Rs("Vinah_PtLivingArrangements_Code") = V.Vinah_PtLivingArrangements_Code & ""
Rs("Vinah_PtUsualAccommodation") = V.Vinah_PtUsualAccommodation & ""
Rs("Vinah_PtUsualAccommodation_Code") = V.Vinah_PtUsualAccommodation_Code & ""
Rs("Vinah_CarerStatus") = V.Vinah_CarerStatus & ""
Rs("Vinah_CarerStatus_Code") = V.Vinah_CarerStatus_Code & ""
Rs("Vinah_CarerRelationship") = V.Vinah_CarerRelationship & ""
Rs("Vinah_CarerRelationship_Code") = V.Vinah_CarerRelationship_Code & ""
Rs("Vinah_CarerResidence") = V.Vinah_CarerResidence & ""
Rs("Vinah_CarerResidence_Code") = V.Vinah_CarerResidence_Code & ""
Rs("Vinah_EpisodeEndReason") = V.Vinah_EpisodeEndReason & ""
Rs("Vinah_EpisodeEndReason_Code") = V.Vinah_EpisodeEndReason_Code & ""
Rs("Vinah_ReferredOutTo") = V.Vinah_ReferredOutTo & ""
Rs("Vinah_ReferredOutTo_Code") = V.Vinah_ReferredOutTo_Code & ""
If IsUDTDate(V.Vinah_ReferredOutTo_Date) Then Rs("Vinah_ReferredOutTo_Date") = V.Vinah_ReferredOutTo_Date Else Rs("Vinah_ReferredOutTo_Date") = Null
If IsUDTDate(V.EpisodeEndDate) Then Rs("Ep_EndDate_VRSS") = V.EpisodeEndDate Else Rs("Ep_EndDate_VRSS") = Null
Rs("Ep_EndedBy") = V.EpisodeClosedBy & ""

Rs("AdvancedCarePlanStatus") = V.AdvancedCarePlanStatus
If IsUDTDate(V.AdvancedCarePlanDate) Then Rs("AdvancedCarePlanDate") = V.AdvancedCarePlanDate Else Rs("AdvancedCarePlanDate") = Null
If IsUDTDate(V.AdvancedCarePlanDateUpdated) Then Rs("AdvancedCarePlanUpdatedDate") = V.AdvancedCarePlanDateUpdated Else Rs("AdvancedCarePlanUpdatedDate") = Null

Rs.Update
Rs.MoveFirst
Save_VRSSEpisodeToDB = Rs("EpisodeID")
Rs.Close

Exit Function


Save_VRSSEpisodeToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSEpisodeToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Save_VRSSAssessmentToDB(V As VRSS_Assess, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_VRSSAssessmentToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM Visit_NonCPAP_Assess WHERE VisitID =" & V.VisitID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving VRSS Assessment data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM Visit_NonCPAP_Assess WHERE VisitID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = V.PatientID
End Select

'Save data
'Rs("PatientInstitution") = V.PatientInstitution
'Rs("PatientDepartment") = V.PatientDepartment
'Rs("AdmissionStatus") = V.AdmissionStatus
If IsUDTDate(V.VisitDate) Then Rs("VisitDate") = V.VisitDate Else Rs("VisitDate") = Null
'If IsDate(V.visitTime) Then Rs("VisitTime") = V.visitTime Else Rs("VisitTime") = Null
'Rs("VisitType") = V.VisitType
'Rs("VisitLocation") = V.VisitLocation
'Rs("VisitSource") = V.VisitSource
Rs("VrssConsultant") = V.VrssConsultant

Rs("RefMO_Name") = V.RefMO_Name
Rs("RefMO_Profession") = V.RefMO_Profession
Rs("RefMO_Institution") = V.RefMO_Institution
Rs("RefMO_Address") = V.RefMO_Address & ""
Rs("RefMO_Suburb") = V.RefMO_Suburb
Rs("RefMO_PCode") = V.RefMO_PCode

If IsUDTDate(V.RefDate) Then Rs("RefMO_RefDate") = V.RefDate Else Rs("RefMO_RefDate") = Null
If IsUDTDate(V.RefReceivedDate) Then Rs("ReferralReceivedDate") = V.RefReceivedDate Else Rs("ReferralReceivedDate") = Null
If IsUDTDate(V.RefAcknowledgedDate) Then Rs("ReferralAcknowledgeDate") = V.RefAcknowledgedDate Else Rs("ReferralAcknowledgeDate") = Null

Rs("RefGP_Name") = V.RefGP_Name & ""
Rs("RefGP_PracticeName") = V.RefGP_PracticeName & ""
Rs("RefGP_Address") = V.RefGP_Address & ""
Rs("RefGP_Suburb") = V.RefGP_Suburb & ""
Rs("RefGP_PCode") = V.RefGP_PCode & ""

Rs("CarerName") = V.Carer_Name
Rs("CarerAddress") = V.Carer_Address & ""
Rs("CarerRelationship") = V.Carer_Relationship & ""
Rs("CarerSuburb") = V.Carer_Suburb
Rs("CarerPostcode") = V.Carer_Postcode
Rs("CarerPhone") = V.Carer_Phone

Rs("Diagnosis_Pri") = V.Diagnosis_Pri
Rs("Diagnosis_Sec") = V.Diagnosis_Sec
Rs("OtherClinicalInfo") = V.OtherClinicalInfo
If IsUDTDate(V.VentStartDate) Then Rs("DecisionToVentDate") = V.VentStartDate Else Rs("DecisionToVentDate") = Null

Rs("Vinah_AccountClass") = V.VINAH_AccountClass & ""
Rs("Vinah_ClaimNo") = V.Vinah_ClaimNo & ""
Rs("Vinah_HealthCondition_Desc") = V.Vinah_HealthCondition_Desc & ""
Rs("Vinah_PtLivingArrangements") = V.Vinah_PtLivingArrangements & ""
Rs("Vinah_PtUsualAccommodation") = V.Vinah_PtUsualAccommodation & ""
Rs("Vinah_CarerRelationship") = V.Vinah_CarerRelationship & ""
Rs("Vinah_CarerResidence") = V.Vinah_CarerResidence & ""
Rs("Vinah_CarerStatus") = V.Vinah_CarerStatus & ""
Rs("Vinah_EpisodeEndReason") = V.Vinah_EpisodeEndReason & ""
Rs("Vinah_Interpreter") = V.Vinah_Interpreter & ""
Rs("Vinah_PreferredLanguage") = V.Vinah_PreferredLanguage & ""
Rs("Vinah_ReferredOutTo") = V.Vinah_ReferredOutTo & ""
If IsUDTDate(V.Vinah_ReferredOutToDate) Then Rs("Vinah_ReferredOutTo_Date") = V.Vinah_ReferredOutToDate Else Rs("Vinah_ReferredOutTo_Date") = Null
Rs("Vinah_EpisodeEndReason") = V.Vinah_EpisodeEndReason & ""
If IsUDTDate(V.EpisodeEndDate_VRSS) Then Rs("EpisodeEndDate_VRSS") = V.EpisodeEndDate_VRSS Else Rs("EpisodeEndDate_VRSS") = Null

Rs.Update
Rs.MoveFirst
Save_VRSSAssessmentToDB = Rs("VisitID")
Rs.Close

Exit Function


Save_VRSSAssessmentToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSAssessmentToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSVentDependentStatus(StatusID As Long) As VRSS_VentDependentStatus

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim D As VRSS_VentDependentStatus
Dim Msg As String

On Error GoTo Get_VRSSVentDependentStatus_Error

'Clear UDT
Dim emptyUDT As VRSS_VentDependentStatus
D = emptyUDT

'Retrieve data
sql = "SELECT * FROM VRSS_VentDependentStatus WHERE StatusID =" & StatusID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving ventilator dependency data - aborted"
    Exit Function
End If

'Get data
D.PatientID = Rs("PatientID")
D.EpisodeID = Rs("EpisodeID")
D.StatusID = StatusID

If IsUDTDate(Rs("StatusDate")) Then D.StatusDate = Rs("StatusDate") Else D.StatusDate = Empty
D.StatusTime = Rs("StatusTime") & ""
D.EnteredBy = Rs("EnteredBy")
D.Status_code = Rs("Status_code")
D.Status_text = Rs("Status_text")

Rs.Close

Get_VRSSVentDependentStatus = D

Exit Function


Get_VRSSVentDependentStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSVentDependentStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function



Friend Function Get_VRSSCurrentEpisodeID(PtID As Long) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_VRSSCurrentEpisodeID_Error

sql = "SELECT EpisodeID FROM VRSS_Episodes WHERE PatientID=" & PtID & " ORDER BY Ref_RecDate DESC"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Get_VRSSCurrentEpisodeID = Rs("EpisodeID")
Else
    Get_VRSSCurrentEpisodeID = -1
End If
Rs.Close

Exit Function


Get_VRSSCurrentEpisodeID_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSCurrentEpisodeID", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Get_VRSSAssessFromDB(VisitID As Long) As VRSS_Assess

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim V As VRSS_Assess
Dim Msg As String

On Error GoTo Get_VRSSAssessFromDB_Error

'Clear UDT
Dim emptyUDT As VRSS_Assess
V = emptyUDT

'Retrieve VRSS assessment data
sql = "SELECT * FROM Visit_NonCPAP_Assess WHERE VisitID =" & VisitID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving VRSS assessment data - aborted"
    Exit Function
End If

'Get data
V.PatientID = Rs("PatientID")
V.VisitID = Rs("VisitID")

If IsNull(Rs("VisitDate")) Then V.VisitDate = Empty Else V.VisitDate = Rs("VisitDate")
V.VrssConsultant = Rs("VrssConsultant") & ""

V.RefMO_Name = Rs("RefMO_Name") & ""
V.RefMO_Profession = Rs("RefMO_Profession") & ""
V.RefMO_Institution = Rs("RefMO_Institution") & ""
V.RefMO_Address = Rs("RefMO_Address") & ""
V.RefMO_Suburb = Rs("RefMO_Suburb") & ""
V.RefMO_PCode = Rs("RefMO_PCode") & ""

If IsNull(Rs("RefMO_RefDate")) Then V.RefDate = Empty Else V.RefDate = Rs("RefMO_RefDate")
If IsNull(Rs("ReferralReceivedDate")) Then V.RefReceivedDate = Empty Else V.RefReceivedDate = Rs("ReferralReceivedDate")
If IsNull(Rs("ReferralAcknowledgeDate")) Then V.RefAcknowledgedDate = Empty Else V.RefAcknowledgedDate = Rs("ReferralAcknowledgeDate")

V.RefGP_Name = Rs("RefGP_Name") & ""
V.RefGP_PracticeName = Rs("RefGP_PracticeName") & ""
V.RefGP_Address = Rs("RefGP_Address") & ""
V.RefGP_Suburb = Rs("RefGP_Suburb") & ""
V.RefGP_PCode = Rs("RefGP_PCode") & ""

V.Carer_Name = Rs("CarerName") & ""
V.Carer_Address = Rs("CarerAddress") & ""
V.Carer_Relationship = Rs("CarerRelationship") & ""
V.Carer_Suburb = Rs("CarerSuburb") & ""
V.Carer_Postcode = Rs("CarerPostcode") & ""
V.Carer_Phone = Rs("CarerPhone") & ""

V.Diagnosis_Pri = Rs("Diagnosis_Pri") & ""
V.Diagnosis_Sec = Rs("Diagnosis_Sec") & ""
V.OtherClinicalInfo = Rs("OtherClinicalInfo") & ""
If IsDate(Rs("DecisionToVentDate")) Then V.VentStartDate = Rs("DecisionToVentDate") Else V.VentStartDate = Empty

V.AdvancedCarePlanStatus = Rs("AdvancedCarePlanStatus") & ""
If IsDate(Rs("AdvancedCarePlanDate")) Then V.AdvancedCarePlanDate = Rs("AdvancedCarePlanDate") Else V.AdvancedCarePlanDate = Empty
If IsDate(Rs("AdvancedCarePlanUpdatedDate")) Then V.AdvancedCarePlanUpdateDate = Rs("AdvancedCarePlanUpdatedDate") Else V.AdvancedCarePlanUpdateDate = Empty

V.VINAH_AccountClass = Rs("Vinah_AccountClass") & ""
V.VINAH_AccountClass_Code = Rs("Vinah_AccountClass_Code") & ""
V.Vinah_ClaimNo = Rs("Vinah_ClaimNo") & ""
V.Vinah_HealthCondition_Desc = Rs("Vinah_HealthCondition_Desc") & ""
V.Vinah_HealthCondition_Code = Rs("Vinah_HealthCondition_Code") & ""
V.Vinah_PtLivingArrangements = Rs("Vinah_PtLivingArrangements") & ""
V.Vinah_PtUsualAccommodation = Rs("Vinah_PtUsualAccommodation") & ""
V.Vinah_CarerRelationship = Rs("Vinah_CarerRelationship") & ""
V.Vinah_CarerResidence = Rs("Vinah_CarerResidence") & ""
V.Vinah_CarerStatus = Rs("Vinah_CarerStatus") & ""
V.Vinah_EpisodeEndReason = Rs("Vinah_EpisodeEndReason") & ""
V.Vinah_Interpreter = Rs("Vinah_Interpreter") & ""
V.Vinah_Interpreter_Code = Rs("Vinah_InterpreterCode") & ""
V.Vinah_PreferredLanguage = Rs("Vinah_PreferredLanguage") & ""
V.Vinah_PreferredLanguage_Code = Rs("Vinah_PreferredLanguage_Code") & ""

V.Vinah_ReferredOutTo = Rs("Vinah_ReferredOutTo") & ""
If IsDate(Rs("Vinah_ReferredOutTo_Date")) Then V.Vinah_ReferredOutToDate = Rs("Vinah_ReferredOutTo_Date") Else V.Vinah_ReferredOutToDate = Empty
V.Vinah_EpisodeEndReason = Rs("Vinah_EpisodeEndReason") & ""
If IsDate(Rs("EpisodeEndDate_VRSS")) Then V.EpisodeEndDate_VRSS = Rs("EpisodeEndDate_VRSS") Else V.EpisodeEndDate_VRSS = Empty

Rs.Close

Get_VRSSAssessFromDB = V

Exit Function


Get_VRSSAssessFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSAssessFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function VRSS_IsNIVdependent(PrescribedUsage As String) As Boolean


Select Case PrescribedUsage
    Case ">16 hrs per day": VRSS_IsNIVdependent = True
    Case Else:              VRSS_IsNIVdependent = False
End Select

End Function

Friend Function VRSS_NIVprescriptions(Pt As PtDemographics1) As VRSS_NIVPrescription()
'Return all NIV prescription info in date desc order

Dim sql As String
Dim Rs As New ADODB.Recordset
Dim Msg As String
Dim i As Integer
Dim P() As VRSS_NIVPrescription

On Error GoTo Error_VRSS_NIVprescriptions

ReDim P(0)

sql = "SELECT VisitDate, PrescribedUsage FROM Visit_NonCPAP_ImpReview WHERE (PatientID = " & Pt.PtID & ") AND (NOT (PrescribedUsage = N'')) ORDER BY VisitDate DESC"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve P(UBound(P) + 1)
    P(UBound(P)).NIVdependent = VRSS_IsNIVdependent(Rs("PrescribedUsage") & "")
    P(UBound(P)).PrescriptionDate = Rs("VisitDate")
    Rs.MoveNext
Loop
Rs.Close

VRSS_NIVprescriptions = P

Exit Function

Error_VRSS_NIVprescriptions:
    Msg = ""
    Call ErrorLog(Msg, "VRSS_NIVprescriptions", "cGeneralDBRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSEpisodeDates(PtID As Long) As VRSS_EpisodeDates()

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim E() As VRSS_EpisodeDates
Dim Msg As String

On Error GoTo Get_VRSSEpisodeDates_Error

ReDim E(0)

sql = "SELECT EpisodeID, Ref_RecDate AS Ep_Start, Ep_EndDate_VRSS FROM VRSS_Episodes WHERE PatientID=" & PtID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve E(UBound(E) + 1)
    E(UBound(E)).EpisodeID = Rs("EpisodeID")
    E(UBound(E)).OpenDate = Rs("Ep_Start")
    If IsUDTDate(Rs("Ep_EndDate_VRSS")) Then E(UBound(E)).CloseDate = Rs("Ep_EndDate_VRSS") Else E(UBound(E)).CloseDate = Empty
    Rs.MoveNext
Loop
Rs.Close

Get_VRSSEpisodeDates = E

Exit Function

Get_VRSSEpisodeDates_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSEpisodeDates", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSEpisodeFromDB(EpisodeID As Long) As VRSS_Episode

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim V As VRSS_Episode
Dim Msg As String

On Error GoTo Get_VRSSEpisodeFromDB_Error

'Clear UDT
Dim emptyUDT As VRSS_Episode
V = emptyUDT

'Retrieve VRSS assessment data
sql = "SELECT * FROM VRSS_Episodes WHERE EpisodeID =" & EpisodeID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    'MsgBox "VRSS episode data not found - aborted"
    Exit Function
End If

'Get data
V.PatientID = Rs("PatientID")
V.EpisodeID = Rs("EpisodeID")

V.EnteredBy = Rs("EntryCreatedBy") & ""
V.EnteredByDate = Rs("EntryCreateDate") & ""
V.EpisodeClosedBy = Rs("Ep_EndedBy") & ""

V.VRSS_Region = Rs("VRSS_Region") & ""
V.Melways = Rs("VRSS_Melways") & ""

If IsNull(Rs("Ref_RecDate")) Then V.EpisodeStartDate = Empty Else V.EpisodeStartDate = Rs("Ref_RecDate")
If IsNull(Rs("Ep_EndDate_VRSS")) Then V.EpisodeEndDate = Empty Else V.EpisodeEndDate = Rs("Ep_EndDate_VRSS")
If IsDate(Rs("Ep_EndDate_VRSS")) Then V.EpisodeEndDate = Rs("Ep_EndDate_VRSS") Else V.EpisodeEndDate = Empty
If IsDate(Rs("Vinah_ReferredOutTo_Date")) Then V.Vinah_ReferredOutTo_Date = Rs("Vinah_ReferredOutTo_Date") Else V.Vinah_ReferredOutTo_Date = Empty

V.RefMO_Name = Rs("RefMO_Name") & ""
V.RefMO_Profession = Rs("RefMO_Profession") & ""
V.RefMO_Institution = Rs("RefMO_Institution") & ""
V.RefMO_Address = Rs("RefMO_Address") & ""
V.RefMO_Suburb = Rs("RefMO_Suburb") & ""
V.RefMO_PCode = Rs("RefMO_PCode") & ""
If IsNull(Rs("Ref_Date")) Then V.RefDate = Empty Else V.RefDate = Rs("Ref_Date")
If IsNull(Rs("Ref_RecDate")) Then V.Ref_ReceivedDate = Empty Else V.Ref_ReceivedDate = Rs("Ref_RecDate")
If IsNull(Rs("Ref_AckDate")) Then V.Ref_AcknowledgeDate = Empty Else V.Ref_AcknowledgeDate = Rs("Ref_AckDate")

V.RefGP_Name = Rs("GP_Name") & ""
V.RefGP_PracticeName = Rs("GP_PracticeName") & ""
V.RefGP_Address = Rs("GP_Address") & ""
V.RefGP_Suburb = Rs("GP_Suburb") & ""
V.RefGP_PCode = Rs("GP_PCode") & ""

V.CarerName = Rs("CarerName") & ""
V.CarerRelationship_VRSS = Rs("CarerRelationship") & ""
V.CarerAddress = Rs("CarerAddress") & ""
V.CarerSuburb = Rs("CarerSuburb") & ""
V.CarerPostcode = Rs("CarerPostcode") & ""
V.CarerPhone = Rs("CarerPhone") & ""

V.Diagnosis_Pri = Rs("Diagnosis_Pri") & ""
V.Diagnosis_Sec = Rs("Diagnosis_Sec") & ""
V.OtherClinicalInfo = Rs("OtherClinicalInfo") & ""
If IsDate(Rs("VentilationStartDate")) Then V.VentilationStartDate = Rs("VentilationStartDate") Else V.VentilationStartDate = Empty

V.VINAH_AccountClass = Rs("Vinah_AccountClass") & ""
V.VINAH_AccountClass_Code = Rs("Vinah_AccountClass_Code") & ""
V.Vinah_ClaimNo = Rs("Vinah_ClaimNo") & ""
V.Vinah_HealthCondition_Desc = Rs("Vinah_HealthCondition_Desc") & ""
V.Vinah_HealthCondition_Code = Rs("Vinah_HealthCondition_Code") & ""
V.Vinah_Interpreter = Rs("Vinah_Interpreter") & ""
V.Vinah_InterpreterCode = Rs("Vinah_InterpreterCode") & ""
V.Vinah_PtLivingArrangements = Rs("Vinah_PtLivingArrangements") & ""
V.Vinah_PtLivingArrangements_Code = Rs("Vinah_PtLivingArrangements_Code") & ""
V.Vinah_PtUsualAccommodation = Rs("Vinah_PtUsualAccommodation") & ""
V.Vinah_PtUsualAccommodation_Code = Rs("Vinah_PtUsualAccommodation_Code") & ""
V.Vinah_CarerStatus = Rs("Vinah_CarerStatus") & ""
V.Vinah_CarerStatus_Code = Rs("Vinah_CarerStatus_Code") & ""
V.Vinah_CarerRelationship = Rs("Vinah_CarerRelationship") & ""
V.Vinah_CarerRelationship_Code = Rs("Vinah_CarerRelationship_Code") & ""
V.Vinah_CarerResidence = Rs("Vinah_CarerResidence") & ""
V.Vinah_CarerResidence_Code = Rs("Vinah_CarerResidence_Code") & ""
V.Vinah_EpisodeEndReason = Rs("Vinah_EpisodeEndReason") & ""
V.Vinah_EpisodeEndReason_Code = Rs("Vinah_EpisodeEndReason_Code") & ""
V.Vinah_ReferredOutTo = Rs("Vinah_ReferredOutTo") & ""
V.Vinah_ReferredOutTo_Code = Rs("Vinah_ReferredOutTo_Code") & ""

Rs.Close

Get_VRSSEpisodeFromDB = V

Exit Function


Get_VRSSEpisodeFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSEpisodeFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSReviewFromDB(VisitID As Long) As VRSS_Review

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim V As VRSS_Review
Dim Msg As String

On Error GoTo Get_VRSSReviewFromDB_Error

'Clear UDT
Dim emptyUDT As VRSS_Review
V = emptyUDT

'Retrieve VRSS assessment data
sql = "SELECT * FROM Visit_NonCPAP_ImpReview WHERE VisitID =" & VisitID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving VRSS review data - aborted"
    Exit Function
End If

'Get data
V.PatientID = Rs("PatientID")
V.EpisodeID = Rs("EpisodeID")
V.VisitID = Rs("VisitID")

If IsDate(Rs("VisitDate")) Then V.VisitDate = Rs("VisitDate") Else V.VisitDate = Empty
If IsDate(Rs("VisitTime")) Then V.visitTime = Rs("VisitTime") Else V.visitTime = Empty
V.VisitLocation = Rs("VisitLocation") & ""
V.AdmissionStatus = Rs("AdmissionStatus") & ""
V.Admission_Duration = Rs("Admission_Duration") & ""
V.VisitType_Purpose = Rs("VisitType") & ""
V.VisitType_Scheduled = Rs("VisitType_Scheduled") & ""
V.Clinician = Rs("Clinician") & ""
V.Clinician_Organisation = Rs("Clinician_Organisation") & ""
V.ClinicalNote = Rs("ClinicalNote") & ""

'Get vinah data
V.Vinah_DeliveryMode = Rs("Vinah_DeliveryMode") & ""
V.Vinah_DeliveryModeCode = Rs("Vinah_DeliveryModeCode") & ""
V.Vinah_DeliveryModeSetting = Rs("Vinah_DeliveryModeSetting") & ""
V.Vinah_DeliveryModeSettingCode = Rs("Vinah_DeliveryModeSettingCode") & ""
V.Vinah_ClientpresentStatus = Rs("Vinah_ClientpresentStatus") & ""
V.Vinah_ClientpresentStatusCode = Rs("Vinah_ClientpresentStatusCode") & ""
V.Vinah_MainPurpose = Rs("Vinah_MainPurpose") & ""
V.Vinah_MainPurposeCode = Rs("Vinah_MainPurposeCode") & ""
V.Vinah_ClinicianCategory = Rs("Vinah_ClinicianCategory") & ""
V.Vinah_ClinicianCategoryCode = Rs("Vinah_ClinicianCategoryCode") & ""
V.Vinah_SessionType = Rs("Vinah_SessionType") & ""
V.Vinah_SessionTypeCode = Rs("Vinah_SessionTypeCode") & ""
'V.Vinah_Interpreter = Rs("Vinah_Interpreter") & ""
'V.Vinah_InterpreterCode = Rs("Vinah_InterpreterCode") & ""
'V.Vinah_PreferredLanguage = Rs("Vinah_PreferredLanguage") & ""
'V.Vinah_PreferredLanguage_Code = Rs("Vinah_PreferredLanguage_Code") & ""

'Mask details
V.MaskManufacturer = Rs("MaskManufacturer") & ""
V.MaskModel = Rs("MaskModel") & ""
V.FrameSize = Rs("FrameSize") & ""
V.HeadGearSize = Rs("HeadGearSize") & ""
V.HeadGearType = Rs("HeadGearType") & ""
V.CushionSize = Rs("CushionSize") & ""
V.CushionType = Rs("CushionType") & ""
V.ChinStrap = Rs("ChinStrap") & ""
V.ChinStrap_Size = Rs("ChinStrap_Size") & ""
V.O2_FiO2 = Rs("O2_FiO2") & ""
V.O2_Flow = Rs("O2_Flow") & ""
V.O2_Site = Rs("O2_Site") & ""
V.Note = Rs("Note") & ""

'Treatment settings
V.Humidifier = Rs("Humidification") & ""
V.PrescribedUsage = Rs("PrescribedUsage") & ""
ReDim V.Settings(0)
If Rs("MachineID") <> 0 Then
    Dim M As Machine_Info
    M = gReg.Get_MachineInfo(Rs("MachineID"))
    V.MachineID = Rs("machineID")
    V.SerialNo = M.SerialNo
    V.VentilationMode = Rs("VentilationMode") & ""
    V.HourMeter = Rs("Hourmeter") & ""
    'Get available settings fields and values
    For i = 1 To UBound(M.Available_Settings)
        ReDim Preserve V.Settings(UBound(V.Settings) + 1)
        V.Settings(UBound(V.Settings)).Name = M.Available_Settings(i).Name
        V.Settings(UBound(V.Settings)).Value = Rs("Sett_" & M.Available_Settings(i).Name).Value & ""
    Next i
End If
Rs.Close

Get_VRSSReviewFromDB = V

Exit Function


Get_VRSSReviewFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSReviewFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSProgram(ProgramID) As Vent_Program

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim P As Vent_Program
Dim Msg As String

On Error GoTo Get_VRSSProgram_Error

sql = "SELECT * FROM VRSS_NIVPrograms WHERE ProgramID = " & ProgramID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    P.AllocationID = Rs("AllocationID")
    P.Date_Entered = Rs("Date_Entered")
    P.Description = Rs("Description")
    P.EnteredBy = Rs("EnteredBy")
    P.ProgramID = ProgramID
End If
Rs.Close

Get_VRSSProgram = P

Exit Function


Get_VRSSProgram_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSProgram", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_VRSSEpisodeIDForReview(ReviewID As Long) As Long
'Finds the episode that the review belongs to
'Returns EpisodeID or 0 if not found

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_VRSSEpisodeIDForReview_Error

'Get list of all existing episodes for selected patient
sql = "SELECT EpisodeID FROM Visit_NonCPAP_ImpReview WHERE VisitID = " & ReviewID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Get_VRSSEpisodeIDForReview = Rs("EpisodeID")
Else
    Get_VRSSEpisodeIDForReview = 0  'should never happen unless an orphaned record
End If
Rs.Close

Exit Function


Get_VRSSEpisodeIDForReview_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSEpisodeIDForReview", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Public Function Get_VRSSEpisodeIDForNewReview(PatientID As Long) As String
'Finds the latest open episode, ir if none open then the latest closed episode
'Returns a string containing a message +/- [EpisodeID]

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim EpisodeIsOpen As Boolean
Dim nOpenEpisodes As Integer
Dim returnValue As String

On Error GoTo Get_VRSSEpisodeIDForNewReview_Error

'Get list of all existing episodes for selected patient
sql = "SELECT EpisodeID FROM VRSS_Episodes WHERE PatientID = " & PatientID & " ORDER BY Ref_RecDate DESC"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.RecordCount
    Case 0
        returnValue = "Error#1: No episode(s) exist for this patient"
    Case Else
        'Check if there is an open episode.
        nOpenEpisodes = 0
        Do While Not Rs.EOF
            If IsVRSSEpisodeOpen(Rs("EpisodeID")) Then
                nOpenEpisodes = nOpenEpisodes + 1
                returnValue = Rs("EpisodeID")
            End If
            Rs.MoveNext
        Loop
        Rs.MoveFirst    'back to the latest episode
        Select Case nOpenEpisodes
            Case 0
                returnValue = "Error#2: Only closed episodes found [" & Rs("EpisodeID") & "]"
            Case 1
                'All good
                returnValue = "Valid#1: [" & Rs("EpisodeID") & "]"
            Case Else:
                'Should never happen
                returnValue = "Error#3: >1 open episodes found"
        End Select
End Select

Get_VRSSEpisodeIDForNewReview = returnValue

Exit Function


Get_VRSSEpisodeIDForNewReview_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSEpisodeIDForNewReview", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSNIVSettingsSets(ProgramID) As Vent_Settings()

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim S() As Vent_Settings
Dim Msg As String
Dim i As Integer
Dim j As Integer

On Error GoTo Get_VRSSNIVSettingsSets_Error

ReDim S(0)
ReDim S(0).Settings(0)

sql = "SELECT * FROM VRSS_NIVSettings WHERE ProgramID = " & ProgramID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    'Get all settings data - filter to fields available for this ventilator later
    ReDim Preserve S(UBound(S) + 1)
    j = 0
    For i = 0 To Rs.Fields.Count - 1
        If InStr(Rs.Fields(i).Name, "Sett_") Then
            j = j + 1
            ReDim Preserve S(UBound(S)).Settings(j)
            S(UBound(S)).Settings(j).Name = Rs.Fields(i).Name
            S(UBound(S)).Settings(j).Value = Rs.Fields(i).Value & ""
        ElseIf Rs.Fields(i).Name = "Mode_VentilationMode" Then
            S(UBound(S)).Mode = Rs.Fields(i).Value & ""
        End If
    Next i
    Rs.MoveNext
Loop
Rs.Close

Get_VRSSNIVSettingsSets = S

Exit Function


Get_VRSSNIVSettingsSets_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSNIVSettingsSets", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSPatientIDFromEpisodeID(EpisodeID) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_VRSSPatientIDFromEpisodeID_Error

sql = "SELECT PatientID FROM VRSS_Episodes WHERE EpisodeID = " & EpisodeID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    Get_VRSSPatientIDFromEpisodeID = Rs("PatientID")
Else
    Get_VRSSPatientIDFromEpisodeID = 0
End If
Rs.Close

Exit Function

Get_VRSSPatientIDFromEpisodeID_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSPatientIDFromEpisodeID", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Get_VRSSNIVSettings(SettingsID) As Vent_Settings

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim V As Vent_Settings
Dim Msg As String

On Error GoTo Get_VRSSNIVSettings_Error


sql = "SELECT * FROM VRSS_NIVSettings WHERE SettingsID = " & SettingsID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then

    'Get all settings fields
    For i = 0 To Rs.Fields.Count - 1
        If Left(Rs.Fields(i).Name, 5) = "Sett_" Or Left(Rs.Fields(i).Name, 6) = "[Sett_" Then
            If Rs.Fields(i).Value = True Then
                ReDim Preserve M.Available_Settings(UBound(M.Available_Settings) + 1)
                M.Available_Settings(UBound(M.Available_Settings)).Name = Mid(Rs.Fields(i).Name, 6)
                M.Available_Settings(UBound(M.Available_Settings)).IsAvailable = Rs.Fields(i).Value
            End If
        ElseIf Left(Rs.Fields(i).Name, 5) = "Mode_" Then
            If Rs.Fields(i).Value = True Then
                ReDim Preserve M.Available_Modes(UBound(M.Available_Modes) + 1)
                M.Available_Modes(UBound(M.Available_Modes)).Name = Mid(Rs.Fields(i).Name, 6)
                M.Available_Modes(UBound(M.Available_Modes)).IsAvailable = Rs.Fields(i).Value
            End If
        End If
    Next i
End If
Rs.Close

Get_VRSSNIVSettings = V

Exit Function


Get_VRSSNIVSettings_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSNIVSettings", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_VRSSContactFromDB(ContactID As Long) As VRSS_ClinicalContact

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim V As VRSS_ClinicalContact
Dim Msg As String

On Error GoTo Get_VRSSContactFromDB_Error

'Clear UDT
Dim emptyUDT As VRSS_ClinicalContact
V = emptyUDT

'Retrieve VRSS assessment data
sql = "SELECT * From VRSS_ClinicalContacts WHERE ContactID = " & ContactID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving VRSS contact data - aborted"
    Exit Function
End If

'Get data
V.PatientID = Rs("PatientID")
V.ContactID = Rs("ContactID")
V.EpisodeID = Rs("EpisodeID")
If IsDate(Rs("Contact_Date")) Then V.Contact_Date = Rs("Contact_Date") Else V.Contact_Date = Empty
If IsDate(Rs("Contact_Time")) Then V.Contact_Time = Rs("Contact_Time") Else V.Contact_Time = Empty
V.Admission_Status = Rs("Admission_Status") & ""
V.Admission_Duration = Rs("Admission_Duration") & ""
V.Contact_Location = Rs("Contact_Location") & ""
V.Contact_Type_Purpose = Rs("Contact_Type_Purpose") & ""
V.Contact_Type_Scheduled = Rs("Contact_Type_Scheduled") & ""

V.Clinician_Name = Rs("Clinician_Name") & ""
V.Clinician_Organisation = Rs("Clinician_Organisation") & ""
V.ClinicalNote = Rs("ClinicalNote") & ""

'Get vinah data
V.Vinah_DeliveryMode = Rs("Vinah_DeliveryMode") & ""
V.Vinah_DeliveryModeCode = Rs("Vinah_DeliveryModeCode") & ""
V.Vinah_DeliveryModeSetting = Rs("Vinah_DeliveryModeSetting") & ""
V.Vinah_DeliveryModeSettingCode = Rs("Vinah_DeliveryModeSettingCode") & ""
V.Vinah_ClientpresentStatus = Rs("Vinah_ClientpresentStatus") & ""
V.Vinah_ClientpresentStatusCode = Rs("Vinah_ClientpresentStatusCode") & ""
V.Vinah_MainPurpose = Rs("Vinah_MainPurpose") & ""
V.Vinah_MainPurposeCode = Rs("Vinah_MainPurposeCode") & ""
V.Vinah_ClinicianCategory = Rs("Vinah_ClinicianCategory") & ""
V.Vinah_ClinicianCategoryCode = Rs("Vinah_ClinicianCategoryCode") & ""
V.Vinah_SessionType = Rs("Vinah_SessionType") & ""
V.Vinah_SessionTypeCode = Rs("Vinah_SessionTypeCode") & ""
V.Vinah_Interpreter = Rs("Vinah_Interpreter") & ""
V.Vinah_InterpreterCode = Rs("Vinah_InterpreterCode") & ""
V.Vinah_PreferredLanguage = Rs("Vinah_PreferredLanguage") & ""
V.Vinah_PreferredLanguage_Code = Rs("Vinah_PreferredLanguage_Code") & ""
V.VINAH_AccountClass = Rs("Vinah_AccountClass") & ""
V.Vinah_FileNumber = Rs("Vinah_FileNumber") & ""

Rs.Close

Get_VRSSContactFromDB = V

Exit Function


Get_VRSSContactFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VRSSContactFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_VRSSContactToDB(c As VRSS_ClinicalContact, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_VRSSContactToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM VRSS_ClinicalContacts WHERE ContactID =" & c.ContactID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving VRSS contact data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM VRSS_ClinicalContacts WHERE ContactID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = c.PatientID
End Select


'Save data
Rs("EpisodeID") = c.EpisodeID

If IsDate(c.Contact_Date) Then Rs("Contact_Date") = c.Contact_Date Else Rs("Contact_Date") = Null
If IsDate(c.Contact_Time) Then Rs("Contact_Time") = c.Contact_Time Else Rs("Contact_Time") = Null
Rs("Admission_Status") = c.Admission_Status
Rs("Admission_Duration") = c.Admission_Duration & ""
Rs("Contact_Location") = c.Contact_Location & ""
Rs("Contact_Type_Purpose") = c.Contact_Type_Purpose & ""
Rs("Contact_Type_Scheduled") = c.Contact_Type_Scheduled & ""

Rs("Clinician_Name") = c.Clinician_Name & ""
Rs("Clinician_Organisation") = c.Clinician_Organisation & ""
Rs("ClinicalNote") = c.ClinicalNote & ""

'Get vinah data
Rs("Vinah_DeliveryMode") = c.Vinah_DeliveryMode & ""
Rs("Vinah_DeliveryModeCode") = c.Vinah_DeliveryModeCode & ""
Rs("Vinah_DeliveryModeSetting") = c.Vinah_DeliveryModeSetting & ""
Rs("Vinah_DeliveryModeSettingCode") = c.Vinah_DeliveryModeSettingCode & ""
Rs("Vinah_ClientpresentStatus") = c.Vinah_ClientpresentStatus & ""
Rs("Vinah_ClientpresentStatusCode") = c.Vinah_ClientpresentStatusCode & ""
Rs("Vinah_MainPurpose") = c.Vinah_MainPurpose & ""
Rs("Vinah_MainPurposeCode") = c.Vinah_MainPurposeCode & ""
Rs("Vinah_ClinicianCategory") = c.Vinah_ClinicianCategory & ""
Rs("Vinah_ClinicianCategoryCode") = c.Vinah_ClinicianCategoryCode & ""
Rs("Vinah_SessionType") = c.Vinah_SessionType & ""
Rs("Vinah_SessionTypeCode") = c.Vinah_SessionTypeCode & ""
Rs("Vinah_Interpreter") = c.Vinah_Interpreter & ""
Rs("Vinah_InterpreterCode") = c.Vinah_InterpreterCode & ""
Rs("Vinah_PreferredLanguage") = c.Vinah_PreferredLanguage & ""
Rs("Vinah_PreferredLanguage_Code") = c.Vinah_PreferredLanguage_Code & ""
Rs("Vinah_AccountClass") = c.VINAH_AccountClass & ""
Rs("Vinah_AccountClassCode") = c.Vinah_AccountClassCode & ""
Rs("Vinah_FileNumber") = c.Vinah_FileNumber & ""

Rs.Update
Rs.Close

'Save_VRSSContactToDB =

Exit Function


Save_VRSSContactToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSContactToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function



Private Function Get_MachineSettingAvailableStatus(E As Machine_Info, SettingName As String) As Boolean

Dim i As Integer
Dim Msg As String

On Error GoTo Get_MachineSettingAvailableStatus_Error

Get_MachineSettingAvailableStatus = False

If UBound(E.Available_Settings) = 0 Then
    Exit Function
End If

For i = 1 To UBound(E.Available_Settings)
    If SettingName = E.Available_Settings(i).Name Then
        Get_MachineSettingAvailableStatus = E.Available_Modes(i).IsAvailable
    End If
Next i

Exit Function

Get_MachineSettingAvailableStatus_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_MachineSettingAvailableStatus", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_VRSSReviewToDB(V As VRSS_Review, Mode As RecordMode) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim i As Integer

On Error GoTo Save_VRSSReviewToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM Visit_NonCPAP_ImpReview WHERE VisitID =" & V.VisitID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving VRSS review data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM Visit_NonCPAP_ImpReview WHERE VisitID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = V.PatientID
End Select

'Save data
Rs("EpisodeID") = V.EpisodeID

If IsDate(V.VisitDate) Then Rs("VisitDate") = V.VisitDate Else Rs("VisitDate") = Null
If IsDate(V.visitTime) Then Rs("VisitTime") = V.visitTime Else Rs("VisitTime") = Null
Rs("VisitLocation") = V.VisitLocation
Rs("AdmissionStatus") = V.AdmissionStatus
Rs("Admission_Duration") = V.Admission_Duration
Rs("VisitType") = V.VisitType_Purpose
Rs("VisitType_Scheduled") = V.VisitType_Scheduled
Rs("Clinician") = V.Clinician
Rs("Clinician_Organisation") = V.Clinician_Organisation

'Get vinah data
Rs("Vinah_DeliveryMode") = V.Vinah_DeliveryMode & ""
Rs("Vinah_DeliveryModeCode") = V.Vinah_DeliveryModeCode & ""
Rs("Vinah_DeliveryModeSetting") = V.Vinah_DeliveryModeSetting & ""
Rs("Vinah_DeliveryModeSettingCode") = V.Vinah_DeliveryModeSettingCode & ""
Rs("Vinah_ClientpresentStatus") = V.Vinah_ClientpresentStatus & ""
Rs("Vinah_ClientpresentStatusCode") = V.Vinah_ClientpresentStatusCode & ""
Rs("Vinah_MainPurpose") = V.Vinah_MainPurpose & ""
Rs("Vinah_MainPurposeCode") = V.Vinah_MainPurposeCode & ""
Rs("Vinah_ClinicianCategory") = V.Vinah_ClinicianCategory & ""
Rs("Vinah_ClinicianCategoryCode") = V.Vinah_ClinicianCategoryCode & ""
Rs("Vinah_SessionType") = V.Vinah_SessionType & ""
Rs("Vinah_SessionTypeCode") = V.Vinah_SessionTypeCode & ""
'Rs("Vinah_Interpreter") = V.Vinah_Interpreter & ""
'Rs("Vinah_InterpreterCode") = V.Vinah_InterpreterCode & ""
'Rs("Vinah_PreferredLanguage") = V.Vinah_PreferredLanguage & ""
'Rs("Vinah_PreferredLanguage_Code") = V.Vinah_PreferredLanguage_Code & ""

'Mask details
Rs("MaskManufacturer") = V.MaskManufacturer
Rs("MaskModel") = V.MaskModel
Rs("FrameSize") = V.FrameSize
Rs("HeadGearSize") = V.HeadGearSize
Rs("HeadGearType") = V.HeadGearType
Rs("CushionSize") = V.CushionSize
Rs("CushionType") = V.CushionType
Rs("ChinStrap") = V.ChinStrap
Rs("ChinStrap_Size") = V.ChinStrap_Size
Rs("O2_FiO2") = V.O2_FiO2
Rs("O2_Flow") = V.O2_Flow
Rs("O2_Site") = V.O2_Site
Rs("Note") = V.Note

'Treatment settings
Rs("Humidification") = V.Humidifier
Rs("PrescribedUsage") = V.PrescribedUsage
If V.MachineID <> 0 Then 'changed by WRR 8/01/2014 due to compile problenm with Profusion add-in
    'Rs("MachineID") = V.EquipID
    Rs("MachineID") = V.MachineID
    Rs("VentilationMode") = V.VentilationMode
    Rs("Hourmeter") = V.HourMeter
    'Clear all settings fields first
    For i = 0 To Rs.Fields.Count - 1
        If Left(Rs(i).Name, 5) = "Sett_" Then Rs(i).Value = ""
    Next i
    'Save settings
    For i = 1 To UBound(V.Settings)
        Rs("Sett_" & V.Settings(i).Name).Value = V.Settings(i).Value
    Next i
Else
    Rs("MachineID") = 0
    Rs("VentilationMode") = ""
    Rs("Hourmeter") = ""
    'Clear all settings fields
    For i = 0 To Rs.Fields.Count - 1
        If Left(Rs(i).Name, 5) = "Sett_" Then Rs(i).Value = ""
    Next i
End If

Rs.Update
Rs.MoveFirst
Save_VRSSReviewToDB = Rs("VisitID")
Rs.Close

Exit Function


Save_VRSSReviewToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_VRSSReviewToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function




Friend Function Save_SlpIDToLinkedCpapRecord(StudyID As Long, VisitID As Long) As Boolean

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_SlpIDToLinkedCpapRecord_Error

sql = "SELECT StudyID FROM Visit_CPAP_ImpReview WHERE VisitID =" & VisitID
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
If Not Rs.EOF Then
    Rs("StudyID") = StudyID
    Rs.Update
Else

End If
Rs.Close
Save_SlpIDToLinkedCpapRecord = True

Exit Function


Save_SlpIDToLinkedCpapRecord_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_SlpIDToLinkedCpapRecord", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Save_SlpIDToLinkedCpapRecord = False
    Exit Function
    Resume
End Function

Friend Function Get_VisitIDForLinkedCpapRecord(StudyID As Long) As Long

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Get_VisitIDForLinkedCpapRecord_Error

sql = "SELECT VisitID FROM Visit_CPAP_ImpReview WHERE StudyID =" & StudyID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.RecordCount
    Case 1: Get_VisitIDForLinkedCpapRecord = Rs("VisitID")
    Case 0: Get_VisitIDForLinkedCpapRecord = 0
    Case Else: MsgBox "More than one linked CPAP record found - contact database admin"
End Select
Rs.Close

Exit Function


Get_VisitIDForLinkedCpapRecord_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_VisitIDForLinkedCpapRecord", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Get_VisitIDForLinkedCpapRecord = 0
    Exit Function
    Resume
End Function

Friend Function Get_CpapFromDB(VisitID As Long) As CPAPClinic

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim c As CPAPClinic
Dim Msg As String

On Error GoTo Get_CpapFromDB_Error

'Clear UDT
Dim emptyUDT As CPAPClinic
c = emptyUDT

'Retrieve CPAP clinic data
sql = "SELECT tbl_Name.*, tbl_RM_Demographics.*, Visit_CPAP_ImpReview.*, "
sql = sql & "tbl_patient.ID, tbl_patient.URNumber, tbl_patient.Type AS Expr1, tbl_patient.DOB, tbl_patient.Gender, "
sql = sql & "tbl_patient.MaritalStatus, tbl_patient.Language, tbl_patient.Religion, tbl_patient.Race, "
sql = sql & "tbl_patient.Nationality, tbl_patient.Citizenship, tbl_patient.MothersID, tbl_patient.DateofDeath, "
sql = sql & "tbl_patient.DeathIndicator, tbl_patient.IdentityUnknown, tbl_patient.PrimaryFacility, "
sql = sql & "tbl_patient.LastModifiedBy, tbl_patient.LastModifiedDate, tbl_patient.MergedUR, "
sql = sql & "tbl_Address.Type AS Expr2, tbl_Address.TypeID AS Expr3, tbl_Address.AddressID, tbl_Address.Address1, tbl_Address.Address2, "
sql = sql & "tbl_Address.Suburb, tbl_Address.State, tbl_Address.Postcode, tbl_Address.Country , tbl_Address.AddressType "
sql = sql & "FROM tbl_patient INNER JOIN "
sql = sql & "tbl_Name ON tbl_patient.ID = tbl_Name.TypeID INNER JOIN Visit_CPAP_ImpReview ON "
sql = sql & "tbl_patient.ID = Visit_CPAP_ImpReview.PatientID LEFT OUTER JOIN tbl_Address ON tbl_patient.ID = "
sql = sql & "tbl_Address.TypeID LEFT OUTER JOIN tbl_RM_Demographics ON tbl_patient.ID = tbl_RM_Demographics.PatientID "
sql = sql & "WHERE (tbl_Name.Type = 'P') AND (tbl_Address.AddressType = '4' OR tbl_Address.AddressType IS NULL) AND (tbl_Address.Type = 'P') AND "
sql = sql & "(tbl_Name.Type = 'P') AND (tbl_patient.Type = 'P') AND (Visit_CPAP_ImpReview.VisitID = " & VisitID & ")"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving CPAP clinic data - aborted"
    Exit Function
End If
    
c.D_DOB = Format(Rs("DOB"), "dd/mm/yyyy")
c.D_Gender = Map_RefTbl(Rs("Gender"), Gender, HL7Code_RMText)
c.D_Address1 = ChangeCase(Rs("Address1"), Titlecase)
c.D_Address2 = ChangeCase(Rs("Address2"), Titlecase)
c.D_FirstName = ChangeCase(Rs("FirstName"), Titlecase)
c.D_Middlename = ChangeCase(Rs("Middlename"), Titlecase)
c.D_Postcode = Rs("Postcode") & ""
c.D_State = ChangeCase(Rs("State"), Titlecase)
c.D_Suburb = Rs("Suburb") & ""
c.D_Surname = ChangeCase((Rs("Surname")), Uppercase)
c.D_UR_PrimaryAustin = Rs("URNumber")
c.D_URforDisplay = Get_PtString(c.PatientID, cHS.CurrentVAED, DisplayUR)
   
c.VisitID = Rs("VisitID")
c.PatientID = Rs("PatientID")

'Visit stuff
If IsDate(Rs("VisitDate")) Then c.Visit_Date = Rs("VisitDate") Else c.Visit_Date = Empty
If IsDate(Rs("VisitTime")) Then c.Visit_Time = Rs("VisitTime") Else c.Visit_Time = Empty
c.Visit_AdmissionStatus = Rs("AdmissionStatus") & ""
c.Visit_Type = Rs("VisitType") & ""
c.Visit_Location = Rs("VisitLocation") & ""
c.Visit_Clinician = Rs("Clinician") & ""
c.Clinic_ReferralMode = Rs("Clinic_ReferralMode") & ""
c.Clinic_ReferredBy = Rs("Clinic_ReferredBy") & ""
If IsDate(Rs("Clinic_ReferralDate")) Then c.Clinic_ReferralDate = Rs("Clinic_ReferralDate") Else c.Clinic_ReferralDate = Empty
c.Clinic_ReportTo = Rs("Clinic_ReportTo") & ""
c.Clinic_ReferralReason = Rs("Clinic_ReferralReason") & ""

'Equipment stuff
If Not IsNull(Rs("MachineID")) Then c.MachineID = Rs("MachineID") Else c.MachineID = 0
c.Mask_Manufacturer = Rs("MaskManufacturer") & ""
c.Mask_Model = Rs("MaskModel") & ""
c.Mask_Frame = Rs("Frame") & ""
c.Mask_HeadgearType = Rs("HeadGearType") & ""
c.Mask_HeadgearSize = Rs("HeadGearSize") & ""
c.Mask_CushionSize = Rs("CushionSize") & ""
c.ChinStrap = Rs("Chinstrap") & ""
c.CPAP_HumidifierType = Rs("CPAP_HumidifierType") & ""

'Treatment settings
c.TreatmentMode = Rs("TreatmentMode") & ""
If IsNull(Rs("CPAP_Set")) Then c.CPAP_Set = "" Else c.CPAP_Set = Rs("CPAP_Set")
If IsNull(Rs("APAP_Min")) Then c.APAP_Min = "" Else c.APAP_Min = Rs("APAP_Min")
If IsNull(Rs("APAP_Max")) Then c.APAP_Max = "" Else c.APAP_Max = Rs("APAP_Max")
If IsNull(Rs("CPAP_Measured")) Then c.CPAP_Measured = "" Else c.CPAP_Measured = Rs("CPAP_Measured")
If IsNull(Rs("Bilevel_EPAP")) Then c.Bilevel_EPAP = "" Else c.Bilevel_EPAP = Rs("Bilevel_EPAP")
If IsNull(Rs("Bilevel_IPAP")) Then c.Bilevel_IPAP = "" Else c.Bilevel_IPAP = Rs("Bilevel_IPAP")
If IsNull(Rs("O2Flow")) Then c.O2_Flow = "" Else c.O2_Flow = Rs("O2Flow")
c.O2_EntrainedLocation = Rs("O2EntrainedSpot") & ""
c.Settings_SetOrPrescribed = Rs("Settings_SetOrPrescribed") & ""

'Adherence stuff
c.HourMeter = Rs("HourMeter") & ""
If IsDate(Rs("CPAP_Commenced")) Then c.CPAP_Commenced = Rs("CPAP_Commenced") Else c.CPAP_Commenced = Empty
c.Adh_Obj_PerNight = Rs("Adh_Obj_PerNight") & ""
c.Adh_Obj_PerWeek = Rs("Adh_Obj_PerWeek") & ""
c.Adh_Obj_Period = Rs("Adh_Obj_Period") & ""
c.Adh_Rep_PerNight = Rs("Adh_Rep_PerNight") & ""
c.Adh_Rep_PerWeek = Rs("Adh_Rep_PerWeek") & ""
c.Adh_Rep_Period = Rs("Adh_Rep_Period") & ""

'DHS stuff
If Not IsNull(Rs("DHS_CPAP_ApprovedForPump")) Then c.DHS_CPAP_ApprovedForPump = Rs("DHS_CPAP_ApprovedForPump") Else c.DHS_CPAP_ApprovedForPump = False
If Not IsNull(Rs("DHS_CPAP_HireTrialFirst")) Then c.DHS_CPAP_HireTrialFirst = Rs("DHS_CPAP_HireTrialFirst") Else c.DHS_CPAP_HireTrialFirst = False
c.DHS_CPAP_Approval_EnteredBy = Rs("DHS_CPAP_Approval_EnteredBy") & ""
If IsDate(Rs("DHS_CPAP_Approval_Date")) Then c.DHS_CPAP_Approval_Date = Rs("DHS_CPAP_Approval_Date") Else c.DHS_CPAP_Approval_Date = Empty
If Not IsNull(Rs("DHS_CPAP_NewOrSecondHand")) Then c.DHS_CPAP_NewOrSecondHand = Rs("DHS_CPAP_NewOrSecondHand") Else c.DHS_CPAP_NewOrSecondHand = False
c.DHS_CPAP_Copayment_Fee = Rs("DHS_CPAP_Copayment_Fee") & ""
c.DHS_CPAP_Copayment_Received = Rs("DHS_CPAP_Copayment_Received") & ""
c.DHS_CPAP_Copayment_Location = Rs("DHS_CPAP_Copayment_Location") & ""
c.DHS_CPAP_Copayment_Method = Rs("DHS_CPAP_Copayment_Method") & ""
If Not IsNull(Rs("DHS_CPAP_Copayment_ReceiptIssued")) Then c.DHS_CPAP_Copayment_ReceiptIssued = Rs("DHS_CPAP_Copayment_ReceiptIssued") Else c.DHS_CPAP_Copayment_ReceiptIssued = False
c.DHS_CPAP_Copayment_TransactedBy = Rs("DHS_CPAP_Copayment_TransactedBy") & ""
If IsDate(Rs("DHS_CPAP_Copayment_TransactDate")) Then c.DHS_CPAP_Copayment_TransactDate = Rs("DHS_CPAP_Copayment_TransactDate") Else c.DHS_CPAP_Copayment_TransactDate = Empty
c.DHS_CPAP_AuthorisingPhysician = Rs("DHS_CPAP_AuthorisingPhysician") & ""
c.BillingItems = Rs("BillingItems") & ""

'Clinical stuff
c.Visit_Note = Rs("Note") & ""
c.ESS = Rs("ESS") & ""
c.CPAP_ReportedAdherencePeriod = Rs("CPAP_ReportedAdherencePeriod") & ""
c.CPAP_Condensation = Rs("CPAP_Condensation") & ""
c.CPAP_Drymouth = Rs("CPAP_Drymouth") & ""
c.CPAP_Drynose = Rs("CPAP_Drynose") & ""
c.CPAP_Pressure = Rs("CPAP_Pressure") & ""
c.CPAP_Runny = Rs("CPAP_Runny") & ""
c.Mask_Eyes = Rs("Mask_Eyes") & ""
c.Mask_Face = Rs("Mask_Face") & ""
c.Mask_Leaks = Rs("Mask_Leaks") & ""
c.Mask_Comfort = Rs("Mask_Comfort") & ""
c.Symptoms_EDS = Rs("Symptoms_EDS") & ""
c.Symptoms_Headache = Rs("Symptoms_Headache") & ""
c.Symptoms_Napping = Rs("Symptoms_Napping") & ""

Get_CpapFromDB = c

Exit Function


Get_CpapFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_CpapFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Get_RFTs(PatientID As Long, StartDate As Date, EndDate As Date) As SetOfRFTs()

Dim Rs As New ADODB.Recordset
Dim sql As String, sSQL As String
Dim Msg As String
Dim R() As SetOfRFTs
Dim D As PtDemographics
Dim dummyD As PtDemographics

On Error GoTo Get_RFTs_Error

ReDim R(0)

sql = "SELECT 'RftResults' AS SourceTable, PatientID, ResultsID, TestType, TestTime, ReportStatus, Campus, MO, TestDate, RequestDate FROM RftResults "
sql = sql & " WHERE (TestDate BETWEEN " & gDBFunctions.Convert_ToSQLQueryDate(StartDate)
sql = sql & " AND " & gDBFunctions.Convert_ToSQLQueryDate(EndDate) & ") AND (PatientID=" & PatientID & ")"
sql = sql & " UNION "
sql = sql & "SELECT 'ExResults' AS SourceTable, PatientID, ResultsID, TestType, TestTime, ReportStatus, Campus, MO, TestDate, RequestDate FROM ExResults "
sql = sql & " WHERE (TestDate BETWEEN " & gDBFunctions.Convert_ToSQLQueryDate(StartDate)
sql = sql & " AND " & gDBFunctions.Convert_ToSQLQueryDate(EndDate) & ") AND (PatientID=" & PatientID & ")"
sql = sql & " ORDER BY TestDate DESC"

Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    If Not IsNull(Rs("PatientID")) Then
        D = dummyD  'Clear D
        D = Pt.Get_DemographicsFromDB(Rs("PatientID"))
        ReDim Preserve R(UBound(R) + 1)
        R(UBound(R)).SourceTable = Rs("SourceTable")
        R(UBound(R)).PatientID = D.ID
        R(UBound(R)).Surname = D.Surname
        R(UBound(R)).Firstname = D.Firstname & ""
        R(UBound(R)).DOB = D.DOB
        R(UBound(R)).Gender = D.Gender
        R(UBound(R)).UR_PrimaryAustin = D.UR_PrimaryAustin
        R(UBound(R)).URforDisplay = D.URforDisplay
        If IsDate(Rs("TestDate")) Then R(UBound(R)).TestDate = Rs("TestDate") Else R(UBound(R)).TestDate = Empty
        If IsDate(Rs("RequestDate")) Then R(UBound(R)).RequestDate = Rs("RequestDate") Else R(UBound(R)).RequestDate = Empty
        R(UBound(R)).TestTime = Rs("TestTime")
        R(UBound(R)).Lab = Rs("Campus") & ""
        R(UBound(R)).RequestingMO = Rs("MO") & ""
        R(UBound(R)).TestType = Rs("TestType")
        R(UBound(R)).ReportStatus = Rs("ReportStatus")
        R(UBound(R)).ResultsID = Rs("ResultsID")
    End If
    Rs.MoveNext
Loop
Rs.Close

Get_RFTs = R

Exit Function

Get_RFTs_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_RFTs", "cGeneralDBRoutines", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Get_RftFromDB(RftID As Long) As RFTs

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim R As RFTs
Dim Msg As String

On Error GoTo Get_RftFromDB_Error

'Clear UDT
Dim emptyUDT As RFTs
R = emptyUDT

'Retrieve Rft data
sql = "SELECT tbl_Name.*, tbl_RM_Demographics.*, RftResults.*, "
sql = sql & "tbl_patient.ID, tbl_patient.URNumber, tbl_patient.Type AS Expr1, tbl_patient.DOB, tbl_patient.Gender, "
sql = sql & "tbl_patient.MaritalStatus, tbl_patient.Language, tbl_patient.Religion, tbl_patient.Race, "
sql = sql & "tbl_patient.Nationality, tbl_patient.Citizenship, tbl_patient.MothersID, tbl_patient.DateofDeath, "
sql = sql & "tbl_patient.DeathIndicator, tbl_patient.IdentityUnknown, tbl_patient.PrimaryFacility, "
sql = sql & "tbl_patient.LastModifiedBy, tbl_patient.LastModifiedDate, tbl_patient.MergedUR, "
sql = sql & "tbl_Address.Type AS Expr2, tbl_Address.TypeID AS Expr3, tbl_Address.AddressID, tbl_Address.Address1, tbl_Address.Address2, "
sql = sql & "tbl_Address.Suburb, tbl_Address.State, tbl_Address.Postcode, tbl_Address.Country , tbl_Address.AddressType "
sql = sql & "FROM tbl_patient INNER JOIN "
sql = sql & "tbl_Name ON tbl_patient.ID = tbl_Name.TypeID INNER JOIN RftResults ON "
sql = sql & "tbl_patient.ID = RftResults.PatientID LEFT OUTER JOIN tbl_Address ON tbl_patient.ID = "
sql = sql & "tbl_Address.TypeID LEFT OUTER JOIN tbl_RM_Demographics ON tbl_patient.ID = tbl_RM_Demographics.PatientID "
sql = sql & "WHERE (tbl_Name.Type = 'P') AND (tbl_Address.AddressType = '4' OR tbl_Address.AddressType IS NULL) AND (tbl_Address.Type = 'P') AND "
sql = sql & "(tbl_Name.Type = 'P') AND (tbl_patient.Type = 'P') AND (RftResults.ResultsID = " & RftID & ")"
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving RFT results - aborted"
    Exit Function
End If
    
R.D_DOB = Format(Rs("DOB"), "dd/mm/yyyy")
R.D_Gender = Map_RefTbl(Rs("Gender"), Gender, HL7Code_RMText)
R.D_Race_Rft = CurrentPt.Get_RaceRft(Rs("PatientID"))
R.D_Address1 = ChangeCase(Rs("Address1"), Titlecase)
R.D_Address2 = ChangeCase(Rs("Address2"), Titlecase)
R.D_FirstName = ChangeCase(Rs("FirstName"), Titlecase)
R.D_Middlename = ChangeCase(Rs("Middlename"), Titlecase)
R.D_Postcode = Rs("Postcode") & ""
R.D_State = ChangeCase(Rs("State"), Titlecase)
R.D_Suburb = Rs("Suburb") & ""
R.D_Surname = ChangeCase((Rs("Surname")), Uppercase)
R.D_UR_PrimaryAustin = Rs("URNumber")
R.D_URforDisplay = Get_PtString(R.PtID, cHS.CurrentVAED, DisplayUR)
   
R.RecordID = Rs("ResultsID")
R.PtID = Rs("PatientID")
R.E_Lab = Rs("Campus") & ""
R.E_RequestingMO = Rs("MO") & ""
R.E_RequestingMO_Pn = Rs("RequestingMO_Pn") & ""
R.E_RefLocation = Rs("RefLocation") & ""
Select Case IsNull(Rs("Request_HealthServiceID"))
    Case True
        R.E_Ref_HealthService = ""
        R.E_Ref_HealthServiceID = 0
    Case False
        R.E_Ref_HealthService = cHS.Get_HSFromVAED(Rs("Request_HealthServiceID"))
        R.E_Ref_HealthServiceID = Rs("Request_HealthServiceID") & ""
End Select
R.E_ReportCopyTo = Rs("Report_CopyTo") & ""
R.TestType = Rs("TestType")
R.E_TestDate = Format(Rs("TestDate"), "dd/mm/yyyy")
If IsDate(Rs("RequestDate")) Then R.E_RequestDate = Format(Rs("RequestDate"), "dd/mm/yyyy") Else R.E_RequestDate = Empty
If IsDate(Rs("TestTime")) Then R.E_TestTime = Format(Rs("TestTime"), "hh:mm") Else R.E_TestTime = Empty
R.Scientist = Rs("scientist") & ""
R.TechComments = Rs("comments") & ""
R.E_Smoke = Rs("Smoke") & ""
R.E_Packyrs = Rs("Packyrs") & ""
R.E_ThisVisit_weight = Rs("Weight") & ""
R.E_ThisVisit_height = Rs("Height_rft") & ""
R.Hb_WaitingForResult = Abs(CInt(Rs("Hb_WaitingForResult")))
R.E_ClinicalNote = Rs("ClinicalNote") & ""
R.E_AdmissionStatus = Rs("AdmissionStatus") & ""
R.E_BillingMO = Rs("BillingMO") & ""
R.E_BillingItems = Rs("BillingItems") & ""
R.E_BillingStatus = Rs("BillingStatus") & ""
R.E_BillingMO_ProviderNumber = Rs("BillingMO_ProviderNumber") & ""
R.D_MedicareNo = GetMedicareNo(R.PtID, Formatted)

R.spir = Rs("spir") & ""
R.BD_1 = Rs("bd1") & ""
R.BD_2 = Rs("bd2") & ""
R.LastBd = Rs("LastBd") & ""

R.preFEV1 = Format(Rs("prefev1"), "0.00")
R.preFVC = Format(Rs("prefvc"), "0.00")
R.preVC = Format(Rs("prevc"), "0.00")
R.postFEV1_1 = Format(Rs("post1fev1"), "0.00")
R.postFVC_1 = Format(Rs("post1fvc"), "0.00")
R.postVC_1 = Format(Rs("post1vc"), "0.00")
R.postFEV1_2 = Format(Rs("post2fev1"), "0.00")
R.postFVC_2 = Format(Rs("post2fvc"), "0.00")
R.postVC_2 = Format(Rs("post2vc"), "0.00")
R.preFEF = Format(Rs("preFEF"), "#0.0")
R.postFEF_1 = Format(Rs("post1FEF"), "#0.0")
R.postFEF_2 = Format(Rs("post2FEF"), "#0.0")

R.tlcoEquip = Rs("tlcoEquip") & ""
R.VA = Format(Rs("va"), "0.0")
R.TLCO = Format(Rs("tlco"), "##.0")
R.KCO = Format(Rs("kco"), "#0.0")

R.Hb = Rs("Hb") & ""
R.HbInfo = Rs("HbInfo") & ""

R.volEquip = Rs("volEquip") & ""
R.FRC = Format(Rs("frc"), "#0.00")
R.TLC = Format(Rs("tlc"), "#0.00")
R.VC = Format(Rs("vc"), "0.00")
R.RV = Format(Rs("rv"), "0.00")
R.Raw = Format(Rs("raw"), "#0.0")
R.Sgaw = Format(Rs("sgaw"), "0.00")
R.Raw_2 = Format(Rs("raw2"), "#0.0")
R.Sgaw_2 = Format(Rs("sgaw2"), "0.00")

R.FiO2_1 = Rs("fio21") & ""
R.pH_1 = Format(Rs("ph1"), "0.00")
R.pCO2_1 = Format(Rs("pco21"), "###")
R.pO2_1 = Format(Rs("po21"), "###")
R.Shunt_1 = Format(Rs("shunt1"), "#0.0")
R.BE_1 = Format(Rs("BE1"), "#0.0")
R.HCO3_1 = Format(Rs("HCO31"), "#0.0")
R.SaO2_1 = Format(Rs("SaO21"), "###")
R.FiO2_2 = Rs("fio22") & ""
R.pH_2 = Format(Rs("ph2"), "0.00")
R.pCO2_2 = Format(Rs("pco22"), "###")
R.pO2_2 = Format(Rs("po22"), "###")
R.Shunt_2 = Format(Rs("shunt2"), "#0.0")
R.BE_2 = Format(Rs("BE2"), "#0.0")
R.HCO3_2 = Format(Rs("HCO32"), "#0.0")
R.SaO2_2 = Format(Rs("SaO22"), "###")

R.MIP = Format(Rs("mips"), "###")
R.MEP = Format(Rs("meps"), "###")
R.SNIP = Format(Rs("SNIP"), "###")

R.prov_FEV1pre = Rs("prefev1") & ""
R.prov_FVCpre = Rs("prefvc") & ""
R.prov_VCpre = Rs("prevc") & ""
R.prov_ChallengeType = Rs("ChallengeType") & ""
R.prov_Control = Rs("control") & ""
R.prov_Dose1 = Rs("Dose1") & ""
R.prov_Dose2 = Rs("Dose2") & ""
R.prov_Dose3 = Rs("Dose3") & ""
R.prov_Dose4 = Rs("Dose4") & ""
R.prov_Dose5 = Rs("Dose5") & ""
R.prov_Dose6 = Rs("Dose6") & ""
R.prov_Dose7 = Rs("Dose7") & ""
R.prov_Dose8 = Rs("Dose8") & ""
R.prov_Dose9 = Rs("Dose9") & ""
R.prov_Post = Rs("post") & ""
R.prov_PostBDType = Rs("PostProvBDType") & ""
R.prov_FEV1_Control = Rs("FEV1-Control") & ""
R.prov_FEV1_Dose1 = Rs("FEV1-1") & ""
R.prov_FEV1_Dose2 = Rs("FEV1-2") & ""
R.prov_FEV1_Dose3 = Rs("FEV1-3") & ""
R.prov_FEV1_Dose4 = Rs("FEV1-4") & ""
R.prov_FEV1_Dose5 = Rs("FEV1-5") & ""
R.prov_FEV1_Dose6 = Rs("FEV1-6") & ""
R.prov_FEV1_Dose7 = Rs("FEV1-7") & ""
R.prov_FEV1_Dose8 = Rs("FEV1-8") & ""
R.prov_FEV1_Dose9 = Rs("FEV1-9") & ""
R.prov_FEV1_Post = Rs("FEV1-post") & ""

R.SkinTestDataString = Rs("SkinDataString") & ""
R.Skin_PanelDescription = Rs("Skin_PanelDescription") & ""

R.AST_TestDataString = Rs("ASTString") & ""

R.AirSaO2String = Rs("AirSaO2String") & ""
R.O2SaO2String = Rs("O2SaO2String") & ""
R.SpeedString = Rs("SpeedString") & ""

R.Report = Rs("Report") & ""
R.Reporter = Rs("Reporter") & ""
If IsDate(Rs("ReportDate")) Then R.ReportDate = Format(Rs("ReportDate"), "dd/mm/yyyy") Else R.ReportDate = Empty
If IsDate(Rs("VerifiedDate")) Then R.VerifiedDate = Format(Rs("VerifiedDate"), "dd/mm/yyyy") Else R.VerifiedDate = Empty
R.ReportStatus = Rs("ReportStatus")
R.VerifiedBy = Rs("VerifiedBy") & ""

Get_RftFromDB = R

Exit Function

Get_RftFromDB_Error:
    Msg = ""
   Call ErrorLog(Msg, "Get_RftFromDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_RftEpisodeData_ForID(ResultsID As Long, SourceTblEnum As DbTables) As RFT_Episode

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim R As RFT_Episode
Dim Msg As String
Dim Tbl As String

On Error GoTo Get_RftEpisodeData_ForID_Error
    
Select Case SourceTblEnum
    Case DbTables.RftResults:   Tbl = "RftResults"
    Case DbTables.ExResults:    Tbl = "ExResults"
End Select

'Retrieve Rft data
sql = "SELECT * FROM " & Tbl & " WHERE ResultsID=" & ResultsID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Rs.EOF Then
    MsgBox "Problem retrieving RFT results - aborted"
    Exit Function
End If
    
R.RecordID = Rs("ResultsID")
R.PtID = Rs("PatientID")
R.SourceTbl = SourceTblEnum
If IsDate(Rs("RequestDate")) Then R.RequestDate = Format(Rs("RequestDate"), "dd/mm/yyyy") Else R.RequestDate = Empty
R.RequestingMO = Rs("MO") & ""
R.RequestingMO_Pn = Rs("RequestingMO_Pn") & ""
Select Case IsNull(Rs("Request_HealthServiceID"))
    Case True
        R.Ref_HealthService = ""
        R.Ref_HealthServiceID = 0
    Case False
        R.Ref_HealthService = cHS.Get_HSFromVAED(Rs("Request_HealthServiceID"))
        R.Ref_HealthServiceID = Rs("Request_HealthServiceID") & ""
End Select
R.RefLocation = Rs("RefLocation") & ""
R.ReportCopyTo = Rs("Report_CopyTo") & ""
R.TestDate = Format(Rs("TestDate"), "dd/mm/yyyy")
If IsDate(Rs("TestTime")) Then R.TestTime = Format(Rs("TestTime"), "hh:mm") Else R.TestTime = Empty
R.Smoke = Rs("Smoke") & ""
R.Packyrs = Rs("Packyrs") & ""
R.ThisVisit_weight = Rs("Weight") & ""
R.LastVisit_weight = CurrentPt.Get_LastRFT_Wt(R.TestDate, R.PtID)
R.ThisVisit_height = Rs("Height_rft") & ""
R.LastVisit_height = CurrentPt.Get_LastRFT_Ht(R.TestDate, R.PtID)
R.Lab = Rs("Campus") & ""
R.D_Race_Rft = CurrentPt.Get_RaceRft(Rs("PatientID"))
R.AdmissionStatus = Rs("AdmissionStatus") & ""
R.ClinicalNote = Rs("ClinicalNote") & ""
R.D_MedicareNo = GetMedicareNo(R.PtID, Formatted)
R.BillingMO = Rs("BillingMO") & ""
R.BillingItems = Rs("BillingItems") & ""
R.BillingStatus = Rs("BillingStatus") & ""
R.BillingMO_ProviderNumber = Rs("BillingMO_ProviderNumber") & ""

Get_RftEpisodeData_ForID = R

Exit Function

Get_RftEpisodeData_ForID_Error:
   Msg = ""
   Call ErrorLog(Msg, "Get_RftEpisodeData_ForID", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_RftEpisodeData_ForDate(PtID As Long, TestDate As Date) As RFT_Episode()
' Episode data will be replicated in each RFT/Exercise record when multiple tests are performed on a single day.
' Will return an array of RFT_episode type - starting with element 1 of an array with lower bound=0

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim R() As RFT_Episode
Dim Msg As String

On Error GoTo Get_RftEpisodeData_ForDate_Error

ReDim R(0)
    
'Retrieve Rft data
sql = "SELECT ResultsID, " & DbTables.RftResults & " AS SourceTblEnum FROM RftResults WHERE PatientID=" & PtID & " AND TestDate=" & gDBFunctions.Convert_ToSQLQueryDate(TestDate)
sql = sql & " UNION ALL "
sql = sql & "SELECT ResultsID, " & DbTables.ExResults & " AS SourceTblEnum FROM ExResults WHERE PatientID=" & PtID & " AND TestDate=" & gDBFunctions.Convert_ToSQLQueryDate(TestDate)
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Do While Not Rs.EOF
    ReDim Preserve R(UBound(R) + 1)
    R(UBound(R)) = Me.Get_RftEpisodeData_ForID(Rs("ResultsID"), Rs("SourceTblEnum"))
    Rs.MoveNext
Loop
Rs.Close

Get_RftEpisodeData_ForDate = R

Exit Function

Get_RftEpisodeData_ForDate_Error:
   Msg = ""
   Call ErrorLog(Msg, "Get_RftEpisodeData_ForDate", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Friend Function Save_RftEpisodeData(R As RFT_Episode, Mode As RecordMode) As Long
'Returns the recordID of the newly created or passed record, or zero if error

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
Dim RecordID As Long

On Error GoTo Save_RftEpisodeData_Error

Select Case Mode
    Case RecordMode.editRecord: RecordID = R.RecordID
    Case RecordMode.newRecord:  RecordID = 0
End Select
    
sql = "SELECT * FROM RftResults WHERE ResultsID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockOptimistic
If Rs.EOF Then
    If Mode = RecordMode.editRecord Then
        MsgBox "Problem saving RFT episode data, recordID not found - aborted"
        Save_RftEpisodeData = 0
        Exit Function
    Else
        Rs.AddNew
        Rs("PatientID") = R.PtID
    End If
Else
    If Rs("PatientID") <> R.PtID Then
        'should never happen
        MsgBox "Problem saving RFT episode data, PatientID mismatch - aborted"
        Save_RftEpisodeData = 0
        Exit Function
    End If
End If
    
If IsUDTDate(R.RequestDate) Then Rs("RequestDate") = R.RequestDate Else Rs("RequestDate") = Null
Rs("MO") = R.RequestingMO
Rs("RequestingMO_Pn") = R.RequestingMO_Pn
Rs("Request_HealthServiceID") = R.Ref_HealthServiceID
Rs("RefLocation") = R.RefLocation
Rs("Report_CopyTo") = R.ReportCopyTo
If IsUDTDate(R.TestDate) Then Rs("TestDate") = R.TestDate Else Rs("TestDate") = Null
If IsUDTDate(R.TestTime) Then Rs("TestTime") = R.TestTime Else Rs("TestTime") = Null
Rs("Smoke") = R.Smoke
Rs("Packyrs") = R.Packyrs
Rs("Weight") = R.ThisVisit_weight
Rs("Height_rft") = R.ThisVisit_height
Rs("Campus") = R.Lab
Rs("AdmissionStatus") = R.AdmissionStatus
Rs("ClinicalNote") = R.ClinicalNote
Rs("BillingMO") = R.BillingMO
Rs("BillingItems") = R.BillingItems
Rs("BillingStatus") = R.BillingStatus
Rs("BillingMO_ProviderNumber") = R.BillingMO_ProviderNumber
Rs.Update
Rs.MoveFirst
RecordID = Rs("ResultsID")
Rs.Close

Save_RftEpisodeData = RecordID

Exit Function

Save_RftEpisodeData_Error:
   Msg = ""
   Call ErrorLog(Msg, "Save_RftEpisodeData", "cPatient", Err.Number, Err.Description)
   Save_RftEpisodeData = 0
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Private Function DoNull(S As String)

If S = "" Then DoNull = Empty Else DoNull = S

End Function
Friend Function Decode_ErgoCPXData(S As String, CPETSystem As String) As CpxData

Dim FirstWord As String
Dim EndOfFirstWord As Integer
Dim CountOfBorders As Integer
Dim Ex As CpxData
Dim IsLineStartsWithEvery As Boolean
Dim IsBlankLine As Boolean
Dim LineOfDataType As sType
Dim Msg As String
Dim Marker As String

If Len(S) = 0 Then
    Ex.LineOfDataType = Header
    'Decode_CpxData = Ex
    'Exit Function
End If

'Get first word
IsBlankLine = Len(Replace(S, " ", "")) = 0

If InStr(S, "¦") Then
    'We are in the body of test data
    'Count number of spaces or "¦" characters to the first word and get the first word
    CountOfBorders = 2
    Do While Mid(S, CountOfBorders, 1) = "¦" Or CountOfBorders = Len(S)
        CountOfBorders = CountOfBorders + 1
    Loop
    If CPETSystem = "Ergo" Then Marker = "¦" Else Marker = " "
    EndOfFirstWord = InStr(CountOfBorders, S, "¦") - 1
    
    FirstWord = Mid(S, CountOfBorders, EndOfFirstWord - (CountOfBorders))    '+ 1))

    If IsNumeric(Left$(FirstWord, 6)) Then
        Ex.LineOfDataType = Exercise
        Ex = Decode_CpxDataGetItems(S, Ex, CPETSystem)
    Else
        Ex.LineOfDataType = Header
    End If

    
Else
    Ex.LineOfDataType = Header
End If

Decode_ErgoCPXData = Ex

'Exit Function

    
End Function

Friend Function Decode_CpxData(S As String, CPETSystem As String) As CpxData
'Decodes a single line of text from the source file

'File format to be decoded as follows (data every 30 sec, warmup period is 'rest'
    'SensorMedics Corporation
    'Yorba Linda, California
    '
    '
    '
    'Date: 03/01/12
    'Metabolic Edit - --CLARKE, Janice - 647557
    '    WorkVE(BTPS)     VO2    VCO2      HR    SpO2      Vt   PetO2  PetCO2      RQ
    '   Watts   L/min   L/min   L/min     BPM       %  Liters    mmHg    mmHg
    'Test Stage - Warmup
    '            11.9   0.197   0.211      74     100   0.734   118.9    29.6    1.08
    '             9.4   0.179   0.170      74     100   0.679   111.8    32.8    0.97
    '            10.7   0.238   0.218      75     100   0.759   110.9    32.6    0.92
    '            12.1   0.214   0.207      73     100   0.550   114.3    31.1    1.00
    'Test Stage - Exercise
    '      10    15.8   0.357   0.318      79      99   0.693   111.0    32.2    0.90
    '      10    19.6   0.434   0.436      84     100   0.908   114.7    30.8    1.00
    '      16    20.5   0.474   0.449      83     100   0.723   113.5    31.0    0.96
    
    '''File format for Medisoft ErgoCard added November 2012 is different - we can detect this system
    '''by looking for ??? in the first line of the text file;
    '''
    ''' ???
    '''
    '''
    '''Last:        BRAZZALE Sex: M
    '''Firstname   : Danny                      Height      : 174.0 cm
    '''MRN     : 650002                         Weight      :  63.0 kg
    '''D.O.B.    : 26/03/1973
    '''Referring :                              FEV         :  3.78 L
    '''Smoker     :                             BTPS fact.  :  1.11
    '''Operator    :                            Humidity    :  42.0 %
    '''Metabolic Exercise of 2/05/2012 to 1:38: - Protocol : AustBioQC4080120    Page : 1
    '''
    '''¦Load   ¦VE     ¦VO2    ¦VCO2   ¦HR     ¦SpO2   ¦TV     ¦PetO2  ¦PetCO2 ¦RER    ¦
    '''¦Watt   ¦L/min  ¦L/min  ¦L/min  ¦#/min  ¦%      ¦L      ¦mm Hg  ¦mm Hg  ¦       ¦
    '''¦ -------¦ - ------¦ - ------¦ - ------¦ - ------¦ - ------¦ - ------¦ - ------¦ - ------¦ - ------¦
    '''Every 30sec
    '''¦     0 ¦  13.0 ¦  0.40 ¦  0.37 ¦    76 ¦    98 ¦  0.59 ¦   109 ¦    38 ¦  0.91 ¦
    '''¦     0 ¦  13.3 ¦  0.38 ¦  0.35 ¦    74 ¦    98 ¦  0.63 ¦   110 ¦    37 ¦  0.93 ¦
    '''¦    40 ¦  25.2 ¦  0.83 ¦  0.76 ¦    89 ¦    95 ¦  0.95 ¦   109 ¦    39 ¦  0.92 ¦
    '''¦    40 ¦  21.4 ¦  0.74 ¦  0.66 ¦    91 ¦    96 ¦  1.07 ¦   106 ¦    41 ¦  0.90 ¦


Dim FirstWord As String
Dim EndOfFirstWord As Integer
Dim CountOfSpaces As Integer
Dim Ex As CpxData
Dim IsLineStartsWithChar As Boolean
Dim IsBlankLine As Boolean
Dim LineOfDataType As sType
Dim Msg As String

On Error GoTo Decode_CpxData_Error
    
If Len(S) = 0 Then
    Ex.LineOfDataType = Header
    Decode_CpxData = Ex
    Exit Function
End If

'Need to identify 3 string types -
    '1. header strings - first word is text or blank line
    '2. warmup data strings (rest) - first word is a number preceeded by 13 spaces
    '3. exercise data strings - first word is a number preceeded by 7 spaces
IsLineStartsWithChar = Left(S, 1) <> " "
IsBlankLine = Len(Replace(S, " ", "")) = 0

If IsLineStartsWithChar Or IsBlankLine Then
    Ex.LineOfDataType = Header
Else
    'Count number of spaces to the first word and get the first word
    CountOfSpaces = 1
    Do While Mid(S, CountOfSpaces, 1) = " " Or CountOfSpaces = Len(S)
        CountOfSpaces = CountOfSpaces + 1
    Loop
    EndOfFirstWord = InStr(CountOfSpaces, S, " ") - 1
    FirstWord = Mid(S, CountOfSpaces, EndOfFirstWord - CountOfSpaces + 1)
    Select Case IsNumeric(FirstWord)
        Case True   'Its a line of data, decode
            Select Case CountOfSpaces
                Case Is < 8: Ex.LineOfDataType = Exercise
                Case Is > 7: Ex.LineOfDataType = Warmup
            End Select
            Ex = Decode_CpxDataGetItems(S, Ex, CPETSystem)
        Case False  'Its a line of text, ignore
            Ex.LineOfDataType = Header

    End Select
End If

Decode_CpxData = Ex

Exit Function


Decode_CpxData_Error:
    Msg = ""
   Call ErrorLog(Msg, "Decode_CpxData", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Private Function Decode_CpxDataGetItems(S As String, X As CpxData, CPETSystem) As CpxData
    'Example data
    'Cols delimited by padding with spaces
    'Ignore load col in exercise period
    'Test Stage - Warmup
    '            11.9   0.197   0.211      74     100   0.734   118.9    29.6    1.08
    '             9.4   0.179   0.170      74     100   0.679   111.8    32.8    0.97
    '            10.7   0.238   0.218      75     100   0.759   110.9    32.6    0.92
    '            12.1   0.214   0.207      73     100   0.550   114.3    31.1    1.00
    'Test Stage - Exercise
    '      10    15.8   0.357   0.318      79      99   0.693   111.0    32.2    0.90
    '      10    19.6   0.434   0.436      84     100   0.908   114.7    30.8    1.00
    '      16    20.5   0.474   0.449      83     100   0.723   113.5    31.0    0.96

Dim Start As Integer
Dim Finish As Integer
Dim i As Integer, j As Integer, z As Integer
Dim Items(0 To 10) As Single
Dim ItemCount As Integer
Dim Msg As String
Dim Marker As String

Select Case CPETSystem
    Case "Ergo"
        Marker = "¦"
    Case "VMAX"
        Marker = " "
End Select


On Error GoTo Decode_CpxDataGetItems_Error

Select Case X.LineOfDataType
    Case sType.Exercise:  ItemCount = -1
    Case sType.Warmup:    ItemCount = 0
End Select

i = 0: Start = 0: Finish = 0
Do
    'Skip spaces till char found
    Do
        i = i + 1
    Loop While Mid(S, i, 1) = Marker
    Start = i
    'Find end of item
    Do While Mid(S, i, 1) <> Marker
        i = i + 1
        If i = Len(S) + 1 Then Exit Do
    Loop
    Finish = i
    ItemCount = ItemCount + 1
    Items(ItemCount) = CSng(Mid(S, Start, Finish - Start))
    i = Finish
    
Loop Until i >= Len(S) - 1 '+ 1

X.VE = Items(1)
X.VO2 = Items(2)
X.VCO2 = Items(3)
X.HR = Items(4)
X.SpO2 = Items(5)
X.Vt = Items(6)
X.PetO2 = Items(7)
X.PetCO2 = Items(8)
X.RQ = Items(9)

Decode_CpxDataGetItems = X

Exit Function


Decode_CpxDataGetItems_Error:
    Msg = ""
   Call ErrorLog(Msg, "Decode_CpxDataGetItems", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Friend Function Save_RftToDB(R As RFTs, Mode As RecordMode) As Long
'Returns the recordID of the new or edited record

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
   
On Error GoTo Save_RftToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM RFTResults WHERE ResultsID =" & R.RecordID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving RFT results"
            Exit Function
        Else
        
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM RFTResults WHERE ResultsID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = R.PtID
End Select
    
Rs("Campus") = R.E_Lab
Rs("MO") = R.E_RequestingMO
Rs("RequestingMO_Pn") = R.E_RequestingMO_Pn
Rs("RefLocation") = R.E_RefLocation & ""
Rs("Request_HealthServiceID") = R.E_Ref_HealthServiceID
Rs("Report_CopyTo") = R.E_ReportCopyTo & ""
Rs("TestType") = R.TestType
Rs("TestDate") = R.E_TestDate
Rs("RequestDate") = R.E_RequestDate
Rs("TestTime") = R.E_TestTime
Rs("scientist") = R.Scientist
Rs("comments") = R.TechComments & ""
Rs("Smoke") = R.E_Smoke & ""
If R.E_Packyrs = "" Then Rs("Packyrs") = Null Else Rs("Packyrs") = R.E_Packyrs
Rs("Weight") = Val(R.E_ThisVisit_weight)
Rs("Height_rft") = Val(R.E_ThisVisit_height)
Rs("Hb_WaitingForResult") = R.Hb_WaitingForResult
Rs("ClinicalNote") = R.E_ClinicalNote
'Rs("Source") = R.Source
Rs("AdmissionStatus") = R.E_AdmissionStatus
Rs("BillingMO") = R.E_BillingMO
Rs("BillingItems") = R.E_BillingItems
Rs("BillingStatus") = R.E_BillingStatus
Rs("BillingMO_ProviderNumber") = R.E_BillingMO_ProviderNumber

Rs("spir") = R.spir
Rs("bd1") = R.BD_1
Rs("bd2") = R.BD_2
Rs("LastBd") = R.LastBd

Rs("prefev1") = DoNull(R.preFEV1)
Rs("prefvc") = DoNull(R.preFVC)
Rs("prevc") = DoNull(R.preVC)
Rs("post1fev1") = DoNull(R.postFEV1_1)
Rs("post1fvc") = DoNull(R.postFVC_1)
Rs("post1vc") = DoNull(R.postVC_1)
Rs("post2fev1") = DoNull(R.postFEV1_2)
Rs("post2fvc") = DoNull(R.postFVC_2)
Rs("post2vc") = DoNull(R.postVC_2)
Rs("preFEF") = DoNull(R.preFEF)
Rs("post1FEF") = DoNull(R.postFEF_1)
Rs("post2FEF") = DoNull(R.postFEF_2)

Rs("tlcoEquip") = R.tlcoEquip
Rs("va") = DoNull(R.VA)
Rs("tlco") = DoNull(R.TLCO)
Rs("kco") = DoNull(R.KCO)
Rs("Hb") = DoNull(R.Hb)
Rs("HbInfo") = R.HbInfo

Rs("volEquip") = R.volEquip
Rs("frc") = DoNull(R.FRC)
Rs("tlc") = DoNull(R.TLC)
Rs("vc") = DoNull(R.VC)
Rs("rv") = DoNull(R.RV)
Rs("raw") = DoNull(R.Raw)
Rs("sgaw") = DoNull(R.Sgaw)
Rs("raw2") = DoNull(R.Raw_2)
Rs("sgaw2") = DoNull(R.Sgaw_2)

Rs("fio21") = DoNull(R.FiO2_1)
Rs("ph1") = DoNull(R.pH_1)
Rs("pco21") = DoNull(R.pCO2_1)
Rs("po21") = DoNull(R.pO2_1)
Rs("shunt1") = DoNull(R.Shunt_1)
Rs("BE1") = DoNull(R.BE_1)
Rs("HCO31") = DoNull(R.HCO3_1)
Rs("SaO21") = DoNull(R.SaO2_1)
Rs("fio22") = DoNull(R.FiO2_2)
Rs("ph2") = DoNull(R.pH_2)
Rs("pco22") = DoNull(R.pCO2_2)
Rs("po22") = DoNull(R.pO2_2)
Rs("shunt2") = DoNull(R.Shunt_2)
Rs("BE2") = DoNull(R.BE_2)
Rs("HCO32") = DoNull(R.HCO3_2)
Rs("SaO22") = DoNull(R.SaO2_2)

Rs("mips") = DoNull(R.MIP)
Rs("meps") = DoNull(R.MEP)
Rs("SNIP") = DoNull(R.SNIP)

Rs("SkinDataString") = R.SkinTestDataString
Rs("Skin_PanelDescription") = R.Skin_PanelDescription

Rs("ASTString") = R.AST_TestDataString & ""

Rs("AirSaO2String") = R.AirSaO2String & ""
Rs("O2SaO2String") = R.O2SaO2String & ""
Rs("SpeedString") = R.SpeedString & ""


Rs("ChallengeType") = R.prov_ChallengeType
Rs("control") = R.prov_Control
Rs("Dose1") = R.prov_Dose1
Rs("Dose2") = R.prov_Dose2
Rs("Dose3") = R.prov_Dose3
Rs("Dose4") = R.prov_Dose4
Rs("Dose5") = R.prov_Dose5
Rs("Dose6") = R.prov_Dose6
Rs("Dose7") = R.prov_Dose7
Rs("Dose8") = R.prov_Dose8
Rs("Dose9") = R.prov_Dose9
Rs("post") = R.prov_Post
Rs("PostProvBDType") = R.prov_PostBDType
Rs("FEV1-Control") = DoNull(R.prov_FEV1_Control)
Rs("FEV1-1") = DoNull(R.prov_FEV1_Dose1)
Rs("FEV1-2") = DoNull(R.prov_FEV1_Dose2)
Rs("FEV1-3") = DoNull(R.prov_FEV1_Dose3)
Rs("FEV1-4") = DoNull(R.prov_FEV1_Dose4)
Rs("FEV1-5") = DoNull(R.prov_FEV1_Dose5)
Rs("FEV1-6") = DoNull(R.prov_FEV1_Dose6)
Rs("FEV1-7") = DoNull(R.prov_FEV1_Dose7)
Rs("FEV1-8") = DoNull(R.prov_FEV1_Dose8)
Rs("FEV1-9") = DoNull(R.prov_FEV1_Dose9)
Rs("FEV1-post") = DoNull(R.prov_FEV1_Post)

Rs("RangeFEV1") = DoNull(R.pnr_FEV1)
Rs("RangeFVC") = DoNull(R.pnr_FVC)
Rs("RangeVC") = DoNull(R.pnr_VC)
Rs("RangeFEF") = DoNull(R.pnr_FEF)
Rs("RangeFER") = DoNull(R.pnr_FER)
Rs("RangeTLCO") = DoNull(R.pnr_TLCO)
Rs("RangeVA") = DoNull(R.pnr_VA)
Rs("RangeKCO") = DoNull(R.pnr_KCO)
Rs("RangeMIP") = DoNull(R.pnr_MIP)
Rs("RangeMEP") = DoNull(R.pnr_MEP)
Rs("RangePH") = DoNull(R.pnr_pH)
Rs("RangePaCO2") = DoNull(R.pnr_PaCO2)
Rs("RangePaO2") = DoNull(R.pnr_PaO2)
Rs("RangeShunt") = DoNull(R.pnr_Shunt)
Rs("RangeRAW") = DoNull(R.pnr_Raw)
Rs("RangeSgAW") = DoNull(R.pnr_Sgaw)
Rs("RangePEF") = DoNull(R.pnr_PEF)
Rs("RangeSaO2") = DoNull(R.pnr_SaO2)
Rs("RangeHCO3") = DoNull(R.pnr_HCO3)
Rs("RangeRVTLC") = DoNull(R.pnr_RVTLC)
Rs("RangeBE") = DoNull(R.pnr_BE)
Rs("RangeFRC") = DoNull(R.pnr_FRC)
Rs("RangeTLC") = DoNull(R.pnr_TLC)
Rs("RangeRV") = DoNull(R.pnr_RV)
Rs("RangeSNIP") = DoNull(R.pnr_SNIP)

Rs("preFEV1%") = DoNull(R.ppn_preFEV1)
Rs("preFVC%") = DoNull(R.ppn_preFVC)
Rs("preVC%") = DoNull(R.ppn_preVC)
Rs("preFEF%") = DoNull(R.ppn_preFEF)
Rs("prePEF%") = DoNull(R.ppn_prePEF)
Rs("post1FEV1%") = DoNull(R.ppn_post1FEV1)
Rs("post1FVC%") = DoNull(R.ppn_post1FVC)
Rs("post1VC%") = DoNull(R.ppn_post1VC)
Rs("post1FEF%") = DoNull(R.ppn_post1FEF)
Rs("post1PEF%") = DoNull(R.ppn_post1PEF)
Rs("post2FEV1%") = DoNull(R.ppn_post2FEV1)
Rs("post2FVC%") = DoNull(R.ppn_post2FVC)
Rs("post2VC%") = DoNull(R.ppn_post2VC)
Rs("post2FEF%") = DoNull(R.ppn_post2FEF)
Rs("post2PEF%") = DoNull(R.ppn_post2PEF)
Rs("TLCO%") = DoNull(R.ppn_TLCO)
Rs("TLCOhb%") = DoNull(R.ppn_TLCOhb)
Rs("KCO%") = DoNull(R.ppn_KCO)
Rs("KCOhb%") = DoNull(R.ppn_KCOhb)
Rs("VA%") = DoNull(R.ppn_VA)
Rs("FRC%") = DoNull(R.ppn_FRC)
Rs("RV%") = DoNull(R.ppn_RV)
Rs("TLC%") = DoNull(R.ppn_TLC)
Rs("MIP%") = DoNull(R.ppn_MIP)
Rs("MEP%") = DoNull(R.ppn_MEP)
Rs("SNIP%") = DoNull(R.ppn_SNIP)

Rs("Report") = R.Report
Rs("Reporter") = R.Reporter
Rs("ReportDate") = R.ReportDate
Rs("VerifiedDate") = R.VerifiedDate
Rs("ReportStatus") = R.ReportStatus
Rs("VerifiedBy") = R.VerifiedBy
Rs("LastEdit") = Now

Rs.Update
Rs.MoveFirst
Save_RftToDB = Rs("ResultsID")
Rs.Close

Exit Function


Save_RftToDB_Error:
    Msg = ""
   Call ErrorLog(Msg, "Save_RftToDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume

End Function

Friend Function Save_SlpToDB(S As SlpResults, Mode As RecordMode) As Long
'Returns the recordID of the new or edited record

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
   
On Error GoTo Save_SlpToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM SlpResults WHERE StudyID =" & S.RecordID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving sleep study results"
            Exit Function
        Else
        
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM SlpResults WHERE StudyID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = S.PatientID
End Select

'Study data (request)
Rs("Campus") = S.Lab
Rs("Studytype") = S.StudyType
Rs("Studydate") = S.StudyDate
If IsDate(S.Request_Date) Then Rs("RequestDate") = S.Request_Date Else Rs("RequestDate") = Null
If IsDate(S.Request_Time) Then Rs("RequestTime") = S.Request_Time Else Rs("RequestTime") = Null
Rs("StudyBy") = S.StudyBy
Rs("ReasonForTest") = S.ReasonForTest
Rs("Admission") = S.Admission
Rs("TreatmentMode") = S.TreatmentMode
Rs("RequestedLocation") = S.RequestedLocation
Rs("Note") = S.Note & ""
Rs("ApptNotes") = S.ApptNotes & ""
Rs("Urgent_scoring") = S.Urgent_scoring
Rs("Urgent_PSG") = S.Urgent_PSG

'Load special procedure info
Rs("Special_TCO2") = S.Special_TCO2
Rs("Special_EMGd") = S.Special_EMGd
Rs("Special_Poes") = S.Special_Poes
Rs("Special_Temp") = S.Special_Temp
Rs("Special_Other") = S.Special_OtherProcedure
Rs("TreatmentOxygen") = S.TreatmentOxygen
Rs("Request_Acknowledged") = S.Request_Acknowledged
Rs("Request_HardCopyRec") = S.Request_HardCopyRec
Rs("Need_NursingCare") = S.Need_NursingCare
Rs("Need_Mobility") = S.Need_Mobility
Rs("Need_Mobility_Type") = S.Need_Mobility_Type
Rs("Need_Interpreter") = S.Need_Interpreter
Rs("Need_Interpreter_Booked") = S.Need_Interpreter_Booked
Rs("Need_Interpreter_Language") = S.Need_Interpreter_Language
Rs("ExistingDiseases") = S.ExistingDiseases
Rs("VRSSPatient") = S.VRSSPatient
Rs("FeeType") = S.FeeType
Rs("HccPensionNumber") = S.HccPensionNumber
Rs("ApproximateWeight") = S.ApproximateWeight

'RequestingMO
Rs("RequestingMO") = S.RequestingMO
Rs("RequestingMO_ProviderNum") = S.RequestingMO_ProviderNum
Rs("ReportDest_2") = S.ReportDest_2 & ""
If IsDate(S.RequestingMO_ReviewDate) Then Rs("RequestingMO_ReviewDate") = S.RequestingMO_ReviewDate Else Rs("RequestingMO_ReviewDate") = Null
Rs("RequestingMO_ReviewLocation") = S.RequestingMO_ReviewLocation & ""
Rs("ReportDest_1") = S.ReportDest_1 & ""
Rs("ReportDest_2") = S.ReportDest_2 & ""

'ReferringMO
Rs("Ref_MO") = S.Ref_MO
Rs("Ref_Address") = S.Ref_Address & ""
If IsDate(S.Ref_Date) Then Rs("Ref_Date") = S.Ref_Date Else Rs("Ref_Date") = Null
'If IsDate(s.Ref_ActivateDate) Then Rs("Ref_ActivationDate") = s.Ref_ActivateDate Else Rs("Ref_ActivationDate") = Null
Rs("Ref_Duration") = S.Ref_Duration
Rs("Ref_ProviderNum") = S.Ref_ProviderNum

'Study data (results)
Rs("Weight") = S.Weight
Rs("Height_slp") = S.Height_slp
Rs("R_Bed") = S.Bed
Rs("OpticalDisk") = S.OpticalDisk
Rs("R_TcCO2_Unit") = S.R_ABG_PtcCO2_Unit
Rs("R_TcCO2_Electrode") = S.R_ABG_PtcCO2_Electrode
Rs("PortableDevice") = S.PortableDevice

'Request status tab (request)
Rs("RequestStatus") = S.Request_Status & ""
Rs("Request_HardCopyRec") = S.Request_HardCopyRec
Rs("RequestMode") = S.RequestMode & ""
Rs("Request_Acknowledged") = S.Request_Acknowledged
Rs("Request_Acknowledged_By") = S.Request_Acknowledged_By

'Reporting
Rs("ReportStatus") = S.ReportStatus
Rs("ScoredBy") = S.ScoredBy
If IsDate(S.ScoredDate) Then Rs("ScoredDate") = S.ScoredDate Else Rs("ScoredDate") = Null
Rs("Reporter_MO") = S.Reporter_MO
If IsDate(S.ReportDate) Then Rs("ReportDate") = S.ReportDate Else Rs("ReportDate") = Null
Rs("VerifiedBy") = S.VerifiedBy
If IsDate(S.VerifiedDate) Then Rs("VerifiedDate") = S.VerifiedDate Else Rs("VerifiedDate") = Null
Rs("TecnicalNote") = S.TechnicalNote
Rs("Report") = S.Report

'Billing data (request)
Rs("BillingMO") = S.BillingMO
Rs("BillingMO_ProviderNumber") = S.BillingMO_ProviderNumber
Rs("BillingItems") = S.BillingItems
Rs("BillingStatus") = S.BillingStatus
If IsDate(S.Payment_ReceivedDate_Medicare) Then Rs("Payment_ReceivedDate_Medicare") = S.Payment_ReceivedDate_Medicare Else Rs("Payment_ReceivedDate_Medicare") = Null
If IsDate(S.Payment_ReceivedDate_Patient) Then Rs("Payment_ReceivedDate_Patient") = S.Payment_ReceivedDate_Patient Else Rs("Payment_ReceivedDate_Patient") = Null
Rs("Payment_AccountNote") = S.Payment_AccountNote & ""
If IsDate(S.Payment_AccountReminderDate) Then Rs("Payment_AccountReminderDate") = S.Payment_AccountReminderDate Else Rs("Payment_AccountReminderDate") = Null
Rs("Payment_Location") = S.Payment_Location
Rs("Payment_Method") = S.Payment_Method
Rs("Payment_OwingAmount_Medicare") = S.Payment_OwingAmount_Medicare
Rs("Payment_OwingAmount_Patient") = S.Payment_OwingAmount_Patient
Rs("Payment_ReceiptIssued") = S.Payment_ReceiptIssued
Rs("Payment_ReceivedAmount_Medicare") = S.Payment_ReceivedAmount_Medicare
Rs("Payment_ReceivedAmount_Patient") = S.Payment_ReceivedAmount_Patient
Rs("Payment_TransactedBy_Medicare") = S.Payment_TransactedBy_Medicare
Rs("Payment_TransactedBy_Patient") = S.Payment_TransactedBy_Patient

'Results
If IsDate(S.Est_LOff) Then Rs("Est_LO") = S.Est_LOff Else Rs("Est_LO") = Null
If IsDate(S.Est_WakeUpTime) Then Rs("Est_WakeUpTime") = S.Est_WakeUpTime Else Rs("Est_WakeUpTime") = Null
Rs("Est_TimeToSleep") = S.Est_TimeToSleep
Rs("Est_SleepDuration") = S.Est_SleepDuration
Rs("R_ESS") = S.ESS
If IsDate(S.R_Recording_Start) Then Rs("R_Recording_Start") = S.R_Recording_Start Else Rs("R_Recording_Start") = Null
If IsDate(S.R_Recording_End) Then Rs("R_Recording_End") = S.R_Recording_End Else Rs("R_Recording_End") = Null
If IsDate(S.R_Report_Start) Then Rs("R_Report_Start") = S.R_Report_Start Else Rs("R_Report_Start") = Null
If IsDate(S.R_Report_End) Then Rs("R_Report_End") = S.R_Report_End Else Rs("R_Report_End") = Null
If IsDate(S.R_LOut) Then Rs("R_LOut") = S.R_LOut Else Rs("R_LOut") = Null
If IsDate(S.R_LOn) Then Rs("R_LOn") = S.R_LOn Else Rs("R_LOn") = Null

Rs("R_AI_RespNrem") = S.R_ArI_RespNrem
Rs("R_AI_RespRem") = S.R_ArI_RespRem
Rs("R_AI_RespAll") = S.R_ArI_RespAll
Rs("R_AI_PlmNrem") = S.R_ArI_PlmNrem
Rs("R_AI_PlmRem") = S.R_ArI_PlmRem
Rs("R_AI_PlmAll") = S.R_ArI_PlmAll
Rs("R_AI_OtherNrem") = S.R_ArI_OtherNrem
Rs("R_AI_OtherRem") = S.R_ArI_OtherRem
Rs("R_AI_OtherAll") = S.R_ArI_OtherAll
Rs("R_AI_AllNrem") = S.R_ArI_AllNrem
Rs("R_AI_AllRem") = S.R_ArI_AllRem
Rs("R_AI_AllAll") = S.R_ArI_AllAll
Rs("R_AI_SupRem") = S.R_ArI_SupRem
Rs("R_AI_SupNrem") = S.R_ArI_SupNrem
Rs("R_AI_SupAll") = S.R_ArI_SupAll
Rs("R_AI_NonSupRem") = S.R_ArI_NonSupRem
Rs("R_AI_NonSupNrem") = S.R_ArI_NonSupNrem
Rs("R_AI_NonSupAll") = S.R_ArI_NonSupAll
Rs("R_AHI_SupNrem") = S.R_AHI_SupNrem
Rs("R_AHI_SupRem") = S.R_AHI_SupRem
Rs("R_AHI_SupAll") = S.R_AHI_SupAll
Rs("R_AHI_NonSupNrem") = S.R_AHI_NonSupNrem
Rs("R_AHI_NonSupRem") = S.R_AHI_NonSupRem
Rs("R_AHI_NonSupAll") = S.R_AHI_NonSupAll
Rs("R_AHI_AllNrem") = S.R_AHI_AllNrem
Rs("R_AHI_AllRem") = S.R_AHI_AllRem
Rs("R_AHI_AllAll") = S.R_AHI_AllAll
Rs("R_HI_SupNrem") = S.R_HI_SupNrem
Rs("R_HI_SupRem") = S.R_HI_SupRem
Rs("R_HI_SupAll") = S.R_HI_SupAll
Rs("R_HI_NonSupNrem") = S.R_HI_NonSupNrem
Rs("R_HI_NonSupRem") = S.R_HI_NonSupRem
Rs("R_HI_NonSupAll") = S.R_HI_NonSupAll
Rs("R_HI_AllNrem") = S.R_HI_AllNrem
Rs("R_HI_AllRem") = S.R_HI_AllRem
Rs("R_HI_AllAll") = S.R_HI_AllAll
Rs("R_OAI_SupNrem") = S.R_OAI_SupNrem
Rs("R_OAI_SupRem") = S.R_OAI_SupRem
Rs("R_OAI_SupAll") = S.R_OAI_SupAll
Rs("R_OAI_NonSupNrem") = S.R_OAI_NonSupNrem
Rs("R_OAI_NonSupRem") = S.R_OAI_NonSupRem
Rs("R_OAI_NonSupAll") = S.R_OAI_NonSupAll
Rs("R_OAI_AllNrem") = S.R_OAI_AllNrem
Rs("R_OAI_AllRem") = S.R_OAI_AllRem
Rs("R_OAI_AllAll") = S.R_OAI_AllAll
Rs("R_CAI_SupNrem") = S.R_CAI_SupNrem
Rs("R_CAI_SupRem") = S.R_CAI_SupRem
Rs("R_CAI_SupAll") = S.R_CAI_SupAll
Rs("R_CAI_NonSupNrem") = S.R_CAI_NonSupNrem
Rs("R_CAI_NonSupRem") = S.R_CAI_NonSupRem
Rs("R_CAI_NonSupAll") = S.R_CAI_NonSupAll
Rs("R_CAI_AllNrem") = S.R_CAI_AllNrem
Rs("R_CAI_AllRem") = S.R_CAI_AllRem
Rs("R_CAI_AllAll") = S.R_CAI_AllAll
Rs("R_MAI_SupNrem") = S.R_MAI_SupNrem
Rs("R_MAI_SupRem") = S.R_MAI_SupRem
Rs("R_MAI_SupAll") = S.R_MAI_SupAll
Rs("R_MAI_NonSupNrem") = S.R_MAI_NonSupNrem
Rs("R_MAI_NonSupRem") = S.R_MAI_NonSupRem
Rs("R_MAI_NonSupAll") = S.R_MAI_NonSupAll
Rs("R_MAI_AllNrem") = S.R_MAI_AllNrem
Rs("R_MAI_AllRem") = S.R_MAI_AllRem
Rs("R_MAI_AllAll") = S.R_MAI_AllAll
Rs("R_PLM_Nrem") = S.R_PLM_Nrem
Rs("R_PLM_Rem") = S.R_PLM_Rem
Rs("R_PLM_All") = S.R_PLM_All
Rs("R_PLM_Wake") = S.R_PLM_Wake
Rs("R_PLM_Ep_All") = S.R_PLM_Ep_All
Rs("R_PLM_Ep_NREM") = S.R_PLM_Ep_NREM
Rs("R_PLM_Ep_REM") = S.R_PLM_Ep_REM
Rs("R_PLM_Ep_Wake") = S.R_PLM_Ep_Wake
Rs("R_NREM1") = S.R_NREM1
Rs("R_NREM2") = S.R_NREM2
Rs("R_NREM3") = S.R_NREM3
Rs("R_NREM4") = S.R_NREM4
Rs("R_NREM12") = S.R_NREM12
Rs("R_NREM34") = S.R_NREM34
Rs("R_NREMTotal") = S.R_NREMTotal
Rs("R_REMTotal") = S.R_REMTotal
Rs("R_TST") = S.R_TST
Rs("R_TST_SupineNREM") = S.R_TST_SupineNREM
Rs("R_TST_NonSupineNREM") = S.R_TST_NonSupineNREM
Rs("R_TST_SupineREM") = S.R_TST_SupineREM
Rs("R_TST_NonSupineREM") = S.R_TST_NonSupineREM
Rs("R_TST_SupineTotal") = S.R_TST_SupineTotal
Rs("R_TST_NonSupineTotal") = S.R_TST_NonSupineTotal
Rs("R_SLat") = S.R_SLat
Rs("R_REMLat") = S.R_REMLat
Rs("R_TDT") = S.R_TDT
Rs("R_SlpEff") = S.R_SlpEff
Rs("R_Awakenings") = S.R_Awakenings
Rs("R_WakeTimeDuringSleep") = S.R_WakeTimeDuringSleep
Rs("R_SaO2_BLAwake") = S.R_SpO2_BLAwake
Rs("R_SaO2_BLAsleep") = S.R_SpO2_BLAsleep
Rs("R_SaO2_MinREM") = S.R_SpO2_MinREM
Rs("R_SaO2_MinNREM") = S.R_SpO2_MinNREM
Rs("R_SaO2_Time95") = S.R_SpO2_Time95
Rs("R_SaO2_Time90") = S.R_SpO2_Time90
Rs("R_SaO2_Time88") = S.R_SpO2_Time88
Rs("R_SaO2_Time85") = S.R_SpO2_Time85
Rs("R_ABG_Conditions1") = S.R_ABG_Conditions1
Rs("R_ABG_Conditions2") = S.R_ABG_Conditions2
Rs("R_ABG_pH1") = S.R_ABG_pH1
Rs("R_ABG_pH2") = S.R_ABG_pH2
Rs("R_ABG_PaO21") = S.R_ABG_PaO21
Rs("R_ABG_PaO22") = S.R_ABG_PaO22
Rs("R_ABG_PaCO21") = S.R_ABG_PaCO21
Rs("R_ABG_PaCO22") = S.R_ABG_PaCO22
Rs("R_ABG_HCO31") = S.R_ABG_HCO31
Rs("R_ABG_HCO32") = S.R_ABG_HCO32
Rs("R_ABG_BE1") = S.R_ABG_BE1
Rs("R_ABG_BE2") = S.R_ABG_BE2
Rs("R_ABG_SaO21") = S.R_ABG_SaO21
Rs("R_ABG_SaO22") = S.R_ABG_SaO22
Rs("R_ABG_FiO21") = S.R_ABG_FiO21
Rs("R_ABG_FiO22") = S.R_ABG_FiO22
Rs("R_ABG_Time1") = S.R_ABG_Time1
Rs("R_ABG_Time2") = S.R_ABG_Time2
Rs("R_PtcCO21") = S.R_ABG_PtcCO21
Rs("R_PtcCO22") = S.R_ABG_PtcCO22
'Rs("R_ABG_PCO2Diff1")=s.R_ABG_PCO2Diff1
'Rs("R_ABG_PCO2Diff2")=s.R_ABG_PCO2Diff2
Rs("R_tcCO2_Unit") = S.R_ABG_PtcCO2_Unit
Rs("R_tcCO2_Electrode") = S.R_ABG_PtcCO2_Electrode
'Rs("R_ABG_PtcCO2_Drift")=s.R_ABG_PtcCO2_Drift
Rs("R_SpO21") = S.R_ABG_SpO21
Rs("R_SpO22") = S.R_ABG_SpO22
'Rs("R_ABG_SO2Diff1")=s.R_ABG_SO2Diff1
'Rs("R_ABG_SO2Diff2")=s.R_ABG_SO2Diff2
Rs("R_RERA_AllAll") = S.R_RERA_AllAll
Rs("R_RERA_SupAll") = S.R_RERA_SupAll
Rs("R_RERA_NonSupAll") = S.R_RERA_NonSupAll
Rs("R_RERA_AllRem") = S.R_RERA_AllRem
Rs("R_RERA_AllNrem") = S.R_RERA_AllNrem
Rs("R_RERA_SupRem") = S.R_RERA_SupRem
Rs("R_RERA_SupNrem") = S.R_RERA_SupNrem
Rs("R_RERA_NonSupRem") = S.R_RERA_NonSupRem
Rs("R_RERA_NonSupNrem") = S.R_RERA_NonSupNrem
Rs("R_RDI_AllAll") = S.R_RDI_AllAll
Rs("R_RDI_SupAll") = S.R_RDI_SupAll
Rs("R_RDI_NonSupAll") = S.R_RDI_NonSupAll
Rs("R_RDI_AllRem") = S.R_RDI_AllRem
Rs("R_RDI_AllNrem") = S.R_RDI_AllNrem
Rs("R_RDI_SupRem") = S.R_RDI_SupRem
Rs("R_RDI_SupNrem") = S.R_RDI_SupNrem
Rs("R_RDI_NonSupRem") = S.R_RDI_NonSupRem
Rs("R_RDI_NonSupNrem") = S.R_RDI_NonSupNrem
Rs("R_Desat4_AllAll") = S.R_ODI4_AllAll
Rs("R_Desat4_AllRem") = S.R_ODI4_AllRem
Rs("R_Desat4_AllNrem") = S.R_ODI4_AllNrem
Rs("R_Desat4_NonSupAll") = S.R_ODI4_NonSupAll
Rs("R_Desat4_NonSupRem") = S.R_ODI4_NonSupRem
Rs("R_Desat4_NonSupNrem") = S.R_ODI4_NonSupNrem
Rs("R_Desat4_SupAll") = S.R_ODI4_SupAll
Rs("R_Desat4_SupRem") = S.R_ODI4_SupRem
Rs("R_Desat4_SupNrem") = S.R_ODI4_SupNrem
Rs("LastEdit") = Now

'Overnight notes (log)
Rs("StudyLog") = S.StudyLog
Rs("StudyPerformedBy") = S.StudyPerformedBy
Rs("StudyLogNote") = S.StudyLogNote

Rs.Update
Rs.MoveFirst
Save_SlpToDB = Rs("StudyID")
Rs.Close

Exit Function


Save_SlpToDB_Error:
   Msg = ""
   Call ErrorLog(Msg, "Save_SlpToDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Delete_TblRecord(RecordID As Long, Tbl As DbTables) As Boolean

Dim sql As String
Dim Rs As New ADODB.Recordset
Dim Response As Integer
Dim Msg As String
Dim Msg1 As String
Dim TblName As String
Dim PrimaryKeyName As String

On Error GoTo Delete_TblRecord_Error

'Set display message
Select Case Tbl
    Case DbTables.SleepAppts
        Msg1 = "this sleep study booking."
        TblName = gDBs.ResMed_203s.Tbl_SleepAppts.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_SleepAppts.PrimaryKey
    Case DbTables.ExResults:
        Msg1 = "this exercise test."
        TblName = gDBs.ResMed_203s.Tbl_ExResults.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_ExResults.PrimaryKey
    Case DbTables.ProgramStatus:
        Msg1 = "this program status record."
        TblName = gDBs.ResMed_203s.Tbl_ProgramStatus.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_ProgramStatus.PrimaryKey
    Case DbTables.RftResults:
        Msg1 = "this test."
        TblName = gDBs.ResMed_203s.Tbl_RftResults.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_RftResults.PrimaryKey
    Case DbTables.SlpResults:
        Msg1 = "this sleep study."
        TblName = gDBs.ResMed_203s.Tbl_SlpResults.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_SlpResults.PrimaryKey
    Case DbTables.Visit_CPAP_ImpReview:
        Msg1 = "this CPAP record."
        TblName = gDBs.ResMed_203s.Tbl_Visit_CPAP_ImpReview.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_Visit_CPAP_ImpReview.PrimaryKey
    Case DbTables.VRSS_Episodes:
        Msg1 = "this VRSS episode record."
        TblName = gDBs.ResMed_203s.Tbl_VRSS_Episodes.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_VRSS_Episodes.PrimaryKey
    Case DbTables.VRSS_VentDependentStatus:
        Msg1 = "this ventilator dependency record."
        TblName = gDBs.ResMed_203s.Tbl_VRSS_VentDependentStatus.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_VRSS_VentDependentStatus.PrimaryKey
    Case DbTables.Visit_NonCPAP_ImpReview:
        Msg1 = "this VRSS review record."
        TblName = gDBs.ResMed_203s.Tbl_Visit_NonCPAP_ImpReview.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_Visit_NonCPAP_ImpReview.PrimaryKey
    Case DbTables.Consults:
        Msg1 = "consult record."
        TblName = gDBs.ResMed_203s.Tbl_Consults.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_Consults.PrimaryKey
    Case DbTables.Documents:
        Msg1 = "document record."
        TblName = gDBs.ResMed_203s.Tbl_Documents.Name
        PrimaryKeyName = gDBs.ResMed_203s.Tbl_Documents.PrimaryKey
End Select

'Display confirmation message
Msg = "Warning - you are about to delete " & Msg1 & vbCrLf & " Continue with delete?"
Response = MsgBox(Msg, vbOKCancel, "Delete record")
If Response = vbOK Then
    
    Call gReportToCerner.DeleteRecord(Tbl, RecordID)
    
    'Retrieve data from the database for the selected patient
    sql = "SELECT * FROM " & TblName & " WHERE " & PrimaryKeyName & " = " & RecordID
    Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
    If Rs.EOF Then
        MsgBox "Record # " & RecordID & " not found, delete aborted.", vbOKOnly, "Delete record"
        Rs.Close
        Delete_TblRecord = False
    Else
        Rs.Delete
        Rs.Close
        MsgBox "Record deleted.", vbOKOnly, "Delete record"
        Screen.MousePointer = 0
        Delete_TblRecord = True
    End If
End If

Screen.MousePointer = 0

Exit Function


Delete_TblRecord_Error:
    Msg = ""
    Call ErrorLog(Msg, "Delete_TblRecord", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function

Friend Function Save_CpapToDB(c As CPAPClinic, Mode As RecordMode) As Long
'Returns the recordID of the new or edited record

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
   
On Error GoTo Save_CpapToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM Visit_CPAP_ImpReview WHERE VisitID =" & c.VisitID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving CPAP Clinic data"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM Visit_CPAP_ImpReview WHERE VisitID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = c.PatientID
End Select

'Visit stuff
If IsDate(c.Visit_Date) Then Rs("VisitDate") = c.Visit_Date Else Rs("VisitDate") = Null
If IsDate(c.Visit_Time) Then Rs("VisitTime") = c.Visit_Time Else Rs("VisitTime") = Null
Rs("AdmissionStatus") = c.Visit_AdmissionStatus
Rs("VisitType") = c.Visit_Type
Rs("VisitLocation") = c.Visit_Location
Rs("Clinician") = c.Visit_Clinician
Rs("Clinic_ReferralMode") = c.Clinic_ReferralMode
Rs("Clinic_ReferredBy") = c.Clinic_ReferredBy
If IsDate(c.Clinic_ReferralDate) Then Rs("Clinic_ReferralDate") = c.Clinic_ReferralDate Else Rs("Clinic_ReferralDate") = Null
Rs("Clinic_ReportTo") = c.Clinic_ReportTo
Rs("Clinic_ReferralReason") = c.Clinic_ReferralReason

'Equipment stuff
Rs("MachineID") = c.MachineID
Rs("MaskManufacturer") = c.Mask_Manufacturer
Rs("MaskModel") = c.Mask_Model
Rs("Frame") = c.Mask_Frame
Rs("HeadGearType") = c.Mask_HeadgearType
Rs("HeadGearSize") = c.Mask_HeadgearSize
Rs("CushionSize") = c.Mask_CushionSize
Rs("Chinstrap") = c.ChinStrap
Rs("CPAP_HumidifierType") = c.CPAP_HumidifierType

'Treatment settings
Rs("TreatmentMode") = c.TreatmentMode
If (c.CPAP_Set = "") Then Rs("CPAP_Set") = Null Else Rs("CPAP_Set") = c.CPAP_Set
If (c.APAP_Min = "") Then Rs("APAP_Min") = Null Else Rs("APAP_Min") = c.APAP_Min
If (c.APAP_Max = "") Then Rs("APAP_Max") = Null Else Rs("APAP_Max") = c.APAP_Max
If (c.CPAP_Measured = "") Then Rs("CPAP_Measured") = Null Else Rs("CPAP_Measured") = c.CPAP_Measured
If (c.Bilevel_EPAP = "") Then Rs("Bilevel_EPAP") = Null Else Rs("Bilevel_EPAP") = c.Bilevel_EPAP
If (c.Bilevel_IPAP = "") Then Rs("Bilevel_IPAP") = Null Else Rs("Bilevel_IPAP") = c.Bilevel_IPAP
If (c.O2_Flow = "") Then Rs("O2Flow") = Null Else Rs("O2Flow") = c.O2_Flow
Rs("O2EntrainedSpot") = c.O2_EntrainedLocation
Rs("Settings_SetOrPrescribed") = c.Settings_SetOrPrescribed

'Adherence stuff
Rs("HourMeter") = c.HourMeter
If IsDate(c.CPAP_Commenced) Then Rs("CPAP_Commenced") = c.CPAP_Commenced Else Rs("CPAP_Commenced") = Null
Rs("Adh_Obj_PerNight") = c.Adh_Obj_PerNight
Rs("Adh_Obj_PerWeek") = c.Adh_Obj_PerWeek
Rs("Adh_Obj_Period") = c.Adh_Obj_Period
Rs("Adh_Rep_PerNight") = c.Adh_Rep_PerNight
Rs("Adh_Rep_PerWeek") = c.Adh_Rep_PerWeek
Rs("Adh_Rep_Period") = c.Adh_Rep_Period

'DHS stuff
Rs("DHS_CPAP_ApprovedForPump") = c.DHS_CPAP_ApprovedForPump
Rs("DHS_CPAP_HireTrialFirst") = c.DHS_CPAP_HireTrialFirst
Rs("DHS_CPAP_Approval_EnteredBy") = c.DHS_CPAP_Approval_EnteredBy
Rs("DHS_CPAP_Approval_Date") = c.DHS_CPAP_Approval_Date
Rs("DHS_CPAP_NewOrSecondHand") = c.DHS_CPAP_NewOrSecondHand
Rs("DHS_CPAP_Copayment_Fee") = c.DHS_CPAP_Copayment_Fee
Rs("DHS_CPAP_Copayment_Received") = c.DHS_CPAP_Copayment_Received
Rs("DHS_CPAP_Copayment_Location") = c.DHS_CPAP_Copayment_Location
Rs("DHS_CPAP_Copayment_Method") = c.DHS_CPAP_Copayment_Method
Rs("DHS_CPAP_Copayment_ReceiptIssued") = c.DHS_CPAP_Copayment_ReceiptIssued
Rs("DHS_CPAP_Copayment_TransactedBy") = c.DHS_CPAP_Copayment_TransactedBy
If IsDate(c.DHS_CPAP_Copayment_TransactDate) Then Rs("DHS_CPAP_Copayment_TransactDate") = c.DHS_CPAP_Copayment_TransactDate Else Rs("DHS_CPAP_Copayment_TransactDate") = Null
Rs("DHS_CPAP_AuthorisingPhysician") = c.DHS_CPAP_AuthorisingPhysician
Rs("BillingItems") = c.BillingItems
Rs("BillingStatus") = c.BillingStatus

'Clinical stuff
Rs("Note") = c.Visit_Note
Rs("ESS") = c.ESS
Rs("CPAP_ReportedAdherencePeriod") = c.CPAP_ReportedAdherencePeriod
Rs("CPAP_Condensation") = c.CPAP_Condensation
Rs("CPAP_Drymouth") = c.CPAP_Drymouth
Rs("CPAP_Drynose") = c.CPAP_Drynose
Rs("CPAP_Pressure") = c.CPAP_Pressure
Rs("CPAP_Runny") = c.CPAP_Runny
Rs("Mask_Eyes") = c.Mask_Eyes
Rs("Mask_Face") = c.Mask_Face
Rs("Mask_Leaks") = c.Mask_Leaks
Rs("Mask_Comfort") = c.Mask_Comfort
Rs("Symptoms_EDS") = c.Symptoms_EDS
Rs("Symptoms_Headache") = c.Symptoms_Headache
Rs("Symptoms_Napping") = c.Symptoms_Napping
Rs("StudyID") = c.LinkedStudyID
Rs.Update
Rs.MoveFirst
Save_CpapToDB = Rs("VisitID")

Rs.Close

Exit Function


Save_CpapToDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Save_CpapToDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Friend Function Save_CpxToDB(X As Cpx, Mode As RecordMode) As Long
'Returns the recordID of the new or edited record

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String
   
On Error GoTo Save_CpxToDB_Error

Select Case Mode
    Case RecordMode.editRecord
        sql = "SELECT * FROM ExResults WHERE ResultsID =" & X.RecordID
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        If Rs.EOF Then
            MsgBox "Problem saving CPX results"
            Exit Function
        End If
    Case RecordMode.newRecord
        sql = "SELECT * FROM ExResults WHERE ResultsID =0"
        Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
        Rs.AddNew
        Rs("PatientID") = X.PtID
End Select
    
Rs("Campus") = X.Lab
Rs("MO") = X.RequestingMO
Rs("RequestingMO_Pn") = X.RequestingMO_Pn
Rs("RefLocation") = X.RefLocation & ""
Rs("Request_HealthServiceID") = X.Ref_HealthServiceID
Rs("Report_CopyTo") = X.ReportCopyTo & ""
Rs("TestType") = X.TestType
Rs("TestDate") = X.TestDate
Rs("RequestDate") = X.RequestDate
Rs("TestTime") = X.TestTime
Rs("scientist") = X.Scientist
Rs("comments") = X.Comments & ""
Rs("LastBD") = X.LastBd
Rs("Smoke") = X.Smoke & ""
If X.Packyrs = "" Then Rs("Packyrs") = Null Else Rs("Packyrs") = X.Packyrs
Rs("Weight") = Val(X.ThisVisit_weight)
Rs("Height_rft") = Val(X.ThisVisit_height)
Rs("ClinicalNote") = X.ClinicalNote & ""
Rs("Source") = X.Source & ""
Rs("AdmissionStatus") = X.AdmissionStatus
Rs("BillingMO") = X.BillingMO & ""
Rs("BillingItems") = X.BillingItems & ""
Rs("BillingStatus") = X.BillingStatus & ""
Rs("BillingMO_ProviderNumber") = X.BillingMO_ProviderNumber & ""

If IsNumeric(X.FEV1) Then Rs("FEV1") = X.FEV1 Else Rs("FEV1") = Null
If IsNumeric(X.FVC) Then Rs("FVC") = X.FVC Else Rs("FVC") = Null

Rs("HRstring") = X.string_HR
Rs("Loadstring") = X.string_Load
Rs("PetCO2string") = X.string_PetCO2
Rs("PetO2string") = X.string_PetO2
Rs("SaO2string") = X.string_SpO2
Rs("VCO2string") = X.string_VCO2
Rs("VEstring") = X.string_VE
Rs("VO2string") = X.string_VO2
Rs("Vtstring") = X.string_Vt

If IsNumeric(X.ABG_BE) Then Rs("BE") = X.ABG_BE Else Rs("BE") = Null
If IsNumeric(X.ABG_HCO3) Then Rs("HCO3") = X.ABG_HCO3 Else Rs("HCO3") = Null
If IsNumeric(X.ABG_PaCO2) Then Rs("pCO2") = X.ABG_PaCO2 Else Rs("pCO2") = Null
If IsNumeric(X.ABG_PaO2) Then Rs("pO2") = X.ABG_PaO2 Else Rs("pO2") = Null
If IsNumeric(X.ABG_pH) Then Rs("pH") = X.ABG_pH Else Rs("pH") = Null
If IsNumeric(X.ABG_SaO2) Then Rs("exSaO2") = X.ABG_SaO2 Else Rs("exSaO2") = Null
If IsNumeric(X.SpO2_Post) Then Rs("PostSaO2") = X.SpO2_Post Else Rs("PostSaO2") = Null

If IsNumeric(X.Borg_Dyspnoea) Then Rs("Dysp") = X.Borg_Dyspnoea Else Rs("Dysp") = Null
If IsNumeric(X.Borg_Legs) Then Rs("Legs") = X.Borg_Legs Else Rs("Legs") = Null
If IsNumeric(X.Borg_Other) Then Rs("Other") = X.Borg_Other Else Rs("Other") = Null
Rs("OtherSympt") = X.Borg_OtherSymptom & ""

Rs("PostBP") = X.BP_Post & ""
Rs("RestBP") = X.BP_Rest & ""

Rs("Report") = X.Report
Rs("Reporter") = X.Reporter
Rs("ReportDate") = X.ReportDate
Rs("VerifiedDate") = X.VerifiedDate
Rs("ReportStatus") = X.ReportStatus
Rs("VerifiedBy") = X.VerifiedBy

Rs.Update
Rs.MoveFirst
Save_CpxToDB = Rs("ResultsID")
Rs.Close

Exit Function

Save_CpxToDB_Error:
    Msg = ""
   Call ErrorLog(Msg, "Save_CpxToDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Save_RMDemoToDB(D As PtDemographics) As Boolean
'Saves non-Medtrak demographic data to TrakLocal

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim Msg As String

On Error GoTo Save_RMDemoToDB_Error

Screen.MousePointer = 11

'Check for an existing record. if not create one
sql = "SELECT * FROM tbl_RM_Demographics WHERE PatientID=" & D.ID
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
Select Case Rs.EOF
    Case False
    Case True
        MsgBox "creating RM record"
        Rs.Close
        sql = "SELECT * FROM tbl_RM_Demographics WHERE PatientID=0"
        Rs.Open sql, ConnectString
        Rs.AddNew
End Select

'Update record
Rs("PatientID") = D.ID
Rs("MedTrakAlertFlag") = D.MedtrakAlert
Rs("gHeight") = D.Height
Rs("race_RFTID") = D.Race_RftID
Rs("LastUpdated") = Now
Rs.Update
Rs.Close
Screen.MousePointer = 0

Exit Function


Save_RMDemoToDB_Error:
    Msg = ""
   Call ErrorLog(Msg, "Save_RMDemoToDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Friend Function Get_DemographicsFromDB(PtID As Long) As PtDemographics

Dim Rs As New ADODB.Recordset
Dim sql As String
Dim sql2 As String
Dim P As PtDemographics
Dim Msg As String

On Error GoTo Get_DemographicsFromDB_Error

'Clear UDT
Dim emptyUDT As PtDemographics
P = emptyUDT

'If PtID=0 then Pt not found so exit
Select Case PtID
    Case 0
        Get_DemographicsFromDB = P
        Exit Function
    Case -1
        P.ID = -1
        Get_DemographicsFromDB = P
        Exit Function
End Select

'Retrieve demo data for primary record
'sql = "SELECT tbl_Address.*, tbl_patient.*, tbl_Name.* "
'sql = sql & "FROM tbl_Name RIGHT OUTER JOIN tbl_patient ON tbl_Name.TypeID = tbl_patient.ID "
'sql = sql & "LEFT OUTER JOIN tbl_Address ON tbl_patient.ID = tbl_Address.TypeID "
'sql = sql & "WHERE (tbl_Name.NameType = 'L') AND (tbl_patient.Type = 'P') AND (tbl_patient.ID = '" & PtID & "') AND (tbl_Address.AddressType = '4' OR tbl_Address.AddressType IS NULL) "
'sql = sql & "AND (tbl_Address.Type = 'P') AND (tbl_Name.Type = 'P')"

sql = "SELECT tbl_Address.*, tbl_patient.*, tbl_Name.* "
sql = sql & "FROM tbl_patient LEFT OUTER JOIN tbl_Address ON tbl_patient.ID = tbl_Address.TypeID LEFT OUTER JOIN "
sql = sql & "tbl_Name ON tbl_patient.ID = tbl_Name.TypeID "
sql = sql & "WHERE (tbl_Name.NameType = 'L') AND (tbl_Name.Type = 'P') AND (tbl_patient.ID = '" & PtID & "') AND "
sql = sql & "(tbl_patient.Type = 'P') AND (tbl_Address.Type = 'P' OR tbl_Address.Type IS NULL) AND (tbl_Address.AddressType = '4' OR "
sql = sql & "tbl_Address.AddressType IS NULL)"

Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
Select Case Rs.RecordCount
    Case 0
        'MsgBox "Should never happen error - Cant find demographics records for this Pt"
        Rs.Close
    Case Else
        'If > 1 record returned then >1 home address and/or >1 name for this patient
        If Rs.RecordCount > 1 Then
            Msg = ""
            Do While Not Rs.EOF
                Msg = Msg & Rs("Surname") & ", " & Rs("FirstName") & "   " & Rs("Address1") & ", " & Rs("Suburb") & vbCrLf
                Rs.MoveNext
            Loop
            Msg = "More than 1 name and/or home address for this patient -" & vbCrLf & vbCrLf & Msg & vbCrLf & "The first record will be selected."
            Msg = Msg & vbCrLf & "Please rectify the problem in Trak."
            MsgBox Msg, vbOKOnly, "ResMed"
            Rs.MoveFirst
        End If

        'Load non-repeating trak data
        P.ID = Rs("ID")
        P.Title = ChangeCase(Rs("Title"), Titlecase)
        P.Firstname = ChangeCase(Rs("FirstName"), Titlecase)
        P.Middlename = ""
        P.Surname = ChangeCase(Rs("Surname"), Uppercase)
        P.DOB = Format(Rs("DOB"), "dd/mm/yyyy")
        P.Gender = Map_RefTbl(Rs("Gender"), Gender, HL7Code_RMText)
        P.Address1 = ChangeCase(Rs("Address1") & "", Titlecase)
        P.Address2 = ChangeCase(Rs("Address2"), Titlecase)
        P.Suburb = ChangeCase(Rs("Suburb"), Titlecase)
        P.State = ChangeCase(Rs("State"), Titlecase)
        P.Postcode = Rs("Postcode") & ""
        
        'Get death info
        Dim dInfo As DeathInfo
        dInfo = Get_DeathInfo(Rs("ID"))
        P.DeathDate = dInfo.DeathDate
        P.DeathStatus = dInfo.DeathStatusString
        
        P.CountryOfBirth = Map_RefTbl(Rs("Nationality"), Nationality, HL7Code_RMText)
        P.Language = Map_RefTbl(Rs("Language"), Language, HL7Code_RMText)
        P.UR_PrimaryAustin = Rs("URNumber")
        P.URforDisplay = Get_PtString(PtID, cHS.CurrentVAED, DisplayUR) ' Get_URforDisplay(gPtID, cHS.CurrentVAED)
        P.AboriginalStatus = Map_RefTbl(Rs("Race") & "", Ethnicity, HL7Code_RMText)
        Rs.Close
    
        'Get medicare num
        P.MedicareNo = GetMedicareNo(P.ID, Formatted)
        
        'Get health insurance
        'p.HealthInsurance = Rs("Private Insurance") & ""
    
        'Get Home landline phone number
        sql = "SELECT TOP 1 tbl_contact.* FROM tbl_contact WHERE Type='P' AND TypeID = '" & PtID & "' "
        sql = sql & "AND ContactType='H' AND NumberType='PH' ORDER BY ContactID ASC"
        Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        If Not Rs.EOF Then P.PhoneHome = Rs("Number") & "" Else P.PhoneHome = Empty
        Rs.Close
        
        'Home home mobile
        sql = "SELECT TOP 1 tbl_contact.* FROM tbl_contact WHERE Type='P' AND TypeID = '" & PtID & "' "
        sql = sql & "AND ContactType='H' AND NumberType='CP' ORDER BY ContactID ASC"
        Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        If Not Rs.EOF Then P.PhoneMobile = Rs("Number") & "" Else P.PhoneMobile = Empty
        Rs.Close
        
        'Work
        sql = "SELECT TOP 1 tbl_contact.* FROM tbl_contact WHERE Type='P' AND TypeID = '" & PtID & "' "
        sql = sql & "AND ContactType='W' AND NumberType='PH' ORDER BY ContactID ASC"
        Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        If Not Rs.EOF Then P.PhoneWork = Rs("Number") & "" Else P.PhoneWork = Empty
        Rs.Close
        
        'Get email address
        sql = "SELECT tbl_contact.* FROM tbl_contact WHERE Type='P' AND TypeID = '" & PtID & "' "
        sql = sql & "AND NumberType='Internet'"
        Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        If Not Rs.EOF Then P.Email = Rs("EmailAddress") & "" Else P.Email = Empty
        Rs.Close
                
        'Get local ResMed demographics
        '  check for >1 record. Flag and need to merge. Choose most recent for this retrieve
        sql = "SELECT List_RacesForRFTs.*, tbl_RM_Demographics.* FROM tbl_RM_Demographics "
        sql = sql & "LEFT OUTER JOIN List_RacesForRFTs ON tbl_RM_Demographics.Race_RFTID = List_RacesForRFTs.RaceID "
        sql = sql & "WHERE (tbl_RM_Demographics.PatientID = '" & PtID & "')"
        sql = sql & "ORDER BY dbo.tbl_RM_Demographics.LastUpdated DESC"
        Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
        Select Case Rs.RecordCount
            Case 0
                'Create an empty record
                Rs.Close
                sql2 = "SELECT * FROM tbl_RM_Demographics WHERE PatientID=0"
                Rs.Open sql2, ConnectString, adOpenStatic, adLockPessimistic
                Rs.AddNew
                Rs("PatientID") = PtID
                Rs("LastUpdated") = Now
                Rs.Update
                Rs.Close
                Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
            Case 1
                'Continue
            Case Else
                Msg = "WARNING" & vbCrLf & vbCrLf & "There are " & Rs.RecordCount & " local ResMed demographic records for this patient (TrakID=" & Rs("PatientID") & ")" & vbCrLf
                Msg = Msg & "There should only be one." & vbCrLf & "Please contact DRSM database adminstrator to arrange merge." & vbCrLf & vbCrLf
                Msg = Msg & "The most recent record will be used."
                MsgBox Msg, vbOKOnly, "Demographics"
                DoEvents
                Rs.MoveFirst
        End Select
        'Get data
        P.OldResmed_SecondaryUR = ""
        P.OldResmed_UR = ""
        If IsNull(Rs("MedtrakAlertFlag")) Then
            P.MedtrakAlert = 0
        Else
            P.MedtrakAlert = Abs(CInt(Rs("MedtrakAlertFlag")))
        End If
        P.Height = Val(Rs("gHeight") & "")
        P.Race_Rft = Rs("Race_Description") & ""
        If IsNull(Rs("RaceID")) Then P.Race_RftID = Empty Else P.Race_RftID = Rs("RaceID")
        If IsDate(Rs("Lastupdated")) Then P.LastRMDemoUpdate = Rs("Lastupdated")
        Rs.Close
End Select

Get_DemographicsFromDB = P

Exit Function

Get_DemographicsFromDB_Error:
    Msg = ""
    Call ErrorLog(Msg, "Get_DemographicsFromDB", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function


Public Function FlowVol_GetFromFile(ImageCtrl As Image, Filename As String) As Boolean
Dim Msg As String

On Error GoTo FlowVol_GetFromFile_Error

ImageCtrl.Tag = ""
ImageCtrl.Picture = LoadPicture(Filename)
ImageCtrl.Tag = Filename

Exit Function


FlowVol_GetFromFile_Error:
    Msg = ""
   Call ErrorLog(Msg, "FlowVol_GetFromFile", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Public Function FlowVol_GetFromDBToImageCtl(localImage As Image, RecordID As Long) As Boolean
    
Dim sql As String
Dim Rs As New ADODB.Recordset
Dim S As New ADODB.Stream
Dim TempFile As String
Dim Msg As String

On Error GoTo FlowVol_GetFromDBToImageCtl_Error

TempFile = Environ("tmp") & "temp.bmp"

sql = "SELECT FlowVolLoop FROM RftResults WHERE ResultsID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields("FlowVolLoop").Value) Then
        S.Type = adTypeBinary
        S.Open
        S.Write Rs.Fields("FlowVolLoop").Value
        S.SaveToFile TempFile, adSaveCreateOverWrite
        localImage.Picture = LoadPicture(TempFile)
        LoadImageFromDB = True
        S.Close
        Kill (TempFile)
    End If
End If
Rs.Close

Exit Function


FlowVol_GetFromDBToImageCtl_Error:
    Msg = ""
   Call ErrorLog(Msg, "FlowVol_GetFromDBToImageCtl", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function

Public Function FlowVol_GetFromDBToFile(Filename As String, RecordID As Long) As Boolean
    
Dim sql As String
Dim Rs As New ADODB.Recordset
Dim S As New ADODB.Stream
Dim Msg As String

On Error GoTo FlowVol_GetFromDBToFile_Error

sql = "SELECT FlowVolLoop FROM RftResults WHERE ResultsID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockReadOnly
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields("FlowVolLoop").Value) Then
        S.Type = adTypeBinary
        S.Open
        S.Write Rs.Fields("FlowVolLoop").Value
        S.SaveToFile Filename, adSaveCreateOverWrite
        FlowVol_GetFromDBToFile = True
        S.Close
    End If
End If
Rs.Close

Exit Function


FlowVol_GetFromDBToFile_Error:
    Msg = ""
   Call ErrorLog(Msg, "FlowVol_GetFromDBToFile", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Public Function FlowVol_SaveToDB(FV As Image, RecordID As Long) As Boolean

Dim sql As String
Dim Rs As New ADODB.Recordset
Dim S As New ADODB.Stream
Dim ImageFile As String
Dim Msg As String

On Error GoTo FlowVol_SaveToDB_Error

Select Case FV.Tag
    Case "":    Exit Function        'Nothing to save
    Case Else:  ImageFile = FV.Tag
End Select

sql = "SELECT FlowVolLoop FROM RftResults WHERE ResultsID=" & RecordID
Rs.Open sql, ConnectString, adOpenStatic, adLockPessimistic
    If FV <> 0 Then
        Set S = New Stream
        S.Type = adTypeBinary
        S.Open
        S.LoadFromFile ImageFile
        Rs("FlowVolLoop") = S.Read
        Set S = Nothing
        Kill ImageFile
    ElseIf FV = 0 Then
        Rs("FlowVolLoop") = Null
    End If
    Rs.Update
Rs.Close
       
Exit Function


FlowVol_SaveToDB_Error:
    Msg = ""
   Call ErrorLog(Msg, "FlowVol_SaveToDB", "cPatient", Err.Number, Err.Description)
   Screen.MousePointer = 0
   Exit Function
   Resume
End Function


Public Function Has_PreviousRFTs(PatientID As Long) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : Has_PreviousRFTs
' Author    : rochpd
' Date      : 23/08/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Dim Msg As String
Dim c As DRSMContacts
Dim i As Integer, j As Integer

On Error GoTo Has_PreviousRFTs_Error

ReDim c.RFTs(0)

'Look for previous RFTs and set previous button
c = CurrentPt.Get_DRSMContacts_V2(CurrentPt.PtID, DRSMContactType.RFTs)
j = 0
For i = 1 To UBound(c.RFTs)
    If UCase(Left(c.RFTs(i).Description, 5)) = "RFT'S" Then j = j + 1
Next i
If j > 1 Then Has_PreviousRFTs = True Else Has_PreviousRFTs = False

Exit Function

Has_PreviousRFTs_Error:
    Msg = ""
    Call ErrorLog(Msg, "Has_PreviousRFTs", "cPatient", Err.Number, Err.Description)
    Screen.MousePointer = 0
    Exit Function
    Resume
End Function
