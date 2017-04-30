Attribute VB_Name = "mUDTDefinitions"


Public Type RftRecord
    ID As Long
    SourceTable As DbTables
End Type

Public Type FindPt_Criteria
    Surname As String
    Firstname As String
    DOB As Date
    Gender As String
    Mode_Surname As FindPt_SearchMode
    Mode_Firstname As FindPt_SearchMode
    Mode_BookingType As ApptType
    RestrictSearch As RestrictPtSearch
End Type
Public Enum RestrictPtSearch
    PreviousRFTs
    PreviousSleepStudies
    PreviousVRSSContacts
    DontRestrict
End Enum
Public Enum ApptType
    SlpBookings
    Appts
End Enum
Public Enum FindPt_SearchMode
    FirstPart
    AnyPart
    exact
End Enum
Public Type PtAlias
    Title As String
    Surname As String
    Firstname As String
    PtID As Long
End Type
Public Enum AlertSourceTable
    Cerner
    MedTrak
End Enum
Public Enum AlertGroup
    Allergy
    Alert
End Enum
Public Type Quarter
    Abbrev As String
    QName As String
    Index As Integer
    StartDate As Date
    EndDate As Date
End Type
Public Enum Plot_Symbol
    Circle_large_open
    Circle_large_closed
    Circle_small_closed
End Enum
Public Enum plot_Linestyle
    Solid = vbSolid
    Dash = vbDash
    DashDot = vbDashDot
    Dot = vbDot
End Enum
Public Enum Plot_Colour
    Count = 6
    Black = vbBlack
    Blue = vbBlue
    red = vbRed
    Green = &H8000&
    brown = &H40C0&
    Grey = &HC0C0C0
End Enum
Public Type Alert
    SourceTable As String
    AlertGroup As AlertGroup
    AlertType_Description As String
    AlertType_Code As String
    Alert_Description As String
    Alert_Code As String
    Alert_reactions As String
    Alert_severity As String
    Alert_Comment As String
    Date_Create As Date
    Date_LastUpdated As Date
    By_Surname As String
    By_FirstName As String
End Type
Public Type GPDetails
    Practitioner_id As String
    Surname As String
    GivenName As String
    Title As String
    PracticeName As String
    Address As String
    Suburb As String
    State As String
    Postcode As String
    Phone As String
    Email As String
End Type
Public Type PtSuburbDetails
    Suburb As String
    Postcode As String
    SLA As String
    Enabled As Boolean
End Type
Public Type DeathInfo
    DeathDate As String
    DeathStatusString As String
End Type

Public Type PtDemographics1
    PtID As Long
    Surname As String
    Firstname As String
    UR As String
    VINAH_AccountClass_Code As String
End Type

Public Type PtDemographics
    ID As Long
    Title As String
    Surname As String
    Firstname As String
    Middlename As String
    Gender As String
    DOB As Date
    UR_PrimaryAustin As String
    URforDisplay As String
    OldResmed_UR As String
    OldResmed_SecondaryUR As String
    Address1 As String
    Address2 As String
    Suburb As String
    State As String
    Postcode As String
    PhoneHome As String
    PhoneWork As String
    PhoneMobile As String
    Email As String
    HealthInsurance As String
    MedicareNo As String
    CountryOfBirth As String  'Country of birth
    Language As String
    Interpreter As String
    Race_Rft As String
    Race_RftID As Long
    AboriginalStatus As String
    DeathStatus As String
    DeathDate As String
    MedtrakAlert As Integer
    Height As Single
    LastRMDemoUpdate As Date
    LastTrakUpdate As Date
End Type

Public Type RftTestIndicators
    Spir_pre As Boolean
    Spir_post1 As Boolean
    Spir_post2 As Boolean
    Spir_any As Boolean
    FlowVols As Boolean
    Challenge_Histamine As Boolean
    Challenge_Methacholine As Boolean
    Challenge_Mannitol As Boolean
    Challenge_any As Boolean
    TLCO As Boolean
    Vols As Boolean
    Abgs_1 As Boolean
    Abgs_2 As Boolean
    Abgs_any As Boolean
    Shunt_1 As Boolean
    Shunt_2 As Boolean
    Shunt_any As Boolean
    Mrps  As Boolean
    ExProv As Boolean
    Hast As Boolean
    SPT As Boolean
    SixMWD As Boolean
    Cpx As Boolean
    Other As Boolean
    Other_description As String
End Type

Public Type SptAllergen
    AllergenID As Integer
    AllergenName As String
    Set_Standard As Boolean
    DisplayOrder As Integer
    Enabled As Boolean
    Archived As Boolean
    GroupName As String
    GroupCode As String
    GroupID As Integer
    GroupColour As String
    GroupVBHexValue As String
    Wheal_mm As String
End Type

Public Type SptAllData
    Allergens() As SptAllergen
    Site As String
    Antihistamine As Boolean
    BetaBlocker As Boolean
End Type

Public Type SptAllergenGroup
    GroupID As Integer
    Code    As String
    Description  As String
    Colour As String
    ColourVBhexvalue As String
End Type

Public Type SptAllergenSet
    SetID As Integer
    Description  As String
End Type

Public Type RFT_Episode
    RecordID As Long
    SourceTbl As DbTables
    PtID As Long
    D_Race_Rft As String
    D_MedicareNo As String
    
    RequestDate As Date
    RequestingMO As String
    RequestingMO_Pn As String
    Ref_HealthService As String
    Ref_HealthServiceID As HSID
    RefLocation As String
    ReportCopyTo As String
    TestDate As Date
    TestTime As Date
    Smoke As String
    Packyrs As String
    ThisVisit_weight As String
    LastVisit_weight As String
    ThisVisit_height As String
    LastVisit_height As String
    Lab As String
    ClinicalNote As String
    AdmissionStatus As String
    BillingMO As String
    BillingItems As String
    BillingStatus As String
    BillingMO_ProviderNumber As String
End Type

Public Type RFTs
    RecordID As Long
    PtID As Long
    
    E_RequestDate As Date
    E_RequestingMO As String
    E_RequestingMO_Pn As String
    E_Ref_HealthService As String
    E_Ref_HealthServiceID As HSID
    E_RefLocation As String
    E_ReportCopyTo As String
    E_TestDate As Date
    E_TestTime As Date
    E_Smoke As String
    E_Packyrs As String
    E_ThisVisit_weight As String
    E_ThisVisit_height As String
    E_LastVisit_height As String
    E_Lab As String
    E_ClinicalNote As String
    'E_Source As String
    E_AdmissionStatus As String
    E_BillingMO As String
    E_BillingItems As String
    E_BillingStatus As String
    E_BillingMO_ProviderNumber As String
    
    TestType As String
    Scientist As String
    TechComments As String
    
    D_Surname As String
    D_FirstName As String
    D_Middlename As String
    D_UR_PrimaryAustin As String
    D_URforDisplay As String
    D_Address1 As String
    D_Address2 As String
    D_Suburb As String
    D_State As String
    D_Postcode As String
    D_DOB As Date
    D_Gender As String
    D_Race_Rft As String
    D_MedicareNo As String
    

    LastBd As String
    spir As String
    BD_1 As String
    BD_2 As String
    preFEV1 As String
    preFVC As String
    preVC As String
    preFEF As String
    postFEV1_1 As String
    postFVC_1 As String
    postVC_1 As String
    postFEF_1 As String
    postFEV1_2 As String
    postFVC_2 As String
    postVC_2 As String
    postFEF_2 As String
    
    tlcoEquip As String
    VA As String
    TLCO As String
    KCO As String
    Hb As String
    HbInfo As String
    Hb_WaitingForResult As Boolean
    
    FRC As String
    TLC As String
    VC As String
    RV As String
    Raw As String
    Raw_2 As String
    Sgaw As String
    Sgaw_2 As String
    volEquip As String
    
    FiO2_1 As String
    pH_1 As String
    pCO2_1 As String
    pO2_1 As String
    Shunt_1 As String
    BE_1 As String
    HCO3_1 As String
    SaO2_1 As String
    
    FiO2_2 As String
    pH_2 As String
    pCO2_2 As String
    pO2_2 As String
    Shunt_2 As String
    BE_2 As String
    HCO3_2 As String
    SaO2_2 As String
    
    MIP As String
    MEP As String
    SNIP As String
    MrpEquip As String
    
    SkinTestDataString As String
    Skin_PanelDescription As String
    AST_TestDataString As String
    
    prov_ChallengeType As String
    prov_Control As String
    prov_Dose1 As String
    prov_Dose2 As String
    prov_Dose3 As String
    prov_Dose4 As String
    prov_Dose5 As String
    prov_Dose6 As String
    prov_Dose7 As String
    prov_Dose8 As String
    prov_Dose9 As String
    prov_Post As String
    prov_PostBDType As String
    prov_FEV1pre As String
    prov_FVCpre As String
    prov_VCpre As String
    prov_FERpre As String
    prov_FEV1_Control As String
    prov_FEV1_Dose1 As String
    prov_FEV1_Dose2 As String
    prov_FEV1_Dose3 As String
    prov_FEV1_Dose4 As String
    prov_FEV1_Dose5 As String
    prov_FEV1_Dose6 As String
    prov_FEV1_Dose7 As String
    prov_FEV1_Dose8 As String
    prov_FEV1_Dose9 As String
    prov_FEV1_Post As String
    
    SpeedString As String
    AirSaO2String As String
    O2SaO2String As String
    

    Report  As String
    Reporter  As String
    ReportDate As Date
    ReportStatus  As String
    VerifiedBy  As String
    VerifiedDate As Date
    LastUpdate As Date
    
    ppn_preFEV1 As String   'ppn = % of predicted normal
    ppn_preFVC As String
    ppn_preVC As String
    ppn_preFEF As String
    ppn_prePEF As String
    ppn_post1FEV1 As String
    ppn_post1VC As String
    ppn_post1FVC As String
    ppn_post1FEF As String
    ppn_post1PEF As String
    ppn_post2FEV1 As String
    ppn_post2FVC As String
    ppn_post2VC As String
    ppn_post2FEF As String
    ppn_post2PEF As String
    ppn_TLCO As String
    ppn_TLCOhb As String
    ppn_KCO As String
    ppn_KCOhb As String
    ppn_VA As String
    ppn_FRC As String
    ppn_TLC As String
    ppn_RV As String
    ppn_MIP As String
    ppn_MEP As String
    ppn_SNIP As String
    
    pnr_FEV1 As String          'pnr = predicted normal range
    pnr_FVC As String
    pnr_VC As String
    pnr_FER As String
    pnr_FEF As String
    pnr_PEF As String
    pnr_TLCO As String
    pnr_KCO As String
    pnr_VA As String
    pnr_TLC As String
    pnr_RV As String
    pnr_FRC As String
    pnr_RVTLC As String
    pnr_Raw As String
    pnr_Sgaw As String
    pnr_MIP As String
    pnr_MEP As String
    pnr_SNIP As String
    pnr_pH As String
    pnr_PaCO2 As String
    pnr_PaO2 As String
    pnr_BE As String
    pnr_HCO3 As String
    pnr_SaO2 As String
    pnr_Shunt As String
    
End Type

Public Type WalkTest
    RecordID As Long
    PtID As Long
    TestDate As Date
    TestTime As Timer
End Type


Public Type Cpx
    RecordID As Long
    PtID As Long
    TestDate As Date
    TestTime As Date
    Lab As String
    ClinicalNote As String
    TestType As String
    Scientist As String
    Comments As String
    
    D_Surname As String
    D_FirstName As String
    D_Middlename As String
    D_UR_PrimaryAustin As String
    D_URforDisplay As String
    D_Address1 As String
    D_Address2 As String
    D_Suburb As String
    D_State As String
    D_Postcode As String
    D_DOB As Date
    D_Gender As String
    D_Race_Rft As String
    
    ThisVisit_weight As String
    ThisVisit_height As String
    LastVisit_height As String
    Smoke As String
    Packyrs As String

    RequestDate As Date
    RequestingMO_Pn As String
    RequestingMO As String
    RefLocation As String
    Ref_HealthService As String
    Ref_HealthServiceID As HSID
    ReportCopyTo As String
    Source As String
    AdmissionStatus As String
    BillingMO As String
    BillingItems As String
    BillingStatus As String
    BillingMO_ProviderNumber As String
    BillingSessionID As String
    MedicareNo As String
    Report  As String
    Reporter  As String
    ReportDate As Date
    ReportStatus  As String
    VerifiedBy  As String
    VerifiedDate As Date
    LastUpdate As Date
    
    LastBd As String
    FEV1 As String
    FVC As String
    BP_Rest As String
    BP_Post As String
    Borg_Dyspnoea As String
    Borg_Legs As String
    Borg_Other As String
    Borg_OtherSymptom As String
    Stages As String
    EndLoad As String
    
    SpO2_Post As String
    ABG_pH As String
    ABG_PaO2 As String
    ABG_PaCO2 As String
    ABG_BE As String
    ABG_SaO2 As String
    ABG_HCO3 As String
    
    string_VO2 As String
    string_VCO2 As String
    string_VE As String
    string_HR As String
    string_SpO2 As String
    string_PetO2 As String
    string_PetCO2 As String
    string_Load As String
    string_Vt As String
    VO2() As String
    VCO2() As String
    VE() As String
    HR() As String
    SpO2() As String
    PetO2() As String
    PetCO2() As String
    Load() As String
    Vt() As String
    MaxVO2 As String
    MaxVO2ml_kg As String
    MaxVCO2 As String
    MaxVE As String
    MaxHR As String
    MaxVt As String
    MaxVeVO2 As String
    MaxO2Pulse As String
    predMaxVE As String
    predMaxVt As String
    predMaxHR As String
    predMaxVO2 As String
    predMaxO2Pulse As String
    predMaxW As String
    predMaxVEppn As String
    predMaxVTppn As String
    predMaxHRppn As String
    predMaxVO2ppn As String
    predMaxWppn As String
    predMaxO2Pulseppn As String
    
    Hb As String
    Hb_Info As String
End Type

Public Type CpxData
    LineOfDataType As sType
    Load As String
    Time As String
    VE As Single
    VO2 As Single
    VCO2 As Single
    HR As Single
    SpO2 As Single
    Vt As Single
    PetO2 As Single
    PetCO2 As Single
    RQ As Single
    VEVO2 As Single
    VEVCO2 As Single
    O2Pulse As Single
End Type

Public Type ReportInfoSection
    TestDate As Date
    TestTime As Date
    Lab As String
    ClinicalNote As String
    ThisVisit_weight As String
    ThisVisit_height As String
    RequestingMO As String
    
    'Resp only
    Resp_Smoke As String
    Resp_Packyrs As String
    Resp_TestType As String
    Resp_RefLocation As String
    Resp_ReportCopyTo As String
    Resp_Source As String
    Resp_LastBd As String
    Resp_Hb As String
    Resp_Hb_Info As String
    
    'Sleep only
    Slp_StudyType As String
    Slp_TreatmentMode As String
    Slp_ReportDest_1 As String
    Slp_ReportDest_2 As String
    
    'CPAP clinic only
    Cpap_ReferralMode As String
    Cpap_ReportTo As String
    Cpap_VisitLocation As String
    Cpap_ReasonForReferral As String
End Type

Public Type ReportSection
    Report As String
    Race_Rft As String
    Scientist As String
    Comments As String
    Reporter As String
    ReportDate As Date
    ReportStatus As String
End Type

Public Type SlpRequest
    PatientID As Long
    StudyID As Long
    ApptID As Long
    
    StudyDate As Date
    ArrivalTime As Date
    Urgent_PSG  As Boolean
    Urgent_scoring As Boolean
    BookedLocation As String
    BookedBy As String
    Admission As Boolean
    Note As String
    ApptNote As String
    BedNumber_Booked As String
    BedNumber_AtPSG As String
    BedIndex As String
    
    RequestDate As Date
    RequestingMO As String
    RequestingMO_ProviderNum As String
    Request_HealthServiceID As HSID
    RequestingMO_ReviewLocation As String
    ReportDest_1 As String
    ReportDest_2 As String
    BilledTo As String
    RequestStatus As String
    Request_HardCopyRec As Boolean
    Interpreter_Language As String
    Interpreter_Booked As Boolean
    Transport_Required As Boolean
    Transport_Booked As Boolean
    EmergencyList As Boolean
    InfoSent As Boolean
    
    StudyType As String
    TreatmentMode As String
    RequestedLocation As String
    Special_TCO2 As Boolean
    Special_Poes As Boolean
    Special_EMGd As Boolean
    Special_Temp As Boolean
    Special_Other As String
    TreatmentOxygen As Boolean
    ClinicalNotes As String
    
    ApptConfirmedDate As Date
    ApptConfirmed As Boolean
    
    Ref_MO As String
    Ref_ProviderNum As String
    Ref_Address As String
    Ref_Duration As String
    Ref_Date As Date
    Ref_ActivateDate As Date
    LastEdit As Date
End Type

Public Type SlpResults
    RecordID As Long
    PatientID As Long
    
    D_Surname As String
    D_FirstName As String
    D_Middlename As String
    D_UR_PrimaryAustin As String
    D_URforDisplay As String
    D_Address1 As String
    D_Address2 As String
    D_Suburb As String
    D_State As String
    D_Postcode As String
    D_PhoneHome As String
    D_PhoneWork As String
    D_PhoneMobile As String
    D_DOB As Variant
    D_Gender As String
    D_Race_Rft As String
    Weight As Single
    Height_slp As Single
    
    StudyDate As Variant
    StudyTime As Variant
    StudyType As String
    Urgent_PSG  As Boolean
    Urgent_scoring As Boolean
    TreatmentMode As String
    TreatmentOxygen As Boolean
    Special_TCO2 As Boolean
    Special_Poes As Boolean
    Special_EMGd As Boolean
    Special_Temp As Boolean
    Special_OtherProcedure As String
    
    Lab As String
    Bed As String
    Note As String
    OpticalDisk As String
    Admission As String
    TechnicalNote As String
    ReasonForTest As String
    
    Report As String
    ReportStatus As String
    ReportDate As Variant
    ReportedNotEnteredDate As Variant
    EnteredNotVerifiedDate As Variant
    VerifiedNotPrintedDate As Variant
    VerifiedAndPrintedDate As Variant
    Reporter_MO As String
    ScoredBy As String
    ScoredDate As Variant
    VerifiedBy As String
    VerifiedDate As Variant
    
    BillingMO As String
    BillingMO_Pn_ID As Integer
    BillingMO_ProviderNumber As String
    BillingAmount As Currency
    BillingStatus As String
    BillingItems As String
    MedicareNo As String
    Payment_ReceivedDate_Medicare As Variant
    Payment_ReceivedDate_Patient As Variant
    Payment_Location As String
    Payment_Method  As String
    Payment_ReceivedAmount_Patient  As String
    Payment_ReceivedAmount_Medicare As String
    Payment_OwingAmount_Patient As String
    Payment_OwingAmount_Medicare As String
    Payment_ReceiptIssued  As String
    Payment_TransactedBy_Patient As String
    Payment_TransactedBy_Medicare As String
    Payment_AccountReminderDate As Variant
    Payment_AccountNote As String
    ReceivedAmount As Currency
    
    Request_Status As String
    Request_HardCopyRec As Boolean
    Request_HealthServiceID As HSID
    Request_Time As Variant
    Request_Date As Variant
    RequestingMO_P_ID As Integer
    RequestingMO As String
    RequestingMO_ProviderNum As String
    RequestedLocation As String
    RequestingMO_ReviewLocation As String
    RequestingMO_ReviewDate As Variant
    RequestMode As String
    ReportDest_1 As String
    ReportDest_2 As String
    Ref_MO As String
    Ref_Address As String
    Ref_Suburb As String
    Ref_PostCode As String
    Ref_Date As Variant
    Ref_ActivateDate As Date
    Ref_Duration As String
    Ref_ProviderNum As String
    Ref_MO_Pn_ID As Integer
    StudyBy As String
    Transport As String
    Technologist As String
    
    R_Recording_Start As String
    R_Recording_End As String
    R_Report_Start As String
    R_Report_End As String
    R_LOut As String
    R_LOn As String
    R_ArI_RespNrem As String
    R_ArI_RespRem As String
    R_ArI_RespAll As String
    R_ArI_PlmNrem As String
    R_ArI_PlmRem As String
    R_ArI_PlmAll As String
    R_ArI_OtherNrem As String
    R_ArI_OtherRem As String
    R_ArI_OtherAll  As String
    R_ArI_AllNrem As String
    R_ArI_AllRem As String
    R_ArI_AllAll As String
    R_ArI_SupRem As String
    R_ArI_SupNrem As String
    R_ArI_SupAll As String
    R_ArI_NonSupRem As String
    R_ArI_NonSupNrem As String
    R_ArI_NonSupAll As String
    R_AHI_SupNrem  As String
    R_AHI_SupRem As String
    R_AHI_SupAll As String
    R_AHI_NonSupNrem As String
    R_AHI_NonSupRem As String
    R_AHI_NonSupAll As String
    R_AHI_AllNrem As String
    R_AHI_AllRem As String
    R_AHI_AllAll As String
    R_RDI_SupNrem  As String
    R_RDI_SupRem As String
    R_RDI_SupAll As String
    R_RDI_NonSupNrem As String
    R_RDI_NonSupRem As String
    R_RDI_NonSupAll As String
    R_RDI_AllNrem As String
    R_RDI_AllRem As String
    R_RDI_AllAll As String
    R_RERA_SupNrem  As String
    R_RERA_SupRem As String
    R_RERA_SupAll As String
    R_RERA_NonSupNrem As String
    R_RERA_NonSupRem As String
    R_RERA_NonSupAll As String
    R_RERA_AllNrem As String
    R_RERA_AllRem As String
    R_RERA_AllAll As String
    R_HI_SupNrem As String
    R_HI_SupRem As String
    R_HI_SupAll As String
    R_HI_NonSupNrem As String
    R_HI_NonSupRem As String
    R_HI_NonSupAll As String
    R_HI_AllNrem As String
    R_HI_AllRem As String
    R_HI_AllAll As String
    R_OAI_SupNrem As String
    R_OAI_SupRem As String
    R_OAI_SupAll As String
    R_OAI_NonSupNrem As String
    R_OAI_NonSupRem As String
    R_OAI_NonSupAll As String
    R_OAI_AllNrem  As String
    R_OAI_AllRem As String
    R_OAI_AllAll As String
    R_CAI_SupNrem As String
    R_CAI_SupRem As String
    R_CAI_SupAll As String
    R_CAI_NonSupNrem As String
    R_CAI_NonSupRem As String
    R_CAI_NonSupAll As String
    R_CAI_AllNrem As String
    R_CAI_AllRem As String
    R_CAI_AllAll As String
    R_MAI_SupNrem As String
    R_MAI_SupRem As String
    R_MAI_SupAll As String
    R_MAI_NonSupNrem As String
    R_MAI_NonSupRem As String
    R_MAI_NonSupAll As String
    R_MAI_AllNrem As String
    R_MAI_AllRem As String
    R_MAI_AllAll As String
    R_ODI4_SupNrem As String
    R_ODI4_SupRem As String
    R_ODI4_SupAll As String
    R_ODI4_NonSupNrem As String
    R_ODI4_NonSupRem As String
    R_ODI4_NonSupAll As String
    R_ODI4_AllNrem  As String
    R_ODI4_AllRem As String
    R_ODI4_AllAll As String
    R_PLM_Nrem  As String
    R_PLM_Rem As String
    R_PLM_All As String
    R_PLM_Wake As String
    R_PLM_Ep_All As String
    R_PLM_Ep_NREM As String
    R_PLM_Ep_REM As String
    R_PLM_Ep_Wake As String
    R_NREM1 As String
    R_NREM2 As String
    R_NREM3 As String
    R_NREM4 As String
    R_NREM12 As String
    R_NREM34 As String
    R_NREMTotal As String
    R_REMTotal As String
    R_TST As String
    R_TST_SupineNREM As String
    R_TST_NonSupineNREM As String
    R_TST_SupineREM As String
    R_TST_NonSupineREM As String
    R_TST_SupineTotal As String
    R_TST_NonSupineTotal As String
    R_SLat As String
    R_REMLat As String
    R_TDT As String
    R_SlpEff As String
    R_Awakenings As String
    R_WakeTimeDuringSleep As String
    R_SpO2_BLAwake As String
    R_SpO2_BLAsleep As String
    R_SpO2_MinREM As String
    R_SpO2_MinNREM As String
    R_SpO2_Time95 As String
    R_SpO2_Time90 As String
    R_SpO2_Time88 As String
    R_SpO2_Time85 As String
    R_ABG_Conditions1 As String
    R_ABG_Conditions2 As String
    R_ABG_pH1 As String
    R_ABG_pH2 As String
    R_ABG_PaO21 As String
    R_ABG_PaO22 As String
    R_ABG_PaCO21 As String
    R_ABG_PaCO22 As String
    R_ABG_HCO31 As String
    R_ABG_HCO32 As String
    R_ABG_BE1 As String
    R_ABG_BE2 As String
    R_ABG_SaO21 As String
    R_ABG_SaO22 As String
    R_ABG_FiO21 As String
    R_ABG_FiO22 As String
    R_ABG_Time1 As String
    R_ABG_Time2 As String
    R_ABG_PtcCO21 As String
    R_ABG_PtcCO22 As String
    R_ABG_PCO2Diff1 As String
    R_ABG_PCO2Diff2 As String
    R_ABG_PtcCO2_Unit As String
    R_ABG_PtcCO2_Electrode As String
    R_ABG_PtcCO2_Drift As String
    R_ABG_SpO21 As String
    R_ABG_SpO22 As String
    R_ABG_SO2Diff1 As String
    R_ABG_SO2Diff2 As String
    R_Graphic As Boolean
    
    Est_SleepDuration As String
    Est_WakeUpTime As Variant
    Est_TimeToSleep As String
    Est_LOff As Variant
    ESS As String
    
    LastEdit As Variant
    Modified As Boolean
    StudyLog As String
    StudyPerformedBy As String
    StudyLogNote As String
    ApproximateWeight As String
    PortableDevice As String
    pdfRequest  As Byte
    RequestTime As Variant
    FeeType As String
    HccPensionNumber As String
    
    Need_NursingCare As Boolean
    Need_Mobility As Boolean
    Need_Mobility_Type As String
    Need_Interpreter As Boolean
    Need_Interpreter_Language As String
    
    ExistingDiseases  As String
    VRSSPatient As Boolean
    Request_Acknowledged As Boolean
    Request_Acknowledged_By As String
    Appt_BookedBy As String
    InfoSent As Boolean
    Emergency As Boolean
    TransportBooked As Boolean
    TransportNeeded As Boolean
    Need_Interpreter_Booked As Boolean
    ApptTime As Variant
    ApptNotes  As String
    ApptConfirmed  As Boolean
    ApptConfirmedDate As Variant
    ApptBedIndex As Boolean
End Type

Public Enum ProvocationTestType
    Error = -1
    Mannitol = 0
    Histamine = 1
    Methacholine = 2
End Enum

Public Enum sType
    Header
    Warmup
    Exercise
End Enum

Public Type DRSMContact
    ID As Long
    SourceTbl As String
    SourceTblEnum As DbTables
    ContactDate As Date
    IsClosed As Boolean
    DateClosed As Date
    Description As String
    Clinician As String
End Type
Public Type DRSMContacts
    Psgs() As DRSMContact
    RFTs() As DRSMContact
    RftsOrCpx() As DRSMContact
    Cpx() As DRSMContact
    CPAP() As DRSMContact
    Vrss_A() As DRSMContact
    Vrss_R() As DRSMContact
    Vrss_Contacts() As DRSMContact
End Type

Public Enum ABGResultSource
    RespLab
    SlpLab
    DomO2Clinic
End Enum

Public Type ABGResult
    TestDate As Date
    TestTime As Date
    Condition As String
    Source As ABGResultSource
    FiO2 As String
    pH As String
    pCO2 As String
    pO2 As String
    Shunt As String
    BE As String
    HCO3 As String
    SaO2 As String
End Type

Public Enum FileTypes
    ftNone = 0
    xls_Excel = 1
    doc_Word = 2
    docx_Word = 3
    pdf_Adobe = 4
    txt_text = 5
    jpeg_image = 5
    emf_image = 6
    rtf_text = 7
    bmp_image = 8
    ico_image = 9
    tif_image = 10
End Enum
