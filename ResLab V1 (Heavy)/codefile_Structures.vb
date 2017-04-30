
#Region "General enum definitions"

Public Enum eUnits
    Respiratory_Laboratory = 1
    Sleep_Laboratory = 2
    CPAP_Clinic = 3
    Victorian_Respiratory_Support_Service = 4
    O2_Therapy_Clinic = 5
End Enum
Public Enum eReportStatus
    Unreported = 1
    For_discussion = 2
    Reported_not_verified = 3
    Verified_not_printed = 4
    Completed = 5
End Enum
Public Enum eNameStringFormats
    Name_UR
    NameForPdfFilename
    SurnameCommaFirstname
    FirstnameSpaceSurname
End Enum
Public Enum eGenders
    Male = 1
    Female = 2
    NotKnown = 3
End Enum
Public Enum eFindBy
    Exact
    First_part
    Any_part
End Enum
Public Enum ePtNameFormats
    SurnameCommaFirstname
    FirstnameSurname
    TitleFirstnameSurname
End Enum
Public Enum ePtIdentifiers
    UR
    PatientID
    NameEtc
End Enum
Public Enum eSaveMode
    NewRecord
    UpdateRecord
End Enum
Public Enum eSortMode
    Descending
    Ascending
End Enum
Public Enum eTestReportTypes
    None = 0
    Routine = 1
    BronchChall = 2
    Walk = 3
    Hast = 4
    Spt = 5
    Sleep = 6
    Cpet = 7
    CpapClinic = 8
    SixMWD = 9
End Enum

#End Region



Public Class AllergenData
    Property allergenid As Integer = 0
    Property allergenname As String = ""
    Property categoryid As Integer = 0
    Property categoryname As String = ""
    Property displayColour As String = ""
    Property enabled As Boolean = True
End Class

Public Class PanelData
    Property panelid As Integer = 0
    Property panelname As String = ""
    Property enabled As Boolean = True
End Class

Public Class PanelMember_AllData
    Property memberid As Integer = 0
    Property panelid As Integer = 0
    Property allergenid As Integer = 0
    Property allergenname As String = ""
    Property displayColour As String = ""
    Property allergengroup As String = ""
    Property allergengroupID As Integer = 0
    Property allergenorder As Integer = 0
End Class

Public Class PanelMember_TableDataOnly
    Property memberid As Integer = 0
    Property panelid As Integer = 0
    Property allergenid As Integer = 0
    Property allergenorder As Integer = 0
End Class

Public Class AllergenCategoryData
    Public Property categoryid As Integer = 0
    Public Property categoryname As String = ""
    Public Property displaycolour As String = ""
    Public Property enabled As Boolean = True
End Class

Public Class class_fields_ur
    Public UR_id As String = "ur_id"
    Public patientID As String = "patientid"
    Public UR As String = "ur"
    Public UR_mergeto As String = "ur_mergeto"
    Public UR_hsid As String = "ur_hsid"
    Public created_inResLab As String = "created_inreslab"
    Public create_date As String = "create_date"
    Public create_by As String = "create_by"
End Class

Public Class class_pred_EquationFields
    Public ReadOnly Test As String = "test"
    Public ReadOnly TestID As String = "testid"
    Public ReadOnly Source As String = "source"
    Public ReadOnly SourceID As String = "sourceid"
    Public ReadOnly Parameter As String = "parameter"
    Public ReadOnly ParameterID As String = "parameterid"
    Public ReadOnly Gender As String = "gender"
    Public ReadOnly GenderID As String = "genderid"
    Public ReadOnly AgeGroup As String = "agegroup"
    Public ReadOnly AgeGroupID As String = "agegroupid"
    Public ReadOnly Age_lower As String = "age_lower"
    Public ReadOnly Age_upper As String = "age_upper"
    Public ReadOnly Age_clipmethod As String = "age_clipmethod"
    Public ReadOnly Age_clipmethodID As String = "age_clipmethodid"
    Public ReadOnly Ht_lower As String = "ht_lower"
    Public ReadOnly Ht_upper As String = "ht_upper"
    Public ReadOnly Ht_clipmethod As String = "ht_clipmethod"
    Public ReadOnly Ht_clipmethodID As String = "ht_clipmethodid"
    Public ReadOnly Wt_lower As String = "wt_lower"
    Public ReadOnly Wt_upper As String = "wt_upper"
    Public ReadOnly Wt_clipmethod As String = "wt_clipmethod"
    Public ReadOnly Wt_clipmethodID As String = "wt_clipmethodid"
    Public ReadOnly Ethnicity As String = "ethnicity"
    Public ReadOnly EthnicityID As String = "ethnicityid"
    Public ReadOnly Equation As String = "equation"
    Public ReadOnly EquationID As String = "equationid"
    Public ReadOnly EquationType As String = "equationtype"
    Public ReadOnly EquationTypeID As String = "equationtypeid"
    Public ReadOnly StatType As String = "stattype"
    Public ReadOnly StatTypeID As String = "stattypeid"
    Public ReadOnly EthnicityCorrectionType As String = "ethnicitycorrectiontype"
    Public ReadOnly Lastedit As String = "lastedit"
    Public ReadOnly LasteditBy As String = "lasteditby"

End Class

Public Class class_pred_SourceFields
    Public ReadOnly SourceID As String = "sourceid"
    Public ReadOnly Source As String = "Source"
    Public ReadOnly Pub_Reference As String = "Pub_Reference"
    Public ReadOnly Pub_Year As String = "Pub_Year"
    Public ReadOnly Lastedit As String = "Lastedit"
    Public ReadOnly LasteditBy As String = "LasteditBy"
End Class

Public Class class_Pref_PredFields
    Public ReadOnly PrefID As String = "prefid"
    Public ReadOnly EquationID_NormalRange As String = "equationid_normalrange"
    Public ReadOnly EquationID_mpv As String = "equationid_mpv"
    Public ReadOnly TestID As String = "testid"
    Public ReadOnly Test As String = "test"
    Public ReadOnly ParameterID As String = "parameterid"
    Public ReadOnly Parameter As String = "parameter"
    Public ReadOnly SourceID As String = "sourceid"
    Public ReadOnly Source As String = "source"
    Public ReadOnly GenderID As String = "genderid"
    Public ReadOnly Gender As String = "gender"
    Public ReadOnly AgeGroupID As String = "agegroupid"
    Public ReadOnly AgeGroup As String = "agegroup"
    Public ReadOnly Age_lower As String = "age_lower"
    Public ReadOnly Age_upper As String = "age_upper"
    Public ReadOnly Age_clipmethod As String = "age_clipmethod"
    Public ReadOnly Age_clipmethodID As String = "age_clipmethodid"
    Public ReadOnly ht_lower As String = "ht_lower"
    Public ReadOnly ht_upper As String = "ht_upper"
    Public ReadOnly ht_clipmethod As String = "ht_clipmethod"
    Public ReadOnly ht_clipmethodID As String = "ht_clipmethodid"
    Public ReadOnly wt_lower As String = "wt_lower"
    Public ReadOnly wt_upper As String = "wt_upper"
    Public ReadOnly wt_clipmethod As String = "wt_clipmethod"
    Public ReadOnly wt_clipmethodID As String = "wt_clipmethodid"
    Public ReadOnly EthnicityID As String = "ethnicityid"
    Public ReadOnly Ethnicity As String = "ethnicity"
    Public ReadOnly StartDate As String = "startdate"
    Public ReadOnly EndDate As String = "enddate"
    Public ReadOnly Lastedit As String = "lastedit"
    Public ReadOnly LasteditBy As String = "lasteditby"
End Class

Public Class class_DemographicFields
    Public ReadOnly PatientID As String = "patientid"
    Public ReadOnly Title As String = "title"
    Public ReadOnly Surname As String = "surname"
    Public ReadOnly Firstname As String = "firstname"
    Public ReadOnly UR As String = "ur"
    Public ReadOnly URandHS As String = "urandhs"
    Public ReadOnly UR_id As String = "ur_id"
    Public ReadOnly UR_hsid As String = "ur_hsid"
    Public ReadOnly URandHS_all_asString As String = "urandhs_all_asstring"
    Public ReadOnly DOB As String = "dob"
    Public ReadOnly Gender As String = "gender"
    Public ReadOnly Gender_code As String = "gender_code"
    Public ReadOnly Gender_forRfts As String = "gender_forrfts"
    Public ReadOnly Gender_forRfts_code As String = "gender_forrfts_code"
    Public ReadOnly Race As String = "race"
    Public ReadOnly Race_code As String = "race_code"
    Public ReadOnly Race_forRfts As String = "race_forrfts"
    Public ReadOnly Race_forRfts_code As String = "race_forrfts_code"
    Public ReadOnly Address_1 As String = "address_1"
    Public ReadOnly Address_2 As String = "address_2"
    Public ReadOnly Suburb As String = "suburb"
    Public ReadOnly State As String = "state"
    Public ReadOnly PostCode As String = "postcode"
    Public ReadOnly Phone_home As String = "phone_home"
    Public ReadOnly Phone_mobile As String = "phone_mobile"
    Public ReadOnly Phone_work As String = "phone_work"
    Public ReadOnly Email As String = "email"
    Public ReadOnly Medicare_No As String = "medicare_no"
    Public ReadOnly Medicare_expirydate As String = "medicare_expirydate"
    Public ReadOnly Death_date As String = "death_date"
    Public ReadOnly Death_indicator As String = "death_indicator"
    Public ReadOnly CountryOfBirth As String = "countryofbirth"
    Public ReadOnly CountryOfBirth_code As String = "countryofbirth_code"
    Public ReadOnly PreferredLanguage As String = "preferredlanguage"
    Public ReadOnly PreferredLanguage_code As String = "preferredlanguage_code"
    Public ReadOnly AboriginalStatus As String = "aboriginalstatus"
    Public ReadOnly AboriginalStatus_code As String = "aboriginalstatus_code"
    Public ReadOnly gpID As String = "gpid"
    Public ReadOnly Lastupdated_date As String = "lastupdated_date"
    Public ReadOnly Lastupdated_by As String = "lastupdated_by"
End Class

Public Class class_Fields_User
    Public ReadOnly personID As String = "personid"
    Public ReadOnly Title As String = "title"
    Public ReadOnly Surname As String = "surname"
    Public ReadOnly Firstname As String = "firstname"
    Public ReadOnly DOB As String = "dob"
    Public ReadOnly Gender As String = "gender"
    Public ReadOnly Profession_category As String = "profession_category"
    Public ReadOnly Address_1 As String = "address_1"
    Public ReadOnly Address_2 As String = "address_2"
    Public ReadOnly Suburb As String = "suburb"
    Public ReadOnly State As String = "state"
    Public ReadOnly PostCode As String = "postcode"
    Public ReadOnly Phone_home As String = "phone_home"
    Public ReadOnly Phone_mobile As String = "phone_mobile"
    Public ReadOnly Phone_work As String = "phone_work"
    Public ReadOnly Email As String = "email"
    Public ReadOnly Institution As String = "institution"
    Public ReadOnly Department As String = "department"
    Public ReadOnly User_name As String = "user_name"
    Public ReadOnly User_password As String = "user_password"
    Public ReadOnly Last_login As String = "last_login"
    Public ReadOnly Lastupdated As String = "lastupdated"
End Class

Public Class class_Fields_app_config
    Public ReadOnly configID As String = "configid"
    Public ReadOnly siteID As String = "site_id"
    Public ReadOnly site_name As String = "site_name"
    Public ReadOnly site_institution As String = "site_institution"
    Public ReadOnly site_state As String = "site_state"
    Public ReadOnly db_type As String = "db_type"
    Public ReadOnly db_name As String = "db_name"
    Public ReadOnly db_servername As String = "db_servername"
    Public ReadOnly db_connectstring As String = "db_connectstring"
    Public ReadOnly pas_mode_local As String = "pas_mode_local"
End Class

Public Class class_Rft_RoutineAndSessionFields
    'Session
    Public ReadOnly TestDate As String = "testdate"
    Public ReadOnly Height As String = "height"
    Public ReadOnly Weight As String = "weight"
    Public ReadOnly Smoke_hx As String = "smoke_hx"
    Public ReadOnly Smoke_yearssmoked As String = "smoke_yearssmoked"
    Public ReadOnly Smoke_cigsperday As String = "smoke_cigsperday"
    Public ReadOnly Smoke_packyears As String = "smoke_packyears"
    Public ReadOnly Req_date As String = "req_date"
    Public ReadOnly Req_time As String = "req_time"
    Public ReadOnly Req_name As String = "req_name"
    Public ReadOnly Req_address As String = "req_address"
    Public ReadOnly Req_phone As String = "req_phone"
    Public ReadOnly Req_fax As String = "req_fax"
    Public ReadOnly Req_email As String = "req_email"
    Public ReadOnly Req_providernumber As String = "req_providernumber"
    Public ReadOnly Report_copyto As String = "report_copyto"
    Public ReadOnly Req_clinicalnotes As String = "req_clinicalnotes"
    Public ReadOnly Req_healthservice As String = "req_healthservice"
    Public ReadOnly AdmissionStatus As String = "admissionstatus"
    Public ReadOnly Billing_billedto As String = "billing_billedto"
    Public ReadOnly Billing_billingMO As String = "billing_billingmo"
    Public ReadOnly Billing_billingMOproviderno As String = "billing_billingmoproviderno"
    Public ReadOnly Pred_SourceIDs As String = "pred_sourceids"
    Public ReadOnly LastUpdated_session As String = "lastupdated_session"
    Public ReadOnly LastUpdatedBy_session As String = "lastupdatedby_session"

    'Routine rft
    Public ReadOnly RftID As String = "rftid"
    Public ReadOnly SessionID As String = "sessionid"
    Public ReadOnly PatientID As String = "patientid"
    Public ReadOnly TestTime As String = "testtime"
    Public ReadOnly TestType As String = "testtype"
    Public ReadOnly Lab As String = "lab"
    Public ReadOnly Scientist As String = "scientist"
    Public ReadOnly BDStatus As String = "bdstatus"
    Public ReadOnly TechnicalNotes As String = "technicalnotes"
    Public ReadOnly LungVolumes_method As String = "lungvolumes_method"

    Public ReadOnly Report_text As String = "report_text"
    Public ReadOnly Report_status As String = "report_status"
    Public ReadOnly Report_reportedby As String = "report_reportedby"
    Public ReadOnly Report_reporteddate As String = "report_reporteddate"
    Public ReadOnly Report_verifiedby As String = "report_verifiedby"
    Public ReadOnly Report_verifieddate As String = "report_verifieddate"

    Public ReadOnly R_bl_Fev1 As String = "r_bl_fev1"
    Public ReadOnly R_bl_Fvc As String = "r_bl_fvc"
    Public ReadOnly R_bl_Vc As String = "r_bl_vc"
    Public ReadOnly R_Bl_Fer As String = "r_bl_fer"
    Public ReadOnly R_bl_Fef2575 As String = "r_bl_fef2575"
    Public ReadOnly R_bl_Pef As String = "r_bl_pef"
    Public ReadOnly R_Bl_Tlco As String = "r_bl_tlco"
    Public ReadOnly R_Bl_Kco As String = "r_bl_kco"
    Public ReadOnly R_Bl_Va As String = "r_bl_va"
    Public ReadOnly R_Bl_Hb As String = "r_bl_hb"
    Public ReadOnly R_Bl_Ivc As String = "r_bl_ivc"
    Public ReadOnly R_Bl_Tlc As String = "r_bl_tlc"
    Public ReadOnly R_Bl_Frc As String = "r_bl_frc"
    Public ReadOnly R_Bl_Rv As String = "r_bl_rv"
    Public ReadOnly R_Bl_LvVc As String = "r_bl_lvvc"
    Public ReadOnly R_Bl_RvTlc As String = "r_bl_rvtlc"
    Public ReadOnly R_Bl_Mip As String = "r_bl_mip"
    Public ReadOnly R_Bl_Mep As String = "r_bl_mep"

    Public ReadOnly R_Bl_FeNO As String = "r_bl_feno"
    Public ReadOnly R_post_Fev1 As String = "r_post_fev1"
    Public ReadOnly R_post_Fvc As String = "r_post_fvc"
    Public ReadOnly R_post_Vc As String = "r_post_vc"
    Public ReadOnly R_Post_Fer As String = "r_post_fer"
    Public ReadOnly R_post_Fef2575 As String = "r_post_fef2575"
    Public ReadOnly R_post_Pef As String = "r_post_pef"
    Public ReadOnly R_post_Condition As String = "r_post_condition"

    Public ReadOnly R_abg1_fio2 As String = "r_abg1_fio2"
    Public ReadOnly R_abg1_ph As String = "r_abg1_ph"
    Public ReadOnly R_abg1_pao2 As String = "r_abg1_pao2"
    Public ReadOnly R_abg1_paco2 As String = "r_abg1_paco2"
    Public ReadOnly R_abg1_sao2 As String = "r_abg1_sao2"
    Public ReadOnly R_abg1_be As String = "r_abg1_be"
    Public ReadOnly R_abg1_hco3 As String = "r_abg1_hco3"
    Public ReadOnly R_abg1_aapo2 As String = "r_abg1_aapo2"
    Public ReadOnly R_abg1_shunt As String = "r_abg1_shunt"
    Public ReadOnly R_abg2_fio2 As String = "r_abg2_fio2"
    Public ReadOnly R_abg2_ph As String = "r_abg2_ph"
    Public ReadOnly R_abg2_pao2 As String = "r_abg2_pao2"
    Public ReadOnly R_abg2_paco2 As String = "r_abg2_paco2"
    Public ReadOnly R_abg2_sao2 As String = "r_abg2_sao2"
    Public ReadOnly R_abg2_be As String = "r_abg2_be"
    Public ReadOnly R_abg2_hco3 As String = "r_abg2_hco3"
    Public ReadOnly R_abg2_aapo2 As String = "r_abg2_aapo2"
    Public ReadOnly R_abg2_shunt As String = "r_abg2_shunt"

    Public ReadOnly R_SpO2_1 As String = "r_spO2_1"
    Public ReadOnly R_SpO2_2 As String = "r_spO2_2"

    Public ReadOnly LastUpdated_rft As String = "lastupdated_rft"
    Public ReadOnly LastUpdatedBy_rft As String = "lastupdatedby_rft"
End Class

Public Class class_fields_WalkAndSession

    'Session
    Public ReadOnly TestDate As String = "testdate"
    Public ReadOnly Height As String = "height"
    Public ReadOnly Weight As String = "weight"
    Public ReadOnly Smoke_hx As String = "smoke_hx"
    Public ReadOnly Smoke_yearssmoked As String = "smoke_yearssmoked"
    Public ReadOnly Smoke_cigsperday As String = "smoke_cigsperday"
    Public ReadOnly Smoke_packyears As String = "smoke_packyears"
    Public ReadOnly Req_date As String = "req_date"
    Public ReadOnly Req_time As String = "req_time"
    Public ReadOnly Req_name As String = "req_name"
    Public ReadOnly Req_address As String = "req_address"
    Public ReadOnly Req_phone As String = "req_phone"
    Public ReadOnly Req_fax As String = "req_fax"
    Public ReadOnly Req_email As String = "req_email"
    Public ReadOnly Req_providernumber As String = "req_providernumber"
    Public ReadOnly Report_copyto As String = "report_copyto"
    Public ReadOnly Req_clinicalnotes As String = "req_clinicalnotes"
    Public ReadOnly Req_healthservice As String = "req_healthservice"
    Public ReadOnly AdmissionStatus As String = "admissionstatus"
    Public ReadOnly Billing_billedto As String = "billing_billedto"
    Public ReadOnly Billing_billingMO As String = "billing_billingmo"
    Public ReadOnly Billing_billingMOproviderno As String = "billing_billingmoproviderno"
    Public ReadOnly Pred_SourceIDs As String = "pred_sourceids"
    Public ReadOnly LastUpdated_session As String = "lastupdated_session"
    Public ReadOnly LastUpdatedBy_session As String = "lastupdatedby_session"

    'Walk
    Public ReadOnly walkID As String = "walkid"
    Public ReadOnly patientID As String = "patientid"
    Public ReadOnly sessionID As String = "sessionid"
    Public ReadOnly TestTime As String = "testtime"
    Public ReadOnly TestType As String = "testtype"
    Public ReadOnly WalkType As String = "walktype"
    Public ReadOnly ProtocolID As String = "protocolid"
    Public ReadOnly Lab As String = "lab"
    Public ReadOnly Scientist As String = "scientist"
    Public ReadOnly BDStatus As String = "bdstatus"
    Public ReadOnly TechnicalNotes As String = "technicalnotes"
    Public ReadOnly Report_text As String = "report_text"
    Public ReadOnly Report_status As String = "report_status"
    Public ReadOnly Report_reportedby As String = "report_reportedby"
    Public ReadOnly Report_reporteddate As String = "report_reporteddate"
    Public ReadOnly Report_verifiedby As String = "report_verifiedby"
    Public ReadOnly Report_verifieddate As String = "report_verifieddate"
    Public ReadOnly DiagnosticCategory As String = "diagnosticcategory"
    Public ReadOnly LastUpdated_walk As String = "lastupdated_walk"
    Public ReadOnly LastUpdatedBy_walk As String = "lastupdatedby_walk"

End Class

Public Class class_fields_Walk_Trial
    Public ReadOnly trialID As String = "trialid"
    Public ReadOnly walkID As String = "walkid"
    Public ReadOnly trial_number As String = "trial_number"
    Public ReadOnly trial_label As String = "trial_label"
    Public ReadOnly trial_distance As String = "trial_distance"
    Public ReadOnly trial_timeofday As String = "trial_timeofday"
    Public ReadOnly LastUpdated_trial As String = "lastupdated_trial"
    Public ReadOnly LastUpdatedBy_trial As String = "lastupdatedby_trial"
End Class

Public Class class_fields_Walk_TrialLevel
    Public ReadOnly levelID As String = "levelid"
    Public ReadOnly trialID As String = "trialid"
    Public ReadOnly time_label As String = "time_label"
    Public ReadOnly time_minute As String = "time_minute"
    Public ReadOnly time_speed_kph As String = "time_speed_kph"
    Public ReadOnly time_gradient As String = "time_gradient"
    Public ReadOnly time_spo2 As String = "time_spo2"
    Public ReadOnly time_hr As String = "time_hr"
    Public ReadOnly time_o2flow As String = "time_o2flow"
    Public ReadOnly time_dyspnoea As String = "time_dyspnoea"
    Public ReadOnly LastUpdated_level As String = "lastupdated_level"
    Public ReadOnly LastUpdatedBy_level As String = "lastupdatedby_level"
End Class

Public Class class_fields_ProvAndSession

    'Session
    Public ReadOnly TestDate As String = "testdate"
    Public ReadOnly Height As String = "height"
    Public ReadOnly Weight As String = "weight"
    Public ReadOnly Smoke_hx As String = "smoke_hx"
    Public ReadOnly Smoke_yearssmoked As String = "smoke_yearssmoked"
    Public ReadOnly Smoke_cigsperday As String = "smoke_cigsperday"
    Public ReadOnly Smoke_packyears As String = "smoke_packyears"
    Public ReadOnly Req_date As String = "req_date"
    Public ReadOnly Req_time As String = "req_time"
    Public ReadOnly Req_name As String = "req_name"
    Public ReadOnly Req_address As String = "req_address"
    Public ReadOnly Req_phone As String = "req_phone"
    Public ReadOnly Req_fax As String = "req_fax"
    Public ReadOnly Req_email As String = "req_email"
    Public ReadOnly Req_providernumber As String = "req_providernumber"
    Public ReadOnly Report_copyto As String = "report_copyto"
    Public ReadOnly Req_clinicalnotes As String = "req_clinicalnotes"
    Public ReadOnly Req_healthservice As String = "req_healthservice"
    Public ReadOnly AdmissionStatus As String = "admissionstatus"
    Public ReadOnly Billing_billedto As String = "billing_billedto"
    Public ReadOnly Billing_billingMO As String = "billing_billingmo"
    Public ReadOnly Billing_billingMOproviderno As String = "billing_billingmoproviderno"
    Public ReadOnly Pred_SourceIDs As String = "pred_sourceids"
    Public ReadOnly LastUpdated_session As String = "lastupdated_session"
    Public ReadOnly LastUpdatedBy_session As String = "lastupdatedby_session"

    'Prov
    Public ReadOnly provID As String = "provid"
    Public ReadOnly patientID As String = "patientid"
    Public ReadOnly sessionID As String = "sessionid"
    Public ReadOnly TestTime As String = "testtime"
    Public ReadOnly TestType As String = "testtype"
    Public ReadOnly Lab As String = "lab"
    Public ReadOnly Scientist As String = "scientist"
    Public ReadOnly BDStatus As String = "bdstatus"
    Public ReadOnly TechnicalNotes As String = "technicalnotes"
    Public ReadOnly Report_text As String = "report_text"
    Public ReadOnly Report_status As String = "report_status"
    Public ReadOnly Report_reportedby As String = "report_reportedby"
    Public ReadOnly Report_reporteddate As String = "report_reporteddate"
    Public ReadOnly Report_verifiedby As String = "report_verifiedby"
    Public ReadOnly Report_verifieddate As String = "report_verifieddate"
    Public ReadOnly R_bl_Fev1 As String = "r_bl_fev1"
    Public ReadOnly R_bl_Fvc As String = "r_bl_fvc"
    Public ReadOnly R_bl_Vc As String = "r_bl_vc"
    Public ReadOnly R_Bl_Fer As String = "r_bl_fer"
    Public ReadOnly R_bl_Fef2575 As String = "r_bl_fef2575"
    Public ReadOnly R_bl_Pef As String = "r_bl_pef"
    Public ReadOnly ProtocolID As String = "protocolid"
    Public ReadOnly Protocol_title As String = "p_title"
    Public ReadOnly Protocol_parameter As String = "p_parameter"
    Public ReadOnly Protocol_parameter_units As String = "p_parameter_units"
    Public ReadOnly Protocol_parameter_response As String = "p_parameter_response"
    Public ReadOnly Protocol_threshold As String = "pd_threshold"
    Public ReadOnly Protocol_pd_decimalplaces As String = "pd_decimalplaces"
    Public ReadOnly Protocol_dose_effect As String = "p_dose_effect"
    Public ReadOnly Protocol_doseunits As String = "p_agent_units"
    Public ReadOnly Protocol_drug As String = "p_agent"
    Public ReadOnly Protocol_post_drug As String = "p_post_drug"
    Public ReadOnly Protocol_method As String = "pd_method"
    Public ReadOnly Protocol_method_reference As String = "p_method_reference"
    Public ReadOnly Pd As String = "pd"
    Public ReadOnly plot_xscaling_type = "plot_xscaling_type"
    Public ReadOnly plot_xtitle As String = "plot_xtitle"
    Public ReadOnly plot_title As String = "plot_title"
    Public ReadOnly plot_ymin As String = "plot_ymin"
    Public ReadOnly plot_ymax As String = "plot_ymax"
    Public ReadOnly plot_ystep As String = "plot_ystep"
    Public ReadOnly LastUpdated_prov As String = "lastupdated_prov"
    Public ReadOnly LastUpdatedBy_prov As String = "lastupdatedby_prov"

End Class

Public Class class_ProvTestDataFields
    Public ReadOnly testdataid As String = "testdataid"
    Public ReadOnly provid As String = "provid"
    Public ReadOnly doseid As String = "doseid"
    Public ReadOnly dose_number As String = "dose_number"
    Public ReadOnly dose_cumulative As String = "dose_cumulative"
    Public ReadOnly dose_discrete As String = "dose_discrete"
    Public ReadOnly result As String = "result"
    Public ReadOnly response As String = "response"
    Public ReadOnly xaxis_label As String = "dose_xaxis_label"
    Public ReadOnly dose_time_min As String = "dose_time_min"
    Public ReadOnly dose_canskip As String = "dose_canskip"
End Class

Public Class class_fields_Cpet_Levels
    Public ReadOnly levelID As String = "levelid"
    Public ReadOnly cpetID As String = "cpetid"
    Public ReadOnly level_number As String = "level_number"
    Public ReadOnly level_time As String = "level_time"
    Public ReadOnly level_workload As String = "level_workload"
    Public ReadOnly level_vo2 As String = "level_vo2"
    Public ReadOnly level_vco2 As String = "level_vco2"
    Public ReadOnly level_ve As String = "level_ve"
    Public ReadOnly level_vt As String = "level_vt"
    Public ReadOnly level_spo2 As String = "level_spo2"
    Public ReadOnly level_hr As String = "level_hr"
    Public ReadOnly level_peto2 As String = "level_peto2"
    Public ReadOnly level_petco2 As String = "level_petco2"
    Public ReadOnly level_rer As String = "level_rer"
    Public ReadOnly level_vevo2 As String = "level_vevo2"
    Public ReadOnly level_vevco2 As String = "level_vevco2"
    Public ReadOnly level_o2pulse As String = "level_o2pulse"
    Public ReadOnly level_ph As String = "level_ph"
    Public ReadOnly level_paco2 As String = "level_paco2"
    Public ReadOnly level_pao2 As String = "level_pao2"
    Public ReadOnly level_hco3 As String = "level_hco3"
    Public ReadOnly level_be As String = "level_be"
    Public ReadOnly level_sao2 As String = "level_sao2"
    Public ReadOnly level_aapo2 As String = "level_aapo2"
    Public ReadOnly level_vdvt As String = "level_vdvt"
    Public ReadOnly level_bp As String = "level_bp"

End Class

Public Class class_fields_CPETAndSessionFields
    'Session
    Public ReadOnly TestDate As String = "testdate"
    Public ReadOnly Height As String = "height"
    Public ReadOnly Weight As String = "weight"
    Public ReadOnly Smoke_hx As String = "smoke_hx"
    Public ReadOnly Smoke_yearssmoked As String = "smoke_yearssmoked"
    Public ReadOnly Smoke_cigsperday As String = "smoke_cigsperday"
    Public ReadOnly Smoke_packyears As String = "smoke_packyears"
    Public ReadOnly Req_date As String = "req_date"
    Public ReadOnly Req_time As String = "req_time"
    Public ReadOnly Req_name As String = "req_name"
    Public ReadOnly Req_address As String = "req_address"
    Public ReadOnly Req_phone As String = "req_phone"
    Public ReadOnly Req_fax As String = "req_fax"
    Public ReadOnly Req_email As String = "req_email"
    Public ReadOnly Req_providernumber As String = "req_providernumber"
    Public ReadOnly Report_copyto As String = "report_copyto"
    Public ReadOnly Req_clinicalnotes As String = "req_clinicalnotes"
    Public ReadOnly Req_healthservice As String = "req_healthservice"
    Public ReadOnly AdmissionStatus As String = "admissionstatus"
    Public ReadOnly Billing_billedto As String = "billing_billedto"
    Public ReadOnly Billing_billingMO As String = "billing_billingmo"
    Public ReadOnly Billing_billingMOproviderno As String = "billing_billingmoproviderno"
    Public ReadOnly Pred_SourceIDs As String = "pred_sourceids"
    Public ReadOnly LastUpdated_session As String = "lastupdated_session"
    Public ReadOnly LastUpdatedBy_session As String = "lastupdatedby_session"

    'CPET
    Public ReadOnly patientID As String = "patientid"
    Public ReadOnly sessionID As String = "sessionid"
    Public ReadOnly cpetID As String = "cpetid"
    Public ReadOnly TestTime As String = "testtime"
    Public ReadOnly TestType As String = "testtype"
    Public ReadOnly Lab As String = "lab"
    Public ReadOnly BDStatus As String = "bdstatus"
    Public ReadOnly Scientist As String = "scientist"
    Public ReadOnly TechnicalNotes As String = "technicalnotes"
    Public ReadOnly Report_text As String = "report_text"
    Public ReadOnly Report_reportedby As String = "report_reportedby"
    Public ReadOnly Report_verifiedby As String = "report_verifiedby"
    Public ReadOnly Report_reporteddate As String = "report_reporteddate"
    Public ReadOnly Report_verifieddate As String = "report_verifieddate"
    Public ReadOnly Report_status As String = "report_status"

    Public ReadOnly r_bp_pre As String = "r_bp_pre"
    Public ReadOnly r_bp_post As String = "r_bp_post"
    Public ReadOnly r_spiro_pre_fev1 As String = "r_spiro_pre_fev1"
    Public ReadOnly r_spiro_pre_fvc As String = "r_spiro_pre_fvc"
    Public ReadOnly r_symptoms_dyspnoea_borg As String = "r_symptoms_dyspnoea_borg"
    Public ReadOnly r_symptoms_legs_borg As String = "r_symptoms_legs_borg"
    Public ReadOnly r_symptoms_other_borg As String = "r_symptoms_other_borg"
    Public ReadOnly r_symptoms_other_text As String = "r_symptoms_other_text"
    Public ReadOnly r_abg_post_fio2 As String = "r_abg_post_fio2"
    Public ReadOnly r_abg_post_ph As String = "r_abg_post_ph"
    Public ReadOnly r_abg_post_pao2 As String = "r_abg_post_pao2"
    Public ReadOnly r_abg_post_paco2 As String = "r_abg_post_paco2"
    Public ReadOnly r_abg_post_be As String = "r_abg_post_be"
    Public ReadOnly r_abg_post_hco3 As String = "r_abg_post_hco3"
    Public ReadOnly r_abg_post_sao2 As String = "r_abg_post_sao2"

    Public ReadOnly r_max_workload As String = "r_max_workload"
    Public ReadOnly r_max_ve As String = "r_max_ve"
    Public ReadOnly r_max_vo2 As String = "r_max_vo2"
    Public ReadOnly r_max_vo2kg As String = "r_max_vo2kg"
    Public ReadOnly r_max_vco2 As String = "r_max_vco2"
    Public ReadOnly r_max_hr As String = "r_max_hr"
    Public ReadOnly r_max_vt As String = "r_max_vt"
    Public ReadOnly r_max_o2pulse As String = "r_max_o2pulse"
    Public ReadOnly r_max_workload_mpv As String = "r_max_workload_mpv"
    Public ReadOnly r_max_ve_mpv As String = "r_max_ve_mpv"
    Public ReadOnly r_max_vo2_mpv As String = "r_max_vo2_mpv"
    Public ReadOnly r_max_vo2kg_mpv As String = "r_max_vo2kg_mpv"
    Public ReadOnly r_max_vco2_mpv As String = "r_max_vco2_mpv"
    Public ReadOnly r_max_hr_mpv As String = "r_max_hr_mpv"
    Public ReadOnly r_max_vt_mpv As String = "r_max_vt_mpv"
    Public ReadOnly r_max_o2pulse_mpv As String = "r_max_o2pulse_mpv"
    Public ReadOnly r_max_workload_lln As String = "r_max_workload_lln"
    Public ReadOnly r_max_ve_lln As String = "r_max_ve_lln"
    Public ReadOnly r_max_vo2_lln As String = "r_max_vo2_lln"
    Public ReadOnly r_max_vo2kg_lln As String = "r_max_vo2kg_lln"
    Public ReadOnly r_max_vco2_lln As String = "r_max_vco2_lln"
    Public ReadOnly r_max_hr_lln As String = "r_max_hr_lln"
    Public ReadOnly r_max_vt_lln As String = "r_max_vt_lln"
    Public ReadOnly r_max_o2pulse_lln As String = "r_max_o2pulse_lln"
    Public ReadOnly LastUpdated_cpet As String = "lastupdated_cpet"
    Public ReadOnly LastUpdatedBy_cpet As String = "lastupdatedby_cpet"
End Class

Public Class class_fields_Spt_Allergens
    Public ReadOnly allergenID As String = "allergenid"
    Public ReadOnly panelmemberID As String = "panelmemberid"
    Public ReadOnly sptID As String = "sptid"
    Public ReadOnly wheal_mm As String = "wheal_mm"
    Public ReadOnly flare_mm As String = "flare_mm"
    Public ReadOnly note As String = "note"
    Public ReadOnly allergen_name As String = "allergen_name"
    Public ReadOnly allergen_category_name As String = "allergen_category_name"
    Public ReadOnly allergen_category_id As String = "allergen_category_id"
    Public ReadOnly allergen_category_colour As String = "allergen_category_colour"
End Class

Public Class class_fields_SptAndSessionFields
    'Session
    Public ReadOnly TestDate As String = "testdate"
    Public ReadOnly Height As String = "height"
    Public ReadOnly Weight As String = "weight"
    Public ReadOnly Smoke_hx As String = "smoke_hx"
    Public ReadOnly Smoke_yearssmoked As String = "smoke_yearssmoked"
    Public ReadOnly Smoke_cigsperday As String = "smoke_cigsperday"
    Public ReadOnly Smoke_packyears As String = "smoke_packyears"
    Public ReadOnly Req_date As String = "req_date"
    Public ReadOnly Req_time As String = "req_time"
    Public ReadOnly Req_name As String = "req_name"
    Public ReadOnly Req_address As String = "req_address"
    Public ReadOnly Req_phone As String = "req_phone"
    Public ReadOnly Req_fax As String = "req_fax"
    Public ReadOnly Req_email As String = "req_email"
    Public ReadOnly Req_providernumber As String = "req_providernumber"
    Public ReadOnly Report_copyto As String = "report_copyto"
    Public ReadOnly Req_clinicalnotes As String = "req_clinicalnotes"
    Public ReadOnly Req_healthservice As String = "req_healthservice"
    Public ReadOnly AdmissionStatus As String = "admissionstatus"
    Public ReadOnly Billing_billedto As String = "billing_billedto"
    Public ReadOnly Billing_billingMO As String = "billing_billingmo"
    Public ReadOnly Billing_billingMOproviderno As String = "billing_billingmoproviderno"
    Public ReadOnly Pred_SourceIDs As String = "pred_sourceids"
    Public ReadOnly LastUpdated_session As String = "lastupdated_session"
    Public ReadOnly LastUpdatedBy_session As String = "lastupdatedby_session"

    'SPT
    Public ReadOnly patientID As String = "patientid"
    Public ReadOnly sessionID As String = "sessionid"
    Public ReadOnly sptID As String = "sptid"
    Public ReadOnly TestTime As String = "testtime"
    Public ReadOnly TestType As String = "testtype"
    Public ReadOnly Lab As String = "lab"
    Public ReadOnly BDStatus As String = "bdstatus"
    Public ReadOnly Scientist As String = "scientist"
    Public ReadOnly TechnicalNotes As String = "technicalnotes"
    Public ReadOnly Report_text As String = "report_text"
    Public ReadOnly Report_reportedby As String = "report_reportedby"
    Public ReadOnly Report_verifiedby As String = "report_verifiedby"
    Public ReadOnly Report_reporteddate As String = "report_reporteddate"
    Public ReadOnly Report_verifieddate As String = "report_verifieddate"
    Public ReadOnly Report_status As String = "report_status"

    Public ReadOnly panelID As String = "panelid"
    Public ReadOnly panel_name As String = "panel_name"
    Public ReadOnly site As String = "site"
    Public ReadOnly medications As String = "medications"

    Public ReadOnly LastUpdated_spt As String = "lastupdated_spt"
    Public ReadOnly LastUpdatedBy_spt As String = "lastupdatedby_spt"
End Class

Public Class class_hast_protocoldata

    Public protocolID As Integer
    Public description As String
    Public pMenuLabel As String
    Public deliverymode_fio2 As String
    Public deliverymode_suppo2 As String
    Public fio2_1 As String
    Public fio2_2 As String
    Public fio2_3 As String
    Public fio2_4 As String
    Public fio2_5 As String
    Public fio2_6 As String
    Public fio2_1_enabled As Boolean
    Public fio2_2_enabled As Boolean
    Public fio2_3_enabled As Boolean
    Public fio2_4_enabled As Boolean
    Public fio2_5_enabled As Boolean
    Public fio2_6_enabled As Boolean
    Public fio2_1_altitude As String
    Public fio2_2_altitude As String
    Public fio2_3_altitude As String
    Public fio2_4_altitude As String
    Public fio2_5_altitude As String
    Public fio2_6_altitude As String
    Public abg_enabled As Boolean
    Public abg_part_enabled As Boolean
    Public protocol_enabled As Boolean
    Public lastedited As String
    Public lasteditby As String

End Class

Public Class class_fields_Hast_Levels
    Public ReadOnly levelID As String = "levelid"
    Public ReadOnly hastID As String = "hastid"
    Public ReadOnly altitude_ft As String = "altitude_ft"
    Public ReadOnly altitude_fio2 As String = "altitude_fio2"
    Public ReadOnly suppO2_flow As String = "suppO2_flow"
    Public ReadOnly r_spo2 As String = "r_spo2"
    Public ReadOnly r_ph As String = "r_ph"
    Public ReadOnly r_paco2 As String = "r_paco2"
    Public ReadOnly r_pao2 As String = "r_pao2"
    Public ReadOnly r_hco3 As String = "r_hco3"
    Public ReadOnly r_be As String = "r_be"
    Public ReadOnly r_sao2 As String = "r_sao2"
    Public ReadOnly r_note As String = "r_note"
    Public ReadOnly row_order As String = "row_order"
End Class

Public Class class_fields_HastAndSessionFields
    'Session
    Public ReadOnly TestDate As String = "testdate"
    Public ReadOnly Height As String = "height"
    Public ReadOnly Weight As String = "weight"
    Public ReadOnly Smoke_hx As String = "smoke_hx"
    Public ReadOnly Smoke_yearssmoked As String = "smoke_yearssmoked"
    Public ReadOnly Smoke_cigsperday As String = "smoke_cigsperday"
    Public ReadOnly Smoke_packyears As String = "smoke_packyears"
    Public ReadOnly Req_date As String = "req_date"
    Public ReadOnly Req_time As String = "req_time"
    Public ReadOnly Req_name As String = "req_name"
    Public ReadOnly Req_address As String = "req_address"
    Public ReadOnly Req_phone As String = "req_phone"
    Public ReadOnly Req_fax As String = "req_fax"
    Public ReadOnly Req_email As String = "req_email"
    Public ReadOnly Req_providernumber As String = "req_providernumber"
    Public ReadOnly Report_copyto As String = "report_copyto"
    Public ReadOnly Req_clinicalnotes As String = "req_clinicalnotes"
    Public ReadOnly Req_healthservice As String = "req_healthservice"
    Public ReadOnly AdmissionStatus As String = "admissionstatus"
    Public ReadOnly Billing_billedto As String = "billing_billedto"
    Public ReadOnly Billing_billingMO As String = "billing_billingmo"
    Public ReadOnly Billing_billingMOproviderno As String = "billing_billingmoproviderno"
    Public ReadOnly Pred_SourceIDs As String = "pred_sourceids"
    Public ReadOnly LastUpdated_session As String = "lastupdated_session"
    Public ReadOnly LastUpdatedBy_session As String = "lastupdatedby_session"

    'HAST
    Public ReadOnly patientID As String = "patientid"
    Public ReadOnly sessionID As String = "sessionid"
    Public ReadOnly hastID As String = "hastid"
    Public ReadOnly TestTime As String = "testtime"
    Public ReadOnly TestType As String = "testtype"
    Public ReadOnly Lab As String = "lab"
    Public ReadOnly BDStatus As String = "bdstatus"
    Public ReadOnly Scientist As String = "scientist"
    Public ReadOnly TechnicalNotes As String = "technicalnotes"
    Public ReadOnly Report_text As String = "report_text"
    Public ReadOnly Report_reportedby As String = "report_reportedby"
    Public ReadOnly Report_verifiedby As String = "report_verifiedby"
    Public ReadOnly Report_reporteddate As String = "report_reporteddate"
    Public ReadOnly Report_verifieddate As String = "report_verifieddate"
    Public ReadOnly Report_status As String = "report_status"

    Public ReadOnly deliverymethod_fio2 As String = "deliverymethod_fio2"
    Public ReadOnly deliverymethod_suppo2 As String = "deliverymethod_suppo2"
    Public ReadOnly protocol_name As String = "protocol_name"
    Public ReadOnly protocol_id As String = "protocol_id"

    Public ReadOnly LastUpdated_hast As String = "lastupdated_hast"
    Public ReadOnly LastUpdatedBy_hast As String = "lastupdatedby_hast"
End Class



Public Class cDatabaseInfo

    Public Enum eTables     '   enum name must match the actual table name - not case sensitive
        FileLocations

        list_AboriginalStatus
        list_accesstypes
        list_AddressType
        list_DeathIndicator
        list_Gender
        list_healthservices
        list_Language
        list_Nationality
        list_permissiontypes
        List_RacesForRFTs
        list_units
        List_ReportStatuses
        List_reportstyles
        List_reporttypes

        log_logins

        pas_pt
        pas_pt_address
        pas_pt_gp
        pas_pt_names
        pas_pt_nok
        pas_pt_ur_numbers
        PatientDetails

        person_permissions
        persons

        Pred_equations
        Pred_gli_lookup
        Pred_lms_coeff
        Pred_lms_equations
        Pred_ref_agegroups
        pred_ref_clipmethods
        Pred_ref_equationtypes
        Pred_ref_ethnicities
        Pred_ref_genders
        Pred_ref_limittypes
        Pred_ref_parameters
        Pred_ref_sources
        Pred_ref_stattypes
        Pred_ref_tests
        Pred_ref_variables

        Prefs_fields
        Prefs_fielditems
        Prefs_report_parameters
        Prefs_reports_strings
        Prefs_reports_styles
        Preferences_pred

        Prov_protocol_doseschedule
        Prov_protocols
        Prov_test
        Prov_testdata

        r_sessions
        r_walktests_v1heavy
        r_walktests_trials
        r_walktests_trials_levels
        r_walktests_protocols
        r_cpet
        r_cpet_levels
        r_spt
        r_spt_allergens
        r_hast
        r_hast_levels
        rft_routine

        site_config
        site_id

        spt_panels
        spt_panelmembers
        spt_allergens
        spt_allergencategories

    End Enum

    Public Function primarykey(eTbl As eTables) As String

        Dim sql As String = "SELECT COLUMN_NAME from information_schema.KEY_COLUMN_USAGE "
        sql = sql & "WHERE CONSTRAINT_NAME LIKE 'PK%' AND TABLE_NAME='" & eTbl.ToString & "' "
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)(0)
            Else
                Return ""
            End If
        Else
            Return ""
        End If

    End Function

    Public Function table_name(eTbl As eTables) As String

        Dim tablename As String = [Enum].GetName(GetType(eTables), eTbl)
        Return tablename

    End Function

    'Public Class cTable
    '    Public Name As String
    '    Public PrimaryKey As String
    '    Public Description As String
    '    Public LookupString As String
    'End Class

    'Public Function Tablename(ByVal Tbl As eTables) As String

    '    Select Case Tbl
    '        Case eTables.site_id : Return Me.lab_id.Name
    '        Case eTables.Rft_Routine : Return Me.Rft_routine.Name
    '        Case eTables.r_sessions : Return Me.r_sessions.Name


    '        Case eTables.FileLocations : Return Me.Filelocations.Name
    '        Case eTables.PatientDetails : Return Me.Patientdetails.Name
    '        Case eTables.Prefs_fields : Return Me.Prefs_fields.Name
    '        Case eTables.Prefs_fielditems : Return Me.Prefs_fielditems.Name
    '        Case eTables.Prefs_reports_strings : Return Me.Prefs_reports_strings.Name
    '        Case eTables.Preferences_pred : Return Me.Preferences_pred.Name
    '        Case eTables.List_ReportStatuses : Return Me.List_ReportStatuses.Name
    '        Case eTables.Pred_equations : Return Me.pred_equations.Name
    '        Case eTables.Pred_ref_tests : Return Me.pred_ref_tests.Name
    '        Case eTables.Pred_ref_parameters : Return Me.pred_ref_parameters.Name
    '        Case eTables.Pred_ref_sources : Return Me.pred_ref_sources.Name
    '        Case eTables.Pred_ref_genders : Return Me.pred_ref_genders.Name
    '        Case eTables.pred_ref_clipmethods : Return Me.pred_ref_clipmethods.Name
    '        Case eTables.Pred_ref_agegroups : Return ""
    '        Case eTables.Prov_test : Return Me.prov_test.Name
    '        Case eTables.Prov_testdata : Return Me.prov_testdata.Name
    '        Case eTables.Prov_protocols : Return Me.prov_protocols.Name
    '        Case eTables.Prov_protocol_doseschedule : Return Me.prov_protocol_doseschedule.Name
    '        Case eTables.spt_panels : Return Me.spt_panels.Name
    '        Case eTables.spt_panelmembers : Return Me.spt_panelmembers.Name
    '        Case eTables.spt_allergencategories : Return Me.spt_allergencategories.Name
    '        Case eTables.spt_allergens : Return Me.spt_allergens.Name
    '        Case eTables.r_walktests_v1heavy : Return Me.r_walktests_v1heavy.Name
    '        Case eTables.r_walktests_trials : Return Me.r_walktests_trials.Name
    '        Case eTables.r_walktests_trials_levels : Return Me.r_walktests_trials_levels.Name
    '        Case eTables.r_cpet : Return Me.r_cpet.Name
    '        Case eTables.r_cpet_levels : Return Me.r_cpet_levels.Name
    '        Case eTables.r_spt : Return Me.r_spt.Name
    '        Case eTables.r_spt_allergens : Return Me.r_spt_allergens.Name
    '        Case eTables.r_hast : Return Me.r_hast.Name
    '        Case eTables.r_hast_levels : Return Me.r_hast_levels.Name
    '        Case eTables.persons : Return Me.persons.Name
    '        Case eTables.person_permissions : Return Me.person_permissions.Name

    '        Case Else : Return ""
    '    End Select

    'End Function

    'Public Function PrimaryKeyName(ByVal Tbl As eTables) As String

    '    Select Case Tbl
    '        Case eTables.site_id : Return Me.lab_id.PrimaryKey
    '        Case eTables.Rft_Routine : Return Me.Rft_routine.PrimaryKey
    '        Case eTables.r_sessions : Return Me.r_sessions.PrimaryKey


    '        Case eTables.FileLocations : Return Me.Filelocations.PrimaryKey
    '        Case eTables.PatientDetails : Return Me.Patientdetails.PrimaryKey
    '        Case eTables.Prefs_fields : Return Me.Prefs_fields.PrimaryKey
    '        Case eTables.Prefs_fielditems : Return Me.Prefs_fielditems.PrimaryKey
    '        Case eTables.Prefs_reports_strings : Return Me.Prefs_reports_strings.PrimaryKey
    '        Case eTables.Preferences_pred : Return Me.Preferences_pred.PrimaryKey
    '        Case eTables.Pred_ref_agegroups : Return Me.pred_ref_agegroups.PrimaryKey
    '        Case eTables.Pred_ref_equationtypes : Return Me.pred_ref_equationtypes.PrimaryKey
    '        Case eTables.Pred_ref_ethnicities : Return Me.pred_ref_ethnicities.PrimaryKey
    '        Case eTables.Pred_ref_genders : Return Me.pred_ref_genders.PrimaryKey
    '        Case eTables.pred_ref_clipmethods : Return Me.pred_ref_clipmethods.PrimaryKey
    '        Case eTables.Pred_ref_limittypes : Return Me.pred_ref_limittypes.PrimaryKey
    '        Case eTables.Pred_ref_parameters : Return Me.pred_ref_parameters.PrimaryKey
    '        Case eTables.Pred_ref_sources : Return Me.pred_ref_sources.PrimaryKey
    '        Case eTables.Pred_ref_stattypes : Return Me.pred_ref_stattypes.PrimaryKey
    '        Case eTables.Pred_ref_tests : Return Me.pred_ref_tests.PrimaryKey
    '        Case eTables.Pred_ref_variables : Return Me.pred_ref_variables.PrimaryKey
    '        Case eTables.List_ReportStatuses : Return Me.List_ReportStatuses.PrimaryKey
    '        Case eTables.Pred_equations : Return Me.pred_equations.PrimaryKey
    '        Case eTables.Prov_test : Return Me.prov_test.PrimaryKey
    '        Case eTables.Prov_testdata : Return Me.prov_testdata.PrimaryKey
    '        Case eTables.Prov_protocols : Return Me.prov_protocols.PrimaryKey
    '        Case eTables.Prov_protocol_doseschedule : Return Me.prov_protocol_doseschedule.PrimaryKey
    '        Case eTables.spt_panels : Return Me.spt_panels.PrimaryKey
    '        Case eTables.spt_panelmembers : Return Me.spt_panelmembers.PrimaryKey
    '        Case eTables.spt_allergencategories : Return Me.spt_allergencategories.PrimaryKey
    '        Case eTables.spt_allergens : Return Me.spt_allergens.PrimaryKey
    '        Case eTables.r_walktests_v1heavy : Return Me.r_walktests_v1heavy.PrimaryKey
    '        Case eTables.r_walktests_trials : Return Me.r_walktests_trials.PrimaryKey
    '        Case eTables.r_walktests_trials_levels : Return Me.r_walktests_trials_levels.PrimaryKey
    '        Case eTables.r_cpet : Return Me.r_cpet.PrimaryKey
    '        Case eTables.r_cpet_levels : Return Me.r_cpet_levels.PrimaryKey
    '        Case eTables.r_spt : Return Me.r_spt.PrimaryKey
    '        Case eTables.r_spt_allergens : Return Me.r_spt_allergens.PrimaryKey
    '        Case eTables.r_hast : Return Me.r_hast.PrimaryKey
    '        Case eTables.r_hast_levels : Return Me.r_hast_levels.PrimaryKey
    '        Case eTables.persons : Return Me.persons.PrimaryKey
    '        Case eTables.person_permissions : Return Me.person_permissions.PrimaryKey

    '        Case Else : Return ""
    '    End Select

    'End Function

    'Public Function DescriptionField(ByVal Tbl As eTables) As String

    '    Select Case Tbl
    '        Case eTables.site_id : Return Me.lab_id.Description
    '        Case eTables.Rft_Routine : Return Me.Rft_routine.Description
    '        Case eTables.r_sessions : Return Me.r_sessions.Description

    '        Case eTables.FileLocations : Return Me.Filelocations.Description
    '        Case eTables.PatientDetails : Return Me.Patientdetails.Description
    '        Case eTables.Prefs_fields : Return Me.Prefs_fields.Description
    '        Case eTables.Prefs_fielditems : Return Me.Prefs_fielditems.Description
    '        Case eTables.Prefs_reports_strings : Return Me.Prefs_reports_strings.Description
    '        Case eTables.Preferences_pred : Return Me.Preferences_pred.Description
    '        Case eTables.Pred_ref_agegroups : Return Me.pred_ref_agegroups.Description
    '        Case eTables.Pred_ref_equationtypes : Return Me.pred_ref_equationtypes.Description
    '        Case eTables.Pred_ref_ethnicities : Return Me.pred_ref_ethnicities.Description
    '        Case eTables.Pred_ref_genders : Return Me.pred_ref_genders.Description
    '        Case eTables.pred_ref_clipmethods : Return Me.pred_ref_clipmethods.Description
    '        Case eTables.Pred_ref_limittypes : Return Me.pred_ref_limittypes.Description
    '        Case eTables.Pred_ref_parameters : Return Me.pred_ref_parameters.Description
    '        Case eTables.Pred_ref_sources : Return Me.pred_ref_sources.Description
    '        Case eTables.Pred_ref_stattypes : Return Me.pred_ref_stattypes.Description
    '        Case eTables.Pred_ref_tests : Return Me.pred_ref_tests.Description
    '        Case eTables.Pred_ref_variables : Return Me.pred_ref_variables.Description
    '        Case eTables.List_ReportStatuses : Return Me.List_ReportStatuses.Description
    '        Case eTables.Pred_equations : Return Me.pred_equations.Description
    '        Case eTables.Prov_test : Return Me.prov_test.Description
    '        Case eTables.Prov_testdata : Return Me.prov_testdata.Description
    '        Case eTables.Prov_protocols : Return Me.prov_protocols.Description
    '        Case eTables.Prov_protocol_doseschedule : Return Me.prov_protocol_doseschedule.Description
    '        Case eTables.spt_panels : Return Me.spt_panels.Description
    '        Case eTables.spt_panelmembers : Return Me.spt_panelmembers.Description
    '        Case eTables.spt_allergencategories : Return Me.spt_allergencategories.Description
    '        Case eTables.spt_allergens : Return Me.spt_allergens.Description
    '        Case eTables.r_walktests_v1heavy : Return Me.r_walktests_v1heavy.Description
    '        Case eTables.r_walktests_trials : Return Me.r_walktests_trials.Description
    '        Case eTables.r_walktests_trials_levels : Return Me.r_walktests_trials_levels.Description
    '        Case eTables.r_cpet : Return Me.r_cpet.Description
    '        Case eTables.r_cpet_levels : Return Me.r_cpet_levels.Description
    '        Case eTables.r_spt : Return Me.r_spt.Description
    '        Case eTables.r_spt_allergens : Return Me.r_spt_allergens.Description
    '        Case eTables.r_hast : Return Me.r_hast.Description
    '        Case eTables.r_hast_levels : Return Me.r_hast_levels.Description
    '        Case eTables.persons : Return Me.persons.Description
    '        Case eTables.person_permissions : Return Me.person_permissions.Description

    '        Case Else : Return ""
    '    End Select

    'End Function

    'Public ReadOnly Property testdata_hast() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "testdata_hast"
    '        t.PrimaryKey = "hast_id"
    '        t.Description = "status"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property testdata_sixmwd() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "testdata_sixmwd"
    '        t.PrimaryKey = "sixmwd_id"
    '        t.Description = "status"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property testdata_treadmill() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "testdata_treadmill"
    '        t.PrimaryKey = "treadmill_id"
    '        t.Description = "status"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property List_ReportStatuses() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "List_ReportStatuses"
    '        t.PrimaryKey = "status_id"
    '        t.Description = "status"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property Prefs_fields() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "Prefs_fields"
    '        t.PrimaryKey = "field_id"
    '        t.Description = "fieldname"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property Prefs_fielditems() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "Prefs_fielditems"
    '        t.PrimaryKey = "prefs_id"
    '        t.Description = "fielditem"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property Prefs_reports_strings() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "Prefs_reports_strings"
    '        t.PrimaryKey = "stringid"
    '        t.Description = "text"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property challengedata() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "challengedata"
    '        t.PrimaryKey = "ChallID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property prov_protocols() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "prov_protocols"
    '        t.PrimaryKey = "protocolID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property prov_protocol_doseschedule() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "prov_protocol_doseschedule"
    '        t.PrimaryKey = "doseID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property

    'Public ReadOnly Property prov_test() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "prov_test"
    '        t.PrimaryKey = "provID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property prov_testdata() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "prov_testdata"
    '        t.PrimaryKey = "testdataID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property Filelocations() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "Filelocations"
    '        t.PrimaryKey = "FileLocID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property Rft_routine() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "Rft_routine"
    '        t.PrimaryKey = "RftID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property Patientdetails() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "PatientDetails"
    '        t.PrimaryKey = "PatientID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property lab_id() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "lab_id"
    '        t.PrimaryKey = "id"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property

    'Public ReadOnly Property Preferences_pred() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "Preferences_pred"
    '        t.PrimaryKey = "prefID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_agegroups() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_agegroups"
    '        t.PrimaryKey = "AgeGroupID"
    '        t.Description = "AgeGroup"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_clipmethods() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_clipmethods"
    '        t.PrimaryKey = "clipmethodID"
    '        t.Description = "Clipmethods"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_equationtypes() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_equationtypes"
    '        t.PrimaryKey = "equationtypeID"
    '        t.Description = "EquationtypeID"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_ethnicities() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_ethnicities"
    '        t.PrimaryKey = "EthnicityID"
    '        t.Description = "Ethnicity"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_genders() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_genders"
    '        t.PrimaryKey = "GenderID"
    '        t.Description = "Gender"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_limittypes() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_limittypes"
    '        t.PrimaryKey = "LimitTypeID"
    '        t.Description = "LimitType"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_parameters() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_parameters"
    '        t.PrimaryKey = "ParameterID"
    '        t.Description = "Parameter"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_sources() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_sources"
    '        t.PrimaryKey = "SourceID"
    '        t.Description = "Source"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_stattypes() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_stattypes"
    '        t.PrimaryKey = "StatTypeID"
    '        t.Description = "StatType"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_tests() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_tests"
    '        t.PrimaryKey = "TestID"
    '        t.Description = "Test"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_ref_variables() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_ref_variables"
    '        t.PrimaryKey = "VariableID"
    '        t.Description = "Variable"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property pred_equations() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "pred_equations"
    '        t.PrimaryKey = "equationID"
    '        t.Description = "Equation"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property spt_panels() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "spt_panels"
    '        t.PrimaryKey = "panelID"
    '        t.Description = "panelName"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property spt_panelmembers() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "spt_panelmembers"
    '        t.PrimaryKey = "memberID"
    '        t.Description = ""
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property spt_allergens() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "spt_allergens"
    '        t.PrimaryKey = "allergenID"
    '        t.Description = "allergenName"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property spt_allergencategories() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "spt_allergencategories"
    '        t.PrimaryKey = "categoryID"
    '        t.Description = "categoryName"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_walktests_v1heavy() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_walktests_v1heavy"
    '        t.PrimaryKey = "walkID"
    '        t.Description = "walktests"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_walktests_trials() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_walktests_trials"
    '        t.PrimaryKey = "trialID"
    '        t.Description = "walktests_trials"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_walktests_trials_levels() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_walktests_trials_levels"
    '        t.PrimaryKey = "levelID"
    '        t.Description = "walktests_trials_levels"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_sessions() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_sessions"
    '        t.PrimaryKey = "sessionID"
    '        t.Description = "rft_session"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_cpet() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_cpet"
    '        t.PrimaryKey = "cpetID"
    '        t.Description = "cpet test"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_cpet_levels() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_cpet_levels"
    '        t.PrimaryKey = "levelID"
    '        t.Description = "cpet raw data table"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_spt() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_spt"
    '        t.PrimaryKey = "sptID"
    '        t.Description = "spt test"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_spt_allergens() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_spt_allergens"
    '        t.PrimaryKey = "allergenID"
    '        t.Description = "spt allergen results"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_hast() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_hast"
    '        t.PrimaryKey = "hastID"
    '        t.Description = "altitude simulation test"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property r_hast_levels() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "r_hast_levels"
    '        t.PrimaryKey = "levelID"
    '        t.Description = "hast level results"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property persons() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "persons"
    '        t.PrimaryKey = "personID"
    '        t.Description = "persons"
    '        Return t
    '    End Get
    'End Property
    'Public ReadOnly Property person_permissions() As cTable
    '    Get
    '        Dim t As New cTable
    '        t.Name = "person_permissions"
    '        t.PrimaryKey = "permissionID"
    '        t.Description = "person_permissions"
    '        Return t
    '    End Get
    'End Property

End Class

