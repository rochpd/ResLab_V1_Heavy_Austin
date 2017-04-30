Option Strict Off
Option Explicit On
Imports System


Public Module module_defs

    

    Public Structure Pred_InputData
        Dim DOB As Date
        Dim Age As Single
        Dim TestDate As Date
        Dim Ht_cm As String
        Dim Wt_kg As String
        Dim GenderID As Integer
        Dim EthnicityID As Integer
        Dim EthnicityCorrectionID As Integer
        Dim ParameterID As Integer
        Dim TestID As Integer
        Dim SourceID As Integer
        Dim AgeGroupID As Integer

        Public Shared Function CreateInstance() As Pred_InputData
            Dim result As New Pred_InputData()
            result.Ht_cm = String.Empty
            result.Wt_kg = String.Empty
            Return result
        End Function
    End Structure



    Public Structure Pred_demo
        Dim Age As Single
        Dim DOB As Date
        Dim Htcm As Single
        Dim Wtkg As Single
        Dim GenderID As Integer
        Dim Gender As String
        Dim EthnicityID As Integer
        Dim Ethnicity As String
        Dim TestDate As Date
        Dim SourcesString As String
        Dim FEV1 As String
        Dim FVC As String
        Dim vo2 As String
    End Structure

    Public Enum StatType
        MPV = 1
        LLN = 2
        ULN = 3
        Range = 4
    End Enum

    Public Enum RftParameterTypes
        FEV1
        FVC
        VC
        FER
        PEF
        FEF2575
        TLCO
        KCO
        VA
        TLC
        FRC
        RV
        RV_TLC
        MIP
        MEP
        SNIP
    End Enum

    Public Enum CalcType
        Standard
        GLI
    End Enum

    Public Structure ComboItem
        Dim Item As String
        Dim ItemID As Long
    End Structure

    Public Enum RefItems
        Genders
        Ethnicities
        AgeGroups
        EquationTypes
        StatisticTypes
        ClipMethods
        Variables
        Parameters
        Tests
        Sources
    End Enum

End Module