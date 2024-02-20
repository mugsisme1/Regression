Imports Nini.Config
Public Class clsReadFIAValues

    'variables for relay.out sections
    Public Shared Header As IConfig
    Public Shared Basic As IConfig
    Public Shared Engine As IConfig

    Public Shared strFIAMDB As String = CStr(Today)
    Public Shared strFIAEXE As String = CStr(Today)
    Public Shared strFIACPY As String = CStr(Today)


    'variables for test values read from relay.out

    'Header Values

    Public Shared strFIACompanyNameTest As String
    Public Shared strFIAProdNameTest As String

    'Basic Values

    Public Shared strFIAAnnualPremiumTest As String
    Public Shared strFIAAgentTest As String

    Public Shared strFIAClient1Test As String
    Public Shared strFIAAge1Test As String
    Public Shared strFIASex1Test As String
    Public Shared strFIAClient2Test As String
    Public Shared strFIAAge2Test As String
    Public Shared strFIASex2Test As String
    Public Shared strFIADOB1Test As String
    Public Shared strFIADOB2Test As String

    Public Shared strFIAStateTest As String
    Public Shared strFIATaxStatusTest As String
    Public Shared strFIAPremiumTaxRateTest As String
    Public Shared strFIAPayoutRateCodeTest As String
    Public Shared strFIAPolicyFormTest As String
    Public Shared strFIASurrChargeYrsTest As String
    Public Shared strFIAChannelCodeTest As String
    Public Shared strFIAMaxIssueAgeTest As String
    Public Shared strFIAMinIntRateTest As String


    Public Shared strFIAGuarPeriodTest As String
    Public Shared strFIATaxBracketTest As String
    Public Shared strFIAProjectedRateTest As String

    Public Shared strFIAWithdrawalTypeTest As String
    Public Shared strFIAWithdrawalFrequencyTest As String
    

    Public Shared strFIASurrenderChargesTest As String
    
    Public Shared strFIAWDPercentTest As String
    Public Shared strFIAAnnualWDAmountTest As String
   

    Public Shared strFIAPremiumEnhancementTest As String
    Public Shared strFIAInitialBeneBaseTest As String
    Public Shared strFIABailoutAnnualCapTest As String
    Public Shared strFIARiderRollUpRateTest As String
    Public Shared strFIARiderChargeTest As String
    Public Shared strFIAAgeAtFirstWDTest As String
    Public Shared strFIAAnnWDLimitGuarTest As String
    Public Shared strFIAAnnWDLimitProjTest As String
    Public Shared strFIAOneYearFixedAllocTest As String
    Public Shared strFIASevenYearFixedAllocTest As String
    Public Shared strFIATenYearFixedAllocTest As String
    Public Shared strFIAAnnCapAllocTest As String
    Public Shared strFIAMonCapAllocTest As String
    Public Shared strFIAPerfTrigAllocTest As String
    Public Shared strFIAOneYearFixedInitialRateTest As String
    Public Shared strFIASevenYearFixedInitialRateTest As String
    Public Shared strFIATenYearFixedInitialRateTest As String
    Public Shared strFIAAnnualCapCapTest As String
    Public Shared strFIAMonthlyCapCapTest As String
    Public Shared strFIAPerfTrigSpecifiedRateTest As String
    Public Shared strFIAYearsToPrintTest As String
    Public Shared strFIASpecPeriodStartDateTest As String
    Public Shared strFIASpecPeriodEndDateTest As String
    Public Shared strFIAFavPeriodStartDateTest As String
    Public Shared strFIAFavPeriodEndDateTest As String
    Public Shared strFIAUnFavPeriodStartDateTest As String
    Public Shared strFIAUnFavPeriodEndDateTest As String
    Public Shared strFIASpecSPChangeTest As String
    Public Shared strFIASpecWDTest As String
    Public Shared strFIASpecAnnCreditRateTest As String
    Public Shared strFIASpecContractValueTest As String
    Public Shared strFIASpecSurrenderValueTest As String
    Public Shared strFIASpecMGSVTest As String
    Public Shared strFIASpecProjBeneBaseTest As String
    Public Shared strFIASpecGuarBeneBaseTest As String
    Public Shared strFIASpecProjWDLimitTest As String
    Public Shared strFIASpecGuarWDLimitTest As String
    Public Shared strFIASevenYearIntRateTest As String
    Public Shared strFIATenYearIntRateTest As String
    Public Shared strFIAMonthlyCapIndexCreditTest As String
    Public Shared strFIAAnnualCapIndexCreditTest As String
    Public Shared strFIAPerfTriggerIndexCreditTest As String
    Public Shared strFIASevenYearAccumValueTest As String
    Public Shared strFIATenYearAccumValueTest As String
    Public Shared strFIAMonthlyCapAccumValueTest As String
    Public Shared strFIAAnnualCapAccumValueTest As String
    Public Shared strFIAPerfTriggerAccumValueTest As String
    Public Shared strFIAContractValueNoWDTest As String
    Public Shared strFIAGuarWDFactorTest As String
    Public Shared strFIAGuarBeneBaseNoWDTest As String
    Public Shared strFIAGuarWDLimitNoWDTest As String
    Public Shared strFIAProjBeneBaseNoWDTest As String
    Public Shared strFIAProjWDLimitNoWDTest As String
    Public Shared strFIAAnnCreditRateNoWDTest As String
    Public Shared strFIAFavSPChangeTest As String
    Public Shared strFIAUnfavSPChangeTest As String
    Public Shared strFIAFavWDTest As String
    Public Shared strFIAUnfavWDTest As String
    Public Shared strFIAFavAnnCreditRateTest As String
    Public Shared strFIAUnfavAnnCreditRateTest As String
    Public Shared strFIAFavContractValueTest As String
    Public Shared strFIAUnfavContractValueTest As String
    Public Shared strFIAFavSurrenderValueTest As String
    Public Shared strFIAUnfavSurrenderValueTest As String
    Public Shared strFIAFavMGSVTest As String
    Public Shared strFIAUnfavMGSVTest As String
    Public Shared strFIAFavProjBeneBaseTest As String
    Public Shared strFIAFavGuarBeneBaseTest As String
    Public Shared strFIAFavProjWDLimitTest As String
    Public Shared strFIAFavGuarWDLimitTest As String
    Public Shared strFIAUnfavProjBeneBaseTest As String
    Public Shared strFIAUnfavGuarBeneBaseTest As String
    Public Shared strFIAUnfavProjWDLimitTest As String
    Public Shared strFIAUnfavGuarWDLimitTest As String
    Public Shared strFIANonForfIntRateTest As String
    Public Shared strFIAMGSVTest As String
    Public Shared strFIAFixedJumboTest As String
    Public Shared strFIAAnnCapJumboTest As String
    Public Shared strFIAMonthlyCapJumboTest As String
    Public Shared strFIAPerfTriggerJumboTest As String
    Public Shared strFIAWDFrequencyTest As String
    Public Shared strFIAGMCVTest As String



    Public Shared bErrorBench As Boolean = False
    Public Shared bErrorTest As Boolean = False

    'Engine Section

    Public Shared strStatusTest As String

    Public Shared strMessage1Test As String
    Public Shared strMessage2Test As String
    Public Shared strMessage3Test As String
    Public Shared strMessage4Test As String
    Public Shared strMessage5Test As String
    Public Shared strMessage6Test As String


    'variables for bench values read from relay.out

    'Header Values

    Public Shared strFIACompanyNameBench As String
    Public Shared strFIAProdNameBench As String

    'Basic Values

    Public Shared strFIAAnnualPremiumBench As String
  
    Public Shared strFIAAgentBench As String

    Public Shared strFIAClient1Bench As String
    Public Shared strFIAAge1Bench As String
    Public Shared strFIASex1Bench As String
    Public Shared strFIAClient2Bench As String
    Public Shared strFIAAge2Bench As String
    Public Shared strFIASex2Bench As String
    Public Shared strFIADOB1Bench As String
    Public Shared strFIADOB2Bench As String

    Public Shared strFIAStateBench As String
    Public Shared strFIATaxStatusBench As String
    Public Shared strFIAPremiumBonusBench As String
    Public Shared strFIAPremiumTaxRateBench As String

    Public Shared strFIAPayoutRateCodeBench As String
    Public Shared strFIAPolicyFormBench As String
    Public Shared strFIASurrChargeYrsBench As String
    Public Shared strFIAChannelCodeBench As String
    Public Shared strFIAMaxIssueAgeBench As String
    Public Shared strFIAMinIntRateBench As String


    Public Shared strFIAGuarPeriodBench As String
    Public Shared strFIATaxBracketBench As String
    Public Shared strFIAProjectedRateBench As String

    Public Shared strFIAWithdrawalTypeBench As String
    Public Shared strFIAWithdrawalFrequencybench As String
   
    Public Shared strFIASurrenderChargesBench As String

    Public Shared strFIAWDPercentBench As String
    Public Shared strFIAAnnualWDAmountBench As String

    Public Shared strFIAPremiumEnhancementBench As String
    Public Shared strFIAInitialBeneBaseBench As String
    Public Shared strFIABailoutAnnualCapBench As String
    Public Shared strFIARiderRollUpRateBench As String
    Public Shared strFIARiderChargeBench As String
    Public Shared strFIAAgeAtFirstWDBench As String
    Public Shared strFIAAnnWDLimitGuarBench As String
    Public Shared strFIAAnnWDLimitProjBench As String
    Public Shared strFIAOneYearFixedAllocBench As String
    Public Shared strFIASevenYearFixedAllocBench As String
    Public Shared strFIATenYearFixedAllocBench As String
    Public Shared strFIAAnnCapAllocBench As String
    Public Shared strFIAMonCapAllocBench As String
    Public Shared strFIAPerfTrigAllocBench As String
    Public Shared strFIAOneYearFixedInitialRateBench As String
    Public Shared strFIASevenYearFixedInitialRateBench As String
    Public Shared strFIATenYearFixedInitialRateBench As String
    Public Shared strFIAAnnualCapCapBench As String
    Public Shared strFIAMonthlyCapCapBench As String
    Public Shared strFIAPerfTrigSpecifiedRateBench As String
    Public Shared strFIAYearsToPrintBench As String
    Public Shared strFIASpecPeriodStartDateBench As String
    Public Shared strFIASpecPeriodEndDateBench As String
    Public Shared strFIAFavPeriodStartDateBench As String
    Public Shared strFIAFavPeriodEndDateBench As String
    Public Shared strFIAUnFavPeriodStartDateBench As String
    Public Shared strFIAUnFavPeriodEndDateBench As String
    Public Shared strFIASpecSPChangeBench As String
    Public Shared strFIASpecWDBench As String
    Public Shared strFIASpecAnnCreditRateBench As String
    Public Shared strFIASpecContractValueBench As String
    Public Shared strFIASpecSurrenderValueBench As String
    Public Shared strFIASpecMGSVBench As String
    Public Shared strFIASpecProjBeneBaseBench As String
    Public Shared strFIASpecGuarBeneBaseBench As String
    Public Shared strFIASpecProjWDLimitBench As String
    Public Shared strFIASpecGuarWDLimitBench As String
    Public Shared strFIASevenYearIntRateBench As String
    Public Shared strFIATenYearIntRateBench As String
    Public Shared strFIAMonthlyCapIndexCreditBench As String
    Public Shared strFIAAnnualCapIndexCreditBench As String
    Public Shared strFIAPerfTriggerIndexCreditBench As String
    Public Shared strFIASevenYearAccumValueBench As String
    Public Shared strFIATenYearAccumValueBench As String
    Public Shared strFIAMonthlyCapAccumValueBench As String
    Public Shared strFIAAnnualCapAccumValueBench As String
    Public Shared strFIAPerfTriggerAccumValueBench As String
    Public Shared strFIAContractValueNoWDBench As String
    Public Shared strFIAGuarWDFactorBench As String
    Public Shared strFIAGuarBeneBaseNoWDBench As String
    Public Shared strFIAGuarWDLimitNoWDBench As String
    Public Shared strFIAProjBeneBaseNoWDBench As String
    Public Shared strFIAProjWDLimitNoWDBench As String
    Public Shared strFIAAnnCreditRateNoWDBench As String
    Public Shared strFIAFavSPChangeBench As String
    Public Shared strFIAUnfavSPChangeBench As String
    Public Shared strFIAFavWDBench As String
    Public Shared strFIAUnfavWDBench As String
    Public Shared strFIAFavAnnCreditRateBench As String
    Public Shared strFIAUnfavAnnCreditRateBench As String
    Public Shared strFIAFavContractValueBench As String
    Public Shared strFIAUnfavContractValueBench As String
    Public Shared strFIAFavSurrenderValueBench As String
    Public Shared strFIAUnfavSurrenderValueBench As String
    Public Shared strFIAFavMGSVBench As String
    Public Shared strFIAUnfavMGSVBench As String
    Public Shared strFIAFavProjBeneBaseBench As String
    Public Shared strFIAFavGuarBeneBaseBench As String
    Public Shared strFIAFavProjWDLimitBench As String
    Public Shared strFIAFavGuarWDLimitBench As String
    Public Shared strFIAUnfavProjBeneBaseBench As String
    Public Shared strFIAUnfavGuarBeneBaseBench As String
    Public Shared strFIAUnfavProjWDLimitBench As String
    Public Shared strFIAUnfavGuarWDLimitBench As String
    Public Shared strFIANonForfIntRateBench As String
    Public Shared strFIAMGSVBench As String
    Public Shared strFIAFixedJumboBench As String
    Public Shared strFIAAnnCapJumboBench As String
    Public Shared strFIAMonthlyCapJumboBench As String
    Public Shared strFIAPerfTriggerJumboBench As String
    Public Shared strFIAWDFrequencyBench As String
    Public Shared strFIAGMCVBench As String

    'Engine Section

    Public Shared strStatusBench As String

    Public Shared strMessage1Bench As String
    Public Shared strMessage2Bench As String
    Public Shared strMessage3Bench As String
    Public Shared strMessage4Bench As String
    Public Shared strMessage5Bench As String
    Public Shared strMessage6Bench As String

    'variables for mismatches
    Public Shared bMisMatch As Boolean
    Public Shared bMismatchAtLeastOnce As Boolean

    'Header Values

    Public Shared strFIACompanyNameMM As String
    Public Shared strFIAProdNameMM As String

    'Basic Values

    Public Shared strFIAAnnualPremiumMM As String
    
    Public Shared strFIAClient1MM As String
    Public Shared strFIAAge1MM As String
    Public Shared strFIASex1MM As String
    Public Shared strFIAClient2MM As String
    Public Shared strFIAAge2MM As String
    Public Shared strFIASex2MM As String
    Public Shared strFIADOB1MM As String
    Public Shared strFIADOB2MM As String

    Public Shared strFIAStateMM As String
    Public Shared strFIATaxStatusMM As String
    Public Shared strFIAPremiumTaxRateMM As String
    Public Shared strFIAPolicyFormMM As String
    Public Shared strFIASurrChargeYrsMM As String
    Public Shared strFIAChannelCodeMM As String

    Public Shared strFIASurrenderChargesMM As String

    Public Shared strFIAWDPercentMM As String
    Public Shared strFIAAnnualWDAmountMM As String


    Public Shared strFIAPremiumEnhancementMM As String
    Public Shared strFIAAgeAtFirstWDMM As String
    Public Shared strFIARiderChargeMM As String
    Public Shared strFIAAnnWDLimitGuarMM As String
    Public Shared strFIAAnnWDLimitProjMM As String
    Public Shared strFIAOneYearFixedAllocMM As String
    Public Shared strFIASevenYearFixedAllocMM As String
    Public Shared strFIATenYearFixedAllocMM As String
    Public Shared strFIAAnnCapAllocMM As String
    Public Shared strFIAMonCapAllocMM As String
    Public Shared strFIAPerfTrigAllocMM As String
    Public Shared strFIAOneYearFixedInitialRateMM As String
    Public Shared strFIASevenYearFixedInitialRateMM As String
    Public Shared strFIATenYearFixedInitialRateMM As String
    Public Shared strFIAAnnualCapCapMM As String
    Public Shared strFIARiderRollupRateMM As String
    Public Shared strFIAInitialBeneBaseMM As String
    Public Shared strFIABailoutAnnualCapMM As String
    Public Shared strFIAperfTrigSpecifiedRateMM As String
    Public Shared strFIAYearsToPrintMM As String
    Public Shared strFIASpecPeriodStartDateMM As String
    Public Shared strFIASpecPeriodEndDateMM As String
    Public Shared strFIAFavPeriodStartDateMM As String
    Public Shared strFIAFavPeriodEndDateMM As String
    Public Shared strFIAUnFavPeriodStartDateMM As String
    Public Shared strFIAUnFavPeriodEndDateMM As String
    Public Shared strFIASpecSPChangeMM As String
    Public Shared strFIASpecWDMM As String
    Public Shared strFIASpecAnnCreditRateMM As String
    Public Shared strFIASpecContractValueMM As String
    Public Shared strFIASpecSurrenderValueMM As String
    Public Shared strFIAMonthlyCapCapMM As String
    Public Shared strFIASpecProjBeneBaseMM As String
    Public Shared strFIASpecGuarBeneBaseMM As String
    Public Shared strFIASpecMGSVMM As String
    Public Shared strFIASpecProjWDLimitMM As String
    Public Shared strFIASevenYearIntRateMM As String
    Public Shared strFIATenYearIntRateMM As String
    Public Shared strFIAMonthlyCapIndexCreditMM As String
    Public Shared strFIASpecGuarWDLimitMM As String
    Public Shared strFIASevenYearAccumValueMM As String
    Public Shared strFIAAnnualCapIndexCreditMM As String
    Public Shared strFIAPerfTriggerIndexCreditMM As String
    Public Shared strFIATenYearAccumValueMM As String
    Public Shared strFIAMonthlyCapAccumValueMM As String
    Public Shared strFIAAnnualCapAccumValueMM As String
    Public Shared strFIAContractValueNoWDMM As String
    Public Shared strFIAGuarBeneBaseNoWDMM As String
    Public Shared strFIAPerfTriggerAccumValueMM As String
    Public Shared strFIAProjBeneBAseNoWDMM As String
    Public Shared strFIAProjWDLimitNoWDMM As String
    Public Shared strFIAAnnCreditRateNoWDMM As String
    Public Shared strFIAGuarWDFactorMM As String
    Public Shared strFIAFavSPChangeMM As String
    Public Shared strFIAUnFavSPChangeMM As String
    Public Shared strFIAGuarWDLimitNoWDMM As String
    Public Shared strFIAFavWDMM As String
    Public Shared strFIAFavAnnCreditRateMM As String
    Public Shared strFIAFavContractValueMM As String
    Public Shared strFIAUnfavWDMM As String
    Public Shared strFIAUnfavAnnCreditRateMM As String
    Public Shared strFIAUnfavContractValueMM As String
    Public Shared strFIAUnfavSurrenderValueMM As String
    Public Shared strFIAFavMGSVMM As String
    Public Shared strFIAUnfavMGSVMM As String
    Public Shared strFIAFavProjBeneBaseMM As String
    Public Shared strFIAFavSurrenderValueMM As String
    Public Shared strFIAUnfavProjBeneBaseMM As String
    Public Shared strFIAFavGuarBeneBaseMM As String
    Public Shared strFIAFavProjWDLimitMM As String
    Public Shared strFIAUnfavGuarBeneBaseMM As String
    Public Shared strFIAUnfavProjWDLimitMM As String
    Public Shared strFIAFavGuarWDLimitMM As String
    Public Shared strFIAUnfavGuarWDLimitMM As String
    Public Shared strFIAFixedJumboMM As String
    Public Shared strFIAAnnCapJumboMM As String
    Public Shared strFIAMonthlyCapJumboMM As String



    Public Shared strFIAPerfTriggerJumboMM As String
    Public Shared strFIAWDFrequencyMM As String
    Public Shared strFIANonForfIntRateMM As String


    Public Shared strFIAMGSVMM As String

    Public Shared strFIAGMCVMM As String
   
    Public Shared strFIARenewalSurrChargesMM As String


    Public Shared bErrorMM As Boolean = False

    'Engine Section

    Public Shared strStatusMM As String

    Public Shared strMessage1MM As String
    Public Shared strMessage2MM As String
    Public Shared strMessage3MM As String
    Public Shared strMessage4MM As String
    Public Shared strMessage5MM As String
    Public Shared strMessage6MM As String


    Public Shared strFIAstreamcountMM As String

    Public Shared strRunNoRunMM As String

    Public Shared strClientXMisMatch(150) As String

    Public Sub ReadFIATestValues(ByVal strTestPath As String)

        'read the values from the test relay.out files

        Dim RelayOutTest = New IniConfigSource(strTestPath)
        Dim i As Integer
        Dim icomma As Integer = 0

        'variables for relay.out sections
        Header = RelayOutTest.Configs("HeaderValues")
        Basic = RelayOutTest.Configs("BasicValues")
        Engine = RelayOutTest.Configs("Engine")

        strStatusTest = Engine.Get("Status")

        'if case does not run
        If strStatusTest <> "0" Then
            Regression.RegressionMain.gbClientDoesntRun = True
            bErrorTest = True
            strMessage1Test = Engine.Get("Message_1")
            If strMessage1Test = Nothing Then strMessage1Test = ""
            If strMessage1Test <> "" Then
                If InStr(strMessage1Test, "WARNING") Then
                    strMessage1Test = strMessage1Test.Remove(7, 2)
                    strMessage1Test = strMessage1Test.Insert(7, "  ")
                Else
                    strMessage1Test = strMessage1Test.Remove(5, 2)
                    strMessage1Test = strMessage1Test.Insert(5, "  ")
                End If
                For i = 0 To strMessage1Test.Length - 1
                    If strMessage1Test.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage1Test = strMessage1Test.Remove(i, 1)
                        strMessage1Test = strMessage1Test.Insert(0, "0")
                    End If
                Next
                strMessage1Test = strMessage1Test.Remove(0, icomma)
                icomma = 0
                strMessage2Test = Engine.Get("Message_2")
                If strMessage2Test = Nothing Then strMessage2Test = ""
            End If
            If strMessage2Test <> "" Then
                If InStr(strMessage2Test, "WARNING") Then
                    strMessage2Test = strMessage2Test.Remove(7, 2)
                    strMessage2Test = strMessage2Test.Insert(7, "  ")
                Else
                    strMessage2Test = strMessage2Test.Remove(5, 2)
                    strMessage2Test = strMessage2Test.Insert(5, "  ")
                End If
                For i = 0 To strMessage2Test.Length - 1
                    If strMessage2Test.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage2Test = strMessage2Test.Remove(i, 1)
                        strMessage2Test = strMessage2Test.Insert(0, "0")
                    End If
                Next
                strMessage2Test = strMessage2Test.Remove(0, icomma)
                icomma = 0
                strMessage3Test = Engine.Get("Message_3")
                If strMessage3Test = Nothing Then strMessage3Test = ""
            End If
            If strMessage3Test <> "" Then
                If InStr(strMessage3Test, "WARNING") Then
                    strMessage3Test = strMessage3Test.Remove(7, 2)
                    strMessage3Test = strMessage3Test.Insert(7, "  ")
                Else
                    strMessage3Test = strMessage3Test.Remove(5, 2)
                    strMessage3Test = strMessage3Test.Insert(5, "  ")
                End If
                For i = 0 To strMessage3Test.Length - 1
                    If strMessage3Test.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage3Test = strMessage3Test.Remove(i, 1)
                        strMessage3Test = strMessage3Test.Insert(0, "0")
                    End If
                Next
                strMessage3Test = strMessage3Test.Remove(0, icomma)
                icomma = 0
                strMessage4Test = Engine.Get("Message_4")
                If strMessage4Test = Nothing Then strMessage4Test = ""
            End If

            If strMessage4Test <> "" Then
                If InStr(strMessage4Test, "WARNING") Then
                    strMessage4Test = strMessage4Test.Remove(7, 2)
                    strMessage4Test = strMessage4Test.Insert(7, "  ")
                Else
                    strMessage4Test = strMessage4Test.Remove(5, 2)
                    strMessage4Test = strMessage4Test.Insert(5, "  ")
                End If
                For i = 0 To strMessage4Test.Length - 1
                    If strMessage4Test.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage4Test = strMessage4Test.Remove(i, 1)
                        strMessage4Test = strMessage4Test.Insert(0, "0")
                    End If
                Next
                strMessage4Test = strMessage4Test.Remove(0, icomma)
                icomma = 0
                strMessage5Test = Engine.Get("Message_5")
                If strMessage5Test = Nothing Then strMessage5Test = ""
            End If
            If strMessage5Test <> "" Then
                If InStr(strMessage5Test, "WARNING") Then
                    strMessage5Test = strMessage5Test.Remove(7, 2)
                    strMessage5Test = strMessage5Test.Insert(7, "  ")
                Else
                    strMessage5Test = strMessage5Test.Remove(5, 2)
                    strMessage5Test = strMessage5Test.Insert(5, "  ")
                End If
                For i = 0 To strMessage5Test.Length - 1
                    If strMessage5Test.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage5Test = strMessage5Test.Remove(i, 1)
                        strMessage5Test = strMessage5Test.Insert(0, "0")
                    End If
                Next
                strMessage5Test = strMessage5Test.Remove(0, icomma)
                icomma = 0
                strMessage6Test = Engine.Get("Message_6")
                If strMessage6Test = Nothing Then strMessage6Test = ""
                If strMessage6Test <> "" Then
                    If InStr(strMessage6Test, "WARNING") Then
                        strMessage6Test = strMessage6Test.Remove(7, 2)
                        strMessage6Test = strMessage6Test.Insert(7, "  ")
                    Else
                        strMessage6Test = strMessage6Test.Remove(5, 2)
                        strMessage6Test = strMessage6Test.Insert(5, "  ")
                    End If
                    For i = 0 To strMessage6Test.Length - 1
                        If strMessage6Test.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage6Test = strMessage6Test.Remove(i, 1)
                            strMessage6Test = strMessage6Test.Insert(0, "0")
                        End If
                    Next
                    strMessage6Test = strMessage5Test.Remove(0, icomma)
                    icomma = 0
                End If
            End If

            'if case does run
        Else
            strMessage1Test = Engine.Get("Message_1")
            If strMessage1Test = Nothing Then strMessage1Test = ""
            If strMessage1Test <> "" Then
                strMessage1Test = strMessage1Test.Remove(7, 2)
                strMessage1Test = strMessage1Test.Insert(7, "  ")
                For i = 0 To strMessage1Test.Length - 1
                    If strMessage1Test.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage1Test = strMessage1Test.Remove(i, 1)
                        strMessage1Test = strMessage1Test.Insert(0, "0")
                    End If
                Next
                strMessage1Test = strMessage1Test.Remove(0, icomma)
                icomma = 0
            End If
            If strMessage1Test <> "" Then
                strMessage2Test = Engine.Get("Message_2")
                If strMessage2Test = Nothing Then strMessage2Test = ""
                If strMessage2Test <> "" Then
                    strMessage2Test = strMessage2Test.Remove(7, 2)
                    strMessage2Test = strMessage2Test.Insert(7, "  ")
                    For i = 0 To strMessage2Test.Length - 1
                        If strMessage2Test.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage2Test = strMessage2Test.Remove(i, 1)
                            strMessage2Test = strMessage2Test.Insert(0, "0")
                        End If
                    Next
                    strMessage2Test = strMessage2Test.Remove(0, icomma)
                    icomma = 0
                End If
            End If
            If strMessage2Test <> "" Then
                strMessage3Test = Engine.Get("Message_3")
                If strMessage3Test = Nothing Then strMessage3Test = ""
                If strMessage3Test <> "" Then
                    strMessage3Test = strMessage3Test.Remove(7, 2)
                    strMessage3Test = strMessage3Test.Insert(7, "  ")
                    For i = 0 To strMessage3Test.Length - 1
                        If strMessage3Test.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage3Test = strMessage3Test.Remove(i, 1)
                            strMessage3Test = strMessage3Test.Insert(0, "0")
                        End If
                    Next
                    strMessage3Test = strMessage3Test.Remove(0, icomma)
                    icomma = 0
                End If
            End If
            If strMessage3Test <> "" Then
                strMessage4Test = Engine.Get("Message_4")
                If strMessage4Test = Nothing Then strMessage4Test = ""
                If strMessage4Test <> "" Then
                    strMessage4Test = strMessage4Test.Remove(7, 2)
                    strMessage4Test = strMessage4Test.Insert(7, "  ")
                    For i = 0 To strMessage4Test.Length - 1
                        If strMessage4Test.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage4Test = strMessage4Test.Remove(i, 1)
                            strMessage4Test = strMessage4Test.Insert(0, "0")
                        End If
                    Next
                    strMessage4Test = strMessage4Test.Remove(0, icomma)
                    icomma = 0
                End If
            End If

            If strMessage4Test <> "" Then
                strMessage5Test = Engine.Get("Message_5")
                If strMessage5Test = Nothing Then strMessage5Test = ""
                If strMessage5Test <> "" Then
                    strMessage5Test = strMessage5Test.Remove(7, 2)
                    strMessage5Test = strMessage5Test.Insert(7, "  ")
                    For i = 0 To strMessage5Test.Length - 1
                        If strMessage5Test.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage5Test = strMessage5Test.Remove(i, 1)
                            strMessage5Test = strMessage5Test.Insert(0, "0")
                        End If
                    Next
                    strMessage5Test = strMessage5Test.Remove(0, icomma)
                    icomma = 0
                End If
            End If
            If strMessage5Test <> "" Then
                strMessage6Test = Engine.Get("Message_6")
                If strMessage6Test = Nothing Then strMessage6Test = ""
                If strMessage6Test <> "" Then
                    strMessage6Test = strMessage6Test.Remove(7, 2)
                    strMessage6Test = strMessage6Test.Insert(7, "  ")
                    For i = 0 To strMessage6Test.Length - 1
                        If strMessage6Test.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage6Test = strMessage6Test.Remove(i, 1)
                            strMessage6Test = strMessage6Test.Insert(0, "0")
                        End If
                    Next
                    strMessage6Test = strMessage6Test.Remove(0, icomma)
                    icomma = 0
                End If
            End If
        End If


        'Header Values
        strFIACompanyNameTest = Header.Get("bridge_CompanyName")

        'Basic Values

        strFIAProdNameTest = Basic.Get("value_ProdName")

        If strFIAProdNameTest = "FIA7L" Then
            strFIAGMCVTest = Basic.Get("value_GMCV")
        End If
        strFIAAnnualPremiumTest = Basic.Get("value_Premium")
        strFIAPremiumEnhancementTest = Basic.Get("value_PremiumEnhancement")
        strFIAInitialBeneBaseTest = Basic.Get("value_InitialBenefitBase")
        strFIABailoutAnnualCapTest = Basic.Get("value_BailoutAnnualCap")
        strFIARiderRollUpRateTest = Basic.Get("value_RiderRollUpRate")
        strFIARiderChargeTest = Basic.Get("value_RiderCharge")
        strFIAAgeAtFirstWDTest = Basic.Get("value_AgeAtFirstWD")
        strFIAAnnWDLimitGuarTest = Basic.Get("value_AnnWDLimitGuar")
        strFIAAnnWDLimitProjTest = Basic.Get("value_AnnWDLimitProj")
        strFIAOneYearFixedAllocTest = Basic.Get("value_OneYearFixedAllocation")
        strFIASevenYearFixedAllocTest = Basic.Get("value_SevenYearFixedAllocation")
        strFIATenYearFixedAllocTest = Basic.Get("value_TenYearFixedAllocation")
        strFIAAnnCapAllocTest = Basic.Get("value_AnnualCapAllocation")
        strFIAMonCapAllocTest = Basic.Get("value_MonthlyCapAllocation")
        strFIAPerfTrigAllocTest = Basic.Get("value_PerfTriggerAllocation")
        strFIAOneYearFixedInitialRateTest = Basic.Get("value_OneYearFixedInitialRate")
        strFIASevenYearFixedInitialRateTest = Basic.Get("value_SevenYearFixedInitialRate")
        strFIATenYearFixedInitialRateTest = Basic.Get("value_TenYearFixedInitialRate")
        strFIAAnnualCapCapTest = Basic.Get("value_AnnualCapCap")
        strFIAMonthlyCapCapTest = Basic.Get("value_MonthlyCapCap")
        strFIAPerfTrigSpecifiedRateTest = Basic.Get("value_PerfTriggerSpecifiedRate")
        strFIAYearsToPrintTest = Basic.Get("value_YearsToPrint")
        strFIASpecPeriodStartDateTest = Basic.Get("value_SpecPeriodStartDate")
        strFIASpecPeriodEndDateTest = Basic.Get("value_SpecPeriodEndDate")
        strFIAFavPeriodStartDateTest = Basic.Get("value_FavPeriodStartDate")
        strFIAFavPeriodEndDateTest = Basic.Get("value_FavPeriodEndDate")
        strFIAUnFavPeriodStartDateTest = Basic.Get("value_UnfavPeriodStartDate")
        strFIAUnFavPeriodEndDateTest = Basic.Get("value_UnfavPeriodEndDate")
        strFIASpecSPChangeTest = Basic.Get("value_SpecSPChange")
        strFIASpecWDTest = Basic.Get("value_SpecWD")
        strFIASpecAnnCreditRateTest = Basic.Get("value_SpecAnnCreditRate")
        strFIASpecContractValueTest = Basic.Get("value_SpecContractValue")
        strFIASpecSurrenderValueTest = Basic.Get("value_SpecSurrenderValue")
        strFIASpecMGSVTest = Basic.Get("value_SpecMGSV")
        strFIASpecProjBeneBaseTest = Basic.Get("value_SpecProjBeneBase")
        strFIASpecGuarBeneBaseTest = Basic.Get("value_SpecGuarBeneBase")
        strFIASpecProjWDLimitTest = Basic.Get("value_SpecProjWDLimit")
        strFIASpecGuarWDLimitTest = Basic.Get("value_SpecGuarWDLimit")
        strFIASevenYearIntRateTest = Basic.Get("value_SevenYearIntRate")
        strFIATenYearIntRateTest = Basic.Get("value_TenYearIntRate")
        strFIAMonthlyCapIndexCreditTest = Basic.Get("value_MonthlyCapIndexCredit")
        strFIAAnnualCapIndexCreditTest = Basic.Get("value_AnnualCapIndexCredit")
        strFIAPerfTriggerIndexCreditTest = Basic.Get("value_PerfTriggerIndexCredit")
        strFIASevenYearAccumValueTest = Basic.Get("value_SevenYearAccumValue")
        strFIATenYearAccumValueTest = Basic.Get("value_TenYearAccumValue")
        strFIAMonthlyCapAccumValueTest = Basic.Get("value_MonthlyCapAccumValue")
        strFIAAnnualCapAccumValueTest = Basic.Get("value_AnnualCapAccumValue")
        strFIAPerfTriggerAccumValueTest = Basic.Get("value_PerfTriggerAccumValue")
        strFIAContractValueNoWDTest = Basic.Get("value_ContractValueNoWD")
        strFIAGuarWDFactorTest = Basic.Get("value_GuarWDFactor")
        strFIAGuarBeneBaseNoWDTest = Basic.Get("value_GuarBeneBaseNoWD")
        strFIAGuarWDLimitNoWDTest = Basic.Get("value_GuarWDLimitNoWD")
        strFIAProjBeneBaseNoWDTest = Basic.Get("value_ProjBeneBaseNoWD")
        strFIAProjWDLimitNoWDTest = Basic.Get("value_ProjWDLimitNoWD")
        strFIAFavSPChangeTest = Basic.Get("value_FavSPChange")
        strFIAUnfavSPChangeTest = Basic.Get("value_UnfavSPChange")
        strFIAFavWDTest = Basic.Get("value_FavWD")
        strFIAUnfavWDTest = Basic.Get("value_UnfavWD")
        strFIAFavAnnCreditRateTest = Basic.Get("value_FavAnnCreditRate")
        strFIAUnfavAnnCreditRateTest = Basic.Get("value_UnfavAnnCreditRate")
        strFIAFavContractValueTest = Basic.Get("value_FavContractValue")
        strFIAUnfavContractValueTest = Basic.Get("value_UnfavContractValue")
        strFIAFavSurrenderValueTest = Basic.Get("value_FavSurrenderValue")
        strFIAUnfavSurrenderValueTest = Basic.Get("value_UnfavSurrenderValue")
        strFIAFavMGSVTest = Basic.Get("value_FavMGSV")
        strFIAUnfavMGSVTest = Basic.Get("value_UnfavMGSV")
        strFIAFavProjBeneBaseTest = Basic.Get("value_FavProjBeneBase")
        strFIAFavGuarBeneBaseTest = Basic.Get("value_FavGuarBeneBase")
        strFIAFavProjWDLimitTest = Basic.Get("value_FavProjWDLimit")
        strFIAFavGuarWDLimitTest = Basic.Get("value_FavGuarWDLimit")
        strFIAUnfavProjBeneBaseTest = Basic.Get("value_UnfavProjBeneBase")
        strFIAUnfavGuarBeneBaseTest = Basic.Get("value_UnfavGuarBeneBase")
        strFIAUnfavProjWDLimitTest = Basic.Get("value_UnfavProjWDLimit")
        strFIAUnfavGuarWDLimitTest = Basic.Get("value_UnfavGuarWDLimit")
        strFIANonForfIntRateTest = Basic.Get("value_FIANonForfIntRate")
        strFIAMGSVTest = Basic.Get("value_FIAMGSV")
        strFIAFixedJumboTest = Basic.Get("value_FixedJumbo")
        strFIAAnnCapJumboTest = Basic.Get("value_AnnCapJumbo")
        strFIAMonthlyCapJumboTest = Basic.Get("value_MonthlyCapJumbo")
        strFIAPerfTriggerJumboTest = Basic.Get("value_PerfTriggerJumbo")
        strFIAWDFrequencyTest = Basic.Get("value_WDFrequency")
        strFIAAnnCreditRateNoWDTest = Basic.Get("value_AnnCreditRateNoWD")

        strFIAClient1Test = Basic.Get("value_Name1")
        strFIAAge1Test = Basic.Get("value_Age1")
        strFIASex1Test = Basic.Get("value_Sex1")
        strFIAClient2Test = Basic.Get("value_Name2")
        strFIAAge2Test = Basic.Get("value_Age2")
        strFIASex2Test = Basic.Get("value_Sex2")
        strFIADOB1Test = Basic.Get("value_DOB1")
        strFIADOB2Test = Basic.Get("value_DOB2")

        strFIAStateTest = Basic.Get("value_State")
        strFIATaxStatusTest = Basic.Get("value_TaxStatus")

        strFIAPremiumTaxRateTest = Basic.Get("value_PremTaxRate")
        strFIAPolicyFormTest = Basic.Get("value_PolicyForm")
        strFIASurrChargeYrsTest = Basic.Get("value_SurrChargeYrs")
        strFIAChannelCodeTest = Basic.Get("value_ChannelCode")



        strFIASurrenderChargesTest = Basic.Get("value_SurrenderCharges")



    End Sub


    Public Sub ReadFIABenchValues(ByVal strBenchPath As String)

        'read the values from the Bench relay.out files

        Dim RelayOutBench = New IniConfigSource(strBenchPath)
        Dim i As Integer
        Dim icomma As Integer = 0

        'variables for relay.out sections
        Header = RelayOutBench.Configs("HeaderValues")
        Basic = RelayOutBench.Configs("BasicValues")
        Engine = RelayOutBench.Configs("Engine")

        strStatusBench = Engine.Get("Status")

        'if case does not run
        If strStatusBench <> "0" Then
            bErrorBench = True
            strMessage1Bench = Engine.Get("Message_1")
            If strMessage1Bench = Nothing Then strMessage1Bench = ""
            If strMessage1Bench <> "" Then
                If InStr(strMessage1Bench, "WARNING") Then
                    strMessage1Bench = strMessage1Bench.Remove(7, 2)
                    strMessage1Bench = strMessage1Bench.Insert(7, "  ")
                Else
                    strMessage1Bench = strMessage1Bench.Remove(5, 2)
                    strMessage1Bench = strMessage1Bench.Insert(5, "  ")
                End If
                For i = 0 To strMessage1Bench.Length - 1
                    If strMessage1Bench.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage1Bench = strMessage1Bench.Remove(i, 1)
                        strMessage1Bench = strMessage1Bench.Insert(0, "0")
                    End If
                Next
                strMessage1Bench = strMessage1Bench.Remove(0, icomma)
                icomma = 0
                strMessage2Bench = Engine.Get("Message_2")
                If strMessage2Bench = Nothing Then strMessage2Bench = ""
            End If
            If strMessage2Bench <> "" Then
                If InStr(strMessage2Bench, "WARNING") Then
                    strMessage2Bench = strMessage2Bench.Remove(7, 2)
                    strMessage2Bench = strMessage2Bench.Insert(7, "  ")
                Else
                    strMessage2Bench = strMessage2Bench.Remove(5, 2)
                    strMessage2Bench = strMessage2Bench.Insert(5, "  ")
                End If
                For i = 0 To strMessage2Bench.Length - 1
                    If strMessage2Bench.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage2Bench = strMessage2Bench.Remove(i, 1)
                        strMessage2Bench = strMessage2Bench.Insert(0, "0")
                    End If
                Next
                strMessage2Bench = strMessage2Bench.Remove(0, icomma)
                icomma = 0
                strMessage3Bench = Engine.Get("Message_3")
                If strMessage3Bench = Nothing Then strMessage3Bench = ""
            End If
            If strMessage3Bench <> "" Then
                If InStr(strMessage3Bench, "WARNING") Then
                    strMessage3Bench = strMessage3Bench.Remove(7, 2)
                    strMessage3Bench = strMessage3Bench.Insert(7, "  ")
                Else
                    strMessage3Bench = strMessage3Bench.Remove(5, 2)
                    strMessage3Bench = strMessage3Bench.Insert(5, "  ")
                End If
                For i = 0 To strMessage3Bench.Length - 1
                    If strMessage3Bench.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage3Bench = strMessage3Bench.Remove(i, 1)
                        strMessage3Bench = strMessage3Bench.Insert(0, "0")
                    End If
                Next
                strMessage3Bench = strMessage3Bench.Remove(0, icomma)
                icomma = 0
                strMessage4Bench = Engine.Get("Message_4")
                If strMessage4Bench = Nothing Then strMessage4Bench = ""
            End If

            If strMessage4Bench <> "" Then
                If InStr(strMessage4Bench, "WARNING") Then
                    strMessage4Bench = strMessage4Bench.Remove(7, 2)
                    strMessage4Bench = strMessage4Bench.Insert(7, "  ")
                Else
                    strMessage4Bench = strMessage4Bench.Remove(5, 2)
                    strMessage4Bench = strMessage4Bench.Insert(5, "  ")
                End If
                For i = 0 To strMessage4Bench.Length - 1
                    If strMessage4Bench.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage4Bench = strMessage4Bench.Remove(i, 1)
                        strMessage4Bench = strMessage4Bench.Insert(0, "0")
                    End If
                Next
                strMessage4Bench = strMessage4Bench.Remove(0, icomma)
                icomma = 0
                strMessage5Bench = Engine.Get("Message_5")
                If strMessage5Bench = Nothing Then strMessage5Bench = ""
            End If
            If strMessage5Bench <> "" Then
                If InStr(strMessage5Bench, "WARNING") Then
                    strMessage5Bench = strMessage5Bench.Remove(7, 2)
                    strMessage5Bench = strMessage5Bench.Insert(7, "  ")
                Else
                    strMessage5Bench = strMessage5Bench.Remove(5, 2)
                    strMessage5Bench = strMessage5Bench.Insert(5, "  ")
                End If
                For i = 0 To strMessage5Bench.Length - 1
                    If strMessage5Bench.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage5Bench = strMessage5Bench.Remove(i, 1)
                        strMessage5Bench = strMessage5Bench.Insert(0, "0")
                    End If
                Next
                strMessage5Bench = strMessage5Bench.Remove(0, icomma)
                icomma = 0
                strMessage6Bench = Engine.Get("Message_6")
                If strMessage6Bench = Nothing Then strMessage6Bench = ""
                If strMessage6Bench <> "" Then
                    If InStr(strMessage6Bench, "WARNING") Then
                        strMessage6Bench = strMessage6Bench.Remove(7, 2)
                        strMessage6Bench = strMessage6Bench.Insert(7, "  ")
                    Else
                        strMessage6Bench = strMessage6Bench.Remove(5, 2)
                        strMessage6Bench = strMessage6Bench.Insert(5, "  ")
                    End If
                    For i = 0 To strMessage6Bench.Length - 1
                        If strMessage6Bench.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage6Bench = strMessage6Bench.Remove(i, 1)
                            strMessage6Bench = strMessage6Bench.Insert(0, "0")
                        End If
                    Next
                    strMessage6Bench = strMessage5Bench.Remove(0, icomma)
                    icomma = 0
                End If
            End If

        Else

            'if case does run

            strMessage1Bench = Engine.Get("Message_1")
            If strMessage1Bench = Nothing Then strMessage1Bench = ""
            If strMessage1Bench <> "" Then
                strMessage1Bench = strMessage1Bench.Remove(7, 2)
                strMessage1Bench = strMessage1Bench.Insert(7, "  ")
                For i = 0 To strMessage1Bench.Length - 1
                    If strMessage1Bench.Substring(i, 1) = "," Then
                        icomma = icomma + 1
                        strMessage1Bench = strMessage1Bench.Remove(i, 1)
                        strMessage1Bench = strMessage1Bench.Insert(0, "0")
                    End If
                Next
                strMessage1Bench = strMessage1Bench.Remove(0, icomma)
                icomma = 0
            End If
            If strMessage1Bench <> "" Then
                strMessage2Bench = Engine.Get("Message_2")
                If strMessage2Bench = Nothing Then strMessage2Bench = ""
                If strMessage2Bench <> "" Then
                    strMessage2Bench = strMessage2Bench.Remove(7, 2)
                    strMessage2Bench = strMessage2Bench.Insert(7, "  ")
                    For i = 0 To strMessage2Bench.Length - 1
                        If strMessage2Bench.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage2Bench = strMessage2Bench.Remove(i, 1)
                            strMessage2Bench = strMessage2Bench.Insert(0, "0")
                        End If
                    Next
                    strMessage2Bench = strMessage2Bench.Remove(0, icomma)
                    icomma = 0
                End If
            End If
            If strMessage2Bench <> "" Then
                strMessage3Bench = Engine.Get("Message_3")
                If strMessage3Bench = Nothing Then strMessage3Bench = ""
                If strMessage3Bench <> "" Then
                    strMessage3Bench = strMessage3Bench.Remove(7, 2)
                    strMessage3Bench = strMessage3Bench.Insert(7, "  ")
                    For i = 0 To strMessage3Bench.Length - 1
                        If strMessage3Bench.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage3Bench = strMessage3Bench.Remove(i, 1)
                            strMessage3Bench = strMessage3Bench.Insert(0, "0")
                        End If
                    Next
                    strMessage3Bench = strMessage3Bench.Remove(0, icomma)
                    icomma = 0
                End If
            End If
            If strMessage3Bench <> "" Then
                strMessage4Bench = Engine.Get("Message_4")
                If strMessage4Bench = Nothing Then strMessage4Bench = ""
                If strMessage4Bench <> "" Then
                    strMessage4Bench = strMessage4Bench.Remove(7, 2)
                    strMessage4Bench = strMessage4Bench.Insert(7, "  ")
                    For i = 0 To strMessage4Bench.Length - 1
                        If strMessage4Bench.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage4Bench = strMessage4Bench.Remove(i, 1)
                            strMessage4Bench = strMessage4Bench.Insert(0, "0")
                        End If
                    Next
                    strMessage4Bench = strMessage4Bench.Remove(0, icomma)
                    icomma = 0
                End If
            End If

            If strMessage4Bench <> "" Then
                strMessage5Bench = Engine.Get("Message_5")
                If strMessage5Bench = Nothing Then strMessage5Bench = ""
                If strMessage5Bench <> "" Then
                    strMessage5Bench = strMessage5Bench.Remove(7, 2)
                    strMessage5Bench = strMessage5Bench.Insert(7, "  ")
                    For i = 0 To strMessage5Bench.Length - 1
                        If strMessage5Bench.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage5Bench = strMessage5Bench.Remove(i, 1)
                            strMessage5Bench = strMessage5Bench.Insert(0, "0")
                        End If
                    Next
                    strMessage5Bench = strMessage5Bench.Remove(0, icomma)
                    icomma = 0
                End If
            End If
            If strMessage5Bench <> "" Then
                strMessage6Bench = Engine.Get("Message_6")
                If strMessage6Bench = Nothing Then strMessage6Bench = ""
                If strMessage6Bench <> "" Then
                    strMessage6Bench = strMessage6Bench.Remove(7, 2)
                    strMessage6Bench = strMessage6Bench.Insert(7, "  ")
                    For i = 0 To strMessage6Bench.Length - 1
                        If strMessage6Bench.Substring(i, 1) = "," Then
                            icomma = icomma + 1
                            strMessage6Bench = strMessage6Bench.Remove(i, 1)
                            strMessage6Bench = strMessage6Bench.Insert(0, "0")
                        End If
                    Next
                    strMessage6Bench = strMessage6Bench.Remove(0, icomma)
                    icomma = 0
                End If
            End If

            'Header Values
            strFIACompanyNameBench = Header.Get("bridge_CompanyName")

            'Basic Values

            strFIAProdNameBench = Basic.Get("value_ProdName")

            If strFIAProdNameBench = "FIA7L" Then
                strFIAGMCVBench = Basic.Get("value_GMCV")
            End If
            strFIAAnnualPremiumBench = Basic.Get("value_Premium")
            strFIAPremiumEnhancementBench = Basic.Get("value_PremiumEnhancement")
            strFIAInitialBeneBaseBench = Basic.Get("value_InitialBenefitBase")
            strFIABailoutAnnualCapBench = Basic.Get("value_BailoutAnnualCap")
            strFIARiderRollUpRateBench = Basic.Get("value_RiderRollUpRate")
            strFIARiderChargeBench = Basic.Get("value_RiderCharge")
            strFIAAgeAtFirstWDBench = Basic.Get("value_AgeAtFirstWD")
            strFIAAnnWDLimitGuarBench = Basic.Get("value_AnnWDLimitGuar")
            strFIAAnnWDLimitProjBench = Basic.Get("value_AnnWDLimitProj")
            strFIAOneYearFixedAllocBench = Basic.Get("value_OneYearFixedAllocation")
            strFIASevenYearFixedAllocBench = Basic.Get("value_SevenYearFixedAllocation")
            strFIATenYearFixedAllocBench = Basic.Get("value_TenYearFixedAllocation")
            strFIAAnnCapAllocBench = Basic.Get("value_AnnualCapAllocation")
            strFIAMonCapAllocBench = Basic.Get("value_MonthlyCapAllocation")
            strFIAPerfTrigAllocBench = Basic.Get("value_PerfTriggerAllocation")
            strFIAOneYearFixedInitialRateBench = Basic.Get("value_OneYearFixedInitialRate")
            strFIASevenYearFixedInitialRateBench = Basic.Get("value_SevenYearFixedInitialRate")
            strFIATenYearFixedInitialRateBench = Basic.Get("value_TenYearFixedInitialRate")
            strFIAAnnualCapCapBench = Basic.Get("value_AnnualCapCap")
            strFIAMonthlyCapCapBench = Basic.Get("value_MonthlyCapCap")
            strFIAPerfTrigSpecifiedRateBench = Basic.Get("value_PerfTriggerSpecifiedRate")
            strFIAYearsToPrintBench = Basic.Get("value_YearsToPrint")
            strFIASpecPeriodStartDateBench = Basic.Get("value_SpecPeriodStartDate")
            strFIASpecPeriodEndDateBench = Basic.Get("value_SpecPeriodEndDate")
            strFIAFavPeriodStartDateBench = Basic.Get("value_FavPeriodStartDate")
            strFIAFavPeriodEndDateBench = Basic.Get("value_FavPeriodEndDate")
            strFIAUnFavPeriodStartDateBench = Basic.Get("value_UnfavPeriodStartDate")
            strFIAUnFavPeriodEndDateBench = Basic.Get("value_UnfavPeriodEndDate")
            strFIASpecSPChangeBench = Basic.Get("value_SpecSPChange")
            strFIASpecWDBench = Basic.Get("value_SpecWD")
            strFIASpecAnnCreditRateBench = Basic.Get("value_SpecAnnCreditRate")
            strFIASpecContractValueBench = Basic.Get("value_SpecContractValue")
            strFIASpecSurrenderValueBench = Basic.Get("value_SpecSurrenderValue")
            strFIASpecMGSVBench = Basic.Get("value_SpecMGSV")
            strFIASpecProjBeneBaseBench = Basic.Get("value_SpecProjBeneBase")
            strFIASpecGuarBeneBaseBench = Basic.Get("value_SpecGuarBeneBase")
            strFIASpecProjWDLimitBench = Basic.Get("value_SpecProjWDLimit")
            strFIASpecGuarWDLimitBench = Basic.Get("value_SpecGuarWDLimit")
            strFIASevenYearIntRateBench = Basic.Get("value_SevenYearIntRate")
            strFIATenYearIntRateBench = Basic.Get("value_TenYearIntRate")
            strFIAMonthlyCapIndexCreditBench = Basic.Get("value_MonthlyCapIndexCredit")
            strFIAAnnualCapIndexCreditBench = Basic.Get("value_AnnualCapIndexCredit")
            strFIAPerfTriggerIndexCreditBench = Basic.Get("value_PerfTriggerIndexCredit")
            strFIASevenYearAccumValueBench = Basic.Get("value_SevenYearAccumValue")
            strFIATenYearAccumValueBench = Basic.Get("value_TenYearAccumValue")
            strFIAMonthlyCapAccumValueBench = Basic.Get("value_MonthlyCapAccumValue")
            strFIAAnnualCapAccumValueBench = Basic.Get("value_AnnualCapAccumValue")
            strFIAPerfTriggerAccumValueBench = Basic.Get("value_PerfTriggerAccumValue")
            strFIAContractValueNoWDBench = Basic.Get("value_ContractValueNoWD")
            strFIAGuarWDFactorBench = Basic.Get("value_GuarWDFactor")
            strFIAGuarBeneBaseNoWDBench = Basic.Get("value_GuarBeneBaseNoWD")
            strFIAGuarWDLimitNoWDBench = Basic.Get("value_GuarWDLimitNoWD")
            strFIAProjBeneBaseNoWDBench = Basic.Get("value_ProjBeneBaseNoWD")
            strFIAProjWDLimitNoWDBench = Basic.Get("value_ProjWDLimitNoWD")
            strFIAFavSPChangeBench = Basic.Get("value_FavSPChange")
            strFIAUnfavSPChangeBench = Basic.Get("value_UnfavSPChange")
            strFIAFavWDBench = Basic.Get("value_FavWD")
            strFIAUnfavWDBench = Basic.Get("value_UnfavWD")
            strFIAFavAnnCreditRateBench = Basic.Get("value_FavAnnCreditRate")
            strFIAUnfavAnnCreditRateBench = Basic.Get("value_UnfavAnnCreditRate")
            strFIAFavContractValueBench = Basic.Get("value_FavContractValue")
            strFIAUnfavContractValueBench = Basic.Get("value_UnfavContractValue")
            strFIAFavSurrenderValueBench = Basic.Get("value_FavSurrenderValue")
            strFIAUnfavSurrenderValueBench = Basic.Get("value_UnfavSurrenderValue")
            strFIAFavMGSVBench = Basic.Get("value_FavMGSV")
            strFIAUnfavMGSVBench = Basic.Get("value_UnfavMGSV")
            strFIAFavProjBeneBaseBench = Basic.Get("value_FavProjBeneBase")
            strFIAFavGuarBeneBaseBench = Basic.Get("value_FavGuarBeneBase")
            strFIAFavProjWDLimitBench = Basic.Get("value_FavProjWDLimit")
            strFIAFavGuarWDLimitBench = Basic.Get("value_FavGuarWDLimit")
            strFIAUnfavProjBeneBaseBench = Basic.Get("value_UnfavProjBeneBase")
            strFIAUnfavGuarBeneBaseBench = Basic.Get("value_UnfavGuarBeneBase")
            strFIAUnfavProjWDLimitBench = Basic.Get("value_UnfavProjWDLimit")
            strFIAUnfavGuarWDLimitBench = Basic.Get("value_UnfavGuarWDLimit")
            strFIANonForfIntRateBench = Basic.Get("value_FIANonForfIntRate")
            strFIAMGSVBench = Basic.Get("value_FIAMGSV")
            strFIAFixedJumboBench = Basic.Get("value_FixedJumbo")
            strFIAAnnCapJumboBench = Basic.Get("value_AnnCapJumbo")
            strFIAMonthlyCapJumboBench = Basic.Get("value_MonthlyCapJumbo")
            strFIAPerfTriggerJumboBench = Basic.Get("value_PerfTriggerJumbo")
            strFIAWDFrequencyBench = Basic.Get("value_WDFrequency")
            strFIAAnnCreditRateNoWDBench = Basic.Get("value_AnnCreditRateNoWD")


            strFIAClient1Bench = Basic.Get("value_Name1")
            strFIAAge1Bench = Basic.Get("value_Age1")
            strFIASex1Bench = Basic.Get("value_Sex1")
            strFIAClient2Bench = Basic.Get("value_Name2")
            strFIAAge2Bench = Basic.Get("value_Age2")
            strFIASex2Bench = Basic.Get("value_Sex2")
            strFIADOB1Bench = Basic.Get("value_DOB1")
            strFIADOB2Bench = Basic.Get("value_DOB2")

            strFIAStateBench = Basic.Get("value_State")
            strFIATaxStatusBench = Basic.Get("value_TaxStatus")

            strFIAPremiumTaxRateBench = Basic.Get("value_PremTaxRate")


            strFIAPolicyFormBench = Basic.Get("value_PolicyForm")
            strFIASurrChargeYrsBench = Basic.Get("value_SurrChargeYrs")
            strFIAChannelCodeBench = Basic.Get("value_ChannelCode")


            strFIASurrenderChargesBench = Basic.Get("value_SurrenderCharges")




            End If
    End Sub

    Public Sub New()

    End Sub
End Class



