Imports Nini.Config

Public Class clsReadValues

    Public Shared Header As IConfig
    Public Shared Basic As IConfig
    Public Shared Engine As IConfig
    Public Shared strFundCodeTest(30) As String
    Public Shared strFundPctTest(30) As String
    Public Shared strFundNameTest(30) As String
    Public Shared strReturnYr1StdTest(30) As String
    Public Shared strReturnAdoptionDateStdTest(30) As String
    Public Shared strReturnYr5StdTest(30) As String
    Public Shared strReturnYr10StdTest(30) As String
    Public Shared strReturnAdoptionStdTest(30) As String
    Public Shared strReturnAdoptionDateStdIGATest(30) As String
    Public Shared strReturnYr1StdGIATest(30) As String
    Public Shared strReturnYr5StdGIATest(30) As String
    Public Shared strReturnYr10StdGIATest(30) As String
    Public Shared strReturnAdoptionStdGIATest(30) As String
    Public Shared strReturnAdoptionDateStdGIATest(30) As String
    Public Shared strReturnYr1NonStdSCTest(30) As String
    Public Shared strReturnYr5NonStdSCTest(30) As String
    Public Shared strReturnYr10NonStdSCTest(30) As String
    Public Shared strReturnAdoptionNonStdSCTest(30) As String
    Public Shared strReturnYr1NonStdSCGIATest(30) As String
    Public Shared strReturnYr5NonStdSCGIATest(30) As String
    Public Shared strReturnYr10NonStdSCGIATest(30) As String
    Public Shared strReturnAdoptionNonStdSCGIATest(30) As String
    Public Shared strReturnYr1NonStdTest(30) As String
    Public Shared strReturnYr5NonStdTest(30) As String
    Public Shared strReturnYr10NonStdTest(30) As String
    Public Shared strReturnAdoptionNonStdTest(30) As String
    Public Shared strReturnYr1NonStdGIATest(30) As String
    Public Shared strReturnYr5NonStdGIATest(30) As String
    Public Shared strReturnYr10NonStdGIATest(30) As String
    Public Shared strReturnAdoptionNonStdGIATest(30) As String
    Public Shared strReturnAdoptionDateNonStdSCTest(30) As String
    Public Shared strReturnAdoptionDateNonStdSCGIATest(30) As String
    Public Shared strReturnInceptionDateNonStdTest(30) As String
    Public Shared strReturnInceptionDateNonStdGIATest(30) As String
    Public Shared strHistPeriodEndingTest(30) As String
    Public Shared strHistCumulativeReturnTest(30) As String
    Public Shared strHistAverageAnnReturnTest(30) As String
    'Public Shared strReturnAdoptionDateNonStdSCTest(30) As String
    Public Shared strReturnAdoptionDateNonStdTest(30) As String
    Public Shared strReturnAdoptionDateNonStdGIATest(30) As String
    Public Shared strCompanyNameTest As String
    Public Shared strClient1Test As String
    Public Shared strAge1Test As String
    Public Shared strAgeOlderTest As String
    Public Shared strIRateTest As String
    Public Shared strSex1Test As String
    Public Shared strInitialDBTest As String
    Public Shared strJointTest As String
    Public Shared strInitialDB2Test As String
    Public Shared strClient2Test As String
    Public Shared strAge2Test As String
    Public Shared strSex2Test As String
    Public Shared strContractTypeTest As String
    Public Shared strSurrChargeYrsTest As String
    Public Shared strHypoNorGTest As String
    Public Shared strZeroNetTest As String
    Public Shared strHypoNetTest As String
    Public Shared strHypoGrossTest As String
    Public Shared strHypoGISRateTest As String
    Public Shared strZeroGISRateTest As String
    Public Shared strExpensesVAMandEOnlyTest As String
    Public Shared strZeroGrowthRateTest As String
    Public Shared strExpensesAdminOnlyTest As String
    Public Shared strFundExpensesVATest As String
    Public Shared strFundExpenseGISTest As String
    Public Shared strFundExpenseEffDateGISTest As String
    Public Shared strVADBBeneRiderChargeTest As String
    Public Shared strVAContractChargeTest As String
    Public Shared strVAContractChargeWaiverLimitTest As String
    Public Shared strFundExpenseEffDateTest As String
    Public Shared strEarlyAccessChargeTest As String
    Public Shared strLivingBenefitRiderChargeTest As String
    Public Shared strInitialPremiumTest As String
    Public Shared strPrintYearsTest As String
    Public Shared strDBTypeTest As String
    Public Shared strFundCountTest As String
    Public Shared strInvestStratTest As String
    Public Shared strIncomeStartAgeTest As String
    Public Shared strIncomeStartAgeJointTest As String
    Public Shared strIncomeStartYearTest As String
    Public Shared strIncomeStartMonthTest As String
    Public Shared strYearsCertainTest As String
    Public Shared strGIAInitialMonthlyPayoutHypoTest As String
    Public Shared strGIAInitialMonthlyPayoutZeroTest As String
    Public Shared strGIASchedInstallmentTest As String
    Public Shared strAnnAmountZeroTest As String
    Public Shared strAnnAmountHypoTest As String
    Public Shared strInstallmentCountZeroTest As String
    Public Shared strInstallmentCountHypoTest As String
    Public Shared strPPDBChargeTest As String
    Public Shared strLIPFactorFirstWDTest As String
    Public Shared strLIPGuarWithdrawalTest As String
    Public Shared strLIPWDStartYearTest As String
    Public Shared strLIPWDStartMonthTest As String
    Public Shared strLIPBenBaseFirstWDTest As String
    Public Shared strTaxExcludableAmtZeroTest As String
    Public Shared strTaxExcludableAmtHypoTest As String
    Public Shared strTaxExcludableAmtHistTest As String
    Public Shared strTaxBracketTest As String
    Public Shared strTaxBasisTest As String

    Public Shared strInvestmentTest As String
    Public Shared strBaseContractValueZeroTest As String
    Public Shared strCombinedSurrValueZeroTest As String
    Public Shared strGISValueZeroTest As String
    Public Shared strAnnualIncomeZeroTest As String
    Public Shared strDeathBenefitZeroTest As String
    Public Shared strTransfertoGISZeroTest As String
    Public Shared strTotalContractValueZeroTest As String
    Public Shared strSurrenderChargesTest As String
    Public Shared strHypoAnnIncomeFloorTest As String
    Public Shared strGIASumGtdAmtTest As String
    Public Shared strHistReturnForPeriodTest As String
    Public Shared strHistAnnIncomeTest As String
    Public Shared strHistAccountGISTest As String
    Public Shared strHistTotalContractValueTest As String
    Public Shared strHistTotalSurrValueTest As String
    Public Shared strHistDeathBenefitGIATest As String
    Public Shared strPPBAHypoRateTest As String
    Public Shared strPPBAZeroRateTest As String
    Public Shared strBenefitBaseZeroRateTest As String
    Public Shared strBenefitBaseHypoRateTest As String
    Public Shared strRollupZeroRateTest As String
    Public Shared strRollupHypoRateTest As String
    Public Shared strLIPAnnIncomeZeroRateTest As String
    Public Shared strLIPAnnIncomeHypoRateTest As String
    Public Shared strLIPResetValueZeroRateTest As String
    Public Shared strLIPContractValueHypoTest As String
    Public Shared strBASEWithdrawalZeroTest As String
    Public Shared strBASEWithdrawalHypoTest As String
    Public Shared strEPRZeroTest As String
    Public Shared strEPRHypoTest As String
    Public Shared strEPRHistTest As String
    Public Shared strDBGIAZeroTest As String
    Public Shared strDBGIAHypoTest As String
    Public Shared strDBLIPZeroTest As String
    Public Shared strDBLIPHypoTest As String
    Public Shared strDBLIPHistTest As String
    Public Shared strDBComboZeroTest As String
    Public Shared strDBComboHypoTest As String
    Public Shared strDBComboHistTest As String
    Public Shared strDBASDBZeroTest As String
    Public Shared strDBASDBHypoTest As String
    Public Shared strDBASDBHistTest As String
    Public Shared strDBRollupZeroTest As String
    Public Shared strDBRollupHypoTest As String
    Public Shared strDBRollupHistTest As String
    Public Shared strDBStandardZeroTest As String
    Public Shared strDBStandardHypoTest As String
    Public Shared strDBStandardHistTest As String

    Public Shared strStatusTest As String
    Public Shared strMessage1Test As String
    Public Shared strMessage2Test As String
    Public Shared strMessage3Test As String
    Public Shared strMessage4Test As String

    Public Shared strFundCodeBench(30) As String
    Public Shared strFundPctBench(30) As String
    Public Shared strFundNameBench(30) As String
    Public Shared strReturnYr1StdBench(30) As String
    Public Shared strReturnAdoptionDateStdBench(30) As String
    Public Shared strReturnYr5StdBench(30) As String
    Public Shared strReturnYr10StdBench(30) As String
    Public Shared strReturnAdoptionStdBench(30) As String
    Public Shared strReturnAdoptionDateStdIGABench(30) As String
    Public Shared strReturnYr1StdGIABench(30) As String
    Public Shared strReturnYr5StdGIABench(30) As String
    Public Shared strReturnYr10StdGIABench(30) As String
    Public Shared strReturnAdoptionStdGIABench(30) As String
    Public Shared strReturnAdoptionDateStdGIABench(30) As String
    Public Shared strReturnYr1NonStdSCBench(30) As String
    Public Shared strReturnYr5NonStdSCBench(30) As String
    Public Shared strReturnYr10NonStdSCBench(30) As String
    Public Shared strReturnAdoptionNonStdSCBench(30) As String
    Public Shared strReturnYr1NonStdSCGIABench(30) As String
    Public Shared strReturnYr5NonStdSCGIABench(30) As String
    Public Shared strReturnYr10NonStdSCGIABench(30) As String
    Public Shared strReturnAdoptionNonStdSCGIABench(30) As String
    Public Shared strReturnYr1NonStdBench(30) As String
    Public Shared strReturnYr5NonStdBench(30) As String
    Public Shared strReturnYr10NonStdBench(30) As String
    Public Shared strReturnAdoptionNonStdBench(30) As String
    Public Shared strReturnYr1NonStdGIABench(30) As String
    Public Shared strReturnYr5NonStdGIABench(30) As String
    Public Shared strReturnYr10NonStdGIABench(30) As String
    Public Shared strReturnAdoptionNonStdGIABench(30) As String
    Public Shared strReturnAdoptionDateNonStdSCBench(30) As String
    Public Shared strReturnAdoptionDateNonStdSCGIABench(30) As String
    Public Shared strReturnInceptionDateNonStdBench(30) As String
    Public Shared strReturnInceptionDateNonStdGIABench(30) As String
    Public Shared strHistPeriodEndingBench(30) As String
    Public Shared strHistCumulativeReturnBench(30) As String
    Public Shared strHistAverageAnnReturnBench(30) As String
    'Public Shared strReturnAdoptionDateNonStdSCBench(30) As String
    Public Shared strReturnAdoptionDateNonStdBench(30) As String
    Public Shared strReturnAdoptionDateNonStdGIABench(30) As String
    Public Shared strCompanyNameBench As String
    Public Shared strClient1Bench As String
    Public Shared strAge1Bench As String
    Public Shared strAgeOlderBench As String
    Public Shared strIRateBench As String
    Public Shared strSex1Bench As String
    Public Shared strInitialDBBench As String
    Public Shared strJointBench As String
    Public Shared strInitialDB2Bench As String
    Public Shared strClient2Bench As String
    Public Shared strAge2Bench As String
    Public Shared strSex2Bench As String
    Public Shared strContractTypeBench As String
    Public Shared strSurrChargeYrsBench As String
    Public Shared strHypoNorGBench As String
    Public Shared strZeroNetBench As String
    Public Shared strHypoNetBench As String
    Public Shared strHypoGrossBench As String
    Public Shared strHypoGISRateBench As String
    Public Shared strZeroGISRateBench As String
    Public Shared strExpensesVAMandEOnlyBench As String
    Public Shared strZeroGrowthRateBench As String
    Public Shared strExpensesAdminOnlyBench As String
    Public Shared strFundExpensesVABench As String
    Public Shared strFundExpenseGISBench As String
    Public Shared strFundExpenseEffDateGISBench As String
    Public Shared strVADBBeneRiderChargeBench As String
    Public Shared strVAContractChargeBench As String
    Public Shared strVAContractChargeWaiverLimitBench As String
    Public Shared strFundExpenseEffDateBench As String
    Public Shared strEarlyAccessChargeBench As String
    Public Shared strLivingBenefitRiderChargeBench As String
    Public Shared strInitialPremiumBench As String
    Public Shared strPrintYearsBench As String
    Public Shared strDBTypeBench As String
    Public Shared strFundCountBench As String
    Public Shared strInvestStratBench As String
    Public Shared strIncomeStartAgeBench As String
    Public Shared strIncomeStartAgeJointBench As String
    Public Shared strIncomeStartYearBench As String
    Public Shared strIncomeStartMonthBench As String
    Public Shared strYearsCertainBench As String
    Public Shared strGIAInitialMonthlyPayoutHypoBench As String
    Public Shared strGIAInitialMonthlyPayoutZeroBench As String
    Public Shared strGIASchedInstallmentBench As String
    Public Shared strAnnAmountZeroBench As String
    Public Shared strAnnAmountHypoBench As String
    Public Shared strInstallmentCountZeroBench As String
    Public Shared strInstallmentCountHypoBench As String
    Public Shared strPPDBChargeBench As String
    Public Shared strLIPFactorFirstWDBench As String
    Public Shared strLIPGuarWithdrawalBench As String
    Public Shared strLIPWDStartYearBench As String
    Public Shared strLIPWDStartMonthBench As String
    Public Shared strLIPBenBaseFirstWDBench As String
    Public Shared strTaxExcludableAmtZeroBench As String
    Public Shared strTaxExcludableAmtHypoBench As String
    Public Shared strTaxExcludableAmtHistBench As String
    Public Shared strTaxBracketBench As String
    Public Shared strTaxBasisBench As String

    Public Shared strInvestmentBench As String
    Public Shared strBaseContractValueZeroBench As String
    Public Shared strCombinedSurrValueZeroBench As String
    Public Shared strGISValueZeroBench As String
    Public Shared strAnnualIncomeZeroBench As String
    Public Shared strDeathBenefitZeroBench As String
    Public Shared strTransfertoGISZeroBench As String
    Public Shared strTotalContractValueZeroBench As String
    Public Shared strSurrenderChargesBench As String
    Public Shared strHypoAnnIncomeFloorBench As String
    Public Shared strGIASumGtdAmtBench As String
    Public Shared strHistReturnForPeriodBench As String
    Public Shared strHistAnnIncomeBench As String
    Public Shared strHistAccountGISBench As String
    Public Shared strHistTotalContractValueBench As String
    Public Shared strHistTotalSurrValueBench As String
    Public Shared strHistDeathBenefitGIABench As String
    Public Shared strPPBAHypoRateBench As String
    Public Shared strPPBAZeroRateBench As String
    Public Shared strBenefitBaseZeroRateBench As String
    Public Shared strBenefitBaseHypoRateBench As String
    Public Shared strRollupZeroRateBench As String
    Public Shared strRollupHypoRateBench As String
    Public Shared strLIPAnnIncomeZeroRateBench As String
    Public Shared strLIPAnnIncomeHypoRateBench As String
    Public Shared strLIPResetValueZeroRateBench As String
    Public Shared strLIPContractValueHypoBench As String
    Public Shared strBASEWithdrawalZeroBench As String
    Public Shared strBASEWithdrawalHypoBench As String
    Public Shared strEPRZeroBench As String
    Public Shared strEPRHypoBench As String
    Public Shared strEPRHistBench As String
    Public Shared strDBGIAZeroBench As String
    Public Shared strDBGIAHypoBench As String
    Public Shared strDBLIPZeroBench As String
    Public Shared strDBLIPHypoBench As String
    Public Shared strDBLIPHistBench As String
    Public Shared strDBComboZeroBench As String
    Public Shared strDBComboHypoBench As String
    Public Shared strDBComboHistBench As String
    Public Shared strDBASDBZeroBench As String
    Public Shared strDBASDBHypoBench As String
    Public Shared strDBASDBHistBench As String
    Public Shared strDBRollupZeroBench As String
    Public Shared strDBRollupHypoBench As String
    Public Shared strDBRollupHistBench As String
    Public Shared strDBStandardZeroBench As String
    Public Shared strDBStandardHypoBench As String
    Public Shared strDBStandardHistBench As String

    Public Shared strStatusBench As String
    Public Shared strMessage1Bench As String
    Public Shared strMessage2Bench As String
    Public Shared strMessage3Bench As String
    Public Shared strMessage4Bench As String

    Public Shared gErrorBench As Boolean = False
    Public Shared gErrorTest As Boolean = False

    'Timestamps

    Public Shared strWFProp As String = CStr(Today)
    Public Shared strGLAICCPY As String = CStr(Today)
    Public Shared strGLICNYCPY As String = CStr(Today)
    Public Shared strAnn1 As String = CStr(Today)
    Public Shared strAnn2 As String = CStr(Today)
    Public Shared strAnn3 As String = CStr(Today)
    Public Shared strAnn4 As String = CStr(Today)
    Public Shared strWFGELA As String = CStr(Today)
    Public Shared strGECLRIC As String = CStr(Today)
    Public Shared strGECLVA1 As String = CStr(Today)
    Public Shared strGECLVA2 As String = CStr(Today)
    Public Shared strGECLVA3 As String = CStr(Today)
    Public Shared strGECLEXE As String = CStr(Today)

    Public Shared strGLAICFIXEDSPIANNVER As String = CStr(Today)
    Public Shared strGLAICFIXEDSPIAANNSUPP As String = CStr(Today)
    Public Shared strGLAICFIXEDSPIACPY As String = CStr(Today)
    Public Shared strGLAICFIXEDSPIAANNSYS As String = CStr(Today)
    Public Shared strGLAICFIXEDSPIAANNPROD As String = CStr(Today)
    Public Shared strGLAICFIXEDSPIAANNRATES As String = CStr(Today)
    Public Shared strGLAICFIXEDSPIAEXE As String = CStr(Today)
    Public Shared strGLICSPDACPY As String = CStr(Today)
    Public Shared strGLICSPDAGNAWINE As String = CStr(Today)
    Public Shared strGLICSPDARATE As String = CStr(Today)
    Public Shared strGLICSPDAEXE As String = CStr(Today)
    Public Shared strGLICSPIANNVER As String = CStr(Today)
    Public Shared strGLICSPIAANNSUPP As String = CStr(Today)
    Public Shared strGLICNYSPIAANNPROD As String = CStr(Today)
    Public Shared strGLICNYSPIACPY As String = CStr(Today)
    Public Shared strGLICNYSPIAANNVER As String = CStr(Today)
    Public Shared strGLICNYSPDAEXE As String = CStr(Today)
    Public Shared strGLICNYSPDAGNAWINE As String = CStr(Today)
    Public Shared strGLICSPIAEXE As String = CStr(Today)
    Public Shared strGLICSPIAANNPROD As String = CStr(Today)
    Public Shared strGLICSPIACPY As String = CStr(Today)
    Public Shared strGLICSPIAANNSYS As String = CStr(Today)
    Public Shared strGLICSPIAANNRATES As String = CStr(Today)
    Public Shared strGLICNYSPDACPY As String = CStr(Today)
    Public Shared strGLICNYSPDARATE As String = CStr(Today)
    Public Shared strGLICNYSPIAANNSYS As String = CStr(Today)
    Public Shared strGLICNYSPIAANNRATES As String = CStr(Today)
    Public Shared strGLICNYSPIAEXE As String = CStr(Today)
    Public Shared strGLICNYSPIAANNSUPP As String = CStr(Today)

    'Mismatches

    Public Shared bMisMatch As Boolean
    Public Shared bMismatchAtLeastOnce As Boolean
    Public Shared strCompanyNameMM As String
    Public Shared strClient1NameMM As String
    Public Shared strClient2NameMM As String
    Public Shared strSex1MM As String
    Public Shared strSex2MM As String
    Public Shared strAge1MM As String
    Public Shared strAge2MM As String
    Public Shared strAgeOlderMM As String
    Public Shared strIRateMM As String
    Public Shared strInitialDBMM As String
    Public Shared strInitialDB2MM As String
    Public Shared strJointMM As String
    Public Shared strContractTypeMM As String
    Public Shared strSurrChargeYrsMM As String
    Public Shared strHypoNorGMM As String
    Public Shared strZeroNetMM As String
    Public Shared strHypoNetMM As String
    Public Shared strHypoGrossMM As String
    Public Shared strHypoGISRateMM As String
    Public Shared strZeroGISRateMM As String
    Public Shared strExpenseVAMandEOnlyMM As String
    Public Shared strZeroGrowthRateMM As String
    Public Shared strExpensesAdminOnlyMM As String
    Public Shared strFundExpensesVAMM As String
    Public Shared strFundExpenseGISMM As String
    Public Shared strFundExpenseEffDateGISMM As String
    Public Shared strVADBBenefitRiderChargeMM As String
    Public Shared strVAContractChargeMM As String
    Public Shared strVAContractChargeWaiverLimitMM As String
    Public Shared strFundExpenseEffDateMM As String

    Public Shared strEarlyAccessChargeMM As String
    Public Shared strLivingBenefitRiderChargeMM As String
    Public Shared strInitialPremiumMM As String
    Public Shared strPrintYearsMM As String
    Public Shared strDBTypeMM As String
    Public Shared strFundCountMM As String
    Public Shared strInvestStratMM As String
    Public Shared strIncomeStartAgeMM As String
    Public Shared strIncomeStartAgeJointMM As String
    Public Shared strIncomeStartYearMM As String
    Public Shared strIncomeStartMonthMM As String
    Public Shared strYearsCertainMM As String
    Public Shared strGIAInitialMonthlyPayoutHypoMM As String
    Public Shared strGIAInitialMonthlyPayoutZeroMM As String
    Public Shared strGIASchedInstallmentMM As String
    Public Shared strAnnAmountZeroMM As String
    Public Shared strAnnAmountHypoMM As String
    Public Shared strInstallmentCountZeroMM As String
    Public Shared strInstallmentCountHypoMM As String
    Public Shared strPPDBChargeMM As String
    Public Shared strLIPFactorFirstWDMM As String
    Public Shared strLIPGuarWithdrawalMM As String
    Public Shared strLIPWDStartYearMM As String
    Public Shared strLIPWDStartMonthMM As String
    Public Shared strLIPBenBaseFirstWDMM As String
    Public Shared strTaxExcludableAmtZeroMM As String
    Public Shared strTaxExcludableAmtHypoMM As String
    Public Shared strTaxExcludableAmtHistMM As String
    Public Shared strTaxBracketMM As String
    Public Shared strTaxBasisMM As String

    Public Shared strInvestmentMM As String
    Public Shared strBaseContractValueZeroMM As String
    Public Shared strCombinedSurrValueZeroMM As String
    Public Shared strGISValueZeroMM As String
    Public Shared strAnnualIncomeZeroMM As String
    Public Shared strDeathBenefitZeroMM As String
    Public Shared strTransfertoGISZeroMM As String
    Public Shared strTotalContractValueZeroMM As String
    Public Shared strSurrenderChargesMM As String
    Public Shared strHypoAnnIncomeFloorMM As String
    Public Shared strGIASumGtdAmtMM As String
    Public Shared strHistReturnForPeriodMM As String
    Public Shared strHistAnnIncomeMM As String
    Public Shared strHistAccountGISMM As String
    Public Shared strHistTotalContractValueMM As String
    Public Shared strHistTotalSurrValueMM As String
    Public Shared strHistDeathBenefitGIAMM As String
    Public Shared strPPBAHypoRateMM As String
    Public Shared strPPBAZeroRateMM As String
    Public Shared strBenefitBaseZeroRateMM As String
    Public Shared strBenefitBaseHypoRateMM As String
    Public Shared strRollupZeroRateMM As String
    Public Shared strRollupHypoRateMM As String
    Public Shared strLIPAnnIncomeZeroRateMM As String
    Public Shared strLIPAnnIncomeHypoRateMM As String
    Public Shared strLIPResetValueZeroRateMM As String
    Public Shared strLIPContractValueHypoMM As String
    Public Shared strBASEWithdrawalZeroMM As String
    Public Shared strBASEWithdrawalHypoMM As String
    Public Shared strEPRZeroMM As String
    Public Shared strEPRHypoMM As String
    Public Shared strEPRHistMM As String
    Public Shared strDBGIAZeroMM As String
    Public Shared strDBGIAHypoMM As String
    Public Shared strDBLIPZeroMM As String
    Public Shared strDBLIPHypoMM As String
    Public Shared strDBLIPHistMM As String
    Public Shared strDBComboZeroMM As String
    Public Shared strDBComboHypoMM As String
    Public Shared strDBComboHistMM As String
    Public Shared strDBASDBZeroMM As String
    Public Shared strDBASDBHypoMM As String
    Public Shared strDBASDBHistMM As String
    Public Shared strDBRollupZeroMM As String
    Public Shared strDBRollupHypoMM As String
    Public Shared strDBRollupHistMM As String
    Public Shared strDBStandardZeroMM As String
    Public Shared strDBStandardHypoMM As String
    Public Shared strDBStandardHistMM As String

    Public Shared strFundCodeMM(30) As String
    Public Shared strFundPctMM(30) As String
    Public Shared strFundNameMM(30) As String
    Public Shared strReturnYr1StdMM(30) As String
    Public Shared strReturnAdoptionDatestdMM(30) As String
    Public Shared strReturnYr5StdMM(30) As String
    Public Shared strReturnYr10StdMM(30) As String
    Public Shared strReturnAdoptionStdMM(30) As String
    Public Shared strReturnAdoptionDatestdIGAMM(30) As String
    Public Shared strReturnYr1StdGIAMM(30) As String
    Public Shared strReturnYr5StdGIAMM(30) As String
    Public Shared strReturnYr10StdGIAMM(30) As String
    Public Shared strReturnAdoptionStdGIAMM(30) As String
    Public Shared strReturnAdoptionDatestdGIAMM(30) As String
    Public Shared strReturnYr1NonStdSCMM(30) As String
    Public Shared strReturnYr5NonStdSCMM(30) As String
    Public Shared strReturnYr10NonStdSCMM(30) As String
    Public Shared strReturnAdoptionNonStdSCMM(30) As String
    Public Shared strReturnYr1NonStdSCGIAMM(30) As String
    Public Shared strReturnYr5NonStdSCGIAMM(30) As String
    Public Shared strReturnYr10NonStdSCGIAMM(30) As String
    Public Shared strReturnAdoptionNonStdSCGIAMM(30) As String
    Public Shared strReturnYr1NonStdMM(30) As String
    Public Shared strReturnYr5NonStdMM(30) As String
    Public Shared strReturnYr10NonStdMM(30) As String
    Public Shared strReturnAdoptionNonStdMM(30) As String
    Public Shared strReturnYr1NonStdGIAMM(30) As String
    Public Shared strReturnYr5NonStdGIAMM(30) As String
    Public Shared strReturnYr10NonStdGIAMM(30) As String
    Public Shared strReturnAdoptionNonStdGIAMM(30) As String
    Public Shared strReturnAdoptionDateNonStdSCMM(30) As String
    Public Shared strReturnAdoptionDateNonStdSCGIAMM(30) As String
    Public Shared strReturnInceptionDateNonStdMM(30) As String
    Public Shared strReturnInceptionDateNonStdGIAMM(30) As String
    Public Shared strHistPeriodEndingMM(30) As String
    Public Shared strHistCumulativeReturnMM(30) As String
    Public Shared strHistAverageAnnReturnMM(30) As String
    Public Shared strReturnAdoptionDateNonStdMM(30) As String
    Public Shared strReturnAdoptionDateNonStdGIAMM(30) As String

    Public Shared strMessage1MM As String
    Public Shared strMessage2MM As String
    Public Shared strMessage3MM As String
    Public Shared strMessage4MM As String

    Public Shared strRunNoRunMM As String

    'Public Shared strClientMisMatchList As String

    'Public Shared strSplitMMList() As String

    Public Shared strClientXMisMatch(150) As String

    Public Sub ReadTestValues(ByVal strTestPath As String)



        Dim RelayOutTest = New IniConfigSource(strTestPath)
        Header = RelayOutTest.Configs("HeaderValues")
        Basic = RelayOutTest.Configs("BasicValues")
        Engine = RelayOutTest.Configs("Engine")

        strStatusTest = Engine.Get("Status")

        If strStatusTest = "1" Then
            gErrorTest = True
            strMessage1Test = Engine.Get("Message_1")
            If strMessage1Test <> "" Then
                strMessage2Test = Engine.Get("Message_2")
            End If
            If strMessage2Test <> "" Then
                strMessage3Test = Engine.Get("Message_3")
            End If
            If strMessage3Test <> "" Then
                strMessage4Test = Engine.Get("Message_4")
            End If
        Else

            strCompanyNameTest = Header.Get("bridge_CompanyName")
            strClient1Test = Header.Get("bridge_Client")
            strAge1Test = Header.Get("bridge_Age")
            strAgeOlderTest = Header.Get("bridge_OlderAge")
            strIRateTest = Header.Get("bridge_InterestRate")
            strSex1Test = Header.Get("bridge_Sex")
            strInitialDBTest = Header.Get("bridge_InitialDB")
            strJointTest = Header.Get("bridge_Joint")
            strInitialDB2Test = Header.Get("bridge_InitialDB2")
            strClient2Test = Header.Get("bridge_Client2")
            strAge2Test = Header.Get("bridge_Age2")
            strSex2Test = Header.Get("bridge_Sex2")
            strContractTypeTest = Header.Get("bridge_ContractType")
            strSurrChargeYrsTest = Header.Get("bridge_SurrenderChargeYears")
            strHypoNorGTest = Header.Get("bridge_HypoNetOrGross")
            strZeroNetTest = Header.Get("bridge_ZeroNet")
            strHypoNetTest = Header.Get("bridge_HypoNet")
            strHypoGrossTest = Header.Get("brige_HypoGross")
            strHypoGISRateTest = Header.Get("bridge_GIA.HypoGISRate")
            strZeroGISRateTest = Header.Get("bridge_ZeroGrowthRate.GIS")
            strExpensesVAMandEOnlyTest = Header.Get("bridge_Expenses.VAMAndEOnly")
            strZeroGrowthRateTest = Header.Get("bridge_ZeroGrowthRate")
            strExpensesAdminOnlyTest = Header.Get("bridge_Expenses.AdminOnly")
            strFundExpensesVATest = Header.Get("bridge_FundExpense.VA")
            strFundExpenseGISTest = Header.Get("bridge_FundExpense.GIS")
            strFundExpenseEffDateGISTest = Header.Get("bridge_FundExpense.EffectiveDate.GIS")
            strVADBBeneRiderChargeTest = Header.Get("bridge_VADeathBenefitRiderCharge")
            strVAContractChargeTest = Header.Get("bridge_VAContractCharge.Amount")
            strVAContractChargeWaiverLimitTest = Header.Get("bridge_VAContractCharge.WaiverLimit")
            strFundExpenseEffDateTest = Header.Get("bridge_FundExpense.EffectiveDate")
            strEarlyAccessChargeTest = Header.Get("bridge_EarlyAccessCharge")
            strLivingBenefitRiderChargeTest = Header.Get("bridge_LivingBenefitRiderCharge")
            strInitialPremiumTest = Header.Get("bridge_InitialPremium")
            strPrintYearsTest = Header.Get("Policy.PrintYears")
            strDBTypeTest = Header.Get("bridge_DeathBenefitType")
            strFundCountTest = Header.Get("bridge_FundCount")
            strInvestStratTest = Header.Get("bridge_InvestmentStrategy")

            For ix = 1 To CInt(strFundCountTest)
                strFundCodeTest(ix) = Header.Get("bridge_FundCode" & ix)
                strFundPctTest(ix) = Header.Get("bridge_FundPct" & ix)
                strFundNameTest(ix) = Header.Get("bridge_FundName" & ix)

                strReturnYr1StdTest(ix) = Header.Get("bridge_ReturnYr1Std" & ix)
                strReturnYr5StdTest(ix) = Header.Get("bridge_ReturnYr5Std" & ix)
                strReturnYr10StdTest(ix) = Header.Get("bridge_ReturnYr10Std" & ix)
                strReturnAdoptionStdTest(ix) = Header.Get("bridge_ReturnSinceAdoptStd" & ix)
                strReturnAdoptionDateStdTest(ix) = Header.Get("bridge_AdoptDateStd" & ix)
                strReturnYr1StdGIATest(ix) = Header.Get("bridge_ReturnYr1StdGIA" & ix)
                strReturnYr5StdGIATest(ix) = Header.Get("bridge_ReturnYr5StdGIA" & ix)
                strReturnYr10StdGIATest(ix) = Header.Get("bridge_ReturnYr10StdGIA" & ix)
                strReturnAdoptionStdGIATest(ix) = Header.Get("bridge_ReturnSinceAdoptStdGIA" & ix)
                strReturnAdoptionDateStdGIATest(ix) = Header.Get("bridge_AdoptDateStdGIA" & ix)

                strReturnYr1NonStdSCTest(ix) = Header.Get("bridge_ReturnYr1NonStdSC" & ix)
                strReturnYr5NonStdSCTest(ix) = Header.Get("bridge_ReturnYr5NonStdSC" & ix)
                strReturnYr10NonStdSCTest(ix) = Header.Get("bridge_ReturnYr10NonStdSC" & ix)
                strReturnAdoptionDateNonStdSCTest(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                strReturnAdoptionDateNonStdSCTest(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                strReturnYr1NonStdSCGIATest(ix) = Header.Get("bridge_ReturnYr1NonStdSCGIA" & ix)
                strReturnYr5NonStdSCGIATest(ix) = Header.Get("bridge_ReturnYr5NonStdSCGIA" & ix)
                strReturnYr10NonStdSCGIATest(ix) = Header.Get("bridge_ReturnYr10NonStdSCGIA" & ix)
                strReturnAdoptionDateNonStdSCGIATest(ix) = Header.Get("bridge_ReturnIncepDateNonStdSCGIA" & ix)

                strReturnYr1NonStdTest(ix) = Header.Get("bridge_ReturnYr1NonStd" & ix)
                strReturnYr5NonStdTest(ix) = Header.Get("bridge_ReturnYr5NonStd" & ix)
                strReturnYr10NonStdTest(ix) = Header.Get("bridge_ReturnYr10NonStd" & ix)
                strReturnAdoptionDateNonStdTest(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                strReturnAdoptionNonStdTest(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                strReturnYr1NonStdGIATest(ix) = Header.Get("bridge_ReturnYr1NonStdGIA" & ix)
                strReturnYr5NonStdGIATest(ix) = Header.Get("bridge_ReturnYr5NonStdGIA" & ix)
                strReturnYr10NonStdGIATest(ix) = Header.Get("bridge_ReturnYr10NonStdGIA" & ix)
                strReturnAdoptionDateNonStdGIATest(ix) = Header.Get("bridge_ReturnSinceAdoptNonStd" & ix)
                strReturnInceptionDateNonStdGIATest(ix) = Header.Get("bridge_ReturnIncepDateNonStdGIA" & ix)

                strHistPeriodEndingTest(ix) = Header.Get("bridge_HistPeriodEnding" & ix)

                strHistCumulativeReturnTest(ix) = Header.Get("bridge_HistCumulativeReturn" & ix)
                strHistAverageAnnReturnTest(ix) = Header.Get("bridge_HistAverageAnnReturn" & ix)

            Next

            strIncomeStartAgeTest = Header.Get("bridge_IncomeStartAge")
            strIncomeStartAgeJointTest = Header.Get("bridge_IncomeStartJointAge")
            strIncomeStartYearTest = Header.Get("bridge_IncomeStartYears")
            strIncomeStartMonthTest = Header.Get("bridge_IncomeStartMonth")
            strYearsCertainTest = Header.Get("bridge_YearsCertain")
            strGIAInitialMonthlyPayoutHypoTest = Header.Get("bridge_GIAInitialMonthlyPayout.Hypo")
            strGIAInitialMonthlyPayoutZeroTest = Header.Get("bridge_GIAInitialMonthlyPayout.Zero")
            strGIASchedInstallmentTest = Header.Get("bridge_GIAScheduledInstallment")
            strAnnAmountZeroTest = Header.Get("bridge_AnnuitizedAmount.Zero")
            strAnnAmountHypoTest = Header.Get("bridge_AnnuitizedAmount.Hypo")
            strInstallmentCountZeroTest = Header.Get("bridge_InstallmentCount.Zero")
            strInstallmentCountHypoTest = Header.Get("bridge_InstallmentCount.Hypo")
            strPPDBChargeTest = Header.Get("bridge_PPDBCharge")
            strLIPFactorFirstWDTest = Header.Get("bridge_LIPWithdrawalLimit")
            strLIPGuarWithdrawalTest = Header.Get("bridge_LIPGuaranteedWithdrawal")
            strLIPWDStartYearTest = Header.Get("bridge_WithdrawalStartYear")
            strLIPWDStartMonthTest = Header.Get("bridge_WithdrawalStartMonth")
            strLIPBenBaseFirstWDTest = Header.Get("bridge_LIPBenefitBaseFirstWD")
            strTaxExcludableAmtZeroTest = Header.Get("bridge_TaxExcludableAmtZero")
            strTaxExcludableAmtHypoTest = Header.Get("bridge_TaxExcludableAmtHypo")
            strTaxExcludableAmtHistTest = Header.Get("bridge_TaxExcludableAmtHist")
            strTaxBracketTest = Header.Get("bridge_TaxBracket")
            strTaxBasisTest = Header.Get("bridge_TaxBasis")


            'BasicValues()

            strInvestmentTest = Basic.Get("value_AnnualPremium")
            strBaseContractValueZeroTest = Basic.Get("value_Account.Zero")
            strCombinedSurrValueZeroTest = Basic.Get("value_Surrender.Zero")
            strGISValueZeroTest = Basic.Get("value_Account.GIS.Zero")
            strAnnualIncomeZeroTest = Basic.Get("value_GIAAnnualPayout.Zero")
            strDeathBenefitZeroTest = Basic.Get("value_DeathBenefit.Zero")
            strTransfertoGISZeroTest = Basic.Get("value_TransferToGIS.Zero")
            strTotalContractValueZeroTest = Basic.Get("value_TotalContractValue.Zero")
            strSurrenderChargesTest = Basic.Get("value_SurrenderCharge")
            strHypoAnnIncomeFloorTest = Basic.Get("value_HypoAnnGtdIncomeFloor")
            strGIASumGtdAmtTest = Basic.Get("value_GIASumGtdAmt")
            strHistReturnForPeriodTest = Basic.Get("value_HistReturnForPeriod")
            strHistAnnIncomeTest = Basic.Get("value_HistAnnIncome")
            strHistAccountGISTest = Basic.Get("value_HistAccountGIS")
            strHistTotalContractValueTest = Basic.Get("value_HistTotalContractValue")
            strHistTotalSurrValueTest = Basic.Get("value_HistTotalSurrValue")
            strHistDeathBenefitGIATest = Basic.Get("value_HistDeathBenefitGIA")
            strPPBAHypoRateTest = Basic.Get("value_PPBAHypoRate")
            strPPBAZeroRateTest = Basic.Get("value_PPBAZeroRate")
            strBenefitBaseZeroRateTest = Basic.Get("value_BenefitBaseZeroRate")
            strBenefitBaseHypoRateTest = Basic.Get("value_BenefitBaseHypoRate")
            strRollupZeroRateTest = Basic.Get("value_RollupZeroRate")
            strRollupHypoRateTest = Basic.Get("value_RollupHypoRate")
            strLIPAnnIncomeZeroRateTest = Basic.Get("value_AnnualIncomeZeroRate")
            strLIPAnnIncomeHypoRateTest = Basic.Get("value_AnnualIncomeHypoRate")
            strLIPResetValueZeroRateTest = Basic.Get("value_LIPResetValueZeroRate")
            strLIPContractValueHypoTest = Basic.Get("value_Account.Hypo")
            strBASEWithdrawalZeroTest = Basic.Get("value_BaseWDZero")
            strBASEWithdrawalHypoTest = Basic.Get("value_BaseWDHypo")
            strEPRZeroTest = Basic.Get("value_EPRDBZero")
            strEPRHypoTest = Basic.Get("value_EPRDBHypo")
            strEPRHistTest = Basic.Get("value_EPRDBHist")
            strDBGIAZeroTest = Basic.Get("value_DBGIAZero")
            strDBGIAHypoTest = Basic.Get("value_DBGIAHypo")
            strDBLIPZeroTest = Basic.Get("value_DBLIPZero")
            strDBLIPHypoTest = Basic.Get("value_DBLIPHypo")
            strDBLIPHistTest = Basic.Get("value_DBLIPHist")
            strDBComboZeroTest = Basic.Get("value_DBComboZero")
            strDBComboHypoTest = Basic.Get("value_DBComboHypo")
            strDBComboHistTest = Basic.Get("value_DBComboHist")
            strDBASDBZeroTest = Basic.Get("value_DBASDBZero")
            strDBASDBHypoTest = Basic.Get("value_DBASDBHypo")
            strDBASDBHistTest = Basic.Get("value_DBASDBHist")
            strDBRollupZeroTest = Basic.Get("value_DBRollupZero")
            strDBRollupHypoTest = Basic.Get("value_DBRollupHypo")
            strDBRollupHistTest = Basic.Get("value_DBRolupHist")
            strDBStandardZeroTest = Basic.Get("value_DBStandardZero")
            strDBStandardHypoTest = Basic.Get("value_DBStandardHypo")
            strDBStandardHistTest = Basic.Get("value_DBStandardHist")
        End If
    End Sub
    Public Sub ReadBenchValues(ByVal strBenchPath As String)

        Dim RelayOutBench = New IniConfigSource(strBenchPath)
        Header = RelayOutBench.Configs("HeaderValues")
        Basic = RelayOutBench.Configs("BasicValues")
        Engine = RelayOutBench.Configs("Engine")

        'need to account for status 0 with warning message
        '****need to split this out, and read if message is warning or error
        'add these strings to compare...

        strStatusBench = Engine.Get("Status")


        If strStatusBench = "1" Then
            gErrorBench = True
            strMessage1Bench = Engine.Get("Message_1")
            If strMessage1Bench <> "" Then
                strMessage2Bench = Engine.Get("Message_2")
            End If
            If strMessage2Bench <> "" Then
                strMessage3Bench = Engine.Get("Message_3")
            End If
            If strMessage3Bench <> "" Then
                strMessage4Bench = Engine.Get("Message_4")
            End If
        Else

            strCompanyNameBench = Header.Get("bridge_CompanyName")
            strClient1Bench = Header.Get("bridge_Client")
            strAge1Bench = Header.Get("bridge_Age")
            strAgeOlderBench = Header.Get("bridge_OlderAge")
            strIRateBench = Header.Get("bridge_InterestRate")
            strSex1Bench = Header.Get("bridge_Sex")
            strInitialDBBench = Header.Get("bridge_InitialDB")
            strJointBench = Header.Get("bridge_Joint")
            strInitialDB2Bench = Header.Get("bridge_InitialDB2")
            strClient2Bench = Header.Get("bridge_Client2")
            strAge2Bench = Header.Get("bridge_Age2")
            strSex2Bench = Header.Get("bridge_Sex2")
            strContractTypeBench = Header.Get("bridge_ContractType")
            strSurrChargeYrsBench = Header.Get("bridge_SurrenderChargeYears")
            strHypoNorGBench = Header.Get("bridge_HypoNetOrGross")
            strZeroNetBench = Header.Get("bridge_ZeroNet")
            strHypoNetBench = Header.Get("bridge_HypoNet")
            strHypoGrossBench = Header.Get("brige_HypoGross")
            strHypoGISRateBench = Header.Get("bridge_GIA.HypoGISRate")
            strZeroGISRateBench = Header.Get("bridge_ZeroGrowthRate.GIS")
            strExpensesVAMandEOnlyBench = Header.Get("bridge_Expenses.VAMAndEOnly")
            strZeroGrowthRateBench = Header.Get("bridge_ZeroGrowthRate")
            strExpensesAdminOnlyBench = Header.Get("bridge_Expenses.AdminOnly")
            strFundExpensesVABench = Header.Get("bridge_FundExpense.VA")
            strFundExpenseGISBench = Header.Get("bridge_FundExpense.GIS")
            strFundExpenseEffDateGISBench = Header.Get("bridge_FundExpense.EffectiveDate.GIS")
            strVADBBeneRiderChargeBench = Header.Get("bridge_VADeathBenefitRiderCharge")
            strVAContractChargeBench = Header.Get("bridge_VAContractCharge.Amount")
            strVAContractChargeWaiverLimitBench = Header.Get("bridge_VAContractCharge.WaiverLimit")
            strFundExpenseEffDateBench = Header.Get("bridge_FundExpense.EffectiveDate")
            strEarlyAccessChargeBench = Header.Get("bridge_EarlyAccessCharge")
            strLivingBenefitRiderChargeBench = Header.Get("bridge_LivingBenefitRiderCharge")
            strInitialPremiumBench = Header.Get("bridge_InitialPremium")
            strPrintYearsBench = Header.Get("Policy.PrintYears")
            strDBTypeBench = Header.Get("bridge_DeathBenefitType")
            strFundCountBench = Header.Get("bridge_FundCount")
            strInvestStratBench = Header.Get("bridge_InvestmentStrategy")


            For ix = 1 To CInt(strFundCountBench)
                strFundCodeBench(ix) = Header.Get("bridge_FundCode" & ix)
                strFundPctBench(ix) = Header.Get("bridge_FundPct" & ix)
                strFundNameBench(ix) = Header.Get("bridge_FundName" & ix)

                strReturnYr1StdBench(ix) = Header.Get("bridge_ReturnYr1Std" & ix)
                strReturnYr5StdBench(ix) = Header.Get("bridge_ReturnYr5Std" & ix)
                strReturnYr10StdBench(ix) = Header.Get("bridge_ReturnYr10Std" & ix)
                strReturnAdoptionStdBench(ix) = Header.Get("bridge_ReturnSinceAdoptStd" & ix)
                strReturnAdoptionDateStdBench(ix) = Header.Get("bridge_AdoptDateStd" & ix)
                strReturnYr1StdGIABench(ix) = Header.Get("bridge_ReturnYr1StdGIA" & ix)
                strReturnYr5StdGIABench(ix) = Header.Get("bridge_ReturnYr5StdGIA" & ix)
                strReturnYr10StdGIABench(ix) = Header.Get("bridge_ReturnYr10StdGIA" & ix)
                strReturnAdoptionStdGIABench(ix) = Header.Get("bridge_ReturnSinceAdoptStdGIA" & ix)
                strReturnAdoptionDateStdGIABench(ix) = Header.Get("bridge_AdoptDateStdGIA" & ix)

                strReturnYr1NonStdSCBench(ix) = Header.Get("bridge_ReturnYr1NonStdSC" & ix)
                strReturnYr5NonStdSCBench(ix) = Header.Get("bridge_ReturnYr5NonStdSC" & ix)
                strReturnYr10NonStdSCBench(ix) = Header.Get("bridge_ReturnYr10NonStdSC" & ix)
                strReturnAdoptionDateNonStdSCBench(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                strReturnAdoptionDateNonStdSCBench(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                strReturnYr1NonStdSCGIABench(ix) = Header.Get("bridge_ReturnYr1NonStdSCGIA" & ix)
                strReturnYr5NonStdSCGIABench(ix) = Header.Get("bridge_ReturnYr5NonStdSCGIA" & ix)
                strReturnYr10NonStdSCGIABench(ix) = Header.Get("bridge_ReturnYr10NonStdSCGIA" & ix)
                strReturnAdoptionDateNonStdSCGIABench(ix) = Header.Get("bridge_ReturnIncepDateNonStdSCGIA" & ix)

                strReturnYr1NonStdBench(ix) = Header.Get("bridge_ReturnYr1NonStd" & ix)
                strReturnYr5NonStdBench(ix) = Header.Get("bridge_ReturnYr5NonStd" & ix)
                strReturnYr10NonStdBench(ix) = Header.Get("bridge_ReturnYr10NonStd" & ix)
                strReturnAdoptionDateNonStdBench(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                strReturnAdoptionNonStdBench(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                strReturnYr1NonStdGIABench(ix) = Header.Get("bridge_ReturnYr1NonStdGIA" & ix)
                strReturnYr5NonStdGIABench(ix) = Header.Get("bridge_ReturnYr5NonStdGIA" & ix)
                strReturnYr10NonStdGIABench(ix) = Header.Get("bridge_ReturnYr10NonStdGIA" & ix)
                strReturnAdoptionDateNonStdGIABench(ix) = Header.Get("bridge_ReturnSinceAdoptNonStd" & ix)
                strReturnInceptionDateNonStdGIABench(ix) = Header.Get("bridge_ReturnIncepDateNonStdGIA" & ix)

                strHistPeriodEndingBench(ix) = Header.Get("bridge_HistPeriodEnding" & ix)

                strHistCumulativeReturnBench(ix) = Header.Get("bridge_HistCumulativeReturn" & ix)
                strHistAverageAnnReturnBench(ix) = Header.Get("bridge_HistAverageAnnReturn" & ix)

            Next

            strIncomeStartAgeBench = Header.Get("bridge_IncomeStartAge")
            strIncomeStartAgeJointBench = Header.Get("bridge_IncomeStartJointAge")
            strIncomeStartYearBench = Header.Get("bridge_IncomeStartYears")
            strIncomeStartMonthBench = Header.Get("bridge_IncomeStartMonth")
            strYearsCertainBench = Header.Get("bridge_YearsCertain")
            strGIAInitialMonthlyPayoutHypoBench = Header.Get("bridge_GIAInitialMonthlyPayout.Hypo")
            strGIAInitialMonthlyPayoutZeroBench = Header.Get("bridge_GIAInitialMonthlyPayout.Zero")
            strGIASchedInstallmentBench = Header.Get("bridge_GIAScheduledInstallment")
            strAnnAmountZeroBench = Header.Get("bridge_AnnuitizedAmount.Zero")
            strAnnAmountHypoBench = Header.Get("bridge_AnnuitizedAmount.Hypo")
            strInstallmentCountZeroBench = Header.Get("bridge_InstallmentCount.Zero")
            strInstallmentCountHypoBench = Header.Get("bridge_InstallmentCount.Hypo")
            strPPDBChargeBench = Header.Get("bridge_PPDBCharge")
            strLIPFactorFirstWDBench = Header.Get("bridge_LIPWithdrawalLimit")
            strLIPGuarWithdrawalBench = Header.Get("bridge_LIPGuaranteedWithdrawal")
            strLIPWDStartYearBench = Header.Get("bridge_WithdrawalStartYear")
            strLIPWDStartMonthBench = Header.Get("bridge_WithdrawalStartMonth")
            strLIPBenBaseFirstWDBench = Header.Get("bridge_LIPBenefitBaseFirstWD")
            strTaxExcludableAmtZeroBench = Header.Get("bridge_TaxExcludableAmtZero")
            strTaxExcludableAmtHypoBench = Header.Get("bridge_TaxExcludableAmtHypo")
            strTaxExcludableAmtHistBench = Header.Get("bridge_TaxExcludableAmtHist")
            strTaxBracketBench = Header.Get("bridge_TaxBracket")
            strTaxBasisBench = Header.Get("bridge_TaxBasis")


            'BasicValues()

            strInvestmentBench = Basic.Get("value_AnnualPremium")
            strBaseContractValueZeroBench = Basic.Get("value_Account.Zero")
            strCombinedSurrValueZeroBench = Basic.Get("value_Surrender.Zero")
            strGISValueZeroBench = Basic.Get("value_Account.GIS.Zero")
            strAnnualIncomeZeroBench = Basic.Get("value_GIAAnnualPayout.Zero")
            strDeathBenefitZeroBench = Basic.Get("value_DeathBenefit.Zero")
            strTransfertoGISZeroBench = Basic.Get("value_TransferToGIS.Zero")
            strTotalContractValueZeroBench = Basic.Get("value_TotalContractValue.Zero")
            strSurrenderChargesBench = Basic.Get("value_SurrenderCharge")
            strHypoAnnIncomeFloorBench = Basic.Get("value_HypoAnnGtdIncomeFloor")
            strGIASumGtdAmtBench = Basic.Get("value_GIASumGtdAmt")
            strHistReturnForPeriodBench = Basic.Get("value_HistReturnForPeriod")
            strHistAnnIncomeBench = Basic.Get("value_HistAnnIncome")
            strHistAccountGISBench = Basic.Get("value_HistAccountGIS")
            strHistTotalContractValueBench = Basic.Get("value_HistTotalContractValue")
            strHistTotalSurrValueBench = Basic.Get("value_HistTotalSurrValue")
            strHistDeathBenefitGIABench = Basic.Get("value_HistDeathBenefitGIA")
            strPPBAHypoRateBench = Basic.Get("value_PPBAHypoRate")
            strPPBAZeroRateBench = Basic.Get("value_PPBAZeroRate")
            strBenefitBaseZeroRateBench = Basic.Get("value_BenefitBaseZeroRate")
            strBenefitBaseHypoRateBench = Basic.Get("value_BenefitBaseHypoRate")
            strRollupZeroRateBench = Basic.Get("value_RollupZeroRate")
            strRollupHypoRateBench = Basic.Get("value_RollupHypoRate")
            strLIPAnnIncomeZeroRateBench = Basic.Get("value_AnnualIncomeZeroRate")
            strLIPAnnIncomeHypoRateBench = Basic.Get("value_AnnualIncomeHypoRate")
            strLIPResetValueZeroRateBench = Basic.Get("value_LIPResetValueZeroRate")
            strLIPContractValueHypoBench = Basic.Get("value_Account.Hypo")
            strBASEWithdrawalZeroBench = Basic.Get("value_BaseWDZero")
            strBASEWithdrawalHypoBench = Basic.Get("value_BaseWDHypo")
            strEPRZeroBench = Basic.Get("value_EPRDBZero")
            strEPRHypoBench = Basic.Get("value_EPRDBHypo")
            strEPRHistBench = Basic.Get("value_EPRDBHist")
            strDBGIAZeroBench = Basic.Get("value_DBGIAZero")
            strDBGIAHypoBench = Basic.Get("value_DBGIAHypo")
            strDBLIPZeroBench = Basic.Get("value_DBLIPZero")
            strDBLIPHypoBench = Basic.Get("value_DBLIPHypo")
            strDBLIPHistBench = Basic.Get("value_DBLIPHist")
            strDBComboZeroBench = Basic.Get("value_DBComboZero")
            strDBComboHypoBench = Basic.Get("value_DBComboHypo")
            strDBComboHistBench = Basic.Get("value_DBComboHist")
            strDBASDBZeroBench = Basic.Get("value_DBASDBZero")
            strDBASDBHypoBench = Basic.Get("value_DBASDBHypo")
            strDBASDBHistBench = Basic.Get("value_DBASDBHist")
            strDBRollupZeroBench = Basic.Get("value_DBRollupZero")
            strDBRollupHypoBench = Basic.Get("value_DBRollupHypo")
            strDBRollupHistBench = Basic.Get("value_DBRolupHist")
            strDBStandardZeroBench = Basic.Get("value_DBStandardZero")
            strDBStandardHypoBench = Basic.Get("value_DBStandardHypo")
            strDBStandardHistBench = Basic.Get("value_DBStandardHist")
        End If
    End Sub

    Public Sub GetTimeStamps()

        'GLAIC VA:

        strWFProp = "WFPROP.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\WFPROP.EXE"))
        strGLAICCPY = "GELA.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\GELA.CPY"))
        strAnn1 = "ANNUITY2.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\ANNUITY.MDB"))
        strAnn2 = "ANNUITY2.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\ANNUITY2.MDB"))
        strAnn3 = "ANNUITY3.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\ANNUITY3.MDB"))
        strAnn4 = "ANNUITY4.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\ANNUITY4.MDB"))
        strWFGELA = "WFGELA.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\WFGELA.MDB"))

        'GLICNY VA:

        strGECLRIC = "GECLRIC.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECLRIC.MDB"))
        strGECLVA1 = "GECLVA.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECLVA.MDB"))
        strGECLVA2 = "GECLVA2.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECLVA2.MDB"))
        strGECLVA3 = "GECLVA3.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECLVA3.MDB"))
        strGECLEXE = "GECLRIC.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECLRIC.EXE"))
        strGLICNYCPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECL.CPY"))

        'GLICNY SPIA:

        strGLICNYSPIAEXE = "WINANN.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\WINANN.EXE"))
        strGLICNYSPIAANNRATES = "ANNRATES.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\ANNRATES.MDB"))
        strGLICNYSPIAANNPROD = "ANNPROD.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\ANNPROD.MDB"))
        strGLICNYSPIAANNSYS = "ANNSYS.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\ANNSYS.MDB"))
        strGLICNYSPIACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECL.CPY"))
        strGLICNYSPIAANNSUPP = "ANNSUPP.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\ANNSUPP.MDB"))
        strGLICNYSPIAANNVER = "ANNVER.MSB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\ANNVER.MDB"))

        'GLICNY SPDA:

        strGLICNYSPDAEXE = "GEDEFANN.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GEDEFANN.EXE"))
        strGLICNYSPDARATE = "RATE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\RATE.MDB"))
        strGLICNYSPDAGNAWINE = "GNAWINE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GNAWINE.MDB"))
        strGLICNYSPDACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GECL.CPY"))

        'GLIC SPIA:

        strGLICSPIAEXE = "WINANN.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\WINANN.EXE"))
        strGLICSPIAANNRATES = "ANNRATES.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\ANNRATES.MDB"))
        strGLICSPIAANNPROD = "ANNPROD.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\ANNPROD.MDB"))
        strGLICSPIAANNSYS = "ANNSYS.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\ANNSYS.MDB"))
        strGLICSPIACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GECA.CPY"))
        strGLICSPIAANNSUPP = "ANNSUPP.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\ANNSUPP.MDB"))
        strGLICSPIANNVER = "ANNVER.MSB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\ANNVER.MDB"))

        'GLIC SPDA:

        strGLICSPDAEXE = "GEDEFANN.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GEDEFANN.EXE"))
        strGLICSPDARATE = "RATE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\RATE.MDB"))
        strGLICSPDAGNAWINE = "GNAWINE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GNAWINE.MDB"))
        strGLICSPDACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GECA.CPY"))

        'GLAICFIXED SPIA:

        strGLAICFIXEDSPIAEXE = "WINANN.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\WINANN.EXE"))
        strGLAICFIXEDSPIAANNRATES = "ANNRATES.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNRATES.MDB"))
        strGLAICFIXEDSPIAANNPROD = "ANNPROD.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNPROD.MDB"))
        strGLAICFIXEDSPIAANNSYS = "ANNSYS.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNSYS.MDB"))
        strGLAICFIXEDSPIACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\FCOL.CPY"))
        strGLAICFIXEDSPIAANNSUPP = "ANNSUPP.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNSUPP.MDB"))
        strGLAICFIXEDSPIANNVER = "ANNVER.MSB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNVER.MDB"))

    End Sub
    Public Sub StoreMismatches()
    End Sub

End Class
