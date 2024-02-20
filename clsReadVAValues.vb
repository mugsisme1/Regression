Imports Nini.Config


Public Class clsReadVAValues

    'variables for relay.out sections
    Public Shared Header As IConfig
    Public Shared Basic As IConfig
    Public Shared Engine As IConfig

    'variables for test values read from relay.out
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
    'for IPR
    Public Shared strHistCumulativeReturnMaxTest(30) As String
    Public Shared strHistAverageAnnReturnMaxTest(30) As String

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
    Public Shared strExpensesTotalBaseContractTest As String
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
    Public Shared strCombinedSurrValueHypoTest As String
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
    Public Shared strIPRHistTotalContractValueCurrTest As String
    Public Shared strIPRHistTotalContractValueMaxTest As String
    Public Shared strIPRHistTotalSurrValueCurrTest As String
    Public Shared strIPRHistTotalSurrValueMaxTest As String
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

    'IPR Hist DB Values
    Public Shared strDBIPRHistContractDBCurrTest As String
    Public Shared strDBIPRHistASDBCurrTest As String
    Public Shared strDBIPRHistBasicDBCurrTest As String
    Public Shared strDBIPRHistRollUpDBCurrTest As String
    Public Shared strDBIPRHistEPRDBCurrTest As String
    Public Shared strDBIPRHistPPDBCurrTest As String

    Public Shared strDBIPRHistContractDBMaxTest As String
    Public Shared strDBIPRHistASDBMaxTest As String
    Public Shared strDBIPRHistBasicDBMaxTest As String
    Public Shared strDBIPRHistRollUpDBMaxTest As String
    Public Shared strDBIPRHistEPRDBMaxTest As String
    Public Shared strDBIPRHistPPDBMaxTest As String

    'IPR Hypo DB Values
    Public Shared strDBIPRHypoContractDBCurrTest As String
    Public Shared strDBIPRHypoASDBCurrTest As String
    Public Shared strDBIPRHypoBasicDBCurrTest As String
    Public Shared strDBIPRHypoRollUpDBCurrTest As String
    Public Shared strDBIPRHypoEPRDBCurrTest As String
    Public Shared strDBIPRHypoPPDBCurrTest As String

    Public Shared strDBIPRHypoContractDBMaxTest As String
    Public Shared strDBIPRHypoASDBMaxTest As String
    Public Shared strDBIPRHypoBasicDBMaxTest As String
    Public Shared strDBIPRHypoRollUpDBMaxTest As String
    Public Shared strDBIPRHypoEPRDBMaxTest As String
    Public Shared strDBIPRHypoPPDBMaxTest As String

    'IPR WD Limit
    Public Shared strIPRWDLimitCurrTest As String
    Public Shared strIPRWDLimitGtdTest As String
    Public Shared strIPRWDLimitHistCurrTest As String
    Public Shared strIPRWDLimitCurrGCTest As String
    Public Shared strIPRWDLimitMaxTest As String
    Public Shared strIPRWDLimitHistMaxTest As String

    Public Shared strRROneHistReturnForPeriodMaxTest As String
    Public Shared strRROneHistReturnForPeriodCurrTest As String

    Public Shared strStatusTest As String
    Public Shared strMessage1Test As String
    Public Shared strMessage2Test As String
    Public Shared strMessage3Test As String
    Public Shared strMessage4Test As String
    Public Shared strMessage5Test As String
    Public Shared strMessage6Test As String

    'IPR Hypo Values
    Public Shared strIPRHypoPPBAMaxTest As String
    Public Shared strIPRHypoMaxAnnValueMaxTest As String
    Public Shared strIPRHypoRollupValueMaxTest As String
    Public Shared strIPRHypoBenefitBaseMaxTest As String
    Public Shared strIPRHypoDBMaxTest As String

    Public Shared strIPRHypoPPBACurrTest As String
    Public Shared strIPRHypoMaxAnnValueCurrTest As String
    Public Shared strIPRHypoRollupValueCurrTest As String
    Public Shared strIPRHypoBenefitBaseCurrTest As String
    Public Shared strIPRHypoDBCurrTest As String

    Public Shared strIPRHypoContractValueCurrTest As String
    Public Shared strIPRHypoSurrenderValueCurrTest As String
    Public Shared strIPRHypoContractValueMaxTest As String
    Public Shared strIPRHypoSurrenderValueMaxTest As String

    'IPR Hist Values
    Public Shared strIPRHistPPBAMaxTest As String
    Public Shared strIPRHistMaxAnnValueMaxTest As String
    Public Shared strIPRHistRollupValueMaxTest As String
    Public Shared strIPRHistBenefitBaseMaxTest As String
    Public Shared strIPRHistDBMaxTest As String

    Public Shared strIPRHistPPBACurrTest As String
    Public Shared strIPRHistMaxAnnValueCurrTest As String
    Public Shared strIPRHistRollupValueCurrTest As String
    Public Shared strIPRHistBenefitBaseCurrTest As String
    Public Shared strIPRHistDBCurrTest As String

    'IPR WD Taken
    Public Shared strIPRWDTakenHypoCurrTest As String
    Public Shared strIPRWDTakenHypoMaxTest As String
    Public Shared strIPRWDTakenHistCurrTest As String
    Public Shared strIPRWDTakenHistMaxTest As String

    'RR One Annual Income, not split between period certain and living only
    Public Shared strRROneAnnualIncomeZeroTest As String
    Public Shared strRROneAnnualIncomeCurrTest As String
    Public Shared strRROneAnnualIncomeHistTest As String

    'Start Wds on Younger Annuitants Bday
    Public Shared strStartWDsYoungerBDTest As String

    'If Living Benefit is IPR
    Public Shared strLivingBenefitRiderTest As String
    Public Shared gbIPRTest As Boolean

    'For MyClearCourse
    'MCC Gtd Payment Floor
    Public Shared strMCCGtdPaymentFloorZeroTest As String
    Public Shared strMCCGtdPaymentFloorHypoTest As String
    Public Shared strMCCGtdPaymentFloorHistTest As String

    'MCC Payment Floor Factor
    Public Shared strMCCGtdPaymentFloorFactorZeroTest As String
    Public Shared strMCCGtdPaymentFloorFactorHypoTest As String
    Public Shared strMCCGtdPaymentFloorFactorHistTest As String

    'MCC First Full Year Income
    Public Shared strMCCFirstFullYrIncomeZeroTest As String
    Public Shared strMCCFirstFullYrIncomeHypoTest As String
    Public Shared strMCCFirstFullYrIncomeHistTest As String

    'MCC Life with Period Certain of
    Public Shared strMCCLifeWithPeriodCertainOfTest As String

    'MCC Guarantee From Plan
    Public Shared strMCCGuaranteeFromPlanTest As String

    'MCC Desired Retirement Age
    Public Shared strMCCDesiredRetirementAgeTest As String

    'MCC Gtd Income Payments at 0% (chart)
    Public Shared strMCCGtdIncomePayments0Test As String

    'MCC First Full Yr Income Hypo Net (chart)
    Public Shared strMCCFirstFullYrIncomeHypoNetTest As String

    'MCC Total Gtd Income Payments at 0% (chart)
    Public Shared strMCCTotalGtdIncomePayments0Test As String

    'MCC Total Income Payments Hypo Net (chart)
    Public Shared strMCCTotalIncomePaymentsHypoNetTest As String

    'MCC Adjustment Account
    Public Shared strMCCAdjustmentAccountZeroTest As String
    Public Shared strMCCAdjustmentAccountHypoTest As String
    Public Shared strMCCAdjustmentAccountHistTest As String

    'MCC Commutation Value
    Public Shared strMCCcommutationValueZeroTest As String
    Public Shared strMCCCommutationValueHypoTest As String
    Public Shared strMCCCommutationValueHistTest As String

    'MCC Historical INCOME Period Returns
    Public Shared strMCCHistIncomePeriodReturnTest As String



    'variables for bench values read from relay.out
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
    Public Shared strReturnInceptionDateNonStdSCBench(30) As String
    Public Shared strReturnAdoptionDateNonStdSCGIABench(30) As String
    Public Shared strReturnInceptionDateNonStdBench(30) As String
    Public Shared strReturnInceptionDateNonStdGIABench(30) As String
    Public Shared strHistPeriodEndingBench(30) As String
    Public Shared strHistCumulativeReturnBench(30) As String
    Public Shared strHistAverageAnnReturnBench(30) As String

    'for IPR
    Public Shared strHistCumulativeReturnMaxBench(30) As String
    Public Shared strHistAverageAnnReturnMaxBench(30) As String

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
    Public Shared strExpensesTotalBaseContractBench As String
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
    Public Shared strCombinedSurrValueHypoBench As String
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
    Public Shared strIPRHistTotalContractValueCurrBench As String
    Public Shared strIPRHistTotalContractValueMaxBench As String
    Public Shared strIPRHistTotalSurrValueCurrBench As String
    Public Shared strIPRHistTotalSurrValueMaxBench As String
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

    'IPR Hist DB Values
    Public Shared strDBIPRHistContractDBCurrBench As String
    Public Shared strDBIPRHistASDBCurrBench As String
    Public Shared strDBIPRHistBasicDBCurrBench As String
    Public Shared strDBIPRHistRollUpDBCurrBench As String
    Public Shared strDBIPRHistEPRDBCurrBench As String
    Public Shared strDBIPRHistPPDBCurrBench As String

    Public Shared strDBIPRHistContractDBMaxBench As String
    Public Shared strDBIPRHistASDBMaxBench As String
    Public Shared strDBIPRHistBasicDBMaxBench As String
    Public Shared strDBIPRHistRollUpDBMaxBench As String
    Public Shared strDBIPRHistEPRDBMaxBench As String
    Public Shared strDBIPRHistPPDBMaxBench As String

    'IPR Hypo DB Values
    Public Shared strDBIPRHypoContractDBCurrBench As String
    Public Shared strDBIPRHypoASDBCurrBench As String
    Public Shared strDBIPRHypoBasicDBCurrBench As String
    Public Shared strDBIPRHypoRollUpDBCurrBench As String
    Public Shared strDBIPRHypoEPRDBCurrBench As String
    Public Shared strDBIPRHypoPPDBCurrBench As String

    Public Shared strDBIPRHypoContractDBMaxBench As String
    Public Shared strDBIPRHypoASDBMaxBench As String
    Public Shared strDBIPRHypoBasicDBMaxBench As String
    Public Shared strDBIPRHypoRollUpDBMaxBench As String
    Public Shared strDBIPRHypoEPRDBMaxBench As String
    Public Shared strDBIPRHypoPPDBMaxBench As String


    Public Shared strIPRWDLimitCurrBench As String
    Public Shared strIPRWDLimitGtdBench As String
    Public Shared strIPRWDLimitHistCurrBench As String
    Public Shared strIPRWDLimitCurrGCBench As String
    Public Shared strIPRWDLimitMaxBench As String
    Public Shared strIPRWDLimitHistMaxBench As String

    Public Shared strRROneHistReturnForPeriodMaxBench As String
    Public Shared strRROneHistReturnForPeriodCurrBench As String

    Public Shared strStatusBench As String
    Public Shared strMessage1Bench As String
    Public Shared strMessage2Bench As String
    Public Shared strMessage3Bench As String
    Public Shared strMessage4Bench As String
    Public Shared strMessage5Bench As String
    Public Shared strMessage6Bench As String

    'IPR Hypo Values
    Public Shared strIPRHypoPPBAMaxBench As String
    Public Shared strIPRHypoMaxAnnValueMaxBench As String
    Public Shared strIPRHypoRollupValueMaxBench As String
    Public Shared strIPRHypoBenefitBaseMaxBench As String
    Public Shared strIPRHypoDBMaxBench As String

    Public Shared strIPRHypoPPBACurrBench As String
    Public Shared strIPRHypoMaxAnnValueCurrBench As String
    Public Shared strIPRHypoRollupValueCurrBench As String
    Public Shared strIPRHypoBenefitBaseCurrBench As String
    Public Shared strIPRHypoDBCurrBench As String

    Public Shared strIPRHypoContractValueCurrBench As String
    Public Shared strIPRHypoSurrenderValueCurrBench As String
    Public Shared strIPRHypoContractValueMaxBench As String
    Public Shared strIPRHypoSurrenderValueMaxBench As String


    'IPR Hist Values
    Public Shared strIPRHistPPBAMaxBench As String
    Public Shared strIPRHistMaxAnnValueMaxBench As String
    Public Shared strIPRHistRollupValueMaxBench As String
    Public Shared strIPRHistBenefitBaseMaxBench As String
    Public Shared strIPRHistDBMaxBench As String

    Public Shared strIPRHistPPBACurrBench As String
    Public Shared strIPRHistMaxAnnValueCurrBench As String
    Public Shared strIPRHistRollupValueCurrBench As String
    Public Shared strIPRHistBenefitBaseCurrBench As String
    Public Shared strIPRHistDBCurrBench As String

    'IPR WD Taken
    Public Shared strIPRWDTakenHypoCurrBench As String
    Public Shared strIPRWDTakenHypoMaxBench As String
    Public Shared strIPRWDTakenHistCurrBench As String
    Public Shared strIPRWDTakenHistMaxBench As String

    'RR One Annual Income, not split between period certain and living only
    Public Shared strRROneAnnualIncomeZeroBench As String
    Public Shared strRROneAnnualIncomeCurrBench As String
    Public Shared strRROneAnnualIncomeHistBench As String

    'Start Wds on Younger Annuitants Bday
    Public Shared strStartWDsYoungerBDBench As String


    Public Shared bErrorBench As Boolean = False
    Public Shared bErrorTest As Boolean = False

    'If Living Benefit is IPR
    Public Shared strLivingBenefitRiderBench As String
    Public Shared gbIPRBench As Boolean

    'For MyClearCourse
    'MCC Gtd Payment Floor
    Public Shared strMCCGtdPaymentFloorZeroBench As String
    Public Shared strMCCGtdPaymentFloorHypoBench As String
    Public Shared strMCCGtdPaymentFloorHistBench As String

    'MCC Payment Floor Factor
    Public Shared strMCCGtdPaymentFloorFactorZeroBench As String
    Public Shared strMCCGtdPaymentFloorFactorHypoBench As String
    Public Shared strMCCGtdPaymentFloorFactorHistBench As String

    'MCC First Full Year Income
    Public Shared strMCCFirstFullYrIncomeZeroBench As String
    Public Shared strMCCFirstFullYrIncomeHypoBench As String
    Public Shared strMCCFirstFullYrIncomeHistBench As String

    'MCC Life with Period Certain of
    Public Shared strMCCLifeWithPeriodCertainOfBench As String

    'MCC Guarantee From Plan
    Public Shared strMCCGuaranteeFromPlanBench As String

    'MCC Desired Retirement Age
    Public Shared strMCCDesiredRetirementAgeBench As String

    'MCC Gtd Income Payments at 0% (chart)
    Public Shared strMCCGtdIncomePayments0Bench As String

    'MCC First Full Yr Income Hypo Net (chart)
    Public Shared strMCCFirstFullYrIncomeHypoNetBench As String

    'MCC Total Gtd Income Payments at 0% (chart)
    Public Shared strMCCTotalGtdIncomePayments0Bench As String

    'MCC Total Income Payments Hypo Net (chart)
    Public Shared strMCCTotalIncomePaymentsHypoNetBench As String

    'MCC Adjustment Account
    Public Shared strMCCAdjustmentAccountZeroBench As String
    Public Shared strMCCAdjustmentAccountHypoBench As String
    Public Shared strMCCAdjustmentAccountHistBench As String

    'MCC Commutation Value
    Public Shared strMCCcommutationValueZeroBench As String
    Public Shared strMCCCommutationValueHypoBench As String
    Public Shared strMCCCommutationValueHistBench As String

    'MCC Historical INCOME Period Returns
    Public Shared strMCCHistIncomePeriodReturnBench As String



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
    Public Shared strFCOLSPDACPY As String = CStr(Today)
    Public Shared strFCOLSPDAGNAWINE As String = CStr(Today)
    Public Shared strFCOLSPDARATE As String = CStr(Today)
    Public Shared strFCOLSPDAEXE As String = CStr(Today)
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

    Public Shared strFIAMDB As String = CStr(Today)
    Public Shared strFIAEXE As String = CStr(Today)
    Public Shared strFIACPY As String = CStr(Today)

    'variables for mismatches
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
    Public Shared strExpensesTotalBaseContractMM As String
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
    Public Shared strCombinedSurrValueHypoMM As String
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
    Public Shared strIPRHistTotalContractValueCurrMM As String
    Public Shared strIPRHistTotalContractValueMaxMM As String
    Public Shared strIPRHistTotalSurrValueCurrMM As String
    Public Shared strIPRHistTotalSurrValueMaxMM As String
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
    Public Shared strReturnInceptionDateNonStdSCMM(30) As String
    Public Shared strReturnAdoptionDateNonStdSCGIAMM(30) As String
    Public Shared strReturnInceptionDateNonStdMM(30) As String
    Public Shared strReturnInceptionDateNonStdGIAMM(30) As String
    Public Shared strHistPeriodEndingMM(30) As String
    Public Shared strHistCumulativeReturnMM(30) As String
    Public Shared strHistAverageAnnReturnMM(30) As String
    Public Shared strReturnAdoptionDateNonStdMM(30) As String
    Public Shared strReturnAdoptionDateNonStdGIAMM(30) As String
    Public Shared strReturnInceptionDateNonStdSCTest(30) As String

    'for IPR
    Public Shared strHistCumulativeReturnMaxMM(30) As String
    Public Shared strHistAverageAnnReturnMaxMM(30) As String

    'RR One Hist DB values
    Public Shared strDBIPRHistBasicDBmaxMM As String
    Public Shared strDBIPRHistASDBCurrMM As String
    Public Shared strDBIPRHistBasicDBCurrMM As String
    Public Shared strDBIPRHistRollUpDBCurrMM As String
    Public Shared strDBIPRHistEPRDBCurrMM As String
    Public Shared strDBIPRHistPPDBCurrMM As String
    Public Shared strDBIPRHistContractDBMaxMM As String
    Public Shared strDBIPRHistASDBMaxMM As String
    Public Shared strDBIPRHistContractDBCurrMM As String
    Public Shared strDBIPRHistRollUpDBMaxMM As String
    Public Shared strDBIPRHistEPRDBMaxMM As String
    Public Shared strDBIPRHistPPDBMaxMM As String

    'RR One Hypo DB values
    Public Shared strDBIPRHypoBasicDBmaxMM As String
    Public Shared strDBIPRHypoASDBCurrMM As String
    Public Shared strDBIPRHypoBasicDBCurrMM As String
    Public Shared strDBIPRHypoRollUpDBCurrMM As String
    Public Shared strDBIPRHypoEPRDBCurrMM As String
    Public Shared strDBIPRHypoPPDBCurrMM As String
    Public Shared strDBIPRHypoContractDBMaxMM As String
    Public Shared strDBIPRHypoASDBMaxMM As String
    Public Shared strDBIPRHypoContractDBCurrMM As String
    Public Shared strDBIPRHypoRollUpDBMaxMM As String
    Public Shared strDBIPRHypoEPRDBMaxMM As String
    Public Shared strDBIPRHypoPPDBMaxMM As String


    'RR One WD Limits
    Public Shared strIPRWDLimitGtdMM As String
    Public Shared strIPRWDLimitCurrMM As String
    Public Shared strIPRWDLimitHistMM As String
    Public Shared strIPRWDLimitCurrGCMM As String
    Public Shared strIPRWDLimitGtdGCMM As String
    Public Shared strIPRWDLimitHistGCMM As String

    'RR One Return for period
    Public Shared strRROneHistReturnForPeriodMaxMM As String
    Public Shared strRROneHistReturnForPeriodCurrMM As String

    'IPR Hypo Values
    Public Shared strIPRHypoPPBAMaxMM As String
    Public Shared strIPRHypoMaxAnnValueMaxMM As String
    Public Shared strIPRHypoRollupValueMaxMM As String
    Public Shared strIPRHypoBenefitBaseMaxMM As String
    Public Shared strIPRHypoDBMaxMM As String

    Public Shared strIPRHypoPPBACurrMM As String
    Public Shared strIPRHypoMaxAnnValueCurrMM As String
    Public Shared strIPRHypoRollupValueCurrMM As String
    Public Shared strIPRHypoBenefitBaseCurrMM As String
    Public Shared strIPRHypoDBCurrMM As String

    Public Shared strIPRHypoContractValueCurrMM As String
    Public Shared strIPRHypoSurrenderValueCurrMM As String
    Public Shared strIPRHypoContractValueMaxMM As String
    Public Shared strIPRHypoSurrenderValueMaxMM As String


    'IPR Hist Values
    Public Shared strIPRHistPPBAMaxMM As String
    Public Shared strIPRHistMaxAnnValueMaxMM As String
    Public Shared strIPRHistRollupValueMaxMM As String
    Public Shared strIPRHistBenefitBaseMaxMM As String
    Public Shared strIPRHistDBMaxMM As String

    Public Shared strIPRHistPPBACurrMM As String
    Public Shared strIPRHistMaxAnnValueCurrMM As String
    Public Shared strIPRHistRollupValueCurrMM As String
    Public Shared strIPRHistBenefitBaseCurrMM As String
    Public Shared strIPRHistDBCurrMM As String

    'IPR WD Taken
    Public Shared strIPRWDTakenHypoCurrMM As String
    Public Shared strIPRWDTakenHypoMaxMM As String
    Public Shared strIPRWDTakenHistCurrMM As String
    Public Shared strIPRWDTakenHistMaxMM As String

    'RR One Annual Income, not split between period certain and living only
    Public Shared strRROneAnnualIncomeZeroMM As String
    Public Shared strRROneAnnualIncomeCurrMM As String
    Public Shared strRROneAnnualIncomeHistMM As String

    'For MyClearCourse
    'MCC Gtd Payment Floor
    Public Shared strMCCGtdPaymentFloorZeroMM As String
    Public Shared strMCCGtdPaymentFloorHypoMM As String
    Public Shared strMCCGtdPaymentFloorHistMM As String

    'MCC Payment Floor Factor
    Public Shared strMCCGtdPaymentFloorFactorZeroMM As String
    Public Shared strMCCGtdPaymentFloorFactorHypoMM As String
    Public Shared strMCCGtdPaymentFloorFactorHistMM As String

    'MCC First Full Year Income
    Public Shared strMCCFirstFullYrIncomeZeroMM As String
    Public Shared strMCCFirstFullYrIncomeHypoMM As String
    Public Shared strMCCFirstFullYrIncomeHistMM As String

    'MCC Life with Period Certain of
    Public Shared strMCCLifeWithPeriodCertainOfMM As String

    'MCC Guarantee From Plan
    Public Shared strMCCGuaranteeFromPlanMM As String

    'MCC Desired Retirement Age
    Public Shared strMCCDesiredRetirementAgeMM As String

    'MCC Gtd Income Payments at 0% (chart)
    Public Shared strMCCGtdIncomePayments0MM As String

    'MCC First Full Yr Income Hypo Net (chart)
    Public Shared strMCCFirstFullYrIncomeHypoNetMM As String

    'MCC Total Gtd Income Payments at 0% (chart)
    Public Shared strMCCTotalGtdIncomePayments0MM As String

    'MCC Total Income Payments Hypo Net (chart)
    Public Shared strMCCTotalIncomePaymentsHypoNetMM As String

    'MCC Adjustment Account
    Public Shared strMCCAdjustmentAccountZeroMM As String
    Public Shared strMCCAdjustmentAccountHypoMM As String
    Public Shared strMCCAdjustmentAccountHistMM As String

    'MCC Commutation Value
    Public Shared strMCCcommutationValueZeroMM As String
    Public Shared strMCCCommutationValueHypoMM As String
    Public Shared strMCCCommutationValueHistMM As String

    'MCC Historical INCOME Period Returns
    Public Shared strMCCHistIncomePeriodReturnMM As String



    'Start Wds on Younger Annuitants Bday
    Public Shared strStartWDsYoungerBDMM As String


    Public Shared strMessage1MM As String
    Public Shared strMessage2MM As String
    Public Shared strMessage3MM As String
    Public Shared strMessage4MM As String
    Public Shared strMessage5MM As String
    Public Shared strMessage6MM As String

    Public Shared strRunNoRunMM As String


    Public Shared strClientXMisMatch(150) As String

    Public Sub ReadTestValues(ByVal strTestPath As String)

        'read the values from the test relay.out files

        Dim RelayOutTest = New IniConfigSource(strTestPath)
        Dim i As Integer
        Dim icomma As Integer = 0



        'variables for relay.out sections
        Header = RelayOutTest.Configs("HeaderValues")
        Basic = RelayOutTest.Configs("BasicValues")
        Engine = RelayOutTest.Configs("Engine")

        strStatusTest = Engine.Get("Status")

        'if client does not run
        If strStatusTest = "1" Then
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
                End If
            End If
        Else

            'if client does run

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
            strHypoGrossTest = Header.Get("bridge_HypoGross")
            strHypoGISRateTest = Header.Get("bridge_GIA.HypoGISRate")
            strZeroGISRateTest = Header.Get("bridge_ZeroGrowthRate.GIS")
            strExpensesVAMandEOnlyTest = Header.Get("bridge_Expenses.VAMAndEOnly")
            strExpensesTotalBaseContractTest = Header.Get("bridge_TotalBaseContractCharges.VA")
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
            strLivingBenefitRiderTest = Header.Get("bridge_LivingBenefitRider")
            If strLivingBenefitRiderTest = "IPR" Then
                gbIPRTest = True
            Else
                gbIPRTest = False
            End If

            For ix = 1 To CInt(strFundCountTest)
                strFundCodeTest(ix) = Header.Get("bridge_FundCode" & ix)
                strFundPctTest(ix) = Header.Get("bridge_FundPct" & ix)
                strFundNameTest(ix) = Header.Get("bridge_FundName" & ix)

                strReturnYr1StdTest(ix) = Header.Get("bridge_ReturnYr1Std" & ix)
                If strReturnYr1StdTest(ix) = "-1000" Then strReturnYr1StdTest(ix) = "N/A"
                strReturnYr5StdTest(ix) = Header.Get("bridge_ReturnYr5Std" & ix)
                If strReturnYr5StdTest(ix) = "-1000" Then strReturnYr5StdTest(ix) = "N/A"
                strReturnYr10StdTest(ix) = Header.Get("bridge_ReturnYr10Std" & ix)
                If strReturnYr10StdTest(ix) = "-1000" Then strReturnYr10StdTest(ix) = "N/A"
                strReturnAdoptionStdTest(ix) = Header.Get("bridge_ReturnSinceAdoptStd" & ix)
                If strReturnAdoptionStdTest(ix) = "-1000" Then strReturnAdoptionStdTest(ix) = "N/A"
                strReturnAdoptionDateStdTest(ix) = Header.Get("bridge_AdoptDateStd" & ix)
                If strReturnAdoptionDateStdTest(ix) = "-1000" Then strReturnAdoptionDateStdTest(ix) = "N/A"
                strReturnYr1StdGIATest(ix) = Header.Get("bridge_ReturnYr1StdGIA" & ix)
                If strReturnYr1StdGIATest(ix) = "-1000" Then strReturnYr1StdGIATest(ix) = "N/A"
                strReturnYr5StdGIATest(ix) = Header.Get("bridge_ReturnYr5StdGIA" & ix)
                If strReturnYr5StdGIATest(ix) = "-1000" Then strReturnYr5StdGIATest(ix) = "N/A"
                strReturnYr10StdGIATest(ix) = Header.Get("bridge_ReturnYr10StdGIA" & ix)
                If strReturnYr10StdGIATest(ix) = "-1000" Then strReturnYr10StdGIATest(ix) = "N/A"
                strReturnAdoptionStdGIATest(ix) = Header.Get("bridge_ReturnSinceAdoptStdGIA" & ix)
                If strReturnAdoptionStdGIATest(ix) = "-1000" Then strReturnAdoptionStdGIATest(ix) = "N/A"
                strReturnAdoptionDateStdGIATest(ix) = Header.Get("bridge_AdoptDateStdGIA" & ix)
                If strReturnAdoptionDateStdGIATest(ix) = "-1000" Then strReturnAdoptionDateStdGIATest(ix) = "N/A"
                strReturnYr1NonStdSCTest(ix) = Header.Get("bridge_ReturnYr1NonStdSC" & ix)
                If strReturnYr1NonStdSCTest(ix) = "-1000" Then strReturnYr1NonStdSCTest(ix) = "N/A"
                strReturnYr5NonStdSCTest(ix) = Header.Get("bridge_ReturnYr5NonStdSC" & ix)
                If strReturnYr5NonStdSCTest(ix) = "-1000" Then strReturnYr5NonStdSCTest(ix) = "N/A"
                strReturnYr10NonStdSCTest(ix) = Header.Get("bridge_ReturnYr10NonStdSC" & ix)
                If strReturnYr10NonStdSCTest(ix) = "-1000" Then strReturnYr10NonStdSCTest(ix) = "N/A"
                strReturnAdoptionDateNonStdSCTest(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                If strReturnAdoptionDateNonStdSCTest(ix) = "-1000" Then strReturnAdoptionDateNonStdSCTest(ix) = "N/A"
                strReturnInceptionDateNonStdSCTest(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                If strReturnInceptionDateNonStdSCTest(ix) = "-1000" Then strReturnInceptionDateNonStdSCTest(ix) = "N/A"
                strReturnYr1NonStdSCGIATest(ix) = Header.Get("bridge_ReturnYr1NonStdSCGIA" & ix)
                strReturnYr5NonStdSCGIATest(ix) = Header.Get("bridge_ReturnYr5NonStdSCGIA" & ix)
                strReturnYr10NonStdSCGIATest(ix) = Header.Get("bridge_ReturnYr10NonStdSCGIA" & ix)
                strReturnAdoptionDateNonStdSCGIATest(ix) = Header.Get("bridge_ReturnIncepDateNonStdSCGIA" & ix)
                strReturnYr1NonStdTest(ix) = Header.Get("bridge_ReturnYr1NonStd" & ix)
                If strReturnYr1NonStdTest(ix) = "-1000" Then strReturnYr1NonStdTest(ix) = "N/A"
                strReturnYr5NonStdTest(ix) = Header.Get("bridge_ReturnYr5NonStd" & ix)
                If strReturnYr5NonStdTest(ix) = "-1000" Then strReturnYr5NonStdTest(ix) = "N/A"
                strReturnYr10NonStdTest(ix) = Header.Get("bridge_ReturnYr10NonStd" & ix)
                If strReturnYr10NonStdTest(ix) = "-1000" Then strReturnYr10NonStdTest(ix) = "N/A"
                strReturnAdoptionDateNonStdTest(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                If strReturnAdoptionDateNonStdTest(ix) = "-1000" Then strReturnAdoptionDateNonStdTest(ix) = "N/A"
                strReturnAdoptionNonStdTest(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                If strReturnAdoptionNonStdTest(ix) = "-1000" Then strReturnAdoptionNonStdTest(ix) = "N/A"
                strReturnYr1NonStdGIATest(ix) = Header.Get("bridge_ReturnYr1NonStdGIA" & ix)
                If strReturnYr1NonStdGIATest(ix) = "-1000" Then strReturnYr1NonStdGIATest(ix) = "N/A"
                strReturnYr5NonStdGIATest(ix) = Header.Get("bridge_ReturnYr5NonStdGIA" & ix)
                If strReturnYr5NonStdGIATest(ix) = "-1000" Then strReturnYr5NonStdGIATest(ix) = "N/A"
                strReturnYr10NonStdGIATest(ix) = Header.Get("bridge_ReturnYr10NonStdGIA" & ix)
                If strReturnYr10NonStdGIATest(ix) = "-1000" Then strReturnYr10NonStdGIATest(ix) = "N/A"
                strReturnAdoptionDateNonStdGIATest(ix) = Header.Get("bridge_ReturnSinceAdoptNonStd" & ix)
                If strReturnAdoptionDateNonStdGIATest(ix) = "-1000" Then strReturnAdoptionDateNonStdGIATest(ix) = "N/A"
                strReturnInceptionDateNonStdGIATest(ix) = Header.Get("bridge_ReturnIncepDateNonStdGIA" & ix)
                If strReturnInceptionDateNonStdGIATest(ix) = "-1000" Then strReturnInceptionDateNonStdGIATest(ix) = "N/A"
                strHistPeriodEndingTest(ix) = Header.Get("bridge_HistPeriodEnding" & ix)
                If strHistPeriodEndingTest(ix) = "-1000" Then strHistPeriodEndingTest(ix) = "N/A"
                strHistCumulativeReturnTest(ix) = Header.Get("bridge_HistCumulativeReturn" & ix)
                If strHistCumulativeReturnTest(ix) = "-1000" Then strHistCumulativeReturnTest(ix) = "N/A"
                strHistAverageAnnReturnTest(ix) = Header.Get("bridge_HistAverageAnnReturn" & ix)
                If strHistAverageAnnReturnTest(ix) = "-1000" Then strHistAverageAnnReturnTest(ix) = "N/A"
                'for IPR
                strHistCumulativeReturnMaxTest(ix) = Header.Get("bridge_HistCumulativeReturnMax" & ix)
                If strHistCumulativeReturnMaxTest(ix) = "-1000" Then strHistCumulativeReturnMaxTest(ix) = "N/A"
                strHistAverageAnnReturnMaxTest(ix) = Header.Get("bridge_HistAverageAnnReturnMax" & ix)
                If strHistAverageAnnReturnMaxTest(ix) = "-1000" Then strHistAverageAnnReturnMaxTest(ix) = "N/A"


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
            strCombinedSurrValueHypoTest = Basic.Get("value_Surrender.Hypo")
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
            strIPRHistTotalContractValueCurrTest = Basic.Get("value_IPRHistTotalContractValueCurr")
            strIPRHistTotalContractValueMaxTest = Basic.Get("value_IPRHistTotalContractValueMax")
            strIPRHistTotalSurrValueCurrTest = Basic.Get("value_IPRHistTotalSurrValueCurr")
            strIPRHistTotalSurrValueMaxTest = Basic.Get("value_IPRHistTotalSurrValueMax")
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

            'IPR Hist DB Values
            strDBIPRHistContractDBCurrTest = Basic.Get("value_IPRHistContractDBCurr")
            strDBIPRHistASDBCurrTest = Basic.Get("value_IPRHistASDBCurr")
            strDBIPRHistBasicDBCurrTest = Basic.Get("value_IPRHistBasicDBCurr")
            strDBIPRHistRollUpDBCurrTest = Basic.Get("value_IPRHistRollUpDBCurr")
            strDBIPRHistEPRDBCurrTest = Basic.Get("value_IPRHistEPRDBCurr")
            strDBIPRHistPPDBCurrTest = Basic.Get("value_IPRHistPPDBCurr")

            strDBIPRHistContractDBMaxTest = Basic.Get("value_IPRHistContractDBMax")
            strDBIPRHistASDBMaxTest = Basic.Get("value_IPRHistASDBMax")
            strDBIPRHistBasicDBMaxTest = Basic.Get("value_IPRHistBasicDBMax")
            strDBIPRHistRollUpDBMaxTest = Basic.Get("value_IPRHistRollUpDBMax")
            strDBIPRHistEPRDBMaxTest = Basic.Get("value_IPRHistEPRDBMax")
            strDBIPRHistPPDBMaxTest = Basic.Get("value_IPRHistPPDBMax")


            'IPR Hypo DB Values
            strDBIPRHypoContractDBCurrTest = Basic.Get("value_IPRHypoContractDBCurr")
            strDBIPRHypoASDBCurrTest = Basic.Get("value_IPRHypoASDBCurr")
            strDBIPRHypoBasicDBCurrTest = Basic.Get("value_IPRHypoBasicDBCurr")
            strDBIPRHypoRollUpDBCurrTest = Basic.Get("value_IPRHypoRollUpDBCurr")
            strDBIPRHypoEPRDBCurrTest = Basic.Get("value_IPRHypoEPRDBCurr")
            strDBIPRHypoPPDBCurrTest = Basic.Get("value_IPRHypoPPDBCurr")

            strDBIPRHypoContractDBMaxTest = Basic.Get("value_IPRHypoContractDBMax")
            strDBIPRHypoASDBMaxTest = Basic.Get("value_IPRHypoASDBMax")
            strDBIPRHypoBasicDBMaxTest = Basic.Get("value_IPRHypoBasicDBMax")
            strDBIPRHypoRollUpDBMaxTest = Basic.Get("value_IPRHypoRollUpDBMax")
            strDBIPRHypoEPRDBMaxTest = Basic.Get("value_IPRHypoEPRDBMax")
            strDBIPRHypoPPDBMaxTest = Basic.Get("value_IPRHypoPPDBMax")

            'IPR WD Taken
            strIPRWDTakenHypoCurrTest = Basic.Get("value_IPRWDTakenHypoCurr")
            strIPRWDTakenHypoMaxTest = Basic.Get("value_IPRWDTakenHypoMax")
            strIPRWDTakenHistCurrTest = Basic.Get("value_IPRWDTakenHistCurr")
            strIPRWDTakenHistMaxTest = Basic.Get("value_IPRWDTakenHistMax")


            'IPR WD Limit

            strIPRWDLimitCurrTest = Basic.Get("value_IPRWDLimitCurr")
            'strIPRWDLimitGtdTest = Basic.Get("value_IPRWDLimitGtd")
            strIPRWDLimitHistCurrTest = Basic.Get("value_IPRWDLimitHist")
            'strIPRWDLimitCurrGCTest = Basic.Get("value_IPRWDLimitCurrGC")
            strIPRWDLimitMaxTest = Basic.Get("value_IPRWDLimitGtdGC")
            strIPRWDLimitHistMaxTest = Basic.Get("value_IPRWDLimitHistGC")

            'RR One Return for Period

            strRROneHistReturnForPeriodMaxTest = Basic.Get("value_RROneHistReturnForPeriodMax")
            strRROneHistReturnForPeriodCurrTest = Basic.Get("value_RROneHistReturnForPeriodCurr")

            'IPR Hypo Values

            strIPRHypoPPBAMaxTest = Basic.Get("value_IPRHypoPPBAMax")
            strIPRHypoMaxAnnValueMaxTest = Basic.Get("value_IPRHypoMaxAnnValueMax")
            strIPRHypoRollupValueMaxTest = Basic.Get("value_IPRHypoRollUpValueMax")
            strIPRHypoBenefitBaseMaxTest = Basic.Get("value_IPRHypoBenefitBaseMax")
            strIPRHypoDBMaxTest = Basic.Get("value_IPRHypoDBMax")

            strIPRHypoPPBACurrTest = Basic.Get("value_IPRHypoPPBACurr")
            strIPRHypoMaxAnnValueCurrTest = Basic.Get("value_IPRHypoMaxAnnValueCurr")
            strIPRHypoRollupValueCurrTest = Basic.Get("value_IPRHypoRollUpValueCurr")
            strIPRHypoBenefitBaseCurrTest = Basic.Get("value_IPRHypoBenefitBaseCurr")
            strIPRHypoDBCurrTest = Basic.Get("value_IPRHypoDBCurr")

            strIPRHypoContractValueCurrTest = Basic.Get("value_IPRHypoContractValueCurr")
            strIPRHypoSurrenderValueCurrTest = Basic.Get("value_IPRHypoSurrenderValueCurr")
            strIPRHypoContractValueMaxTest = Basic.Get("value_IPRHypoContractValueMax")
            strIPRHypoSurrenderValueMaxTest = Basic.Get("value_IPRHypoSurrenderValueMax")


            'IPR Hist Values

            strIPRHistPPBAMaxTest = Basic.Get("value_IPRHistPPBAMax")
            strIPRHistMaxAnnValueMaxTest = Basic.Get("value_IPRHistMaxAnnValueMax")
            strIPRHistRollupValueMaxTest = Basic.Get("value_IPRHistRollUpValueMax")
            strIPRHistBenefitBaseMaxTest = Basic.Get("value_IPRHistBenefitBaseMax")
            strIPRHistDBMaxTest = Basic.Get("value_IPRHistDBMax")

            strIPRHistPPBACurrTest = Basic.Get("value_IPRHistPPBACurr")
            strIPRHistMaxAnnValueCurrTest = Basic.Get("value_IPRHistMaxAnnValueCurr")
            strIPRHistRollupValueCurrTest = Basic.Get("value_IPRHistRollUpValueCurr")
            strIPRHistBenefitBaseCurrTest = Basic.Get("value_IPRHistBenefitBaseCurr")
            strIPRHistDBCurrTest = Basic.Get("value_IPRHistDBCurr")

            'RR One Annual Income, not split between period certain and living only
            strRROneAnnualIncomeZeroTest = Basic.Get("value_RROneAnnualIncomeZero")
            strRROneAnnualIncomeCurrTest = Basic.Get("value_RROneAnnualIncomeCurr")
            strRROneAnnualIncomeHistTest = Basic.Get("value_RROneAnnualIncomeHist")

            'Start Wds on Younger Annuitants Bday
            strStartWDsYoungerBDTest = Basic.Get("value_StartWDsYoungerBD")

            'For MyClearCourse
            strMCCGtdPaymentFloorZeroTest = Basic.Get("value_MCCGtdPaymentFloor.Zero")
            strMCCGtdPaymentFloorHypoTest = Basic.Get("value_MCCGtdPaymentFloor.Hypo")
            strMCCGtdPaymentFloorHistTest = Basic.Get("value_MCCGtdPaymentFloor.Hist")

            strMCCGtdPaymentFloorFactorZeroTest = Basic.Get("value_MCCGtdPaymentFloorFactor.Zero")
            strMCCGtdPaymentFloorFactorHypoTest = Basic.Get("value_MCCGtdPaymentFloorFactor.Hypo")
            strMCCGtdPaymentFloorFactorHistTest = Basic.Get("value_MCCGtdPaymentFloorFactor.Hist")

            strMCCFirstFullYrIncomeZeroTest = Basic.Get("value_MCCFirstFullYearIncome.Zero")
            strMCCFirstFullYrIncomeHypoTest = Basic.Get("value_MCCFirstFullYearIncome.Hypo")
            strMCCFirstFullYrIncomeHistTest = Basic.Get("value_MCCFirstFullYearIncome.Hist")

            'MCC Life with Period Certain of
            strMCCLifeWithPeriodCertainOfTest = Basic.Get("value_MCCLifeWithPeriodCertainOf")

            'MCC Guarantee From Plan
            strMCCGuaranteeFromPlanTest = Basic.Get("value_MCCGuaranteeFromPlan")

            'MCC Desired Retirement Age
            strMCCDesiredRetirementAgeTest = Basic.Get("value_DesiredRetirementAge")

            'MCC Gtd Income Payments at 0% (chart)
            strMCCGtdIncomePayments0Test = Basic.Get("value_MCCGtdIncomePayments0%")

            'MCC First Full Yr Income Hypo Net (chart)
            strMCCFirstFullYrIncomeHypoNetTest = Basic.Get("value_MCCFirstFullYearIncomeNet")

            'MCC Total Gtd Income Payments at 0% (chart)
            strMCCTotalGtdIncomePayments0Test = Basic.Get("value_MCCTotalGtdPayments0%")

            'MCC Total Income Payments Hypo Net (chart)
            strMCCTotalIncomePaymentsHypoNetTest = Basic.Get("value_MCCTotalIncomePaymentsNet")

            'MCC Adjustment Account
            strMCCAdjustmentAccountZeroTest = Basic.Get("value_MCCAdjustmetnAccount.Zero")
            strMCCAdjustmentAccountHypoTest = Basic.Get("value_MCCAdjustmentAccount.Hypo")
            strMCCAdjustmentAccountHistTest = Basic.Get("value_MCCAdjustmentAccount.Hist")

            'MCC Commutation Value
            strMCCcommutationValueZeroTest = Basic.Get("value_MCCCommutationValue.Zero")
            strMCCCommutationValueHypoTest = Basic.Get("value_MCCCommutationValue.Hypo")
            strMCCCommutationValueHistTest = Basic.Get("value_MCCCommutationValue.Hist")

            'MCC Historical INCOME Period Returns
            strMCCHistIncomePeriodReturnTest = Basic.Get("value_MCCHistIncomeReturn")


        End If
    End Sub
    Public Sub ReadBenchValues(ByVal strBenchPath As String)

        'read the values from the bench relay.out files

        Dim RelayOutBench = New IniConfigSource(strBenchPath)
        Dim i As Integer
        Dim icomma As Integer = 0

        'variables for relay.out sections
        Header = RelayOutBench.Configs("HeaderValues")
        Basic = RelayOutBench.Configs("BasicValues")
        Engine = RelayOutBench.Configs("Engine")



        strStatusBench = Engine.Get("Status")

        'if client does not run
        If strStatusBench = "1" Then
            Regression.RegressionMain.gbClientDoesntRun = True
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
            strHypoGrossBench = Header.Get("bridge_HypoGross")
            strHypoGISRateBench = Header.Get("bridge_GIA.HypoGISRate")
            strZeroGISRateBench = Header.Get("bridge_ZeroGrowthRate.GIS")
            strExpensesVAMandEOnlyBench = Header.Get("bridge_Expenses.VAMAndEOnly")
            strZeroGrowthRateBench = Header.Get("bridge_ZeroGrowthRate")
            strExpensesAdminOnlyBench = Header.Get("bridge_Expenses.AdminOnly")
            strExpensesTotalBaseContractBench = Header.Get("bridge_TotalBaseContractCharges.VA")
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
            strLivingBenefitRiderBench = Header.Get("bridge_LivingBenefitRider")
            If strLivingBenefitRiderBench = "IPR" Then
                gbIPRBench = True
            Else
                gbIPRBench = False
            End If

            For ix = 1 To CInt(strFundCountBench)
                strFundCodeBench(ix) = Header.Get("bridge_FundCode" & ix)
                strFundPctBench(ix) = Header.Get("bridge_FundPct" & ix)
                strFundNameBench(ix) = Header.Get("bridge_FundName" & ix)

                strReturnYr1StdBench(ix) = Header.Get("bridge_ReturnYr1Std" & ix)
                If strReturnYr1StdBench(ix) = "-1000" Then strReturnYr1StdBench(ix) = "N/A"
                strReturnYr5StdBench(ix) = Header.Get("bridge_ReturnYr5Std" & ix)
                If strReturnYr5StdBench(ix) = "-1000" Then strReturnYr5StdBench(ix) = "N/A"
                strReturnYr10StdBench(ix) = Header.Get("bridge_ReturnYr10Std" & ix)
                If strReturnYr10StdBench(ix) = "-1000" Then strReturnYr10StdBench(ix) = "N/A"
                strReturnAdoptionStdBench(ix) = Header.Get("bridge_ReturnSinceAdoptStd" & ix)
                If strReturnAdoptionStdBench(ix) = "-1000" Then strReturnAdoptionStdBench(ix) = "N/A"
                strReturnAdoptionDateStdBench(ix) = Header.Get("bridge_AdoptDateStd" & ix)
                If strReturnAdoptionDateStdBench(ix) = "-1000" Then strReturnAdoptionDateStdBench(ix) = "N/A"
                strReturnYr1StdGIABench(ix) = Header.Get("bridge_ReturnYr1StdGIA" & ix)
                If strReturnYr1StdGIABench(ix) = "-1000" Then strReturnYr1StdGIABench(ix) = "N/A"
                strReturnYr5StdGIABench(ix) = Header.Get("bridge_ReturnYr5StdGIA" & ix)
                If strReturnYr5StdGIABench(ix) = "-1000" Then strReturnYr5StdGIABench(ix) = "N/A"
                strReturnYr10StdGIABench(ix) = Header.Get("bridge_ReturnYr10StdGIA" & ix)
                If strReturnYr10StdGIABench(ix) = "-1000" Then strReturnYr10StdGIABench(ix) = "N/A"
                strReturnAdoptionStdGIABench(ix) = Header.Get("bridge_ReturnSinceAdoptStdGIA" & ix)
                If strReturnAdoptionStdGIABench(ix) = "-1000" Then strReturnAdoptionStdGIABench(ix) = "N/A"
                strReturnAdoptionDateStdGIABench(ix) = Header.Get("bridge_AdoptDateStdGIA" & ix)
                If strReturnAdoptionDateStdGIABench(ix) = "-1000" Then strReturnAdoptionDateStdGIABench(ix) = "N/A"
                strReturnYr1NonStdSCBench(ix) = Header.Get("bridge_ReturnYr1NonStdSC" & ix)
                If strReturnYr1NonStdSCBench(ix) = "-1000" Then strReturnYr1NonStdSCBench(ix) = "N/A"
                strReturnYr5NonStdSCBench(ix) = Header.Get("bridge_ReturnYr5NonStdSC" & ix)
                If strReturnYr5NonStdSCBench(ix) = "-1000" Then strReturnYr5NonStdSCBench(ix) = "N/A"
                strReturnYr10NonStdSCBench(ix) = Header.Get("bridge_ReturnYr10NonStdSC" & ix)
                If strReturnYr10NonStdSCBench(ix) = "-1000" Then strReturnYr10NonStdSCBench(ix) = "N/A"
                strReturnAdoptionDateNonStdSCBench(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                If strReturnAdoptionDateNonStdSCBench(ix) = "-1000" Then strReturnAdoptionDateNonStdSCBench(ix) = "N/A"
                strReturnInceptionDateNonStdSCBench(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                If strReturnInceptionDateNonStdSCBench(ix) = "-1000" Then strReturnInceptionDateNonStdSCBench(ix) = "N/A"
                strReturnYr1NonStdSCGIABench(ix) = Header.Get("bridge_ReturnYr1NonStdSCGIA" & ix)
                strReturnYr5NonStdSCGIABench(ix) = Header.Get("bridge_ReturnYr5NonStdSCGIA" & ix)
                strReturnYr10NonStdSCGIABench(ix) = Header.Get("bridge_ReturnYr10NonStdSCGIA" & ix)
                strReturnAdoptionDateNonStdSCGIABench(ix) = Header.Get("bridge_ReturnIncepDateNonStdSCGIA" & ix)
                strReturnYr1NonStdBench(ix) = Header.Get("bridge_ReturnYr1NonStd" & ix)
                If strReturnYr1NonStdBench(ix) = "-1000" Then strReturnYr1NonStdBench(ix) = "N/A"
                strReturnYr5NonStdBench(ix) = Header.Get("bridge_ReturnYr5NonStd" & ix)
                If strReturnYr5NonStdBench(ix) = "-1000" Then strReturnYr5NonStdBench(ix) = "N/A"
                strReturnYr10NonStdBench(ix) = Header.Get("bridge_ReturnYr10NonStd" & ix)
                If strReturnYr10NonStdBench(ix) = "-1000" Then strReturnYr10NonStdBench(ix) = "N/A"
                strReturnAdoptionDateNonStdBench(ix) = Header.Get("bridge_ReturnSinceAdoptNonStdSC" & ix)
                If strReturnAdoptionDateNonStdBench(ix) = "-1000" Then strReturnAdoptionDateNonStdBench(ix) = "N/A"
                strReturnAdoptionNonStdBench(ix) = Header.Get("bridge_ReturnIncepDateNonStdSC" & ix)
                If strReturnAdoptionNonStdBench(ix) = "-1000" Then strReturnAdoptionNonStdBench(ix) = "N/A"
                strReturnYr1NonStdGIABench(ix) = Header.Get("bridge_ReturnYr1NonStdGIA" & ix)
                If strReturnYr1NonStdGIABench(ix) = "-1000" Then strReturnYr1NonStdGIABench(ix) = "N/A"
                strReturnYr5NonStdGIABench(ix) = Header.Get("bridge_ReturnYr5NonStdGIA" & ix)
                If strReturnYr5NonStdGIABench(ix) = "-1000" Then strReturnYr5NonStdGIABench(ix) = "N/A"
                strReturnYr10NonStdGIABench(ix) = Header.Get("bridge_ReturnYr10NonStdGIA" & ix)
                If strReturnYr10NonStdGIABench(ix) = "-1000" Then strReturnYr10NonStdGIABench(ix) = "N/A"
                strReturnAdoptionDateNonStdGIABench(ix) = Header.Get("bridge_ReturnSinceAdoptNonStd" & ix)
                If strReturnAdoptionDateNonStdGIABench(ix) = "-1000" Then strReturnAdoptionDateNonStdGIABench(ix) = "N/A"
                strReturnInceptionDateNonStdGIABench(ix) = Header.Get("bridge_ReturnIncepDateNonStdGIA" & ix)
                If strReturnInceptionDateNonStdGIABench(ix) = "-1000" Then strReturnInceptionDateNonStdGIABench(ix) = "N/A"
                strHistPeriodEndingBench(ix) = Header.Get("bridge_HistPeriodEnding" & ix)
                If strHistPeriodEndingBench(ix) = "-1000" Then strHistPeriodEndingBench(ix) = "N/A"
                strHistCumulativeReturnBench(ix) = Header.Get("bridge_HistCumulativeReturn" & ix)
                If strHistCumulativeReturnBench(ix) = "-1000" Then strHistCumulativeReturnBench(ix) = "N/A"
                strHistAverageAnnReturnBench(ix) = Header.Get("bridge_HistAverageAnnReturn" & ix)
                If strHistAverageAnnReturnBench(ix) = "-1000" Then strHistAverageAnnReturnBench(ix) = "N/A"
                'for IPR
                strHistCumulativeReturnMaxBench(ix) = Header.Get("bridge_HistCumulativeReturnMax" & ix)
                If strHistCumulativeReturnMaxBench(ix) = "-1000" Then strHistCumulativeReturnMaxBench(ix) = "N/A"
                strHistAverageAnnReturnMaxBench(ix) = Header.Get("bridge_HistAverageAnnReturnMax" & ix)
                If strHistAverageAnnReturnMaxBench(ix) = "-1000" Then strHistAverageAnnReturnMaxBench(ix) = "N/A"


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
            strCombinedSurrValueHypoBench = Basic.Get("value_Surrender.Hypo")
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
            strIPRHistTotalContractValueCurrBench = Basic.Get("value_IPRHistTotalContractValueCurr")
            strIPRHistTotalContractValueMaxBench = Basic.Get("value_IPRHistTotalContractValueMax")
            strIPRHistTotalSurrValueCurrBench = Basic.Get("value_IPRHistTotalSurrValueCurr")
            strIPRHistTotalSurrValueMaxBench = Basic.Get("value_IPRHistTotalSurrValueMax")
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

            'IPR Hist DB values
            strDBIPRHistContractDBCurrBench = Basic.Get("value_IPRHistContractDBCurr")
            strDBIPRHistASDBCurrBench = Basic.Get("value_IPRHistASDBCurr")
            strDBIPRHistBasicDBCurrBench = Basic.Get("value_IPRHistBasicDBCurr")
            strDBIPRHistRollUpDBCurrBench = Basic.Get("value_IPRHistRollUpDBCurr")
            strDBIPRHistEPRDBCurrBench = Basic.Get("value_IPRHistEPRDBCurr")
            strDBIPRHistPPDBCurrBench = Basic.Get("value_IPRHistPPDBCurr")

            strDBIPRHistContractDBMaxBench = Basic.Get("value_IPRHistContractDBMax")
            strDBIPRHistASDBMaxBench = Basic.Get("value_IPRHistASDBMax")
            strDBIPRHistBasicDBMaxBench = Basic.Get("value_IPRHistBasicDBMax")
            strDBIPRHistRollUpDBMaxBench = Basic.Get("value_IPRHistRollUpDBMax")
            strDBIPRHistEPRDBMaxBench = Basic.Get("value_IPRHistEPRDBMax")
            strDBIPRHistPPDBMaxBench = Basic.Get("value_IPRHistPPDBMax")

            'IPR Hypo DB values
            strDBIPRHypoContractDBCurrBench = Basic.Get("value_IPRHypoContractDBCurr")
            strDBIPRHypoASDBCurrBench = Basic.Get("value_IPRHypoASDBCurr")
            strDBIPRHypoBasicDBCurrBench = Basic.Get("value_IPRHypoBasicDBCurr")
            strDBIPRHypoRollUpDBCurrBench = Basic.Get("value_IPRHypoRollUpDBCurr")
            strDBIPRHypoEPRDBCurrBench = Basic.Get("value_IPRHypoEPRDBCurr")
            strDBIPRHypoPPDBCurrBench = Basic.Get("value_IPRHypoPPDBCurr")

            strDBIPRHypoContractDBMaxBench = Basic.Get("value_IPRHypoContractDBMax")
            strDBIPRHypoASDBMaxBench = Basic.Get("value_IPRHypoASDBMax")
            strDBIPRHypoBasicDBMaxBench = Basic.Get("value_IPRHypoBasicDBMax")
            strDBIPRHypoRollUpDBMaxBench = Basic.Get("value_IPRHypoRollUpDBMax")
            strDBIPRHypoEPRDBMaxBench = Basic.Get("value_IPRHypoEPRDBMax")
            strDBIPRHypoPPDBMaxBench = Basic.Get("value_IPRHypoPPDBMax")


            'IPR WD Taken
            strIPRWDTakenHypoCurrBench = Basic.Get("value_IPRWDTakenHypoCurr")
            strIPRWDTakenHypoMaxBench = Basic.Get("value_IPRWDTakenHypoMax")
            strIPRWDTakenHistCurrBench = Basic.Get("value_IPRWDTakenHistCurr")
            strIPRWDTakenHistMaxBench = Basic.Get("value_IPRWDTakenHistMax")

            'IPR WD Limit

            strIPRWDLimitCurrBench = Basic.Get("value_IPRWDLimitCurr")
            'strIPRWDLimitGtdTest = Basic.Get("value_IPRWDLimitGtd")
            strIPRWDLimitHistCurrBench = Basic.Get("value_IPRWDLimitHist")
            'strIPRWDLimitCurrGCTest = Basic.Get("value_IPRWDLimitCurrGC")
            strIPRWDLimitMaxBench = Basic.Get("value_IPRWDLimitGtdGC")
            strIPRWDLimitHistMaxBench = Basic.Get("value_IPRWDLimitHistGC")

            'RR One Return for Period

            strRROneHistReturnForPeriodMaxBench = Basic.Get("value_RROneHistReturnForPeriodMax")
            strRROneHistReturnForPeriodCurrBench = Basic.Get("value_RROneHistReturnForPeriodCurr")

            'IPR Hypo Values

            strIPRHypoPPBAMaxBench = Basic.Get("value_IPRHypoPPBAMax")
            strIPRHypoMaxAnnValueMaxBench = Basic.Get("value_IPRHypoMaxAnnValueMax")
            strIPRHypoRollupValueMaxBench = Basic.Get("value_IPRHypoRollUpValueMax")
            strIPRHypoBenefitBaseMaxBench = Basic.Get("value_IPRHypoBenefitBaseMax")
            strIPRHypoDBMaxBench = Basic.Get("value_IPRHypoDBMax")

            strIPRHypoPPBACurrBench = Basic.Get("value_IPRHypoPPBACurr")
            strIPRHypoMaxAnnValueCurrBench = Basic.Get("value_IPRHypoMaxAnnValueCurr")
            strIPRHypoRollupValueCurrBench = Basic.Get("value_IPRHypoRollUpValueCurr")
            strIPRHypoBenefitBaseCurrBench = Basic.Get("value_IPRHypoBenefitBaseCurr")
            strIPRHypoDBCurrBench = Basic.Get("value_IPRHypoDBCurr")

            strIPRHypoContractValueCurrBench = Basic.Get("value_IPRHypoContractValueCurr")
            strIPRHypoSurrenderValueCurrBench = Basic.Get("value_IPRHypoSurrenderValueCurr")
            strIPRHypoContractValueMaxBench = Basic.Get("value_IPRHypoContractValueMax")
            strIPRHypoSurrenderValueMaxBench = Basic.Get("value_IPRHypoSurrenderValueMax")

            'IPR Hist Values

            strIPRHistPPBAMaxBench = Basic.Get("value_IPRHistPPBAMax")
            strIPRHistMaxAnnValueMaxBench = Basic.Get("value_IPRHistMaxAnnValueMax")
            strIPRHistRollupValueMaxBench = Basic.Get("value_IPRHistRollUpValueMax")
            strIPRHistBenefitBaseMaxBench = Basic.Get("value_IPRHistBenefitBaseMax")
            strIPRHistDBMaxBench = Basic.Get("value_IPRHistDBMax")

            strIPRHistPPBACurrBench = Basic.Get("value_IPRHistPPBACurr")
            strIPRHistMaxAnnValueCurrBench = Basic.Get("value_IPRHistMaxAnnValueCurr")
            strIPRHistRollupValueCurrBench = Basic.Get("value_IPRHistRollUpValueCurr")
            strIPRHistBenefitBaseCurrBench = Basic.Get("value_IPRHistBenefitBaseCurr")
            strIPRHistDBCurrBench = Basic.Get("value_IPRHistDBCurr")

            'RR One Annual Income, not split between period certain and living only
            strRROneAnnualIncomeZeroBench = Basic.Get("value_RROneAnnualIncomeZero")
            strRROneAnnualIncomeCurrBench = Basic.Get("value_RROneAnnualIncomeCurr")
            strRROneAnnualIncomeHistBench = Basic.Get("value_RROneAnnualIncomeHist")

            'Start Wds on Younger Annuitants Bday
            strStartWDsYoungerBDBench = Basic.Get("value_StartWDsYoungerBD")


            'For MyClearCourse
            'MCC Gtd Payment Floor
            strMCCGtdPaymentFloorZeroBench = Basic.Get("value_MCCGtdPaymentFloor.Zero")
            strMCCGtdPaymentFloorHypoBench = Basic.Get("value_MCCGtdPaymentFloor.Hypo")
            strMCCGtdPaymentFloorHistBench = Basic.Get("value_MCCGtdPaymentFloor.Hist")

            'MCC Payment Floor Factor
            strMCCGtdPaymentFloorFactorZeroBench = Basic.Get("value_MCCGtdPaymentFloorFactor.Zero")
            strMCCGtdPaymentFloorFactorHypoBench = Basic.Get("value_MCCGtdPaymentFloorFactor.Hypo")
            strMCCGtdPaymentFloorFactorHistBench = Basic.Get("value_MCCGtdPaymentFloorFactor.Hist")

            'MCC First Full Yr Income
            strMCCFirstFullYrIncomeZeroBench = Basic.Get("value_MCCFirstFullYearIncome.Zero")
            strMCCFirstFullYrIncomeHypoBench = Basic.Get("value_MCCFirstFullYearIncome.Hypo")
            strMCCFirstFullYrIncomeHistBench = Basic.Get("value_MCCFirstFullYearIncome.Hist")

            'MCC Life with Period Certain of
            strMCCLifeWithPeriodCertainOfBench = Basic.Get("value_MCCLifeWithPeriodCertainOf")

            'MCC Guarantee From Plan
            strMCCGuaranteeFromPlanBench = Basic.Get("value_MCCGuaranteeFromPlan")

            'MCC Desired Retirement Age
            strMCCDesiredRetirementAgeBench = Basic.Get("value_DesiredRetirementAge")

            'MCC Gtd Income Payments at 0% (chart)
            strMCCGtdIncomePayments0Bench = Basic.Get("value_MCCGtdIncomePayments0%")

            'MCC First Full Yr Income Hypo Net (chart)
            strMCCFirstFullYrIncomeHypoNetBench = Basic.Get("value_MCCFirstFullYearIncomeNet")

            'MCC Total Gtd Income Payments at 0% (chart)
            strMCCTotalGtdIncomePayments0Bench = Basic.Get("value_MCCTotalGtdPayments0%")

            'MCC Total Income Payments Hypo Net (chart)
            strMCCTotalIncomePaymentsHypoNetBench = Basic.Get("value_MCCTotalIncomePaymentsNet")

            'MCC Adjustment Account
            strMCCAdjustmentAccountZeroBench = Basic.Get("value_MCCAdjustmetnAccount.Zero")
            strMCCAdjustmentAccountHypoBench = Basic.Get("value_MCCAdjustmentAccount.Hypo")
            strMCCAdjustmentAccountHistBench = Basic.Get("value_MCCAdjustmentAccount.Hist")


            'MCC Commutation Value
            strMCCcommutationValueZeroBench = Basic.Get("value_MCCCommutationValue.Zero")
            strMCCCommutationValueHypoBench = Basic.Get("value_MCCCommutationValue.Hypo")
            strMCCCommutationValueHistBench = Basic.Get("value_MCCCommutationValue.Hist")

            'MCC Historical INCOME Period Returns
            strMCCHistIncomePeriodReturnBench = Basic.Get("value_MCCHistIncomeReturn")

        End If
    End Sub

    Public Sub GetTimeStamps()

        'GLAIC VA:

        strWFProp = "WFPROP.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\WFPROP.EXE"))
        strGLAICCPY = "GELA.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\GELA.CPY"))
        strAnn1 = "ANNUITY.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GELADATA\ANNUITY.MDB"))
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
        strGLICNYSPIAANNVER = "ANNVER.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\ANNVER.MDB"))

        'GLICNY SPDA:

        strGLICNYSPDAEXE = "GEDefAnn.exe  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECLDATA\GEDefAnn.exe"))
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
        strGLICSPIANNVER = "ANNVER.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\ANNVER.MDB"))

        'GLIC SPDA:

        strGLICSPDAEXE = "GEDefAnn.exe  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GEDefAnn.exe"))
        strGLICSPDARATE = "RATE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\RATE.MDB"))
        strGLICSPDAGNAWINE = "GNAWINE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GNAWINE.MDB"))
        strGLICSPDACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GECA.CPY"))

        'GLACI SPDA(MVA):

        strFCOLSPDAEXE = "GEDefAnn.exe  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GEDefAnn.exe"))
        strFCOLSPDARATE = "RATE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\RATE.MDB"))
        strFCOLSPDAGNAWINE = "GNAWINE.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GNAWINE.MDB"))
        strFCOLSPDACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\GECADATA\GECA.CPY"))

        'GLAICFIXED SPIA:

        strGLAICFIXEDSPIAEXE = "WINANN.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\WINANN.EXE"))
        strGLAICFIXEDSPIAANNRATES = "ANNRATES.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNRATES.MDB"))
        strGLAICFIXEDSPIAANNPROD = "ANNPROD.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNPROD.MDB"))
        strGLAICFIXEDSPIAANNSYS = "ANNSYS.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNSYS.MDB"))
        strGLAICFIXEDSPIACPY = "GECL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\FCOL.CPY"))
        strGLAICFIXEDSPIAANNSUPP = "ANNSUPP.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNSUPP.MDB"))
        strGLAICFIXEDSPIANNVER = "ANNVER.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\ANNVER.MDB"))

        'FIA

        strFIAEXE = "FIA.EXE  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\FIA.EXE"))
        strFIACPY = "FCOL.CPY  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\FCOL.CPY"))
        strFIAMDB = "FIA.MDB  " & CStr(FileSystem.FileDateTime("C:\WinFlex6\FCOLDATA\FIA.MDB"))


    End Sub
    Public Sub StoreMismatches()
    End Sub

End Class
