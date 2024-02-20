Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
'Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports System.Windows.Forms
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Nini.Config
Imports Scripting
Imports System.Diagnostics
'Imports System.ComponentModel
Imports System
Imports System.Reflection
Imports System.Threading
Imports System.Xml
'Imports Namespaces
Imports SR = System.Reflection


Public Class RegressionMain
    Public Shared gstrCase(100) As String
    Public Shared gstrpathProduct As String
    Public Shared gstrCompCode(100) As String
    Public Shared username As String
    Public Shared gbSPIAEffectiveDate As Boolean
    Public Shared gbSPDAEffectiveDate As Boolean
    Public Shared gbVAHistoricalDate As Boolean
    Public Shared gstrSPIARateDate As String
    Public Shared gstrSPDARateDate As String
    Public Shared gstrVAHistDate As String
    Public Shared gbVASaveAge(100) As Boolean 'if age is saved on particular client
    Public Shared gbVASaveAgeChecked As Boolean 'if checked yes to save age
    Public Shared gbFIASaveAge(100) As Boolean 'if age is saved on particular client
    Public Shared gbFIASaveAgeChecked As Boolean 'if checked yes to save age
    Public Shared gstrDOB2Original As String
    Public Shared gstrDOB1Original As String
    Public Shared gstrDOB2New As String
    Public Shared gstrDOB1New As String
    Public Shared gstrVASaveAgeDOB1 As String
    Public Shared gstrVASaveAgeDOB2 As String
    Public Shared gstrFIASaveAgeDOB1 As String
    Public Shared gstrFIASaveAgeDOB2 As String
    Public Shared gbRateCancel As Boolean
    Public Shared gbVASaveAgeCancel As Boolean
    Public Shared gbFIASaveAgeCancel As Boolean
    Public Shared gbVAPrevHistoricalCancel As Boolean
    Public Shared gbClientDoesntRun As Boolean
    Private strsplitMMList(100) As String
    Private strsplitmmlistFIATool(100) As String
    Private strClientXMisMatch(100) As String
    Private strClientXMisMatchFIATool(100) As String
    Private strClientMisMatchList As String
    Private strClientMisMatchListFIATool As String
    Dim ib As Integer = 0
    Public Shared gstrpath As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression"
    Dim ReadTest As New clsReadVAValues()
    Dim ReadBench As New clsReadVAValues()
    Dim TimeStamps As New clsReadVAValues()
    Dim ReadSPIATest As New clsReadSPIAValues()
    Dim ReadSPIABench As New clsReadSPIAValues()
    Dim ReadSPDATest As New clsReadSPDAValues()
    Dim ReadSPDABench As New clsReadSPDAValues()
    Dim ReadFIATest As New clsReadFIAValues()
    Dim ReadFIABench As New clsReadFIAValues()
    Dim bPalomaOnly As Boolean
    Dim FIATool As New ReadFIARelayINI()

    Private WithEvents TestWorker As System.ComponentModel.BackgroundWorker

    Private Sub RegressionMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Form3.Enabled = True Then Form3.Close()
        If Form4.Enabled = True Then Form4.Close()
    End Sub

    Public Sub RegressionMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim win As System.Security.Principal.WindowsIdentity
        win = System.Security.Principal.WindowsIdentity.GetCurrent
        username = win.Name.Substring(win.Name.IndexOf("\") + 1)
    End Sub
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim CS_NOCLOSE As Integer = Int32.Parse("200", Globalization.NumberStyles.HexNumber)
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = CS_NOCLOSE
            Return cp
        End Get
    End Property
    Private Sub CompareValuesVA(ByVal ic As Integer, ByVal ib As Integer, ByVal strcomp() As String)

        Dim bFundCountDifferent As Boolean
        Dim ix As Integer = 0
        Dim imax As Integer = 0
        Dim iAdd As Integer = 0
        Dim iShort As Integer = 0

        'set mismatch flag to false
        clsReadVAValues.bMisMatch = False

        'go through each value in relay.out to compare between bench and test, if mismatch, create strings for datagrid

        'below 4 are for warning and error messages
        If clsReadVAValues.bErrorBench = True And clsReadVAValues.bErrorTest = True Then
            If (String.Compare(clsReadVAValues.strMessage1Test, clsReadVAValues.strMessage1Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage1MM = ib & ",Message 1," & clsReadVAValues.strMessage1Bench & "," & clsReadVAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage2Test, clsReadVAValues.strMessage2Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage2MM = ib & ",Message 2," & clsReadVAValues.strMessage2Bench & "," & clsReadVAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage3Test, clsReadVAValues.strMessage3Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage3MM = ib & ",Message 3," & clsReadVAValues.strMessage3Bench & "," & clsReadVAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage4Test, clsReadVAValues.strMessage4Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage4MM = ib & ",Message 4," & clsReadVAValues.strMessage4Bench & "," & clsReadVAValues.strMessage4Test & "&"
            End If

            'Set Error Messages back to nothing

            clsReadVAValues.strMessage1Bench = ""
            clsReadVAValues.strMessage1Test = ""
            clsReadVAValues.strMessage2Bench = ""
            clsReadVAValues.strMessage2Test = ""
            clsReadVAValues.strMessage3Bench = ""
            clsReadVAValues.strMessage3Test = ""
            clsReadVAValues.strMessage4Bench = ""
            clsReadVAValues.strMessage4Test = ""
            clsReadVAValues.strMessage5Bench = ""
            clsReadVAValues.strMessage5Test = ""
            clsReadVAValues.strMessage6Bench = ""
            clsReadVAValues.strMessage6Test = ""


            'see if one runs and one doesnt
        ElseIf clsReadVAValues.bErrorBench = True And clsReadVAValues.bErrorTest = False Then
            clsReadVAValues.bMisMatch = True
            clsReadVAValues.strRunNoRunMM = ib & ",Test runs/Bench doesn't run,,&"

            If gbVASaveAge(ib) Then
                If gstrDOB1New <> "" Then
                    clsReadVAValues.bMisMatch = True
                    gstrVASaveAgeDOB1 = ib & ",DOB1," & gstrDOB1Original & "," & gstrDOB1New & "&"
                End If
            End If

            If gbVASaveAge(ib) Then
                If gstrDOB2New <> "" Then
                    clsReadVAValues.bMisMatch = True
                    gstrVASaveAgeDOB2 = ib & ",DOB2," & gstrDOB2Original & "," & gstrDOB2New & "&"
                End If
            End If

           
        ElseIf clsReadVAValues.bErrorBench = False And clsReadVAValues.bErrorTest = True Then
            clsReadVAValues.bMisMatch = True
            clsReadVAValues.strRunNoRunMM = ib & ",Bench runs/Test doesn't run,,&"

            If gbVASaveAge(ib) Then
                If gstrDOB1New <> "" Then
                    clsReadVAValues.bMisMatch = True
                    gstrVASaveAgeDOB1 = ib & ",DOB1," & gstrDOB1Original & "," & gstrDOB1New & "&"
                End If
            End If

            If gbVASaveAge(ib) Then
                If gstrDOB2New <> "" Then
                    clsReadVAValues.bMisMatch = True
                    gstrVASaveAgeDOB2 = ib & ",DOB2," & gstrDOB2Original & "," & gstrDOB2New & "&"
                End If
            End If

        Else
            'if both cases run, then compare...

            If (String.Compare(clsReadVAValues.strMessage1Test, clsReadVAValues.strMessage1Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage1MM = ib & ",Message1," & clsReadVAValues.strMessage1Bench & "," & clsReadVAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage2Test, clsReadVAValues.strMessage2Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage2MM = ib & ",Message2," & clsReadVAValues.strMessage2Bench & "," & clsReadVAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage3Test, clsReadVAValues.strMessage3Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage3MM = ib & ",Message3," & clsReadVAValues.strMessage3Bench & "," & clsReadVAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage4Test, clsReadVAValues.strMessage4Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage4MM = ib & ",Message4," & clsReadVAValues.strMessage4Bench & "," & clsReadVAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage5Test, clsReadVAValues.strMessage5Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage5MM = ib & ",Message5," & clsReadVAValues.strMessage5Bench & "," & clsReadVAValues.strMessage5Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strMessage6Test, clsReadVAValues.strMessage6Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMessage6MM = ib & ",Message6," & clsReadVAValues.strMessage6Bench & "," & clsReadVAValues.strMessage6Test & "&"
            End If

            'Set Error Messages back to nothing

            clsReadVAValues.strMessage1Bench = ""
            clsReadVAValues.strMessage1Test = ""
            clsReadVAValues.strMessage2Bench = ""
            clsReadVAValues.strMessage2Test = ""
            clsReadVAValues.strMessage3Bench = ""
            clsReadVAValues.strMessage3Test = ""
            clsReadVAValues.strMessage4Bench = ""
            clsReadVAValues.strMessage4Test = ""
            clsReadVAValues.strMessage5Bench = ""
            clsReadVAValues.strMessage5Test = ""
            clsReadVAValues.strMessage6Bench = ""
            clsReadVAValues.strMessage6Test = ""


            If gbVASaveAge(ib) Then
                If gstrDOB1New <> "" Then
                    clsReadVAValues.bMisMatch = True
                    gstrVASaveAgeDOB1 = ib & ",DOB1," & gstrDOB1Original & "," & gstrDOB1New & "&"
                End If
            End If

            If gbVASaveAge(ib) Then
                If gstrDOB2New <> "" Then
                    clsReadVAValues.bMisMatch = True
                    gstrVASaveAgeDOB2 = ib & ",DOB2," & gstrDOB2Original & "," & gstrDOB2New & "&"
                End If
            End If

            If (String.Compare(clsReadVAValues.strCompanyNameTest, clsReadVAValues.strCompanyNameBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strCompanyNameMM = ib & ",Company Name," & clsReadVAValues.strCompanyNameBench & "," & clsReadVAValues.strCompanyNameTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strClient1Test, clsReadVAValues.strClient1Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strClient1NameMM = ib & ",Client1.Name, " & clsReadVAValues.strClient1Bench & "," & clsReadVAValues.strClient1Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strAge1Test, clsReadVAValues.strAge1Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strAge1MM = ib & ",Client1.Age," & clsReadVAValues.strAge1Bench & "," & clsReadVAValues.strAge1Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strAgeOlderTest, clsReadVAValues.strAgeOlderBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strAgeOlderMM = ib & ",Age.Older," & clsReadVAValues.strAgeOlderBench & "," & clsReadVAValues.strAgeOlderTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strIRateTest, clsReadVAValues.strIRateBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strIRateMM = ib & ",Interest.Rate," & clsReadVAValues.strIRateBench & "," & clsReadVAValues.strIRateTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strSex1Test, clsReadVAValues.strSex1Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strSex1MM = ib & ",Client1.Sex," & clsReadVAValues.strSex1Bench & "," & clsReadVAValues.strSex1Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strInitialDBTest, clsReadVAValues.strInitialDBBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strInitialDBMM = ib & ",Initial DB," & clsReadVAValues.strInitialDBBench & "," & clsReadVAValues.strInitialDBTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strJointTest, clsReadVAValues.strJointBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strJointMM = ib & ",Joint?," & clsReadVAValues.strJointBench & "," & clsReadVAValues.strJointTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strInitialDB2Test, clsReadVAValues.strInitialDB2Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strInitialDB2MM = ib & ",Initial DB2," & clsReadVAValues.strInitialDB2Bench & "," & clsReadVAValues.strInitialDB2Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strClient2Test, clsReadVAValues.strClient2Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strClient2NameMM = ib & ",Client2.Name," & clsReadVAValues.strClient2Bench & "," & clsReadVAValues.strClient2Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strAge2Test, clsReadVAValues.strAge2Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strAge2MM = ib & ",Client2.Age," & clsReadVAValues.strAge2Bench & "," & clsReadVAValues.strAge2Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strSex2Test, clsReadVAValues.strSex2Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strSex2MM = ib & ",Client2.Sex," & clsReadVAValues.strSex2Bench & "," & clsReadVAValues.strSex2Test & "&"
            End If

            If (String.Compare(clsReadVAValues.strContractTypeTest, clsReadVAValues.strContractTypeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strContractTypeMM = ib & ",Contract.Type," & clsReadVAValues.strContractTypeBench & "," & clsReadVAValues.strContractTypeTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strSurrChargeYrsTest, clsReadVAValues.strSurrChargeYrsBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strSurrChargeYrsMM = ib & ",Surr.Charge.Yrs," & clsReadVAValues.strSurrChargeYrsBench & "," & clsReadVAValues.strSurrChargeYrsTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strHypoNorGTest, clsReadVAValues.strHypoNorGBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strHypoNorGMM = ib & ",Hypo.Rate.NetOrGross," & clsReadVAValues.strHypoNorGBench & "," & clsReadVAValues.strHypoNorGTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strZeroNetTest, clsReadVAValues.strZeroNetBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strZeroNetMM = ib & ",NetRate.Zero," & clsReadVAValues.strZeroNetBench & "," & clsReadVAValues.strZeroNetTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strHypoNetTest, clsReadVAValues.strHypoNetBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strHypoNetMM = ib & ",Hypo.Rate.Net," & clsReadVAValues.strHypoNetBench & "," & clsReadVAValues.strHypoNetTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strHypoGrossTest, clsReadVAValues.strHypoGrossBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strHypoGrossMM = ib & ",Hypo.Rate.Gross," & clsReadVAValues.strHypoGrossBench & "," & clsReadVAValues.strHypoGrossTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strHypoGISRateTest, clsReadVAValues.strHypoGISRateBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strHypoGISRateMM = ib & ",Hypo.Rate.GIS," & clsReadVAValues.strHypoGISRateBench & "," & clsReadVAValues.strHypoGISRateTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strZeroGISRateTest, clsReadVAValues.strZeroGISRateBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strZeroGISRateMM = ib & ",Zero.Rate.GIS," & clsReadVAValues.strZeroGISRateBench & "," & clsReadVAValues.strZeroGISRateTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strExpensesVAMandEOnlyTest, clsReadVAValues.strExpensesVAMandEOnlyBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strExpenseVAMandEOnlyMM = ib & ",Expenses.VA.M&EOnly," & clsReadVAValues.strExpensesVAMandEOnlyBench & "," & clsReadVAValues.strExpensesVAMandEOnlyTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strExpensesTotalBaseContractTest, clsReadVAValues.strExpensesTotalBaseContractBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strExpenseVAMandEOnlyMM = ib & ",VA.TotBaseContChrgs," & clsReadVAValues.strExpensesTotalBaseContractBench & "," & clsReadVAValues.strExpensesTotalBaseContractTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strZeroGrowthRateTest, clsReadVAValues.strZeroGrowthRateBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strZeroGrowthRateMM = ib & ",Zero.Rate.Growth," & clsReadVAValues.strZeroGrowthRateBench & "," & clsReadVAValues.strZeroGrowthRateTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strExpensesAdminOnlyTest, clsReadVAValues.strExpensesAdminOnlyBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strExpensesAdminOnlyMM = ib & ",Expenses.AdminOnly," & clsReadVAValues.strExpensesAdminOnlyBench & "," & clsReadVAValues.strExpensesAdminOnlyTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strFundExpensesVATest, clsReadVAValues.strFundExpensesVABench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strFundExpensesVAMM = ib & ",Expenses.VA.Fund," & clsReadVAValues.strFundExpensesVABench & "," & clsReadVAValues.strFundExpensesVATest & "&"
            End If

            If (String.Compare(clsReadVAValues.strFundExpenseGISTest, clsReadVAValues.strFundExpenseGISBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strFundExpenseGISMM = ib & ",Expenses.Fund.GIS," & clsReadVAValues.strFundExpenseGISBench & "," & clsReadVAValues.strFundExpenseGISTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strFundExpenseEffDateGISTest, clsReadVAValues.strFundExpenseEffDateGISBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strFundExpenseEffDateGISMM = ib & ",Expenses.Fund.GIS.EffDate," & clsReadVAValues.strFundExpenseEffDateGISBench & "," & clsReadVAValues.strFundExpenseEffDateGISTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strVADBBeneRiderChargeTest, clsReadVAValues.strVADBBeneRiderChargeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strVADBBenefitRiderChargeMM = ib & ",DBRider.Charge.VA," & clsReadVAValues.strVADBBeneRiderChargeBench & "," & clsReadVAValues.strVADBBeneRiderChargeTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strVAContractChargeTest, clsReadVAValues.strVAContractChargeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strVAContractChargeMM = ib & ",VA.Contract.Charge," & clsReadVAValues.strVAContractChargeBench & "," & clsReadVAValues.strVAContractChargeTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strVAContractChargeWaiverLimitTest, clsReadVAValues.strVAContractChargeWaiverLimitBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strVAContractChargeWaiverLimitMM = ib & ",ContrChrg.VA.WaiverLimit," & clsReadVAValues.strVAContractChargeWaiverLimitBench & "," & clsReadVAValues.strVAContractChargeWaiverLimitTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strFundExpenseEffDateTest, clsReadVAValues.strFundExpenseEffDateBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strFundExpenseEffDateMM = ib & ",Expenses.Fund.EffDate," & clsReadVAValues.strFundExpenseEffDateBench & "," & clsReadVAValues.strFundExpenseEffDateTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strEarlyAccessChargeTest, clsReadVAValues.strEarlyAccessChargeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strEarlyAccessChargeMM = ib & ",EarlyAccess.Charge," & clsReadVAValues.strEarlyAccessChargeBench & "," & clsReadVAValues.strEarlyAccessChargeTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strLivingBenefitRiderChargeTest, clsReadVAValues.strLivingBenefitRiderChargeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLivingBenefitRiderChargeMM = ib & ",RiderChrg.LivingBenefit," & clsReadVAValues.strLivingBenefitRiderChargeBench & "," & clsReadVAValues.strLivingBenefitRiderChargeTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strInitialPremiumTest, clsReadVAValues.strInitialPremiumBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strInitialPremiumMM = ib & ",Premium.Initial," & clsReadVAValues.strInitialPremiumBench & "," & clsReadVAValues.strInitialPremiumTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strDBTypeTest, clsReadVAValues.strDBTypeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strDBTypeMM = ib & ",DeathBen.Type," & clsReadVAValues.strDBTypeBench & "," & clsReadVAValues.strDBTypeTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strFundCountTest), CStr(clsReadVAValues.strFundCountBench), True)) <> 0 Then
                bFundCountDifferent = True
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strFundCountMM = ib & ",Fund.Count," & clsReadVAValues.strFundCountBench & "," & clsReadVAValues.strFundCountTest & "&"
            End If

            If (String.Compare(clsReadVAValues.strInvestStratTest, clsReadVAValues.strInvestStratBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strInvestStratMM = ib & ",Invest.Strategy.Type," & clsReadVAValues.strInvestStratBench & "," & clsReadVAValues.strInvestStratTest & "&"
            End If

            If bFundCountDifferent = False Then

                For ix = 1 To clsReadVAValues.strFundCountTest

                    If (String.Compare(clsReadVAValues.strFundCodeTest(ix), clsReadVAValues.strFundCodeBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strFundCodeTest(ix)
                        Dim strBench As String = clsReadVAValues.strFundCodeBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strFundCodeMM(ix) = clsReadVAValues.strFundCodeMM(ix) & ib & ",Fund.Code " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strFundPctTest(ix), clsReadVAValues.strFundPctBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strFundPctTest(ix)
                        Dim strBench As String = clsReadVAValues.strFundPctBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strFundPctMM(ix) = clsReadVAValues.strFundPctMM(ix) & ib & ",Fund.Percent " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strFundNameTest(ix), clsReadVAValues.strFundNameBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strFundNameTest(ix)
                        Dim strBench As String = clsReadVAValues.strFundNameBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strFundNameMM(ix) = clsReadVAValues.strFundNameMM(ix) & ib & ",Fund.Name " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr1StdTest(ix), clsReadVAValues.strReturnYr1StdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr1StdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr1StdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If


                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr1StdMM(ix) = clsReadVAValues.strReturnYr1StdMM(ix) & ib & ",FundRet.1Yr.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr5StdTest(ix), clsReadVAValues.strReturnYr5StdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr5StdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr5StdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If
                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr5StdMM(ix) = clsReadVAValues.strReturnYr5StdMM(ix) & ib & ",FundRet.5Yr.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr10StdTest(ix), clsReadVAValues.strReturnYr10StdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr10StdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr10StdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr10StdMM(ix) = clsReadVAValues.strReturnYr10StdMM(ix) & ib & ",FundRet10Yr.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionStdTest(ix), clsReadVAValues.strReturnAdoptionStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionStdMM(ix) = clsReadVAValues.strReturnAdoptionStdMM(ix) & ib & ",FundRet.SinceAdopt.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionDateStdTest(ix), clsReadVAValues.strReturnAdoptionDateStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionDateStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionDateStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDatestdMM(ix) = clsReadVAValues.strReturnAdoptionDatestdMM(ix) & ib & ",FundRet.AdoptDate.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr1StdGIATest(ix), clsReadVAValues.strReturnYr1StdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr1StdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr1StdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If
                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr1StdGIAMM(ix) = clsReadVAValues.strReturnYr1StdGIAMM(ix) & ib & ",FundRet.1Yr.GIA.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr5StdGIATest(ix), clsReadVAValues.strReturnYr5StdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr5StdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr5StdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr5StdGIAMM(ix) = clsReadVAValues.strReturnYr5StdGIAMM(ix) & ib & ",FundRet.5Yr.GIA.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr10StdGIATest(ix), clsReadVAValues.strReturnYr10StdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr10StdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr10StdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If


                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr10StdGIAMM(ix) = clsReadVAValues.strReturnYr10StdGIAMM(ix) & ib & ",FundRet.10Yr.GIA.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionStdGIATest(ix), clsReadVAValues.strReturnAdoptionStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionStdGIAMM(ix) = clsReadVAValues.strReturnAdoptionStdGIAMM(ix) & ib & ",FundRet.SinceAdopt.GIA.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionDateStdGIATest(ix), clsReadVAValues.strReturnAdoptionDateStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionDateStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionDateStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDatestdGIAMM(ix) = clsReadVAValues.strReturnAdoptionDatestdGIAMM(ix) & ib & ",FundRet.AdoptDate.GIA.Stdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr1NonStdSCTest(ix), clsReadVAValues.strReturnYr1NonStdSCBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr1NonStdSCTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr1NonStdSCBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr1NonStdSCMM(ix) = clsReadVAValues.strReturnYr1NonStdSCMM(ix) & ib & ",FundRet.1Yr.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr5NonStdSCTest(ix), clsReadVAValues.strReturnYr5NonStdSCBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr5NonStdSCTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr5NonStdSCBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr5NonStdSCMM(ix) = clsReadVAValues.strReturnYr5NonStdSCMM(ix) & ib & ",FundRet.5Yr.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr10NonStdSCTest(ix), clsReadVAValues.strReturnYr10NonStdSCBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr10NonStdSCTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr10NonStdSCBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr10NonStdSCMM(ix) = clsReadVAValues.strReturnYr10NonStdSCMM(ix) & ib & ",FundRet.10Yr.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionNonStdSCTest(ix), clsReadVAValues.strReturnAdoptionNonStdSCBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionNonStdSCTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionNonStdSCBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix) = clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix) & ib & ",FundRet.SinceAdopt.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionDateNonStdSCTest(ix), clsReadVAValues.strReturnAdoptionDateNonStdSCBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionDateNonStdSCTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionDateNonStdSCBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix) = clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix) & ib & ",FundRet.AdoptDate.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnInceptionDateNonStdSCTest(ix), clsReadVAValues.strReturnInceptionDateNonStdSCBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnInceptionDateNonStdSCTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnInceptionDateNonStdSCBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnInceptionDateNonStdSCMM(ix) = clsReadVAValues.strReturnInceptionDateNonStdSCMM(ix) & ib & ",FundRet.InceptDate.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr1NonStdSCGIATest(ix), clsReadVAValues.strReturnYr1NonStdSCGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr1NonStdSCGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr1NonStdSCGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr1NonStdSCGIAMM(ix) = clsReadVAValues.strReturnYr1NonStdSCGIAMM(ix) & ib & ",FundRet.1Yr.GIA.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr5NonStdSCGIATest(ix), clsReadVAValues.strReturnYr5NonStdSCGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr5NonStdSCGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr5NonStdSCGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr5NonStdSCGIAMM(ix) = clsReadVAValues.strReturnYr5NonStdSCGIAMM(ix) & ib & ",FundRet.5Yr.GIA.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr10NonStdSCGIATest(ix), clsReadVAValues.strReturnYr10NonStdSCGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr10NonStdSCGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr10NonStdSCGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr10NonStdSCGIAMM(ix) = clsReadVAValues.strReturnYr10NonStdSCGIAMM(ix) & ib & ",FundRet.10Yr.GIA.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionNonStdSCGIATest(ix), clsReadVAValues.strReturnAdoptionNonStdSCGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionNonStdSCGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionNonStdSCGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionNonStdSCGIAMM(ix) = clsReadVAValues.strReturnAdoptionNonStdSCGIAMM(ix) & ib & ",FundRet.SinceAdopt.GIA.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionDateNonStdSCGIATest(ix), clsReadVAValues.strReturnAdoptionDateNonStdSCGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionDateNonStdSCGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionDateNonStdSCGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDateNonStdSCGIAMM(ix) = clsReadVAValues.strReturnAdoptionDateNonStdSCGIAMM(ix) & ib & ",FundRet.AdoptDate.GIA.NonStdzed.Surr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr1NonStdTest(ix), clsReadVAValues.strReturnYr1NonStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr1NonStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr1NonStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr1NonStdMM(ix) = clsReadVAValues.strReturnYr1NonStdMM(ix) & ib & ",FundRet.1Yr.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr5NonStdTest(ix), clsReadVAValues.strReturnYr5NonStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr5NonStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr5NonStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr5NonStdMM(ix) = clsReadVAValues.strReturnYr5NonStdMM(ix) & ib & ",FundRet.5Yr.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr10NonStdTest(ix), clsReadVAValues.strReturnYr10NonStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr10NonStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr10NonStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr10NonStdMM(ix) = clsReadVAValues.strReturnYr10NonStdMM(ix) & ib & ",FundRet.10Yr.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionDateNonStdTest(ix), clsReadVAValues.strReturnAdoptionDateNonStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionDateNonStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionDateNonStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDateNonStdMM(ix) = clsReadVAValues.strReturnAdoptionDateNonStdMM(ix) & ib & ",FundRet.AdoptDate.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionNonStdTest(ix), clsReadVAValues.strReturnAdoptionNonStdBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionNonStdTest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionNonStdBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionNonStdMM(ix) = clsReadVAValues.strReturnAdoptionNonStdMM(ix) & ib & ",FundRet.SinceAdopt.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr1NonStdGIATest(ix), clsReadVAValues.strReturnYr1NonStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr1NonStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr1NonStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr1NonStdGIAMM(ix) = clsReadVAValues.strReturnYr1NonStdGIAMM(ix) & ib & ",Fund.Ret.1Yr.GIA.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr5NonStdGIATest(ix), clsReadVAValues.strReturnYr5NonStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr5NonStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr5NonStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr5NonStdGIAMM(ix) = clsReadVAValues.strReturnYr5NonStdGIAMM(ix) & ib & ",Fund.Ret.5Yr.GIA.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnYr10NonStdGIATest(ix), clsReadVAValues.strReturnYr10NonStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnYr10NonStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnYr10NonStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnYr10NonStdGIAMM(ix) = clsReadVAValues.strReturnYr10NonStdGIAMM(ix) & ib & ",Fund.Ret.10Yr.GIA.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnAdoptionDateNonStdGIATest(ix), clsReadVAValues.strReturnAdoptionDateNonStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnAdoptionDateNonStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnAdoptionDateNonStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnAdoptionDateNonStdGIAMM(ix) = clsReadVAValues.strReturnAdoptionDateNonStdGIAMM(ix) & ib & ",FundRet.AdoptDate.GIA.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strReturnInceptionDateNonStdGIATest(ix), clsReadVAValues.strReturnInceptionDateNonStdGIABench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strReturnInceptionDateNonStdGIATest(ix)
                        Dim strBench As String = clsReadVAValues.strReturnInceptionDateNonStdGIABench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strReturnInceptionDateNonStdGIAMM(ix) = clsReadVAValues.strReturnInceptionDateNonStdGIAMM(ix) & ib & ",FundRet.InceptDate.GIA.NonStdzed " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strHistPeriodEndingTest(ix), clsReadVAValues.strHistPeriodEndingBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strHistPeriodEndingTest(ix)
                        Dim strBench As String = clsReadVAValues.strHistPeriodEndingBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strHistPeriodEndingMM(ix) = clsReadVAValues.strHistPeriodEndingMM(ix) & ib & ",Fund.EndOfContractYear " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strHistCumulativeReturnTest(ix), clsReadVAValues.strHistCumulativeReturnBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strHistCumulativeReturnTest(ix)
                        Dim strBench As String = clsReadVAValues.strHistCumulativeReturnBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strHistCumulativeReturnMM(ix) = clsReadVAValues.strHistCumulativeReturnMM(ix) & ib & ",Fund.CumulativeRet.Curr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strHistAverageAnnReturnTest(ix), clsReadVAValues.strHistAverageAnnReturnBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strHistAverageAnnReturnTest(ix)
                        Dim strBench As String = clsReadVAValues.strHistAverageAnnReturnBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strHistAverageAnnReturnMM(ix) = clsReadVAValues.strHistAverageAnnReturnMM(ix) & ib & ",Fund.Avrge.AnnRet.Curr " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadVAValues.strHistCumulativeReturnMaxTest(ix), clsReadVAValues.strHistCumulativeReturnMaxBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strHistCumulativeReturnMaxTest(ix)
                        Dim strBench As String = clsReadVAValues.strHistCumulativeReturnMaxBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strHistCumulativeReturnMaxMM(ix) = clsReadVAValues.strHistCumulativeReturnMaxMM(ix) & ib & ",Fund.CumulativeRet.Max " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If


                    If (String.Compare(clsReadVAValues.strHistAverageAnnReturnMaxTest(ix), clsReadVAValues.strHistAverageAnnReturnMaxBench(ix), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strHistAverageAnnReturnMaxTest(ix)
                        Dim strBench As String = clsReadVAValues.strHistAverageAnnReturnMaxBench(ix)

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strHistAverageAnnReturnMaxMM(ix) = clsReadVAValues.strHistAverageAnnReturnMaxMM(ix) & ib & ",Fund.Avrge.AnnRet.Max " & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                Next ix
            End If

            If (String.Compare(CStr(clsReadVAValues.strIncomeStartAgeTest), CStr(clsReadVAValues.strIncomeStartAgeBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strIncomeStartAgeMM = ib & ",Income.StartAge," & clsReadVAValues.strIncomeStartAgeBench & "," & clsReadVAValues.strIncomeStartAgeTest & "&"
            Else

            End If

            If (String.Compare(CStr(clsReadVAValues.strIncomeStartAgeJointTest), CStr(clsReadVAValues.strIncomeStartAgeJointBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strIncomeStartAgeJointMM = ib & ",JointtIncome.StartAge," & clsReadVAValues.strIncomeStartAgeJointBench & "," & clsReadVAValues.strIncomeStartAgeJointTest & "&"
            Else

            End If

            If (String.Compare(CStr(clsReadVAValues.strIncomeStartYearTest), CStr(clsReadVAValues.strIncomeStartYearBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strIncomeStartYearMM = ib & ",Income.StartYr," & clsReadVAValues.strIncomeStartYearBench & "," & clsReadVAValues.strIncomeStartYearTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strIncomeStartMonthTest), CStr(clsReadVAValues.strIncomeStartMonthBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strIncomeStartMonthMM = ib & ",Income.StartMnth," & clsReadVAValues.strIncomeStartMonthBench & "," & clsReadVAValues.strIncomeStartMonthTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strYearsCertainTest), CStr(clsReadVAValues.strYearsCertainBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strYearsCertainMM = ib & ",Certain.Yrs," & clsReadVAValues.strYearsCertainBench & "," & clsReadVAValues.strYearsCertainTest & "&"
            End If

            If clsReadVAValues.gbIPRBench = False And clsReadVAValues.gbIPRTest = False Then

                If (String.Compare(CStr(clsReadVAValues.strGIAInitialMonthlyPayoutHypoTest), CStr(clsReadVAValues.strGIAInitialMonthlyPayoutHypoBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strGIAInitialMonthlyPayoutHypoMM = ib & ",Hypo.GIA.GuarIncomeFloor.Mthly.Curr," & clsReadVAValues.strGIAInitialMonthlyPayoutHypoBench & "," & clsReadVAValues.strGIAInitialMonthlyPayoutHypoTest & "&"
                End If

                If (String.Compare(CStr(clsReadVAValues.strGIAInitialMonthlyPayoutZeroTest), CStr(clsReadVAValues.strGIAInitialMonthlyPayoutZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strGIAInitialMonthlyPayoutZeroMM = ib & ",Hypo.GIA.GuarIncomeFloor.Mthly.Zero," & clsReadVAValues.strGIAInitialMonthlyPayoutZeroBench & "," & clsReadVAValues.strGIAInitialMonthlyPayoutZeroTest & "&"
                End If

                If (String.Compare(CStr(clsReadVAValues.strGIASchedInstallmentTest), CStr(clsReadVAValues.strGIASchedInstallmentBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strGIASchedInstallmentMM = ib & ",Hypo.GIA.MonthlySysTransfers," & clsReadVAValues.strGIASchedInstallmentBench & "," & clsReadVAValues.strGIASchedInstallmentTest & "&"
                End If

                If (String.Compare(CStr(clsReadVAValues.strAnnAmountZeroTest), CStr(clsReadVAValues.strAnnAmountZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strAnnAmountZeroMM = ib & ",Hypo.GIA.AnnuitizedAmt.Zero," & clsReadVAValues.strAnnAmountZeroBench & "," & clsReadVAValues.strAnnAmountZeroTest & "&"
                End If

                If (String.Compare(CStr(clsReadVAValues.strAnnAmountHypoTest), CStr(clsReadVAValues.strAnnAmountHypoBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strAnnAmountHypoMM = ib & ",Hypo.GIA.AnnuitizedAmt.Curr," & clsReadVAValues.strAnnAmountHypoBench & "," & clsReadVAValues.strAnnAmountHypoTest & "&"
                End If

                If (String.Compare(CStr(clsReadVAValues.strInstallmentCountZeroTest), CStr(clsReadVAValues.strInstallmentCountZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strInstallmentCountZeroMM = ib & ",Hypo.GIA.InstallmentCount.Zero," & clsReadVAValues.strInstallmentCountZeroBench & "," & clsReadVAValues.strInstallmentCountZeroTest & "&"
                End If

                If (String.Compare(CStr(clsReadVAValues.strInstallmentCountHypoTest), CStr(clsReadVAValues.strInstallmentCountHypoBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strInstallmentCountHypoMM = ib & ",Hypo.GIA.InstallmentCount.Curr," & clsReadVAValues.strInstallmentCountHypoBench & "," & clsReadVAValues.strInstallmentCountHypoTest & "&"
                End If
            End If

            If (String.Compare(CStr(clsReadVAValues.strPPDBChargeTest), CStr(clsReadVAValues.strPPDBChargeBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strPPDBChargeMM = ib & ",Charge.PPDB," & clsReadVAValues.strPPDBChargeBench & "," & clsReadVAValues.strPPDBChargeTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPFactorFirstWDTest), CStr(clsReadVAValues.strLIPFactorFirstWDBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPFactorFirstWDMM = ib & ",LIP.WDLimit.1stWD," & clsReadVAValues.strLIPFactorFirstWDBench & "," & clsReadVAValues.strLIPFactorFirstWDTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPGuarWithdrawalTest), CStr(clsReadVAValues.strLIPGuarWithdrawalBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPGuarWithdrawalMM = ib & ",LIP.Guar.WD," & clsReadVAValues.strLIPGuarWithdrawalBench & "," & clsReadVAValues.strLIPGuarWithdrawalTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPWDStartYearTest), CStr(clsReadVAValues.strLIPWDStartYearBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPWDStartYearMM = ib & ",LIP.WD.StartYr," & clsReadVAValues.strLIPWDStartYearBench & "," & clsReadVAValues.strLIPWDStartYearTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPWDStartMonthTest), CStr(clsReadVAValues.strLIPWDStartMonthBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPWDStartMonthMM = ib & ",LIP.WDStart.Mth," & clsReadVAValues.strLIPWDStartMonthBench & "," & clsReadVAValues.strLIPWDStartMonthTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPBenBaseFirstWDTest), CStr(clsReadVAValues.strLIPBenBaseFirstWDBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPBenBaseFirstWDMM = ib & ",LIP.BenBase.1stWD," & clsReadVAValues.strLIPBenBaseFirstWDBench & "," & clsReadVAValues.strLIPBenBaseFirstWDTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strTaxExcludableAmtZeroTest), CStr(clsReadVAValues.strTaxExcludableAmtZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strTaxExcludableAmtZeroMM = ib & ",Hypo.TaxExclAmt.Zero," & clsReadVAValues.strTaxExcludableAmtZeroBench & "," & clsReadVAValues.strTaxExcludableAmtZeroTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strTaxExcludableAmtHypoTest), CStr(clsReadVAValues.strTaxExcludableAmtHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strTaxExcludableAmtHypoMM = ib & ",Hypo.TaxExclAmt.Curr," & clsReadVAValues.strTaxExcludableAmtHypoBench & "," & clsReadVAValues.strTaxExcludableAmtHypoTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strTaxExcludableAmtHistTest), CStr(clsReadVAValues.strTaxExcludableAmtHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strTaxExcludableAmtHistMM = ib & ",Hist.TaxExclAmt," & clsReadVAValues.strTaxExcludableAmtHistBench & "," & clsReadVAValues.strTaxExcludableAmtHistTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strTaxBracketTest), CStr(clsReadVAValues.strTaxBracketBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strTaxBracketMM = ib & ",Tax.Bracket," & clsReadVAValues.strTaxBracketBench & "," & clsReadVAValues.strTaxBracketTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strTaxBasisTest), CStr(clsReadVAValues.strTaxBasisBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strTaxBasisMM = ib & ",Tax.CostBasis," & clsReadVAValues.strTaxBasisBench & "," & clsReadVAValues.strTaxBasisTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strInvestmentTest), CStr(clsReadVAValues.strInvestmentBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strInvestmentTest
                Dim strBench As String = clsReadVAValues.strInvestmentBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strInvestmentMM = clsReadVAValues.strInvestmentMM & ib & ",Premium.Annual," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If clsReadVAValues.gbIPRBench = False And clsReadVAValues.gbIPRTest = False Then
                If (String.Compare(CStr(clsReadVAValues.strBaseContractValueZeroTest), CStr(clsReadVAValues.strBaseContractValueZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    'if dont match, then put into array to find exactly which ones dont match, and make sure arrays have same  # of elements
                    Dim strTest As String = clsReadVAValues.strBaseContractValueZeroTest
                    Dim strBench As String = clsReadVAValues.strBaseContractValueZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If
                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strBaseContractValueZeroMM = clsReadVAValues.strBaseContractValueZeroMM & ib & ",Hypo.BaseContrValue.Zero, " & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strCombinedSurrValueZeroTest), CStr(clsReadVAValues.strCombinedSurrValueZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strCombinedSurrValueZeroTest
                    Dim strBench As String = clsReadVAValues.strCombinedSurrValueZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strCombinedSurrValueZeroMM = clsReadVAValues.strCombinedSurrValueZeroMM & ib & ",Hypo.CombSurrValue.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strCombinedSurrValueHypoTest), CStr(clsReadVAValues.strCombinedSurrValueHypoBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strCombinedSurrValueHypoTest
                    Dim strBench As String = clsReadVAValues.strCombinedSurrValueHypoBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strCombinedSurrValueHypoMM = clsReadVAValues.strCombinedSurrValueHypoMM & ib & ",Hypo.CombSurrValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strGISValueZeroTest), CStr(clsReadVAValues.strGISValueZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strGISValueZeroTest
                    Dim strBench As String = clsReadVAValues.strGISValueZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strGISValueZeroMM = clsReadVAValues.strGISValueZeroMM & ib & ",Hypo.GuarIncomeSubAcctValue.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strAnnualIncomeZeroTest), CStr(clsReadVAValues.strAnnualIncomeZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strAnnualIncomeZeroTest
                    Dim strBench As String = clsReadVAValues.strAnnualIncomeZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strAnnualIncomeZeroMM = clsReadVAValues.strAnnualIncomeZeroMM & ib & ",Hypo.AnnualIncome.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDeathBenefitZeroTest), CStr(clsReadVAValues.strDeathBenefitZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDeathBenefitZeroTest
                    Dim strBench As String = clsReadVAValues.strDeathBenefitZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDeathBenefitZeroMM = clsReadVAValues.strDeathBenefitZeroMM & ib & ",Hypo.DeathBen.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strTransfertoGISZeroTest), CStr(clsReadVAValues.strTransfertoGISZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strTransfertoGISZeroTest
                    Dim strBench As String = clsReadVAValues.strTransfertoGISZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strTransfertoGISZeroMM = clsReadVAValues.strTransfertoGISZeroMM & ib & ",Hypo.AnnTransferToGIS.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strTotalContractValueZeroTest), CStr(clsReadVAValues.strTotalContractValueZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strTotalContractValueZeroTest
                    Dim strBench As String = clsReadVAValues.strTotalContractValueZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strTotalContractValueZeroMM = clsReadVAValues.strTotalContractValueZeroMM & ib & ",Hypo.TotContractValue.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If
            If (String.Compare(CStr(clsReadVAValues.strSurrenderChargesTest), CStr(clsReadVAValues.strSurrenderChargesBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strSurrenderChargesTest
                Dim strBench As String = clsReadVAValues.strSurrenderChargesBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strSurrenderChargesMM = clsReadVAValues.strSurrenderChargesMM & ib & ",Charges.Surrender," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strHypoAnnIncomeFloorTest), CStr(clsReadVAValues.strHypoAnnIncomeFloorBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strHypoAnnIncomeFloorMM = ib & ",Hypo.AnnGuarIncomeFloor," & clsReadVAValues.strHypoAnnIncomeFloorBench & "," & clsReadVAValues.strHypoAnnIncomeFloorTest & "&"
            End If

            If clsReadVAValues.gbIPRBench = False And clsReadVAValues.gbIPRTest = False Then

                If (String.Compare(CStr(clsReadVAValues.strGIASumGtdAmtTest), CStr(clsReadVAValues.strGIASumGtdAmtBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strGIASumGtdAmtTest
                    Dim strBench As String = clsReadVAValues.strGIASumGtdAmtBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strGIASumGtdAmtMM = clsReadVAValues.strGIASumGtdAmtMM & ib & ",Hypo.GIATotalGtdAmt," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If

            If (String.Compare(CStr(clsReadVAValues.strHistReturnForPeriodTest), CStr(clsReadVAValues.strHistReturnForPeriodBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strHistReturnForPeriodTest
                Dim strBench As String = clsReadVAValues.strHistReturnForPeriodBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strHistReturnForPeriodMM = clsReadVAValues.strHistReturnForPeriodMM & ib & ",Hist.ReturnForPeriod," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strHistAnnIncomeTest), CStr(clsReadVAValues.strHistAnnIncomeBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strHistAnnIncomeTest
                Dim strBench As String = clsReadVAValues.strHistAnnIncomeBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strHistAnnIncomeMM = clsReadVAValues.strHistAnnIncomeMM & ib & ",Hist.AnnualIncome," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strHistAccountGISTest), CStr(clsReadVAValues.strHistAccountGISBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strHistAccountGISTest
                Dim strBench As String = clsReadVAValues.strHistAccountGISBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strHistAccountGISMM = clsReadVAValues.strHistAccountGISMM & ib & ",Hist.TotalContrValue," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If clsReadVAValues.gbIPRBench = True And clsReadVAValues.gbIPRTest = True Then
                If (String.Compare(CStr(clsReadVAValues.strIPRHistTotalContractValueCurrTest), CStr(clsReadVAValues.strIPRHistTotalContractValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistTotalContractValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHistTotalContractValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistTotalContractValueCurrMM = clsReadVAValues.strIPRHistTotalContractValueCurrMM & ib & ",Hist.ContrValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If clsReadVAValues.gbIPRBench = True And clsReadVAValues.gbIPRTest = True Then

                    If (String.Compare(CStr(clsReadVAValues.strIPRHistTotalSurrValueCurrTest), CStr(clsReadVAValues.strIPRHistTotalSurrValueCurrBench), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strIPRHistTotalSurrValueCurrTest
                        Dim strBench As String = clsReadVAValues.strIPRHistTotalSurrValueCurrBench

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strIPRHistTotalSurrValueCurrMM = clsReadVAValues.strIPRHistTotalSurrValueCurrMM & ib & ",Hist.SurrValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(CStr(clsReadVAValues.strIPRHistTotalContractValueMaxTest), CStr(clsReadVAValues.strIPRHistTotalContractValueMaxBench), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strIPRHistTotalContractValueMaxTest
                        Dim strBench As String = clsReadVAValues.strIPRHistTotalContractValueMaxBench

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strIPRHistTotalContractValueMaxMM = clsReadVAValues.strIPRHistTotalContractValueMaxMM & ib & ",Hist.ContrValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(CStr(clsReadVAValues.strIPRHistTotalSurrValueMaxTest), CStr(clsReadVAValues.strIPRHistTotalSurrValueMaxBench), True)) <> 0 Then
                        clsReadVAValues.bMisMatch = True

                        Dim strTest As String = clsReadVAValues.strIPRHistTotalSurrValueMaxTest
                        Dim strBench As String = clsReadVAValues.strIPRHistTotalSurrValueMaxBench

                        Dim SplitTest = Split(strTest, ",")
                        Dim SplitBench = Split(strBench, ",")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadVAValues.strIPRHistTotalSurrValueMaxMM = clsReadVAValues.strIPRHistTotalSurrValueMaxMM & ib & ",Hist.SurrValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If
                End If
            End If

            If clsReadVAValues.gbIPRBench = False And clsReadVAValues.gbIPRTest = False Then

                If (String.Compare(CStr(clsReadVAValues.strHistDeathBenefitGIATest), CStr(clsReadVAValues.strHistDeathBenefitGIABench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strHistDeathBenefitGIATest
                    Dim strBench As String = clsReadVAValues.strHistDeathBenefitGIABench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strHistDeathBenefitGIAMM = clsReadVAValues.strHistDeathBenefitGIAMM & ib & ",Hist.DB.GIA," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If

            If (String.Compare(CStr(clsReadVAValues.strPPBAHypoRateTest), CStr(clsReadVAValues.strPPBAHypoRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strPPBAHypoRateMM = ib & ",Hypo.PurchPayBenAmt.Curr," & clsReadVAValues.strPPBAHypoRateBench & "," & clsReadVAValues.strPPBAHypoRateTest & "&"
            Else

            End If

            If (String.Compare(CStr(clsReadVAValues.strPPBAZeroRateTest), CStr(clsReadVAValues.strPPBAZeroRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strPPBAZeroRateMM = ib & ",Hypo.PurchPayBenAmt.Zero," & clsReadVAValues.strPPBAZeroRateBench & "," & clsReadVAValues.strPPBAZeroRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strBenefitBaseZeroRateTest), CStr(clsReadVAValues.strBenefitBaseZeroRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strBenefitBaseZeroRateMM = ib & ",Hypo.BenBase.Zero," & clsReadVAValues.strBenefitBaseZeroRateBench & "," & clsReadVAValues.strBenefitBaseZeroRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strBenefitBaseHypoRateTest), CStr(clsReadVAValues.strBenefitBaseHypoRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strBenefitBaseHypoRateMM = ib & ",Hypo.BenBase.Curr," & clsReadVAValues.strBenefitBaseHypoRateBench & "," & clsReadVAValues.strBenefitBaseHypoRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strRollupZeroRateTest), CStr(clsReadVAValues.strRollupZeroRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strRollupZeroRateMM = ib & ",Hypo.RollupValue.Zero," & clsReadVAValues.strRollupZeroRateBench & "," & clsReadVAValues.strRollupZeroRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strRollupHypoRateTest), CStr(clsReadVAValues.strRollupHypoRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strRollupHypoRateMM = ib & ",Hypo.RollupValue.Curr," & clsReadVAValues.strRollupHypoRateBench & "," & clsReadVAValues.strRollupHypoRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPAnnIncomeZeroRateTest), CStr(clsReadVAValues.strLIPAnnIncomeZeroRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPAnnIncomeZeroRateMM = ib & ",Hypo.AnnIncome.Zero," & clsReadVAValues.strLIPAnnIncomeZeroRateBench & "," & clsReadVAValues.strLIPAnnIncomeZeroRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPAnnIncomeHypoRateTest), CStr(clsReadVAValues.strLIPAnnIncomeHypoRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPAnnIncomeHypoRateMM = ib & ",Hypo.AnnIncome.Curr," & clsReadVAValues.strLIPAnnIncomeHypoRateBench & "," & clsReadVAValues.strLIPAnnIncomeHypoRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strLIPResetValueZeroRateTest), CStr(clsReadVAValues.strLIPResetValueZeroRateBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strLIPResetValueZeroRateMM = ib & ",Hypo.MaxAnnvsyValue(Reset).Zero," & clsReadVAValues.strLIPResetValueZeroRateBench & "," & clsReadVAValues.strLIPResetValueZeroRateTest & "&"
            End If


            If (String.Compare(CStr(clsReadVAValues.strLIPContractValueHypoTest), CStr(clsReadVAValues.strLIPContractValueHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strLIPContractValueHypoTest
                Dim strBench As String = clsReadVAValues.strLIPContractValueHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strLIPContractValueHypoMM = clsReadVAValues.strLIPContractValueHypoMM & ib & ",Hypo.ContrValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(CStr(clsReadVAValues.strBASEWithdrawalZeroTest), CStr(clsReadVAValues.strBASEWithdrawalZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strBASEWithdrawalZeroTest
                Dim strBench As String = clsReadVAValues.strBASEWithdrawalZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strBASEWithdrawalZeroMM = clsReadVAValues.strBASEWithdrawalZeroMM & ib & ",Hypo.Withdrawal.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strBASEWithdrawalHypoTest), CStr(clsReadVAValues.strBASEWithdrawalHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strBASEWithdrawalHypoTest
                Dim strBench As String = clsReadVAValues.strBASEWithdrawalHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strBASEWithdrawalHypoMM = clsReadVAValues.strBASEWithdrawalHypoMM & ib & ",Hypo.Withdrawal.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strEPRZeroTest), CStr(clsReadVAValues.strEPRZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strEPRZeroMM = ib & ",Hypo.EPRDB.Zero," & clsReadVAValues.strEPRZeroBench & "," & clsReadVAValues.strEPRZeroTest & "&"
            End If

            If (String.Compare(CStr(clsReadVAValues.strEPRHypoTest), CStr(clsReadVAValues.strEPRHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strEPRHypoMM = ib & ",Hypo.EPRDB.Curr," & clsReadVAValues.strEPRHypoBench & "," & clsReadVAValues.strEPRHypoTest & "&"
            End If

            'If (String.Compare(CStr(clsReadVAValues.strEPRHistTest), CStr(clsReadVAValues.strEPRHistBench), True)) <> 0 Then
            '    clsReadVAValues.bMisMatch = True
            '    clsReadVAValues.strEPRHistMM = ib & ",EPR Hist," & clsReadVAValues.strEPRHistBench & "," & clsReadVAValues.strEPRHistTest & "&"
            'End If


            If (String.Compare(CStr(clsReadVAValues.strEPRHistTest), CStr(clsReadVAValues.strEPRHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strEPRHistTest
                Dim strBench As String = clsReadVAValues.strEPRHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strEPRHistMM = clsReadVAValues.strEPRHistMM & ib & ",Hist.EPRDB," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If clsReadVAValues.gbIPRBench = False And clsReadVAValues.gbIPRTest = False Then

                If (String.Compare(CStr(clsReadVAValues.strDBGIAZeroTest), CStr(clsReadVAValues.strDBGIAZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBGIAZeroTest
                    Dim strBench As String = clsReadVAValues.strDBGIAZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBGIAZeroMM = clsReadVAValues.strDBGIAZeroMM & ib & ",Hypo.DB.DETAIL.GIA.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBGIAHypoTest), CStr(clsReadVAValues.strDBGIAHypoBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBGIAHypoTest
                    Dim strBench As String = clsReadVAValues.strDBGIAHypoBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBGIAHypoMM = clsReadVAValues.strDBGIAHypoMM & ib & ",Hypo.DB.DETAIL.GIA.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBLIPZeroTest), CStr(clsReadVAValues.strDBLIPZeroBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBLIPZeroTest
                    Dim strBench As String = clsReadVAValues.strDBLIPZeroBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBLIPZeroMM = clsReadVAValues.strDBLIPZeroMM & ib & ",Hypo.DB.DETAIL.LIP.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBLIPHypoTest), CStr(clsReadVAValues.strDBLIPHypoBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBLIPHypoTest
                    Dim strBench As String = clsReadVAValues.strDBLIPHypoBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBLIPHypoMM = clsReadVAValues.strDBLIPHypoMM & ib & ",Hypo.DB.DETAIL.LIP.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBLIPHistTest), CStr(clsReadVAValues.strDBLIPHistBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBLIPHistTest
                    Dim strBench As String = clsReadVAValues.strDBLIPHistBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBLIPHistMM = clsReadVAValues.strDBLIPHistMM & ib & ",Hist.DB.DETAIL.LIP," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If
            If (String.Compare(CStr(clsReadVAValues.strDBComboZeroTest), CStr(clsReadVAValues.strDBComboZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBComboZeroTest
                Dim strBench As String = clsReadVAValues.strDBComboZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBComboZeroMM = clsReadVAValues.strDBComboZeroMM & ib & ",Hypo.DB.DETAIL.Combo.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBComboHypoTest), CStr(clsReadVAValues.strDBComboHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBComboHypoTest
                Dim strBench As String = clsReadVAValues.strDBComboHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBComboHypoMM = clsReadVAValues.strDBComboHypoMM & ib & ",Hypo.DB.DETAIL.Combo.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBComboHistTest), CStr(clsReadVAValues.strDBComboHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBComboHistTest
                Dim strBench As String = clsReadVAValues.strDBComboHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBComboHistMM = clsReadVAValues.strDBComboHistMM & ib & ",Hist.DB.DETAIL.Combo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBASDBZeroTest), CStr(clsReadVAValues.strDBASDBZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBASDBZeroTest
                Dim strBench As String = clsReadVAValues.strDBASDBZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBASDBZeroMM = clsReadVAValues.strDBASDBZeroMM & ib & ",Hypo.DB.DETAIL.ASDB.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBASDBHypoTest), CStr(clsReadVAValues.strDBASDBHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBASDBHypoTest
                Dim strBench As String = clsReadVAValues.strDBASDBHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBASDBHypoMM = clsReadVAValues.strDBASDBHypoMM & ib & ",Hypo.DB.DETAIL.ASDB.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBASDBHistTest), CStr(clsReadVAValues.strDBASDBHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBASDBHistTest
                Dim strBench As String = clsReadVAValues.strDBASDBHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBASDBHistMM = clsReadVAValues.strDBASDBHistMM & ib & ",Hist.DB.DETAIL.ASDB," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBRollupZeroTest), CStr(clsReadVAValues.strDBRollupZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBRollupZeroTest
                Dim strBench As String = clsReadVAValues.strDBRollupZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBRollupZeroMM = clsReadVAValues.strDBRollupZeroMM & ib & ",Hypo.DB.DETAIL.Rollup.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBRollupHypoTest), CStr(clsReadVAValues.strDBRollupHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBRollupHypoTest
                Dim strBench As String = clsReadVAValues.strDBRollupHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBRollupHypoMM = clsReadVAValues.strDBRollupHypoMM & ib & ",Hypo.DB.DETAIL.Rollup.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBRollupHistTest), CStr(clsReadVAValues.strDBRollupHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBRollupHistTest
                Dim strBench As String = clsReadVAValues.strDBRollupHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBRollupHistMM = clsReadVAValues.strDBRollupHistMM & ib & ",Hist.DB.DETAIL.Rollup," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBStandardZeroTest), CStr(clsReadVAValues.strDBStandardZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBStandardZeroTest
                Dim strBench As String = clsReadVAValues.strDBStandardZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBStandardZeroMM = clsReadVAValues.strDBStandardZeroMM & ib & ",Hypo.DB.DETAIL.Std.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBStandardHypoTest), CStr(clsReadVAValues.strDBStandardHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBStandardHypoTest
                Dim strBench As String = clsReadVAValues.strDBStandardHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBStandardHypoMM = clsReadVAValues.strDBStandardHypoMM & ib & ",Hypo.DB.DETAIL.Std.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strDBStandardHistTest), CStr(clsReadVAValues.strDBStandardHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strDBStandardHistTest
                Dim strBench As String = clsReadVAValues.strDBStandardHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strDBStandardHistMM = clsReadVAValues.strDBStandardHistMM & ib & ",Hist.DB.DETAIL.Std," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'IPR Hist DB values
            If clsReadVAValues.gbIPRBench = True And clsReadVAValues.gbIPRTest = True Then

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistContractDBCurrTest), CStr(clsReadVAValues.strDBIPRHistContractDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistContractDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistContractDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistContractDBCurrMM = clsReadVAValues.strDBIPRHistContractDBCurrMM & ib & ",Hist.ContrDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistASDBCurrTest), CStr(clsReadVAValues.strDBIPRHistASDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistASDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistASDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistASDBCurrMM = clsReadVAValues.strDBIPRHistASDBCurrMM & ib & ",Hist.ASDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistBasicDBCurrTest), CStr(clsReadVAValues.strDBIPRHistBasicDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistBasicDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistBasicDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistBasicDBCurrMM = clsReadVAValues.strDBIPRHistBasicDBCurrMM & ib & ",Hist.BasicDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistRollUpDBCurrTest), CStr(clsReadVAValues.strDBIPRHistRollUpDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistRollUpDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistRollUpDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistRollUpDBCurrMM = clsReadVAValues.strDBIPRHistRollUpDBCurrMM & ib & ",Hist.RollUp.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistEPRDBCurrTest), CStr(clsReadVAValues.strDBIPRHistEPRDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistEPRDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistEPRDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistEPRDBCurrMM = clsReadVAValues.strDBIPRHistEPRDBCurrMM & ib & ",Hist.EPRDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistPPDBCurrTest), CStr(clsReadVAValues.strDBIPRHistPPDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistPPDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistPPDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistPPDBCurrMM = clsReadVAValues.strDBIPRHistPPDBCurrMM & ib & ",Hist.PPDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistContractDBMaxTest), CStr(clsReadVAValues.strDBIPRHistContractDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistContractDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistContractDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistContractDBMaxMM = clsReadVAValues.strDBIPRHistContractDBMaxMM & ib & ",Hist.ContrDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistASDBMaxTest), CStr(clsReadVAValues.strDBIPRHistASDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistASDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistASDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistASDBMaxMM = clsReadVAValues.strDBIPRHistASDBMaxMM & ib & ",Hist.ASDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistBasicDBMaxTest), CStr(clsReadVAValues.strDBIPRHistBasicDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistBasicDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistBasicDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistBasicDBmaxMM = clsReadVAValues.strDBIPRHistBasicDBmaxMM & ib & ",Hist.BasicDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistRollUpDBMaxTest), CStr(clsReadVAValues.strDBIPRHistRollUpDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistRollUpDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistRollUpDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistRollUpDBMaxMM = clsReadVAValues.strDBIPRHistRollUpDBMaxMM & ib & ",Hist.RollUp.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistEPRDBMaxTest), CStr(clsReadVAValues.strDBIPRHistEPRDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistEPRDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistEPRDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistEPRDBMaxMM = clsReadVAValues.strDBIPRHistEPRDBMaxMM & ib & ",Hist.EPRDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHistPPDBMaxTest), CStr(clsReadVAValues.strDBIPRHistPPDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHistPPDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHistPPDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHistPPDBMaxMM = clsReadVAValues.strDBIPRHistPPDBMaxMM & ib & ",Hist.PPDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                'IPR Hypo DB values

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoContractDBCurrTest), CStr(clsReadVAValues.strDBIPRHypoContractDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoContractDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoContractDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoContractDBCurrMM = clsReadVAValues.strDBIPRHypoContractDBCurrMM & ib & ",Hypo.ContrDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoASDBCurrTest), CStr(clsReadVAValues.strDBIPRHypoASDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoASDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoASDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoASDBCurrMM = clsReadVAValues.strDBIPRHypoASDBCurrMM & ib & ",Hypo.ASDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoBasicDBCurrTest), CStr(clsReadVAValues.strDBIPRHypoBasicDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoBasicDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoBasicDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoBasicDBCurrMM = clsReadVAValues.strDBIPRHypoBasicDBCurrMM & ib & ",Hypo.BasicDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoRollUpDBCurrTest), CStr(clsReadVAValues.strDBIPRHypoRollUpDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoRollUpDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoRollUpDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoRollUpDBCurrMM = clsReadVAValues.strDBIPRHypoRollUpDBCurrMM & ib & ",Hypo.RollUp.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoEPRDBCurrTest), CStr(clsReadVAValues.strDBIPRHypoEPRDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoEPRDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoEPRDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoEPRDBCurrMM = clsReadVAValues.strDBIPRHypoEPRDBCurrMM & ib & ",Hypo.EPRDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoPPDBCurrTest), CStr(clsReadVAValues.strDBIPRHypoPPDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoPPDBCurrTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoPPDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoPPDBCurrMM = clsReadVAValues.strDBIPRHypoPPDBCurrMM & ib & ",Hypo.PPDB.DETAIL.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoContractDBMaxTest), CStr(clsReadVAValues.strDBIPRHypoContractDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoContractDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoContractDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoContractDBMaxMM = clsReadVAValues.strDBIPRHypoContractDBMaxMM & ib & ",Hypo.ContrDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoASDBMaxTest), CStr(clsReadVAValues.strDBIPRHypoASDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoASDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoASDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoASDBMaxMM = clsReadVAValues.strDBIPRHypoASDBMaxMM & ib & ",Hypo.ASDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoBasicDBMaxTest), CStr(clsReadVAValues.strDBIPRHypoBasicDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoBasicDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoBasicDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoBasicDBmaxMM = clsReadVAValues.strDBIPRHypoBasicDBmaxMM & ib & ",Hypo.BasicDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoRollUpDBMaxTest), CStr(clsReadVAValues.strDBIPRHypoRollUpDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoRollUpDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoRollUpDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoRollUpDBMaxMM = clsReadVAValues.strDBIPRHypoRollUpDBMaxMM & ib & ",Hypo.RollUp.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoEPRDBMaxTest), CStr(clsReadVAValues.strDBIPRHypoEPRDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoEPRDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoEPRDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoEPRDBMaxMM = clsReadVAValues.strDBIPRHypoEPRDBMaxMM & ib & ",Hypo.EPRDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strDBIPRHypoPPDBMaxTest), CStr(clsReadVAValues.strDBIPRHypoPPDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strDBIPRHypoPPDBMaxTest
                    Dim strBench As String = clsReadVAValues.strDBIPRHypoPPDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strDBIPRHypoPPDBMaxMM = clsReadVAValues.strDBIPRHypoPPDBMaxMM & ib & ",Hypo.PPDB.DETAIL.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                'IPR WD Limits

                If (String.Compare(CStr(clsReadVAValues.strIPRWDLimitCurrTest), CStr(clsReadVAValues.strIPRWDLimitCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDLimitCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRWDLimitCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDLimitCurrMM = clsReadVAValues.strIPRWDLimitCurrMM & ib & ",Hypo.GuarWDLimit.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRWDLimitGtdTest), CStr(clsReadVAValues.strIPRWDLimitGtdBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDLimitGtdTest
                    Dim strBench As String = clsReadVAValues.strIPRWDLimitGtdBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDLimitGtdMM = clsReadVAValues.strIPRWDLimitGtdMM & ib & ",Hypo.GuarWDLimit.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRWDLimitHistCurrTest), CStr(clsReadVAValues.strIPRWDLimitHistCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDLimitHistCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRWDLimitHistCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDLimitHistMM = clsReadVAValues.strIPRWDLimitHistMM & ib & ",Hist.GuarWDLimit.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                'IPR WD Taken

                If (String.Compare(CStr(clsReadVAValues.strIPRWDTakenHypoCurrTest), CStr(clsReadVAValues.strIPRWDTakenHypoCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDTakenHypoCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRWDTakenHypoCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDTakenHypoCurrMM = clsReadVAValues.strIPRWDTakenHypoCurrMM & ib & ",Hypo.WDTaken.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRWDTakenHypoMaxTest), CStr(clsReadVAValues.strIPRWDTakenHypoMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDTakenHypoMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRWDTakenHypoMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDTakenHypoMaxMM = clsReadVAValues.strIPRWDTakenHypoMaxMM & ib & ",Hypo.WDTaken.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRWDTakenHistCurrTest), CStr(clsReadVAValues.strIPRWDTakenHistCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDTakenHistCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRWDTakenHistCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDTakenHistCurrMM = clsReadVAValues.strIPRWDTakenHistCurrMM & ib & ",Hist.WDTaken.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                If (String.Compare(CStr(clsReadVAValues.strIPRWDTakenHistMaxTest), CStr(clsReadVAValues.strIPRWDTakenHistMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDTakenHistMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRWDTakenHistMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDTakenHistMaxMM = clsReadVAValues.strIPRWDTakenHistMaxMM & ib & ",Hist.WDTaken.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If

            If clsReadVAValues.gbIPRBench = True And clsReadVAValues.gbIPRTest = True Then
                If (String.Compare(CStr(clsReadVAValues.strIPRWDLimitMaxTest), CStr(clsReadVAValues.strIPRWDLimitMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDLimitMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRWDLimitMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDLimitGtdGCMM = clsReadVAValues.strIPRWDLimitGtdGCMM & ib & ",Hypo.GuarWDLimit.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRWDLimitHistMaxTest), CStr(clsReadVAValues.strIPRWDLimitHistMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRWDLimitHistMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRWDLimitHistMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRWDLimitHistGCMM = clsReadVAValues.strIPRWDLimitHistGCMM & ib & ",Hist.GuarWDLimit.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If
            'RR One Return for Period

            If clsReadVAValues.gbIPRBench = True And clsReadVAValues.gbIPRTest = True Then

                If (String.Compare(CStr(clsReadVAValues.strRROneHistReturnForPeriodCurrTest), CStr(clsReadVAValues.strRROneHistReturnForPeriodCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strRROneHistReturnForPeriodCurrTest
                    Dim strBench As String = clsReadVAValues.strRROneHistReturnForPeriodCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strRROneHistReturnForPeriodCurrMM = clsReadVAValues.strRROneHistReturnForPeriodCurrMM & ib & ",Hist.RetForPd.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strRROneHistReturnForPeriodMaxTest), CStr(clsReadVAValues.strRROneHistReturnForPeriodMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strRROneHistReturnForPeriodMaxTest
                    Dim strBench As String = clsReadVAValues.strRROneHistReturnForPeriodMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strRROneHistReturnForPeriodMaxMM = clsReadVAValues.strRROneHistReturnForPeriodMaxMM & ib & ",Hist.RetForPd.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                'IPR Hypo Values

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoPPBAMaxTest), CStr(clsReadVAValues.strIPRHypoPPBAMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoPPBAMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoPPBAMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoPPBAMaxMM = clsReadVAValues.strIPRHypoPPBAMaxMM & ib & ",Hypo.PurchPayBemAmt.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoMaxAnnValueMaxTest), CStr(clsReadVAValues.strIPRHypoMaxAnnValueMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoMaxAnnValueMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoMaxAnnValueMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoMaxAnnValueMaxMM = clsReadVAValues.strIPRHypoMaxAnnValueMaxMM & ib & ",Hypo.MaxAnnverValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoRollupValueMaxTest), CStr(clsReadVAValues.strIPRHypoRollupValueMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoRollupValueMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoRollupValueMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoRollupValueMaxMM = clsReadVAValues.strIPRHypoRollupValueMaxMM & ib & ",Hypo.RollupValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoBenefitBaseMaxTest), CStr(clsReadVAValues.strIPRHypoBenefitBaseMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoBenefitBaseMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoBenefitBaseMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoBenefitBaseMaxMM = clsReadVAValues.strIPRHypoBenefitBaseMaxMM & ib & ",Hypo.BenBase.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoDBMaxTest), CStr(clsReadVAValues.strIPRHypoDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoDBMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoDBMaxMM = clsReadVAValues.strIPRHypoDBMaxMM & ib & ",Hypo.DB.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoPPBACurrTest), CStr(clsReadVAValues.strIPRHypoPPBACurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoPPBACurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoPPBACurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoPPBACurrMM = clsReadVAValues.strIPRHypoPPBACurrMM & ib & ",Hypo.PurchPayBenAmt.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoMaxAnnValueCurrTest), CStr(clsReadVAValues.strIPRHypoMaxAnnValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoMaxAnnValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoMaxAnnValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoMaxAnnValueCurrMM = clsReadVAValues.strIPRHypoMaxAnnValueCurrMM & ib & ",Hypo.MaxAnnverValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoRollupValueCurrTest), CStr(clsReadVAValues.strIPRHypoRollupValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoRollupValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoRollupValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoRollupValueCurrMM = clsReadVAValues.strIPRHypoRollupValueCurrMM & ib & ",Hypo.RollupValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoBenefitBaseCurrTest), CStr(clsReadVAValues.strIPRHypoBenefitBaseCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoBenefitBaseCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoBenefitBaseCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoBenefitBaseCurrMM = clsReadVAValues.strIPRHypoBenefitBaseCurrMM & ib & ",Hypo.BenBase.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoDBCurrTest), CStr(clsReadVAValues.strIPRHypoDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoDBCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoDBCurrMM = clsReadVAValues.strIPRHypoDBCurrMM & ib & ",Hypo.DB.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoContractValueCurrTest), CStr(clsReadVAValues.strIPRHypoContractValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoContractValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoContractValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoContractValueCurrMM = clsReadVAValues.strIPRHypoContractValueCurrMM & ib & ",Hypo.ContValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoSurrenderValueCurrTest), CStr(clsReadVAValues.strIPRHypoSurrenderValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoSurrenderValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoSurrenderValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoSurrenderValueCurrMM = clsReadVAValues.strIPRHypoSurrenderValueCurrMM & ib & ",Hypo.SurrValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoContractValueMaxTest), CStr(clsReadVAValues.strIPRHypoContractValueMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoContractValueMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoContractValueMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoContractValueMaxMM = clsReadVAValues.strIPRHypoContractValueMaxMM & ib & ",Hypo.ContValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHypoSurrenderValueMaxTest), CStr(clsReadVAValues.strIPRHypoSurrenderValueMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHypoSurrenderValueMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHypoSurrenderValueMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHypoSurrenderValueMaxMM = clsReadVAValues.strIPRHypoSurrenderValueMaxMM & ib & ",Hypo.SurrValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If


                'IPR Hist Values

                If (String.Compare(CStr(clsReadVAValues.strIPRHistPPBAMaxTest), CStr(clsReadVAValues.strIPRHistPPBAMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistPPBAMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHistPPBAMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistPPBAMaxMM = clsReadVAValues.strIPRHistPPBAMaxMM & ib & ",Hist.PurchPayBen.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistMaxAnnValueMaxTest), CStr(clsReadVAValues.strIPRHistMaxAnnValueMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistMaxAnnValueMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHistMaxAnnValueMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistMaxAnnValueMaxMM = clsReadVAValues.strIPRHistMaxAnnValueMaxMM & ib & ",Hist.MaxAnnValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistRollupValueMaxTest), CStr(clsReadVAValues.strIPRHistRollupValueMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistRollupValueMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHistRollupValueMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistRollupValueMaxMM = clsReadVAValues.strIPRHistRollupValueMaxMM & ib & ",Hist.RollUpValue.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistBenefitBaseMaxTest), CStr(clsReadVAValues.strIPRHistBenefitBaseMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistBenefitBaseMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHistBenefitBaseMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistBenefitBaseMaxMM = clsReadVAValues.strIPRHistBenefitBaseMaxMM & ib & ",Hist.BenBase.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistDBMaxTest), CStr(clsReadVAValues.strIPRHistDBMaxBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistDBMaxTest
                    Dim strBench As String = clsReadVAValues.strIPRHistDBMaxBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistDBMaxMM = clsReadVAValues.strIPRHistDBMaxMM & ib & ",Hist.DB.Max," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistPPBACurrTest), CStr(clsReadVAValues.strIPRHistPPBACurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistPPBACurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHistPPBACurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistPPBACurrMM = clsReadVAValues.strIPRHistPPBACurrMM & ib & ",Hist.PurchPayBen.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistMaxAnnValueCurrTest), CStr(clsReadVAValues.strIPRHistMaxAnnValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistMaxAnnValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHistMaxAnnValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistMaxAnnValueCurrMM = clsReadVAValues.strIPRHistMaxAnnValueCurrMM & ib & ",Hist.MaxAnnValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistRollupValueCurrTest), CStr(clsReadVAValues.strIPRHistRollupValueCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistRollupValueCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHistRollupValueCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistRollupValueCurrMM = clsReadVAValues.strIPRHistRollupValueCurrMM & ib & ",Hist.RollupValue.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistBenefitBaseCurrTest), CStr(clsReadVAValues.strIPRHistBenefitBaseCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistBenefitBaseCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHistBenefitBaseCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistBenefitBaseCurrMM = clsReadVAValues.strIPRHistBenefitBaseCurrMM & ib & ",Hist.BenBase.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If

                If (String.Compare(CStr(clsReadVAValues.strIPRHistDBCurrTest), CStr(clsReadVAValues.strIPRHistDBCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strIPRHistDBCurrTest
                    Dim strBench As String = clsReadVAValues.strIPRHistDBCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strIPRHistDBCurrMM = clsReadVAValues.strIPRHistDBCurrMM & ib & ",Hist.DB.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If
            End If
            'RR One Annual Income not split up between period certain and living only yet
            If (String.Compare(CStr(clsReadVAValues.strRROneAnnualIncomeZeroTest), CStr(clsReadVAValues.strRROneAnnualIncomeZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strRROneAnnualIncomeZeroTest
                Dim strBench As String = clsReadVAValues.strRROneAnnualIncomeZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strRROneAnnualIncomeZeroMM = clsReadVAValues.strRROneAnnualIncomeZeroMM & ib & ",Hypo.AnnualIncome.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strRROneAnnualIncomeHistTest), CStr(clsReadVAValues.strRROneAnnualIncomeHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strRROneAnnualIncomeHistTest
                Dim strBench As String = clsReadVAValues.strRROneAnnualIncomeHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strRROneAnnualIncomeHistMM = clsReadVAValues.strRROneAnnualIncomeHistMM & ib & ",Hist.AnnualIncome," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'For MCC
            'Gtd Payment Floor Zero
            If (String.Compare(CStr(clsReadVAValues.strMCCGtdPaymentFloorZeroTest), CStr(clsReadVAValues.strMCCGtdPaymentFloorZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCGtdPaymentFloorZeroTest
                Dim strBench As String = clsReadVAValues.strMCCGtdPaymentFloorZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCGtdPaymentFloorZeroMM = clsReadVAValues.strMCCGtdPaymentFloorZeroMM & ib & ",Hypo.MCC.GtdPaymentFloor.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'Gtd Payment Floor Curr
            If (String.Compare(CStr(clsReadVAValues.strMCCGtdPaymentFloorHypoTest), CStr(clsReadVAValues.strMCCGtdPaymentFloorHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCGtdPaymentFloorHypoTest
                Dim strBench As String = clsReadVAValues.strMCCGtdPaymentFloorHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCGtdPaymentFloorHypoMM = clsReadVAValues.strMCCGtdPaymentFloorHypoMM & ib & ",Hypo.MCC.GtdPaymentFloor.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'Gtd Payment Floor Hist
            If (String.Compare(CStr(clsReadVAValues.strMCCGtdPaymentFloorHistTest), CStr(clsReadVAValues.strMCCGtdPaymentFloorHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCGtdPaymentFloorHistTest
                Dim strBench As String = clsReadVAValues.strMCCGtdPaymentFloorHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCGtdPaymentFloorHistMM = clsReadVAValues.strMCCGtdPaymentFloorHistMM & ib & ",Hypo.MCC.GtdPaymentFloor.Hist," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            ' MCC Gtd Payment Floor Factor Zero
            If (String.Compare(CStr(clsReadVAValues.strMCCGtdPaymentFloorFactorZeroTest), CStr(clsReadVAValues.strMCCGtdPaymentFloorFactorZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCGtdPaymentFloorFactorZeroTest
                Dim strBench As String = clsReadVAValues.strMCCGtdPaymentFloorFactorZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCGtdPaymentFloorFactorZeroMM = clsReadVAValues.strMCCGtdPaymentFloorFactorZeroMM & ib & ",Hypo.MCC.GtdPaymentFloorFactor.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            ' MCC Gtd Payment Floor Factor Curr
            If (String.Compare(CStr(clsReadVAValues.strMCCGtdPaymentFloorFactorHypoTest), CStr(clsReadVAValues.strMCCGtdPaymentFloorFactorHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCGtdPaymentFloorFactorHypoTest
                Dim strBench As String = clsReadVAValues.strMCCGtdPaymentFloorFactorHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCGtdPaymentFloorFactorHypoMM = clsReadVAValues.strMCCGtdPaymentFloorFactorHypoMM & ib & ",Hypo.MCC.GtdPaymentFloorFactor.Hypo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            ' MCC Gtd Payment Floor Factor Hist
            If (String.Compare(CStr(clsReadVAValues.strMCCGtdPaymentFloorFactorHistTest), CStr(clsReadVAValues.strMCCGtdPaymentFloorFactorHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCGtdPaymentFloorFactorHistTest
                Dim strBench As String = clsReadVAValues.strMCCGtdPaymentFloorFactorHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCGtdPaymentFloorFactorHistMM = clsReadVAValues.strMCCGtdPaymentFloorFactorHistMM & ib & ",Hist.MCC.GtdPaymentFloorFactor.Hist," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'MCC First Full Yr Income Zero
            If (String.Compare(clsReadVAValues.strMCCFirstFullYrIncomeZeroTest, clsReadVAValues.strMCCFirstFullYrIncomeZeroBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCFirstFullYrIncomeZeroMM = ib & ",MCC.FirstFullYrIncome.Zero," & clsReadVAValues.strMCCFirstFullYrIncomeZeroBench & "," & clsReadVAValues.strMCCFirstFullYrIncomeZeroTest & "&"
            End If

            'MCC First Full Yr Income Hypo
            If (String.Compare(clsReadVAValues.strMCCFirstFullYrIncomeHypoTest, clsReadVAValues.strMCCFirstFullYrIncomeHypoBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCFirstFullYrIncomeHypoMM = ib & ",MCC.FirstFullYrIncome.Hypo," & clsReadVAValues.strMCCFirstFullYrIncomeHypoBench & "," & clsReadVAValues.strMCCFirstFullYrIncomeHypoTest & "&"
            End If

            'MCC First Full Yr Income Hist
            If (String.Compare(clsReadVAValues.strMCCFirstFullYrIncomeHistTest, clsReadVAValues.strMCCFirstFullYrIncomeHistBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCFirstFullYrIncomeHistMM = ib & ",MCC.FirstFullYrIncome.Hist," & clsReadVAValues.strMCCFirstFullYrIncomeHistBench & "," & clsReadVAValues.strMCCFirstFullYrIncomeHistTest & "&"
            End If

            'MCC Life with Period Certain Of
            If (String.Compare(clsReadVAValues.strMCCLifeWithPeriodCertainOfTest, clsReadVAValues.strMCCLifeWithPeriodCertainOfBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCLifeWithPeriodCertainOfMM = ib & ",MCC.LifeWithPeriodCertainOf," & clsReadVAValues.strMCCLifeWithPeriodCertainOfBench & "," & clsReadVAValues.strMCCLifeWithPeriodCertainOfTest & "&"
            End If

            'MCC Guarantee From Plan
            If (String.Compare(clsReadVAValues.strMCCGuaranteeFromPlanTest, clsReadVAValues.strMCCGuaranteeFromPlanBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCGuaranteeFromPlanMM = ib & ",MCC.GuaranteeFromPlan," & clsReadVAValues.strMCCGuaranteeFromPlanBench & "," & clsReadVAValues.strMCCGuaranteeFromPlanTest & "&"
            End If

            'MCC Desired Retirement Age
            If (String.Compare(clsReadVAValues.strMCCDesiredRetirementAgeTest, clsReadVAValues.strMCCDesiredRetirementAgeBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCDesiredRetirementAgeMM = ib & ",MCC.DesiredRetirementAge," & clsReadVAValues.strMCCDesiredRetirementAgeBench & "," & clsReadVAValues.strMCCDesiredRetirementAgeTest & "&"
            End If

            'MCC Gtd Income Payments at 0% (chart)
            If (String.Compare(clsReadVAValues.strMCCGtdIncomePayments0Test, clsReadVAValues.strMCCGtdIncomePayments0Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCGtdIncomePayments0MM = ib & ",MCC.GtdIncomePayments0%," & clsReadVAValues.strMCCGtdIncomePayments0Bench & "," & clsReadVAValues.strMCCGtdIncomePayments0Test & "&"
            End If


            'MCC First Full Yr Income Hypo Net (chart)
            If (String.Compare(clsReadVAValues.strMCCFirstFullYrIncomeHypoNetTest, clsReadVAValues.strMCCFirstFullYrIncomeHypoNetBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCFirstFullYrIncomeHypoNetMM = ib & ",MCC.FirstFullYearIncomeHypoNet," & clsReadVAValues.strMCCFirstFullYrIncomeHypoNetBench & "," & clsReadVAValues.strMCCFirstFullYrIncomeHypoNetTest & "&"
            End If
            
            'MCC Total Gtd Income Payments at 0% (chart)
            If (String.Compare(clsReadVAValues.strMCCTotalGtdIncomePayments0Test, clsReadVAValues.strMCCTotalGtdIncomePayments0Bench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCTotalGtdIncomePayments0MM = ib & ",MCC.TotalGtdIncomePayments0%," & clsReadVAValues.strMCCTotalGtdIncomePayments0Bench & "," & clsReadVAValues.strMCCTotalGtdIncomePayments0Test & "&"
            End If


            'MCC Total Income Payments Hypo Net (chart)
            If (String.Compare(clsReadVAValues.strMCCTotalIncomePaymentsHypoNetTest, clsReadVAValues.strMCCTotalIncomePaymentsHypoNetBench, True)) <> 0 Then
                clsReadVAValues.bMisMatch = True
                clsReadVAValues.strMCCTotalIncomePaymentsHypoNetMM = ib & ",MCC.TotalIncomePaymentsHypoNet," & clsReadVAValues.strMCCTotalIncomePaymentsHypoNetBench & "," & clsReadVAValues.strMCCTotalIncomePaymentsHypoNetTest & "&"
            End If

            'MCC Adjustment Account

            If (String.Compare(CStr(clsReadVAValues.strMCCAdjustmentAccountZeroTest), CStr(clsReadVAValues.strMCCAdjustmentAccountZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCAdjustmentAccountZeroTest
                Dim strBench As String = clsReadVAValues.strMCCAdjustmentAccountZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCAdjustmentAccountZeroMM = clsReadVAValues.strMCCAdjustmentAccountZeroMM & ib & ",Hypo.MCC.AdjustmentAccount.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strMCCAdjustmentAccountHypoTest), CStr(clsReadVAValues.strMCCAdjustmentAccountHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCAdjustmentAccountHypoTest
                Dim strBench As String = clsReadVAValues.strMCCAdjustmentAccountHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCAdjustmentAccountHypoMM = clsReadVAValues.strMCCAdjustmentAccountHypoMM & ib & ",Hypo.MCC.AdjustmentAccount.Hypo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strMCCAdjustmentAccountHistTest), CStr(clsReadVAValues.strMCCAdjustmentAccountHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCAdjustmentAccountHistTest
                Dim strBench As String = clsReadVAValues.strMCCAdjustmentAccountHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCAdjustmentAccountHistMM = clsReadVAValues.strMCCAdjustmentAccountHistMM & ib & ",Hist.MCC.AdjustmentAccount," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'MCC Commutation Value

            If (String.Compare(CStr(clsReadVAValues.strMCCcommutationValueZeroTest), CStr(clsReadVAValues.strMCCcommutationValueZeroBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCcommutationValueZeroTest
                Dim strBench As String = clsReadVAValues.strMCCcommutationValueZeroBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCcommutationValueZeroMM = clsReadVAValues.strMCCcommutationValueZeroMM & ib & ",Hypo.MCC.CommutationValue.Zero," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strMCCCommutationValueHypoTest), CStr(clsReadVAValues.strMCCCommutationValueHypoBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCCommutationValueHypoTest
                Dim strBench As String = clsReadVAValues.strMCCCommutationValueHypoBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCCommutationValueHypoMM = clsReadVAValues.strMCCCommutationValueHypoMM & ib & ",Hypo.MCC.CommutationValue.Hypo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadVAValues.strMCCCommutationValueHistTest), CStr(clsReadVAValues.strMCCCommutationValueHistBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCCommutationValueHistTest
                Dim strBench As String = clsReadVAValues.strMCCCommutationValueHistBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCCommutationValueHistMM = clsReadVAValues.strMCCCommutationValueHistMM & ib & ",Hypo.MCC.CommutationValue.Hist," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            'MCC Hist Income Return
            If (String.Compare(CStr(clsReadVAValues.strMCCHistIncomePeriodReturnTest), CStr(clsReadVAValues.strMCCHistIncomePeriodReturnBench), True)) <> 0 Then
                clsReadVAValues.bMisMatch = True

                Dim strTest As String = clsReadVAValues.strMCCHistIncomePeriodReturnTest
                Dim strBench As String = clsReadVAValues.strMCCHistIncomePeriodReturnBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadVAValues.strMCCHistIncomePeriodReturnMM = clsReadVAValues.strMCCHistIncomePeriodReturnMM & ib & ",Hist.MCC.IncomePeriodReturn," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If



            If clsReadVAValues.gbIPRBench = True And clsReadVAValues.gbIPRTest = True Then

                If (String.Compare(CStr(clsReadVAValues.strRROneAnnualIncomeCurrTest), CStr(clsReadVAValues.strRROneAnnualIncomeCurrBench), True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True

                    Dim strTest As String = clsReadVAValues.strRROneAnnualIncomeCurrTest
                    Dim strBench As String = clsReadVAValues.strRROneAnnualIncomeCurrBench

                    Dim SplitTest = Split(strTest, ",")
                    Dim SplitBench = Split(strBench, ",")

                    'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                    If UBound(SplitTest) < UBound(SplitBench) Then
                        iShort = UBound(SplitTest)
                        ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                        For iAdd = iShort To UBound(SplitTest)
                            SplitTest(iAdd) = "0"
                        Next
                    ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                        iShort = UBound(SplitBench)
                        ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                        For iAdd = iShort To UBound(SplitBench)
                            SplitBench(iAdd) = "0"
                        Next
                    End If

                    For iElement = 0 To UBound(SplitBench)
                        If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                            clsReadVAValues.strRROneAnnualIncomeCurrMM = clsReadVAValues.strRROneAnnualIncomeCurrMM & ib & ",Hypo.AnnualIncome.Curr," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                        End If
                    Next
                End If



                If (String.Compare(clsReadVAValues.strStartWDsYoungerBDTest, clsReadVAValues.strStartWDsYoungerBDBench, True)) <> 0 Then
                    clsReadVAValues.bMisMatch = True
                    clsReadVAValues.strStartWDsYoungerBDMM = ib & ",Start WDs Younger BD?," & clsReadVAValues.strStartWDsYoungerBDBench & "," & clsReadVAValues.strStartWDsYoungerBDTest & "&"
                End If


            End If
        End If


        'if cases don't match, create the mismatch reports
        If clsReadVAValues.bMisMatch = True Or gbVASaveAge(ib) Then
            clsReadVAValues.bMismatchAtLeastOnce = True
            SaveNewBench(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
            CreateVAMismatchReport(ib)
            strClientMisMatchList = strClientMisMatchList & "," & ib
            strsplitMMList = Split(strClientMisMatchList, ",")
            'WriteMismatchList(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If
        'copy the Test.pdf to the test folder under the client folder, whether there are mismatches or not
        If kclbClientList.GetItemChecked(ib - 1) Then
            CopyPDF(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If

        'write the stats to the file
        WriteMatchStatus(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp, clsReadVAValues.bMisMatch, bNewBench:=False)



    End Sub
    Private Sub CompareValuesSPDA(ByVal ic As Integer, ByVal ib As Integer, ByVal strcomp() As String)


        Dim ix As Integer = 0
        Dim imax As Integer = 0
        Dim iAdd As Integer = 0
        Dim iShort As Integer = 0

        'set mismatch flag to false
        clsReadSPDAValues.bMisMatch = False

        'go through each value in relay.out to compare between bench and test, if mismatch, create strings for datagrid

        'below 20 are for warning and error messages
        If clsReadSPDAValues.bErrorBench = True And clsReadSPDAValues.bErrorTest = True Then
            If (String.Compare(clsReadSPDAValues.strMessage1Test, clsReadSPDAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage1MM = ib & ",Message 1," & clsReadSPDAValues.strMessage1Bench & "," & clsReadSPDAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage2Test, clsReadSPDAValues.strMessage2Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage2MM = ib & ",Message 2," & clsReadSPDAValues.strMessage2Bench & "," & clsReadSPDAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage3Test, clsReadSPDAValues.strMessage3Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage3MM = ib & ",Message 3," & clsReadSPDAValues.strMessage3Bench & "," & clsReadSPDAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage4Test, clsReadSPDAValues.strMessage4Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage4MM = ib & ",Message 4," & clsReadSPDAValues.strMessage4Bench & "," & clsReadSPDAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage5Test, clsReadSPDAValues.strMessage5Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage5MM = ib & ",Message 5," & clsReadSPDAValues.strMessage5Bench & "," & clsReadSPDAValues.strMessage5Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage6Test, clsReadSPDAValues.strMessage6Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage6MM = ib & ",Message 6," & clsReadSPDAValues.strMessage6Bench & "," & clsReadSPDAValues.strMessage6Test & "&"
            End If

            'Set messages back to nothing
            clsReadSPDAValues.strMessage1Bench = ""
            clsReadSPDAValues.strMessage1Test = ""
            clsReadSPDAValues.strMessage2Bench = ""
            clsReadSPDAValues.strMessage2Test = ""
            clsReadSPDAValues.strMessage3Bench = ""
            clsReadSPDAValues.strMessage3Test = ""
            clsReadSPDAValues.strMessage4Bench = ""
            clsReadSPDAValues.strMessage4Test = ""
            clsReadSPDAValues.strMessage5Bench = ""
            clsReadSPDAValues.strMessage5Test = ""
            clsReadSPDAValues.strMessage6Bench = ""
            clsReadSPDAValues.strMessage6Test = ""

            'see if one runs and one doesnt
        ElseIf clsReadSPDAValues.bErrorBench = True And clsReadSPDAValues.bErrorTest = False Then
            clsReadSPDAValues.bMisMatch = True
            clsReadSPDAValues.strRunNoRunMM = ib & ",Test runs/Bench doesn't run,,&"

            If (String.Compare(clsReadSPDAValues.strMessage1Test, clsReadSPDAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage1MM = ib & ",Message1," & clsReadSPDAValues.strMessage1Bench & "," & clsReadSPDAValues.strMessage1Test & "&"
            End If

            clsReadSPDAValues.strMessage1Test = ""
            clsReadSPDAValues.strMessage1Bench = ""

        ElseIf clsReadSPDAValues.bErrorBench = False And clsReadSPDAValues.bErrorTest = True Then
            clsReadSPDAValues.bMisMatch = True
            clsReadSPDAValues.strRunNoRunMM = ib & ",Bench runs/Test doesn't run,,&"

            If (String.Compare(clsReadSPDAValues.strMessage1Test, clsReadSPDAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage1MM = ib & ",Message1," & clsReadSPDAValues.strMessage1Bench & "," & clsReadSPDAValues.strMessage1Test & "&"
            End If

            'If (String.Compare(clsReadSPDAValues.strMessage2Test, clsReadSPDAValues.strMessage2Bench, True)) <> 0 Then
            '    clsReadSPDAValues.bMisMatch = True
            '    clsReadSPDAValues.strMessage2MM = ib & ",Message2," & clsReadSPDAValues.strMessage2Bench & "," & clsReadSPDAValues.strMessage2Test & "&"
            'End If
            clsReadSPDAValues.strMessage1Test = ""
            clsReadSPDAValues.strMessage1Bench = ""
        Else
            'if both cases run, then compare...

            If (String.Compare(clsReadSPDAValues.strMessage1Test, clsReadSPDAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage1MM = ib & ",Message1," & clsReadSPDAValues.strMessage1Bench & "," & clsReadSPDAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage2Test, clsReadSPDAValues.strMessage2Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage2MM = ib & ",Message2," & clsReadSPDAValues.strMessage2Bench & "," & clsReadSPDAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage3Test, clsReadSPDAValues.strMessage3Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage3MM = ib & ",Message3," & clsReadSPDAValues.strMessage3Bench & "," & clsReadSPDAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage4Test, clsReadSPDAValues.strMessage4Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage4MM = ib & ",Message4," & clsReadSPDAValues.strMessage4Bench & "," & clsReadSPDAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strMessage5Test, clsReadSPDAValues.strMessage5Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage5MM = ib & ",Message5," & clsReadSPDAValues.strMessage5Bench & "," & clsReadSPDAValues.strMessage5Test & "&"
            End If
            If (String.Compare(clsReadSPDAValues.strMessage6Test, clsReadSPDAValues.strMessage6Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strMessage6MM = ib & ",Message6," & clsReadSPDAValues.strMessage6Bench & "," & clsReadSPDAValues.strMessage6Test & "&"
            End If

            'Set messages back to nothing
            clsReadSPDAValues.strMessage1Bench = ""
            clsReadSPDAValues.strMessage1Test = ""
            clsReadSPDAValues.strMessage2Bench = ""
            clsReadSPDAValues.strMessage2Test = ""
            clsReadSPDAValues.strMessage3Bench = ""
            clsReadSPDAValues.strMessage3Test = ""
            clsReadSPDAValues.strMessage4Bench = ""
            clsReadSPDAValues.strMessage4Test = ""
            clsReadSPDAValues.strMessage5Bench = ""
            clsReadSPDAValues.strMessage5Test = ""
            clsReadSPDAValues.strMessage6Bench = ""
            clsReadSPDAValues.strMessage6Test = ""

            If (String.Compare(clsReadSPDAValues.strSPDACompanyNameTest, clsReadSPDAValues.strSPDACompanyNameBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDACompanyNameMM = ib & ",Company Name," & clsReadSPDAValues.strSPDACompanyNameBench & "," & clsReadSPDAValues.strSPDACompanyNameTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAClient1Test, clsReadSPDAValues.strSPDAClient1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAClient1MM = ib & ",Client1 Name, " & clsReadSPDAValues.strSPDAClient1Bench & "," & clsReadSPDAValues.strSPDAClient1Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAAge1Test, clsReadSPDAValues.strSPDAAge1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAAge1MM = ib & ",Client1 Age," & clsReadSPDAValues.strSPDAAge1Bench & "," & clsReadSPDAValues.strSPDAAge1Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDASex1Test, clsReadSPDAValues.strSPDASex1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDASex1MM = ib & ",Client1 Sex," & clsReadSPDAValues.strSPDASex1Bench & "," & clsReadSPDAValues.strSPDASex1Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDADOB1Test, clsReadSPDAValues.strSPDADOB1Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDADOB1MM = ib & ",Client1 DOB," & clsReadSPDAValues.strSPDADOB1Bench & "," & clsReadSPDAValues.strSPDADOB1Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAClient2Test, clsReadSPDAValues.strSPDAClient2Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAClient2MM = ib & ",Client2 Name, " & clsReadSPDAValues.strSPDAClient2Bench & "," & clsReadSPDAValues.strSPDAClient2Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAAge2Test, clsReadSPDAValues.strSPDAAge2Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAAge2MM = ib & ",Client2 Age," & clsReadSPDAValues.strSPDAAge2Bench & "," & clsReadSPDAValues.strSPDAAge2Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDASex2Test, clsReadSPDAValues.strSPDASex2Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDASex2MM = ib & ",Client2 Sex," & clsReadSPDAValues.strSPDASex2Bench & "," & clsReadSPDAValues.strSPDASex2Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDADOB2Test, clsReadSPDAValues.strSPDADOB2Bench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDADOB2MM = ib & ",Client2 DOB," & clsReadSPDAValues.strSPDADOB2Bench & "," & clsReadSPDAValues.strSPDADOB2Test & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDACompanyNameTest, clsReadSPDAValues.strSPDACompanyNameBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDACompanyNameMM = ib & ",Comp Name," & clsReadSPDAValues.strSPDACompanyNameBench & "," & clsReadSPDAValues.strSPDACompanyNameTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAProdNameTest, clsReadSPDAValues.strSPDAProdNameBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAProdNameMM = ib & ",Prod Name," & clsReadSPDAValues.strSPDAProdNameBench & "," & clsReadSPDAValues.strSPDAProdNameTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAProdNameTest, clsReadSPDAValues.strSPDAProdNameBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAProdNameMM = ib & ",Prod Name Short," & clsReadSPDAValues.strSPDAProdNameBench & "," & clsReadSPDAValues.strSPDAProdNameTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDATaxStatusTest, clsReadSPDAValues.strSPDATaxStatusBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDATaxStatusMM = ib & ",Tax Status," & clsReadSPDAValues.strSPDATaxStatusBench & "," & clsReadSPDAValues.strSPDATaxStatusTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAPremiumBonusTest, clsReadSPDAValues.strSPDAPremiumBonusBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAPremiumBonusMM = ib & ",Premium Bonus," & clsReadSPDAValues.strSPDAPremiumBonusBench & "," & clsReadSPDAValues.strSPDAPremiumBonusTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAPremiumTaxRateTest, clsReadSPDAValues.strSPDAPremiumTaxRateBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAPremiumTaxRateMM = ib & ",Premium Tax Rate," & clsReadSPDAValues.strSPDAPremiumTaxRateBench & "," & clsReadSPDAValues.strSPDAPremiumTaxRateTest & "&"
            End If


            If (String.Compare(clsReadSPDAValues.strSPDAStateTest, clsReadSPDAValues.strSPDAStateBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAStateMM = ib & ",State," & clsReadSPDAValues.strSPDAStateBench & "," & clsReadSPDAValues.strSPDAStateTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDABailoutRateTest, clsReadSPDAValues.strSPDABailoutRateBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDABailoutRateMM = ib & ",Bailout Rate," & clsReadSPDAValues.strSPDABailoutRateBench & "," & clsReadSPDAValues.strSPDABailoutRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDATaxStatusTest), CStr(clsReadSPDAValues.strSPDATaxStatusBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDATaxStatusMM = ib & ",Type of Funds," & clsReadSPDAValues.strSPDATaxStatusBench & "," & clsReadSPDAValues.strSPDATaxStatusTest & "&"
            Else

            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAPolicyFormTest), CStr(clsReadSPDAValues.strSPDAPolicyFormBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAPolicyFormMM = ib & ",Policy Form," & clsReadSPDAValues.strSPDAPolicyFormBench & "," & clsReadSPDAValues.strSPDAPolicyFormTest & "&"
            Else

            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDASurrChargeYrsTest), CStr(clsReadSPDAValues.strSPDASurrChargeYrsBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDASurrChargeYrsMM = ib & ",Surr Charge Years," & clsReadSPDAValues.strSPDASurrChargeYrsBench & "," & clsReadSPDAValues.strSPDASurrChargeYrsTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAChannelCodeTest), CStr(clsReadSPDAValues.strSPDAChannelCodeBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAChannelCodeMM = ib & ",Channel Code," & clsReadSPDAValues.strSPDAChannelCodeBench & "," & clsReadSPDAValues.strSPDAChannelCodeTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAMaxIssueAgeTest), CStr(clsReadSPDAValues.strSPDAMaxIssueAgeBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAMaxIssueAgeMM = ib & ",Max Issue Age," & clsReadSPDAValues.strSPDAMaxIssueAgeBench & "," & clsReadSPDAValues.strSPDAMaxIssueAgeTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDADeclaredRateTest), CStr(clsReadSPDAValues.strSPDADeclaredRateBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDADeclaredRateMM = ib & ",Declared Rate," & clsReadSPDAValues.strSPDADeclaredRateBench & "," & clsReadSPDAValues.strSPDADeclaredRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDARederminationYearTest), CStr(clsReadSPDAValues.strSPDARederminationYearBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDARederminationYearMM = ib & ",Redermination Year," & clsReadSPDAValues.strSPDARederminationYearBench & "," & clsReadSPDAValues.strSPDARederminationYearTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAGuarPeriodTest), CStr(clsReadSPDAValues.strSPDAGuarPeriodBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAGuarPeriodMM = ib & ",Guar Period," & clsReadSPDAValues.strSPDAGuarPeriodBench & "," & clsReadSPDAValues.strSPDAGuarPeriodTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDATaxBracketTest), CStr(clsReadSPDAValues.strSPDATaxBracketBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDATaxBracketMM = ib & ",Tax Bracket," & clsReadSPDAValues.strSPDATaxBracketBench & "," & clsReadSPDAValues.strSPDATaxBracketTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAProjectedRateTest), CStr(clsReadSPDAValues.strSPDAProjectedRateBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAProjectedRateMM = ib & ",Projected Rate," & clsReadSPDAValues.strSPDAProjectedRateBench & "," & clsReadSPDAValues.strSPDAProjectedRateTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAAnnuitizationAgeTest), CStr(clsReadSPDAValues.strSPDAAnnuitizationAgeBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAAnnuitizationAgeMM = ib & ",Annuitization Age," & clsReadSPDAValues.strSPDAAnnuitizationAgeBench & "," & clsReadSPDAValues.strSPDAAnnuitizationAgeTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAWithdrawalTypeTest), CStr(clsReadSPDAValues.strSPDAWithdrawalTypeBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAWithdrawalTypeMM = ib & ",Withdrawal Type," & clsReadSPDAValues.strSPDAWithdrawalTypeBench & "," & clsReadSPDAValues.strSPDAWithdrawalTypeTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAWithdrawalFrequencyTest), CStr(clsReadSPDAValues.strSPDAWithdrawalFrequencyBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAWithdrawalFrequencyMM = ib & ",Withdrawal Frequency," & clsReadSPDAValues.strSPDAWithdrawalFrequencyBench & "," & clsReadSPDAValues.strSPDAWithdrawalFrequencyTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAAutoInterestStartYearTest), CStr(clsReadSPDAValues.strSPDAAutoInterestStartYearBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAAutoInterestStartYearMM = ib & ",Auto Interest Start Year," & clsReadSPDAValues.strSPDAAutoInterestStartYearBench & "," & clsReadSPDAValues.strSPDAAutoInterestStartYearTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAAutoInterestStopYearTest), CStr(clsReadSPDAValues.strSPDAAutoInterestStopYearBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAAutoInterestStopYearMM = ib & ",Auto Interest Stop Year," & clsReadSPDAValues.strSPDAAutoInterestStopYearBench & "," & clsReadSPDAValues.strSPDAAutoInterestStopYearTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAAnnualPremiumTest), CStr(clsReadSPDAValues.strSPDAAnnualPremiumBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAAnnualPremiumMM = ib & ",Annual Premium," & clsReadSPDAValues.strSPDAAnnualPremiumBench & "," & clsReadSPDAValues.strSPDAAnnualPremiumTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDACashBenefitTest), CStr(clsReadSPDAValues.strSPDACashBenefitBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDACashBenefitMM = ib & ",Cash Benefit," & clsReadSPDAValues.strSPDACashBenefitBench & "," & clsReadSPDAValues.strSPDACashBenefitTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDADeathBenefitTest), CStr(clsReadSPDAValues.strSPDADeathBenefitBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDADeathBenefitMM = ib & ",Death Benefit," & clsReadSPDAValues.strSPDADeathBenefitBench & "," & clsReadSPDAValues.strSPDADeathBenefitTest & "&"
            End If


            If (String.Compare(CStr(clsReadSPDAValues.strSPDASurrenderChargesTest), CStr(clsReadSPDAValues.strSPDASurrenderChargesBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDASurrenderChargesTest
                Dim strBench As String = clsReadSPDAValues.strSPDASurrenderChargesBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDASurrenderChargesMM = clsReadSPDAValues.strSPDASurrenderChargesMM & ib & ",Surrender Charges," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAHypoInterestRatesTest), CStr(clsReadSPDAValues.strSPDAHypoInterestRatesBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAHypoInterestRatesTest
                Dim strBench As String = clsReadSPDAValues.strSPDAHypoInterestRatesBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAHypoInterestRatesMM = clsReadSPDAValues.strSPDAHypoInterestRatesMM & ib & ",Hypo Interest Rates," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAGuarInterestRatesTest), CStr(clsReadSPDAValues.strSPDAGuarInterestRatesBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAGuarInterestRatesTest
                Dim strBench As String = clsReadSPDAValues.strSPDAGuarInterestRatesBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAGuarInterestRatesMM = clsReadSPDAValues.strSPDAGuarInterestRatesMM & ib & ",Guar Interest Rates," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAHypoPartialWDTest), CStr(clsReadSPDAValues.strSPDAHypoPartialWDBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAHypoPartialWDTest
                Dim strBench As String = clsReadSPDAValues.strSPDAHypoPartialWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAHypoPartialWDMM = clsReadSPDAValues.strSPDAHypoPartialWDMM & ib & ",Hypo Partial WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAGuarPartialWDTest), CStr(clsReadSPDAValues.strSPDAGuarPartialWDBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAGuarPartialWDTest
                Dim strBench As String = clsReadSPDAValues.strSPDAGuarPartialWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAGuarPartialWDMM = clsReadSPDAValues.strSPDAGuarPartialWDMM & ib & ",Guar Partial WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAWDPercentTest), CStr(clsReadSPDAValues.strSPDAWDPercentBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAWDPercentTest
                Dim strBench As String = clsReadSPDAValues.strSPDAWDPercentBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAWDPercentMM = clsReadSPDAValues.strSPDAWDPercentMM & ib & ",WD Percent," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAAnnualWDAmountTest), CStr(clsReadSPDAValues.strSPDAAnnualWDAmountBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAAnnualWDAmountTest
                Dim strBench As String = clsReadSPDAValues.strSPDAAnnualWDAmountBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAAnnualWDAmountMM = clsReadSPDAValues.strSPDAAnnualWDAmountMM & ib & ",Annual WD Amount," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(CStr(clsReadSPDAValues.strSPDAHypoCVTest), CStr(clsReadSPDAValues.strSPDAHypoCVBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAHypoCVTest
                Dim strBench As String = clsReadSPDAValues.strSPDAHypoCVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAHypoCVMM = clsReadSPDAValues.strSPDAHypoCVMM & ib & ",Hypo CV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAGuarCVTest), CStr(clsReadSPDAValues.strSPDAGuarCVBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAGuarCVTest
                Dim strBench As String = clsReadSPDAValues.strSPDAGuarCVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAGuarCVMM = clsReadSPDAValues.strSPDAGuarCVMM & ib & ",Guar CV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAHypoSVTest), CStr(clsReadSPDAValues.strSPDAHypoSVBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAHypoSVTest
                Dim strBench As String = clsReadSPDAValues.strSPDAHypoSVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAHypoSVMM = clsReadSPDAValues.strSPDAHypoSVMM & ib & ",Hypo SV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDAGuarSVTest), CStr(clsReadSPDAValues.strSPDAGuarSVBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDAGuarSVTest
                Dim strBench As String = clsReadSPDAValues.strSPDAGuarSVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDAGuarSVMM = clsReadSPDAValues.strSPDAGuarSVMM & ib & ",Guar SV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            'Added for MVA



            If (String.Compare(clsReadSPDAValues.strSPDAMGSVTest, clsReadSPDAValues.strSPDAMGSVBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAMGSVMM = ib & ",MVA.SurrValue.MinGuar," & clsReadSPDAValues.strSPDAMGSVBench & "," & clsReadSPDAValues.strSPDAMGSVTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAMGIRTest, clsReadSPDAValues.strSPDAMGIRBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAMGIRMM = ib & ",MVA.IntRate.MinGuar," & clsReadSPDAValues.strSPDAMGIRBench & "," & clsReadSPDAValues.strSPDAMGIRTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDANonForfRateTest, clsReadSPDAValues.strSPDANonForfRateBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDANonForfRateMM = ib & ",MVA.IntRate.NonForf," & clsReadSPDAValues.strSPDANonForfRateBench & "," & clsReadSPDAValues.strSPDANonForfRateTest & "&"
            End If

            If (String.Compare(clsReadSPDAValues.strSPDAJumboRatesTest, clsReadSPDAValues.strSPDAJumboRatesBench, True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True
                clsReadSPDAValues.strSPDAJumboRatesMM = ib & ",MVA.JumboRates?," & clsReadSPDAValues.strSPDAJumboRatesBench & "," & clsReadSPDAValues.strSPDAJumboRatesTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPDAValues.strSPDARenewalSurrChargesTest), CStr(clsReadSPDAValues.strSPDARenewalSurrChargesBench), True)) <> 0 Then
                clsReadSPDAValues.bMisMatch = True

                Dim strTest As String = clsReadSPDAValues.strSPDARenewalSurrChargesTest
                Dim strBench As String = clsReadSPDAValues.strSPDARenewalSurrChargesBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPDAValues.strSPDARenewalSurrChargesMM = clsReadSPDAValues.strSPDARenewalSurrChargesMM & ib & ",MVA.Renewal.SurrenderCharges," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

        End If
        'if cases don't match, create the mismatch reports
        If clsReadSPDAValues.bMisMatch = True Then
            clsReadSPDAValues.bMismatchAtLeastOnce = True
            SaveNewBench(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
            CreateSPDAMismatchReport(ib)
            strClientMisMatchList = strClientMisMatchList & "," & ib
            strsplitMMList = Split(strClientMisMatchList, ",")
            'WriteMismatchList(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If
        'copy the Test.pdf to the test folder under the client folder, whether there are mismatches or not
        If kclbClientList.GetItemChecked(ib - 1) Then
            CopyPDF(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If
        'write the stats to the file
        WriteMatchStatus(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp, clsReadSPDAValues.bMisMatch, bNewBench:=False)

    End Sub
    Private Sub CompareValuesFIA(ByVal ic As Integer, ByVal ib As Integer, ByVal strcomp() As String)


        Dim ix As Integer = 0
        Dim imax As Integer = 0
        Dim iAdd As Integer = 0
        Dim iShort As Integer = 0

        'set mismatch flag to false
        clsReadFIAValues.bMisMatch = False

        'go through each value in relay.out to compare between bench and test, if mismatch, create strings for datagrid

        'below 20 are for warning and error messages
        If clsReadFIAValues.bErrorBench = True And clsReadFIAValues.bErrorTest = True Then
            If (String.Compare(clsReadFIAValues.strMessage1Test, clsReadFIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage1MM = ib & ",Message 1," & clsReadFIAValues.strMessage1Bench & "," & clsReadFIAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage2Test, clsReadFIAValues.strMessage2Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage2MM = ib & ",Message 2," & clsReadFIAValues.strMessage2Bench & "," & clsReadFIAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage3Test, clsReadFIAValues.strMessage3Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage3MM = ib & ",Message 3," & clsReadFIAValues.strMessage3Bench & "," & clsReadFIAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage4Test, clsReadFIAValues.strMessage4Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage4MM = ib & ",Message 4," & clsReadFIAValues.strMessage4Bench & "," & clsReadFIAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage5Test, clsReadFIAValues.strMessage5Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage5MM = ib & ",Message 5," & clsReadFIAValues.strMessage5Bench & "," & clsReadFIAValues.strMessage5Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage6Test, clsReadFIAValues.strMessage6Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage6MM = ib & ",Message 6," & clsReadFIAValues.strMessage6Bench & "," & clsReadFIAValues.strMessage6Test & "&"
            End If

            'Set messages back to nothing
            clsReadFIAValues.strMessage1Bench = ""
            clsReadFIAValues.strMessage1Test = ""
            clsReadFIAValues.strMessage2Bench = ""
            clsReadFIAValues.strMessage2Test = ""
            clsReadFIAValues.strMessage3Bench = ""
            clsReadFIAValues.strMessage3Test = ""
            clsReadFIAValues.strMessage4Bench = ""
            clsReadFIAValues.strMessage4Test = ""
            clsReadFIAValues.strMessage5Bench = ""
            clsReadFIAValues.strMessage5Test = ""
            clsReadFIAValues.strMessage6Bench = ""
            clsReadFIAValues.strMessage6Test = ""

            'see if one runs and one doesnt
        ElseIf clsReadFIAValues.bErrorBench = True And clsReadFIAValues.bErrorTest = False Then
            clsReadFIAValues.bMisMatch = True
            clsReadFIAValues.strRunNoRunMM = ib & ",Test runs/Bench doesn't run,,&"

            If gbFIASaveAge(ib) Then
                If gstrDOB1New <> "" Then
                    clsReadFIAValues.bMisMatch = True
                    gstrFIASaveAgeDOB1 = ib & ",DOB1," & gstrDOB1Original & "," & gstrDOB1New & "&"
                End If
            End If

            If gbFIASaveAge(ib) Then
                If gstrDOB2New <> "" Then
                    clsReadFIAValues.bMisMatch = True
                    gstrFIASaveAgeDOB2 = ib & ",DOB2," & gstrDOB2Original & "," & gstrDOB2New & "&"
                End If
            End If

            If (String.Compare(clsReadFIAValues.strMessage1Test, clsReadFIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage1MM = ib & ",Message1," & clsReadFIAValues.strMessage1Bench & "," & clsReadFIAValues.strMessage1Test & "&"
            End If

            clsReadFIAValues.strMessage1Test = ""
            clsReadFIAValues.strMessage1Bench = ""

        ElseIf clsReadFIAValues.bErrorBench = False And clsReadFIAValues.bErrorTest = True Then
            clsReadFIAValues.bMisMatch = True
            clsReadFIAValues.strRunNoRunMM = ib & ",Bench runs/Test doesn't run,,&"

            If gbFIASaveAge(ib) Then
                If gstrDOB1New <> "" Then
                    clsReadFIAValues.bMisMatch = True
                    gstrFIASaveAgeDOB1 = ib & ",DOB1," & gstrDOB1Original & "," & gstrDOB1New & "&"
                End If
            End If

            If gbFIASaveAge(ib) Then
                If gstrDOB2New <> "" Then
                    clsReadFIAValues.bMisMatch = True
                    gstrFIASaveAgeDOB2 = ib & ",DOB2," & gstrDOB2Original & "," & gstrDOB2New & "&"
                End If
            End If

            If (String.Compare(clsReadFIAValues.strMessage1Test, clsReadFIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage1MM = ib & ",Message1," & clsReadFIAValues.strMessage1Bench & "," & clsReadFIAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage1Test, clsReadFIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage1MM = ib & ",Message1," & clsReadFIAValues.strMessage1Bench & "," & clsReadFIAValues.strMessage1Test & "&"
            End If

            clsReadFIAValues.strMessage1Test = ""
            clsReadFIAValues.strMessage1Bench = ""
        Else
            'if both cases run, then compare...

            If (String.Compare(clsReadFIAValues.strMessage1Test, clsReadFIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage1MM = ib & ",Message1," & clsReadFIAValues.strMessage1Bench & "," & clsReadFIAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage2Test, clsReadFIAValues.strMessage2Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage2MM = ib & ",Message2," & clsReadFIAValues.strMessage2Bench & "," & clsReadFIAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage3Test, clsReadFIAValues.strMessage3Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage3MM = ib & ",Message3," & clsReadFIAValues.strMessage3Bench & "," & clsReadFIAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage4Test, clsReadFIAValues.strMessage4Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage4MM = ib & ",Message4," & clsReadFIAValues.strMessage4Bench & "," & clsReadFIAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strMessage5Test, clsReadFIAValues.strMessage5Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage5MM = ib & ",Message5," & clsReadFIAValues.strMessage5Bench & "," & clsReadFIAValues.strMessage5Test & "&"
            End If
            If (String.Compare(clsReadFIAValues.strMessage6Test, clsReadFIAValues.strMessage6Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage6MM = ib & ",Message6," & clsReadFIAValues.strMessage6Bench & "," & clsReadFIAValues.strMessage6Test & "&"
            End If

            'Set messages back to nothing
            clsReadFIAValues.strMessage1Bench = ""
            clsReadFIAValues.strMessage1Test = ""
            clsReadFIAValues.strMessage2Bench = ""
            clsReadFIAValues.strMessage2Test = ""
            clsReadFIAValues.strMessage3Bench = ""
            clsReadFIAValues.strMessage3Test = ""
            clsReadFIAValues.strMessage4Bench = ""
            clsReadFIAValues.strMessage4Test = ""
            clsReadFIAValues.strMessage5Bench = ""
            clsReadFIAValues.strMessage5Test = ""
            clsReadFIAValues.strMessage6Bench = ""
            clsReadFIAValues.strMessage6Test = ""

            If gbFIASaveAge(ib) Then
                If gstrDOB1New <> "" Then
                    clsReadFIAValues.bMisMatch = True
                    gstrFIASaveAgeDOB1 = ib & ",DOB1," & gstrDOB1Original & "," & gstrDOB1New & "&"
                End If
            End If

            If gbFIASaveAge(ib) Then
                If gstrDOB2New <> "" Then
                    clsReadFIAValues.bMisMatch = True
                    gstrFIASaveAgeDOB2 = ib & ",DOB2," & gstrDOB2Original & "," & gstrDOB2New & "&"
                End If
            End If

            If (String.Compare(clsReadFIAValues.strMessage1Test, clsReadFIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strMessage1MM = ib & ",Message1," & clsReadFIAValues.strMessage1Bench & "," & clsReadFIAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIACompanyNameTest, clsReadFIAValues.strFIACompanyNameBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIACompanyNameMM = ib & ",Company Name," & clsReadFIAValues.strFIACompanyNameBench & "," & clsReadFIAValues.strFIACompanyNameTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAClient1Test, clsReadFIAValues.strFIAClient1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAClient1MM = ib & ",Client1 Name, " & clsReadFIAValues.strFIAClient1Bench & "," & clsReadFIAValues.strFIAClient1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAAge1Test, clsReadFIAValues.strFIAAge1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAge1MM = ib & ",Client1 Age," & clsReadFIAValues.strFIAAge1Bench & "," & clsReadFIAValues.strFIAAge1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIASex1Test, clsReadFIAValues.strFIASex1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASex1MM = ib & ",Client1 Sex," & clsReadFIAValues.strFIASex1Bench & "," & clsReadFIAValues.strFIASex1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIADOB1Test, clsReadFIAValues.strFIADOB1Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIADOB1MM = ib & ",Client1 DOB," & clsReadFIAValues.strFIADOB1Bench & "," & clsReadFIAValues.strFIADOB1Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAClient2Test, clsReadFIAValues.strFIAClient2Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAClient2MM = ib & ",Client2 Name, " & clsReadFIAValues.strFIAClient2Bench & "," & clsReadFIAValues.strFIAClient2Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAAge2Test, clsReadFIAValues.strFIAAge2Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAge2MM = ib & ",Client2 Age," & clsReadFIAValues.strFIAAge2Bench & "," & clsReadFIAValues.strFIAAge2Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIASex2Test, clsReadFIAValues.strFIASex2Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASex2MM = ib & ",Client2 Sex," & clsReadFIAValues.strFIASex2Bench & "," & clsReadFIAValues.strFIASex2Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIADOB2Test, clsReadFIAValues.strFIADOB2Bench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIADOB2MM = ib & ",Client2 DOB," & clsReadFIAValues.strFIADOB2Bench & "," & clsReadFIAValues.strFIADOB2Test & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAProdNameTest, clsReadFIAValues.strFIAProdNameBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAProdNameMM = ib & ",Prod Name," & clsReadFIAValues.strFIAProdNameBench & "," & clsReadFIAValues.strFIAProdNameTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIATaxStatusTest, clsReadFIAValues.strFIATaxStatusBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIATaxStatusMM = ib & ",Tax Status," & clsReadFIAValues.strFIATaxStatusBench & "," & clsReadFIAValues.strFIATaxStatusTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAPremiumTaxRateTest, clsReadFIAValues.strFIAPremiumTaxRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAPremiumTaxRateMM = ib & ",Premium Tax Rate," & clsReadFIAValues.strFIAPremiumTaxRateBench & "," & clsReadFIAValues.strFIAPremiumTaxRateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAStateTest, clsReadFIAValues.strFIAStateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAStateMM = ib & ",State," & clsReadFIAValues.strFIAStateBench & "," & clsReadFIAValues.strFIAStateTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualPremiumTest), CStr(clsReadFIAValues.strFIAAnnualPremiumBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAnnualPremiumMM = ib & ",Annual Premium," & clsReadFIAValues.strFIAAnnualPremiumBench & "," & clsReadFIAValues.strFIAAnnualPremiumTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAPremiumEnhancementTest), CStr(clsReadFIAValues.strFIAPremiumEnhancementBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAPremiumEnhancementMM = ib & ",Premium Enhancement," & clsReadFIAValues.strFIAPremiumEnhancementBench & "," & clsReadFIAValues.strFIAPremiumEnhancementTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAInitialBeneBaseTest), CStr(clsReadFIAValues.strFIAInitialBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAInitialBeneBaseMM = ib & ",Initial Benefit Base," & clsReadFIAValues.strFIAInitialBeneBaseBench & "," & clsReadFIAValues.strFIAInitialBeneBaseTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIABailoutAnnualCapTest, clsReadFIAValues.strFIABailoutAnnualCapBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIABailoutAnnualCapMM = ib & ",Bailout Annual Cap," & clsReadFIAValues.strFIABailoutAnnualCapBench & "," & clsReadFIAValues.strFIABailoutAnnualCapTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIARiderRollUpRateTest, clsReadFIAValues.strFIARiderRollUpRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIARiderRollUpRateMM = ib & ",Rider RollUp Rate," & clsReadFIAValues.strFIARiderRollUpRateBench & "," & clsReadFIAValues.strFIARiderRollUpRateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIARiderChargeTest, clsReadFIAValues.strFIARiderChargeBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIARiderChargeMM = ib & ",Rider Charge," & clsReadFIAValues.strFIARiderChargeBench & "," & clsReadFIAValues.strFIARiderChargeTest & "&"
            End If

            If clsReadFIAValues.strFIAProdNameTest = "FIA7L" Then
                If (String.Compare(clsReadFIAValues.strFIAGMCVTest, clsReadFIAValues.strFIAGMCVBench, True)) <> 0 Then
                    clsReadFIAValues.bMisMatch = True
                    clsReadFIAValues.strFIAGMCVMM = ib & ",GMCV," & clsReadFIAValues.strFIAGMCVBench & "," & clsReadFIAValues.strFIAGMCVTest & "&"
                End If
            End If
            If (String.Compare(clsReadFIAValues.strFIAAgeAtFirstWDTest, clsReadFIAValues.strFIAAgeAtFirstWDBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAgeAtFirstWDMM = ib & ",Age at 1st WD," & clsReadFIAValues.strFIAAgeAtFirstWDBench & "," & clsReadFIAValues.strFIAAgeAtFirstWDTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAAnnWDLimitGuarTest, clsReadFIAValues.strFIAAnnWDLimitGuarBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAnnWDLimitGuarMM = ib & ",Guar Ann WD Limit," & clsReadFIAValues.strFIAAnnWDLimitGuarBench & "," & clsReadFIAValues.strFIAAnnWDLimitGuarTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAAnnWDLimitProjTest, clsReadFIAValues.strFIAAnnWDLimitProjBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAnnWDLimitProjMM = ib & ",Proj Ann WD Limit," & clsReadFIAValues.strFIAAnnWDLimitProjBench & "," & clsReadFIAValues.strFIAAnnWDLimitProjTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAOneYearFixedAllocTest, clsReadFIAValues.strFIAOneYearFixedAllocBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAOneYearFixedAllocMM = ib & ",1 Yr Fixed Alloc," & clsReadFIAValues.strFIAOneYearFixedAllocBench & "," & clsReadFIAValues.strFIAOneYearFixedAllocTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIASevenYearFixedAllocTest, clsReadFIAValues.strFIASevenYearFixedAllocBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASevenYearFixedAllocMM = ib & ",7 Yr Fixed Alloc," & clsReadFIAValues.strFIASevenYearFixedAllocBench & "," & clsReadFIAValues.strFIASevenYearFixedAllocTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIATenYearFixedAllocTest, clsReadFIAValues.strFIATenYearFixedAllocBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIATenYearFixedAllocMM = ib & ",10 Yr Fixed Alloc," & clsReadFIAValues.strFIATenYearFixedAllocBench & "," & clsReadFIAValues.strFIATenYearFixedAllocTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAAnnCapAllocTest, clsReadFIAValues.strFIAAnnCapAllocBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAnnCapAllocMM = ib & ",Ann Cap Alloc," & clsReadFIAValues.strFIAAnnCapAllocBench & "," & clsReadFIAValues.strFIAAnnCapAllocTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAMonCapAllocTest, clsReadFIAValues.strFIAMonCapAllocBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAMonCapAllocMM = ib & ",Mon Cap Alloc," & clsReadFIAValues.strFIAMonCapAllocBench & "," & clsReadFIAValues.strFIAMonCapAllocTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAPerfTrigAllocTest, clsReadFIAValues.strFIAPerfTrigAllocBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAPerfTrigAllocMM = ib & ",Perf Trig Alloc," & clsReadFIAValues.strFIAPerfTrigAllocBench & "," & clsReadFIAValues.strFIAPerfTrigAllocTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAOneYearFixedInitialRateTest, clsReadFIAValues.strFIAOneYearFixedInitialRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAOneYearFixedInitialRateMM = ib & ",1 Yr Fixed Init Rate," & clsReadFIAValues.strFIAOneYearFixedInitialRateBench & "," & clsReadFIAValues.strFIAOneYearFixedInitialRateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIASevenYearFixedInitialRateTest, clsReadFIAValues.strFIASevenYearFixedInitialRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASevenYearFixedInitialRateMM = ib & ",7 Yr Fixed Init Rate," & clsReadFIAValues.strFIASevenYearFixedInitialRateBench & "," & clsReadFIAValues.strFIASevenYearFixedInitialRateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIATenYearFixedInitialRateTest, clsReadFIAValues.strFIATenYearFixedInitialRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIATenYearFixedInitialRateMM = ib & ",10 Yr Fixed Init Rate," & clsReadFIAValues.strFIATenYearFixedInitialRateBench & "," & clsReadFIAValues.strFIATenYearFixedInitialRateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAAnnualCapCapTest, clsReadFIAValues.strFIAAnnualCapCapBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAnnualCapCapMM = ib & ",Annual Cap Cap," & clsReadFIAValues.strFIAAnnualCapCapBench & "," & clsReadFIAValues.strFIAAnnualCapCapTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAMonthlyCapCapTest, clsReadFIAValues.strFIAMonthlyCapCapBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAMonthlyCapCapMM = ib & ",Monthly Cap Cap," & clsReadFIAValues.strFIAMonthlyCapCapBench & "," & clsReadFIAValues.strFIAMonthlyCapCapTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAPerfTrigSpecifiedRateTest, clsReadFIAValues.strFIAPerfTrigSpecifiedRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAperfTrigSpecifiedRateMM = ib & ",Perf Trig Spec Rate," & clsReadFIAValues.strFIAPerfTrigSpecifiedRateBench & "," & clsReadFIAValues.strFIAPerfTrigSpecifiedRateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAYearsToPrintTest, clsReadFIAValues.strFIAYearsToPrintBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAYearsToPrintMM = ib & ",Yrs to Print," & clsReadFIAValues.strFIAYearsToPrintBench & "," & clsReadFIAValues.strFIAYearsToPrintTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIASpecPeriodStartDateTest, clsReadFIAValues.strFIASpecPeriodStartDateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASpecPeriodStartDateMM = ib & ",Spec Period Start Date," & clsReadFIAValues.strFIASpecPeriodStartDateBench & "," & clsReadFIAValues.strFIASpecPeriodStartDateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIASpecPeriodEndDateTest, clsReadFIAValues.strFIASpecPeriodEndDateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASpecPeriodEndDateMM = ib & ",Spec Period End Date," & clsReadFIAValues.strFIASpecPeriodEndDateBench & "," & clsReadFIAValues.strFIASpecPeriodEndDateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAFavPeriodStartDateTest, clsReadFIAValues.strFIAFavPeriodStartDateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAFavPeriodStartDateMM = ib & ",Fav Period Start Date," & clsReadFIAValues.strFIAFavPeriodStartDateBench & "," & clsReadFIAValues.strFIAFavPeriodStartDateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAFavPeriodEndDateTest, clsReadFIAValues.strFIAFavPeriodEndDateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAFavPeriodEndDateMM = ib & ",Fav Period End Date," & clsReadFIAValues.strFIAFavPeriodEndDateBench & "," & clsReadFIAValues.strFIAFavPeriodEndDateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAUnFavPeriodStartDateTest, clsReadFIAValues.strFIAUnFavPeriodStartDateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAUnFavPeriodStartDateMM = ib & ",UnFav Period Start Date," & clsReadFIAValues.strFIAUnFavPeriodStartDateBench & "," & clsReadFIAValues.strFIAUnFavPeriodStartDateTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIAUnFavPeriodEndDateTest, clsReadFIAValues.strFIAUnFavPeriodEndDateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAUnFavPeriodEndDateMM = ib & ",UnFav Period End Date," & clsReadFIAValues.strFIAUnFavPeriodEndDateBench & "," & clsReadFIAValues.strFIAUnFavPeriodEndDateTest & "&"
            End If






            If (String.Compare(CStr(clsReadFIAValues.strFIASpecAnnCreditRateTest), CStr(clsReadFIAValues.strFIASpecAnnCreditRateBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecAnnCreditRateTest
                Dim strBench As String = clsReadFIAValues.strFIASpecAnnCreditRateBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecAnnCreditRateMM = clsReadFIAValues.strFIASpecAnnCreditRateMM & ib & ",Spec Ann Credit Rate," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnCreditRateNoWDTest), CStr(clsReadFIAValues.strFIAAnnCreditRateNoWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAAnnCreditRateNoWDTest
                Dim strBench As String = clsReadFIAValues.strFIAAnnCreditRateNoWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAAnnCreditRateNoWDMM = clsReadFIAValues.strFIAAnnCreditRateNoWDMM & ib & ",Ann Credit Rate NoWD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavAnnCreditRateTest), CStr(clsReadFIAValues.strFIAFavAnnCreditRateBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavAnnCreditRateTest
                Dim strBench As String = clsReadFIAValues.strFIAFavAnnCreditRateBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavAnnCreditRateMM = clsReadFIAValues.strFIAFavAnnCreditRateMM & ib & ",Fav Ann Credit Rate," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavAnnCreditRateTest), CStr(clsReadFIAValues.strFIAUnfavAnnCreditRateBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavAnnCreditRateTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavAnnCreditRateBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavAnnCreditRateMM = clsReadFIAValues.strFIAUnfavAnnCreditRateMM & ib & ",Unfav Ann Credit Rate," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecContractValueTest), CStr(clsReadFIAValues.strFIASpecContractValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecContractValueTest
                Dim strBench As String = clsReadFIAValues.strFIASpecContractValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecContractValueMM = clsReadFIAValues.strFIASpecContractValueMM & ib & ",Spec CV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(CStr(clsReadFIAValues.strFIAFavContractValueTest), CStr(clsReadFIAValues.strFIAFavContractValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavContractValueTest
                Dim strBench As String = clsReadFIAValues.strFIAFavContractValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavContractValueMM = clsReadFIAValues.strFIAFavContractValueMM & ib & ",Fav CV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavContractValueTest), CStr(clsReadFIAValues.strFIAUnfavContractValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavContractValueTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavContractValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavContractValueMM = clsReadFIAValues.strFIAUnfavContractValueMM & ib & ",Unfav CV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecSurrenderValueTest), CStr(clsReadFIAValues.strFIASpecSurrenderValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecSurrenderValueTest
                Dim strBench As String = clsReadFIAValues.strFIASpecSurrenderValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecSurrenderValueMM = clsReadFIAValues.strFIASpecSurrenderValueMM & ib & ",Spec SV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavSurrenderValueTest), CStr(clsReadFIAValues.strFIAFavSurrenderValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavSurrenderValueTest
                Dim strBench As String = clsReadFIAValues.strFIAFavSurrenderValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavSurrenderValueMM = clsReadFIAValues.strFIAFavSurrenderValueMM & ib & ",Fav SV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavSurrenderValueTest), CStr(clsReadFIAValues.strFIAUnfavSurrenderValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavSurrenderValueTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavSurrenderValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavSurrenderValueMM = clsReadFIAValues.strFIAUnfavSurrenderValueMM & ib & ",Unfav SV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecMGSVTest), CStr(clsReadFIAValues.strFIASpecMGSVBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecMGSVTest
                Dim strBench As String = clsReadFIAValues.strFIASpecMGSVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecMGSVMM = clsReadFIAValues.strFIASpecMGSVMM & ib & ",Spec MGSV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavMGSVTest), CStr(clsReadFIAValues.strFIAFavMGSVBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavMGSVTest
                Dim strBench As String = clsReadFIAValues.strFIAFavMGSVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavMGSVMM = clsReadFIAValues.strFIAFavMGSVMM & ib & ",Fav MGSV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavMGSVTest), CStr(clsReadFIAValues.strFIAUnfavMGSVBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavMGSVTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavMGSVBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavMGSVMM = clsReadFIAValues.strFIAUnfavMGSVMM & ib & ",Unfav MGSV," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecProjBeneBaseTest), CStr(clsReadFIAValues.strFIASpecProjBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecProjBeneBaseTest
                Dim strBench As String = clsReadFIAValues.strFIASpecProjBeneBaseBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecProjBeneBaseMM = clsReadFIAValues.strFIASpecProjBeneBaseMM & ib & ",Spec Proj Bene Base," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavProjBeneBaseTest), CStr(clsReadFIAValues.strFIAFavProjBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavProjBeneBaseTest
                Dim strBench As String = clsReadFIAValues.strFIAFavProjBeneBaseBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavProjBeneBaseMM = clsReadFIAValues.strFIAFavProjBeneBaseMM & ib & ",Fav Proj Bene Base," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavProjBeneBaseTest), CStr(clsReadFIAValues.strFIAUnfavProjBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavProjBeneBaseTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavProjBeneBaseBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavProjBeneBaseMM = clsReadFIAValues.strFIAUnfavProjBeneBaseMM & ib & ",Unfav Proj Bene Base," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecGuarBeneBaseTest), CStr(clsReadFIAValues.strFIASpecGuarBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecGuarBeneBaseTest
                Dim strBench As String = clsReadFIAValues.strFIASpecGuarBeneBaseBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecGuarBeneBaseMM = clsReadFIAValues.strFIASpecGuarBeneBaseMM & ib & ",Spec Guar Bene Base," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavGuarBeneBaseTest), CStr(clsReadFIAValues.strFIAFavGuarBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavGuarBeneBaseTest
                Dim strBench As String = clsReadFIAValues.strFIAFavGuarBeneBaseBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavGuarBeneBaseMM = clsReadFIAValues.strFIAFavGuarBeneBaseMM & ib & ",Fav Guar Bene Base," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavGuarBeneBaseTest), CStr(clsReadFIAValues.strFIAUnfavGuarBeneBaseBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavGuarBeneBaseTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavGuarBeneBaseBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavGuarBeneBaseMM = clsReadFIAValues.strFIAUnfavGuarBeneBaseMM & ib & ",Unfav Guar Bene Base," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecGuarWDLimitTest), CStr(clsReadFIAValues.strFIASpecGuarWDLimitBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecGuarWDLimitTest
                Dim strBench As String = clsReadFIAValues.strFIASpecGuarWDLimitBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecGuarWDLimitMM = clsReadFIAValues.strFIASpecGuarWDLimitMM & ib & ",Spec Guar WD Limit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavGuarWDLimitTest), CStr(clsReadFIAValues.strFIAFavGuarWDLimitBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavGuarWDLimitTest
                Dim strBench As String = clsReadFIAValues.strFIAFavGuarWDLimitBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavGuarWDLimitMM = clsReadFIAValues.strFIAFavGuarWDLimitMM & ib & ",Fav Guar WD Limit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavGuarWDLimitTest), CStr(clsReadFIAValues.strFIAUnfavGuarWDLimitBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavGuarWDLimitTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavGuarWDLimitBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavGuarWDLimitMM = clsReadFIAValues.strFIAUnfavGuarWDLimitMM & ib & ",Unfav Guar WD Limit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecProjWDLimitTest), CStr(clsReadFIAValues.strFIASpecProjWDLimitBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecProjWDLimitTest
                Dim strBench As String = clsReadFIAValues.strFIASpecProjWDLimitBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecProjWDLimitMM = clsReadFIAValues.strFIASpecProjWDLimitMM & ib & ",Spec Proj WD Limit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavProjWDLimitTest), CStr(clsReadFIAValues.strFIAFavProjWDLimitBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavProjWDLimitTest
                Dim strBench As String = clsReadFIAValues.strFIAFavProjWDLimitBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavProjWDLimitMM = clsReadFIAValues.strFIAFavProjWDLimitMM & ib & ",Fav Proj WD Limit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavProjWDLimitTest), CStr(clsReadFIAValues.strFIAUnfavProjWDLimitBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavProjWDLimitTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavProjWDLimitBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavProjWDLimitMM = clsReadFIAValues.strFIAUnfavProjWDLimitMM & ib & ",Unfav Proj WD Limit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(CStr(clsReadFIAValues.strFIAMonthlyCapIndexCreditTest), CStr(clsReadFIAValues.strFIAMonthlyCapIndexCreditBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAMonthlyCapIndexCreditTest
                Dim strBench As String = clsReadFIAValues.strFIAMonthlyCapIndexCreditBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAMonthlyCapIndexCreditMM = clsReadFIAValues.strFIAMonthlyCapIndexCreditMM & ib & ",Monthly Cap Index Credit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualCapIndexCreditTest), CStr(clsReadFIAValues.strFIAAnnualCapIndexCreditBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAAnnualCapIndexCreditTest
                Dim strBench As String = clsReadFIAValues.strFIAAnnualCapIndexCreditBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAAnnualCapIndexCreditMM = clsReadFIAValues.strFIAAnnualCapIndexCreditMM & ib & ",Annual Cap Index Credit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAPerfTriggerIndexCreditTest), CStr(clsReadFIAValues.strFIAPerfTriggerIndexCreditBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAPerfTriggerIndexCreditTest
                Dim strBench As String = clsReadFIAValues.strFIAPerfTriggerIndexCreditBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAPerfTriggerIndexCreditMM = clsReadFIAValues.strFIAPerfTriggerIndexCreditMM & ib & ",Perf Trigger Index Credit," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASevenYearIntRateTest), CStr(clsReadFIAValues.strFIASevenYearIntRateBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASevenYearIntRateTest
                Dim strBench As String = clsReadFIAValues.strFIASevenYearIntRateBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASevenYearIntRateMM = clsReadFIAValues.strFIASevenYearIntRateMM & ib & ",7 Yr Int Rate," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIATenYearIntRateTest), CStr(clsReadFIAValues.strFIATenYearIntRateBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIATenYearIntRateTest
                Dim strBench As String = clsReadFIAValues.strFIATenYearIntRateBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIATenYearIntRateMM = clsReadFIAValues.strFIATenYearIntRateMM & ib & ",10 Yr Int Rate," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASevenYearAccumValueTest), CStr(clsReadFIAValues.strFIASevenYearAccumValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASevenYearAccumValueTest
                Dim strBench As String = clsReadFIAValues.strFIASevenYearAccumValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASevenYearAccumValueMM = clsReadFIAValues.strFIASevenYearAccumValueMM & ib & ",7 Year Accum Value," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIATenYearAccumValueTest), CStr(clsReadFIAValues.strFIATenYearAccumValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIATenYearAccumValueTest
                Dim strBench As String = clsReadFIAValues.strFIATenYearAccumValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIATenYearAccumValueMM = clsReadFIAValues.strFIATenYearAccumValueMM & ib & ",10 Year Accum Value," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAMonthlyCapAccumValueTest), CStr(clsReadFIAValues.strFIAMonthlyCapAccumValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAMonthlyCapAccumValueTest
                Dim strBench As String = clsReadFIAValues.strFIAMonthlyCapAccumValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAMonthlyCapAccumValueMM = clsReadFIAValues.strFIAMonthlyCapAccumValueMM & ib & ",Montly Cap Accum Value," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualCapAccumValueTest), CStr(clsReadFIAValues.strFIAAnnualCapAccumValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAAnnualCapAccumValueTest
                Dim strBench As String = clsReadFIAValues.strFIAAnnualCapAccumValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAAnnualCapAccumValueMM = clsReadFIAValues.strFIAAnnualCapAccumValueMM & ib & ",Annual Cap Accum Value," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAPerfTriggerAccumValueTest), CStr(clsReadFIAValues.strFIAPerfTriggerAccumValueBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAPerfTriggerAccumValueTest
                Dim strBench As String = clsReadFIAValues.strFIAPerfTriggerAccumValueBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAPerfTriggerAccumValueMM = clsReadFIAValues.strFIAPerfTriggerAccumValueMM & ib & ",Perf Trigger Accum Value," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAContractValueNoWDTest), CStr(clsReadFIAValues.strFIAContractValueNoWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAContractValueNoWDTest
                Dim strBench As String = clsReadFIAValues.strFIAContractValueNoWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAContractValueNoWDMM = clsReadFIAValues.strFIAContractValueNoWDMM & ib & ",CV No WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAGuarBeneBaseNoWDTest), CStr(clsReadFIAValues.strFIAGuarBeneBaseNoWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAGuarBeneBaseNoWDTest
                Dim strBench As String = clsReadFIAValues.strFIAGuarBeneBaseNoWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAGuarBeneBaseNoWDMM = clsReadFIAValues.strFIAGuarBeneBaseNoWDMM & ib & ",Guar BeneBase No WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAProjBeneBaseNoWDTest), CStr(clsReadFIAValues.strFIAProjBeneBaseNoWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAProjBeneBaseNoWDTest
                Dim strBench As String = clsReadFIAValues.strFIAProjBeneBaseNoWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAProjBeneBAseNoWDMM = clsReadFIAValues.strFIAProjBeneBAseNoWDMM & ib & ",Proj BeneBase No WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAGuarWDLimitNoWDTest), CStr(clsReadFIAValues.strFIAGuarWDLimitNoWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAGuarWDLimitNoWDTest
                Dim strBench As String = clsReadFIAValues.strFIAGuarWDLimitNoWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAGuarWDLimitNoWDMM = clsReadFIAValues.strFIAGuarWDLimitNoWDMM & ib & ",Guar WD Limit No WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAProjWDLimitNoWDTest), CStr(clsReadFIAValues.strFIAProjWDLimitNoWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAProjWDLimitNoWDTest
                Dim strBench As String = clsReadFIAValues.strFIAProjWDLimitNoWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAProjWDLimitNoWDMM = clsReadFIAValues.strFIAProjWDLimitNoWDMM & ib & ",Proj WD Limit No WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAGuarWDFactorTest), CStr(clsReadFIAValues.strFIAGuarWDFactorBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAGuarWDFactorTest
                Dim strBench As String = clsReadFIAValues.strFIAGuarWDFactorBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAGuarWDFactorMM = clsReadFIAValues.strFIAGuarWDFactorMM & ib & ",Guar WD Factor," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecSPChangeTest), CStr(clsReadFIAValues.strFIASpecSPChangeBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecSPChangeTest
                Dim strBench As String = clsReadFIAValues.strFIASpecSPChangeBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecSPChangeMM = clsReadFIAValues.strFIASpecSPChangeMM & ib & ",Spec SP Change," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavSPChangeTest), CStr(clsReadFIAValues.strFIAFavSPChangeBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavSPChangeTest
                Dim strBench As String = clsReadFIAValues.strFIAFavSPChangeBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavSPChangeMM = clsReadFIAValues.strFIAFavSPChangeMM & ib & ",Fav SP Change," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavSPChangeTest), CStr(clsReadFIAValues.strFIAUnfavSPChangeBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavSPChangeTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavSPChangeBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnFavSPChangeMM = clsReadFIAValues.strFIAUnFavSPChangeMM & ib & ",UnFav SP Change," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecWDTest), CStr(clsReadFIAValues.strFIASpecWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASpecWDTest
                Dim strBench As String = clsReadFIAValues.strFIASpecWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASpecWDMM = clsReadFIAValues.strFIASpecWDMM & ib & ",Spec WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavWDTest), CStr(clsReadFIAValues.strFIAFavWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFavWDTest
                Dim strBench As String = clsReadFIAValues.strFIAFavWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFavWDMM = clsReadFIAValues.strFIAFavWDMM & ib & ",Fav WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavWDTest), CStr(clsReadFIAValues.strFIAUnfavWDBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavWDTest
                Dim strBench As String = clsReadFIAValues.strFIAUnfavWDBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAUnfavWDMM = clsReadFIAValues.strFIAUnfavWDMM & ib & ",UnFav WD," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFixedJumboTest), CStr(clsReadFIAValues.strFIAFixedJumboBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAFixedJumboTest
                Dim strBench As String = clsReadFIAValues.strFIAFixedJumboBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAFixedJumboMM = clsReadFIAValues.strFIAFixedJumboMM & ib & ",Fixed Jumbo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAMonthlyCapJumboTest), CStr(clsReadFIAValues.strFIAMonthlyCapJumboBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAMonthlyCapJumboTest
                Dim strBench As String = clsReadFIAValues.strFIAMonthlyCapJumboBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAMonthlyCapJumboMM = clsReadFIAValues.strFIAMonthlyCapJumboMM & ib & ",Monthly Cap Jumbo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnCapJumboTest), CStr(clsReadFIAValues.strFIAAnnCapJumboBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAAnnCapJumboTest
                Dim strBench As String = clsReadFIAValues.strFIAAnnCapJumboBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAAnnCapJumboMM = clsReadFIAValues.strFIAAnnCapJumboMM & ib & ",Annual Cap Jumbo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAPerfTriggerJumboTest), CStr(clsReadFIAValues.strFIAPerfTriggerJumboBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAPerfTriggerJumboTest
                Dim strBench As String = clsReadFIAValues.strFIAPerfTriggerJumboBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAPerfTriggerJumboMM = clsReadFIAValues.strFIAPerfTriggerJumboMM & ib & ",Perf Trigger Jumbo," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If



            If (String.Compare(clsReadFIAValues.strFIAWDFrequencyTest, clsReadFIAValues.strFIAWDFrequencyBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAWDFrequencyMM = ib & ",WD Frequency," & clsReadFIAValues.strFIAWDFrequencyBench & "," & clsReadFIAValues.strFIAWDFrequencyTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIATaxStatusTest), CStr(clsReadFIAValues.strFIATaxStatusBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIATaxStatusMM = ib & ",Type of Funds," & clsReadFIAValues.strFIATaxStatusBench & "," & clsReadFIAValues.strFIATaxStatusTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAPolicyFormTest), CStr(clsReadFIAValues.strFIAPolicyFormBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAPolicyFormMM = ib & ",Policy Form," & clsReadFIAValues.strFIAPolicyFormBench & "," & clsReadFIAValues.strFIAPolicyFormTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASurrChargeYrsTest), CStr(clsReadFIAValues.strFIASurrChargeYrsBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIASurrChargeYrsMM = ib & ",Surr Charge Years," & clsReadFIAValues.strFIASurrChargeYrsBench & "," & clsReadFIAValues.strFIASurrChargeYrsTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAChannelCodeTest), CStr(clsReadFIAValues.strFIAChannelCodeBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAChannelCodeMM = ib & ",Channel Code," & clsReadFIAValues.strFIAChannelCodeBench & "," & clsReadFIAValues.strFIAChannelCodeTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualPremiumTest), CStr(clsReadFIAValues.strFIAAnnualPremiumBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAAnnualPremiumMM = ib & ",Annual Premium," & clsReadFIAValues.strFIAAnnualPremiumBench & "," & clsReadFIAValues.strFIAAnnualPremiumTest & "&"
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASurrenderChargesTest), CStr(clsReadFIAValues.strFIASurrenderChargesBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIASurrenderChargesTest
                Dim strBench As String = clsReadFIAValues.strFIASurrenderChargesBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIASurrenderChargesMM = clsReadFIAValues.strFIASurrenderChargesMM & ib & ",Surrender Charges," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAWDPercentTest), CStr(clsReadFIAValues.strFIAWDPercentBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAWDPercentTest
                Dim strBench As String = clsReadFIAValues.strFIAWDPercentBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAWDPercentMM = clsReadFIAValues.strFIAWDPercentMM & ib & ",WD Percent," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualWDAmountTest), CStr(clsReadFIAValues.strFIAAnnualWDAmountBench), True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True

                Dim strTest As String = clsReadFIAValues.strFIAAnnualWDAmountTest
                Dim strBench As String = clsReadFIAValues.strFIAAnnualWDAmountBench

                Dim SplitTest = Split(strTest, ",")
                Dim SplitBench = Split(strBench, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadFIAValues.strFIAAnnualWDAmountMM = clsReadFIAValues.strFIAAnnualWDAmountMM & ib & ",Annual WD Amount," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(clsReadFIAValues.strFIAMGSVTest, clsReadFIAValues.strFIAMGSVBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIAMGSVMM = ib & ",MVA.SurrValue.MinGuar," & clsReadFIAValues.strFIAMGSVBench & "," & clsReadFIAValues.strFIAMGSVTest & "&"
            End If

            If (String.Compare(clsReadFIAValues.strFIANonForfIntRateTest, clsReadFIAValues.strFIANonForfIntRateBench, True)) <> 0 Then
                clsReadFIAValues.bMisMatch = True
                clsReadFIAValues.strFIANonForfIntRateMM = ib & ",MVA.IntRate.NonForf," & clsReadFIAValues.strFIANonForfIntRateBench & "," & clsReadFIAValues.strFIANonForfIntRateTest & "&"
            End If

            End If
            'if cases don't match, create the mismatch reports
            If clsReadFIAValues.bMisMatch = True Then
                clsReadFIAValues.bMismatchAtLeastOnce = True
                SaveNewBench(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
                CreateFIAMismatchReport(ib)
                strClientMisMatchList = strClientMisMatchList & "," & ib
                strsplitMMList = Split(strClientMisMatchList, ",")
                'WriteMismatchList(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
            End If
            'copy the Test.pdf to the test folder under the client folder, whether there are mismatches or not
            'If kclbClientList.GetItemChecked(ib - 1) Then
            '    CopyPDF(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
            'End If
            'write the stats to the file
            WriteMatchStatus(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp, clsReadFIAValues.bMisMatch, bNewBench:=False)


    End Sub
    Private Sub CompareValuesFIATOOL(ByVal ic As Integer, ByVal ib As Integer, ByVal strcomp() As String)


        Dim ix As Integer = 0
        Dim imax As Integer = 0
        Dim iAdd As Integer = 0
        Dim iShort As Integer = 0

        'set mismatch flag to false
        ReadFIARelayINI.bMisMatchTool = False

        'Compare current WinFlex run (test) to FIA TOOL.......

        'If (String.Compare(clsReadFIAValues.strFIAAnnWDLimitProjTest, ReadFIARelayINI.strFIAAnnWDLimitProjTool, True)) <> 0 Then
        '    readfiarelayini.bmismatchtool = True
        '    clsReadFIAValues.strFIAAnnWDLimitProjMM = ib & ",Proj Ann WD Limit," & ReadFIARelayINI.strFIAAnnWDLimitProjTool & "," & clsReadFIAValues.strFIAAnnWDLimitProjTest & "&"
        'End If

        If ReadFIARelayINI.strFIAINIProdName = "FIA 7 Yr" Then
            If (String.Compare(clsReadFIAValues.strFIAGMCVTest, ReadFIARelayINI.strFIAGMCVTool, True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True
                ReadFIARelayINI.strFIAGMCVToolMM = ib & ",GMCV," & ReadFIARelayINI.strFIAGMCVTool & "," & clsReadFIAValues.strFIAGMCVTest & "&"
            End If
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASpecAnnCreditRateTest), CStr(ReadFIARelayINI.strFIASpecAnnCreditRateTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASpecAnnCreditRateTest
            Dim strTool As String = ReadFIARelayINI.strFIASpecAnnCreditRateTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASpecAnnCreditRateToolMM = ReadFIARelayINI.strFIASpecAnnCreditRateToolMM & ib & ",Spec Ann Credit Rate," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAFavAnnCreditRateTest), CStr(ReadFIARelayINI.strFIAFavAnnCreditRateTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAFavAnnCreditRateTest
            Dim strTool As String = ReadFIARelayINI.strFIAFavAnnCreditRateTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAFavAnnCreditRateToolMM = ReadFIARelayINI.strFIAFavAnnCreditRateToolMM & ib & ",Fav Ann Credit Rate," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavAnnCreditRateTest), CStr(ReadFIARelayINI.strFIAUnfavAnnCreditRateTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAUnfavAnnCreditRateTest
            Dim strTool As String = ReadFIARelayINI.strFIAUnfavAnnCreditRateTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAUnfavAnnCreditRateToolMM = ReadFIARelayINI.strFIAUnfavAnnCreditRateToolMM & ib & ",Unfav Ann Credit Rate," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASpecContractValueTest), CStr(ReadFIARelayINI.strFIASpecContractValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASpecContractValueTest
            Dim strTool As String = ReadFIARelayINI.strFIASpecContractValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASpecContractValueToolMM = ReadFIARelayINI.strFIASpecContractValueToolMM & ib & ",Spec CV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If


        If (String.Compare(CStr(clsReadFIAValues.strFIAFavContractValueTest), CStr(ReadFIARelayINI.strFIAFavContractValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAFavContractValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAFavContractValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAFavContractValueToolMM = ReadFIARelayINI.strFIAFavContractValueToolMM & ib & ",Fav CV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavContractValueTest), CStr(ReadFIARelayINI.strFIAUnfavContractValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAUnfavContractValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAUnfavContractValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAUnfavContractValueToolMM = ReadFIARelayINI.strFIAUnfavContractValueToolMM & ib & ",Unfav CV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASpecSurrenderValueTest), CStr(ReadFIARelayINI.strFIASpecSurrenderValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASpecSurrenderValueTest
            Dim strTool As String = ReadFIARelayINI.strFIASpecSurrenderValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASpecSurrenderValueToolMM = ReadFIARelayINI.strFIASpecSurrenderValueToolMM & ib & ",Spec SV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAFavSurrenderValueTest), CStr(ReadFIARelayINI.strFIAFavSurrenderValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAFavSurrenderValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAFavSurrenderValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAFavSurrenderValueToolMM = ReadFIARelayINI.strFIAFavSurrenderValueToolMM & ib & ",Fav SV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavSurrenderValueTest), CStr(ReadFIARelayINI.strFIAUnfavSurrenderValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAUnfavSurrenderValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAUnfavSurrenderValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM = ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM & ib & ",Unfav SV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASpecMGSVTest), CStr(ReadFIARelayINI.strFIASpecMGSVTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASpecMGSVTest
            Dim strTool As String = ReadFIARelayINI.strFIASpecMGSVTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASpecMGSVToolMM = ReadFIARelayINI.strFIASpecMGSVToolMM & ib & ",Spec MGSV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAFavMGSVTest), CStr(ReadFIARelayINI.strFIAFavMGSVTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAFavMGSVTest
            Dim strTool As String = ReadFIARelayINI.strFIAFavMGSVTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAFavMGSVToolMM = ReadFIARelayINI.strFIAFavMGSVToolMM & ib & ",Fav MGSV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavMGSVTest), CStr(ReadFIARelayINI.strFIAUnfavMGSVTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAUnfavMGSVTest
            Dim strTool As String = ReadFIARelayINI.strFIAUnfavMGSVTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAUnfavMGSVToolMM = ReadFIARelayINI.strFIAUnfavMGSVToolMM & ib & ",Unfav MGSV," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If ReadFIARelayINI.strFIAINIINcomeRider = "Yes" Then
            If (String.Compare(CStr(clsReadFIAValues.strFIASpecProjBeneBaseTest), CStr(ReadFIARelayINI.strFIASpecProjBeneBaseTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIASpecProjBeneBaseTest
                Dim strTool As String = ReadFIARelayINI.strFIASpecProjBeneBaseTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIASpecProjBeneBaseToolMM = ReadFIARelayINI.strFIASpecProjBeneBaseToolMM & ib & ",Spec Proj Bene Base," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavProjBeneBaseTest), CStr(ReadFIARelayINI.strFIAFavProjBeneBaseTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAFavProjBeneBaseTest
                Dim strTool As String = ReadFIARelayINI.strFIAFavProjBeneBaseTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAFavProjBeneBaseToolMM = ReadFIARelayINI.strFIAFavProjBeneBaseToolMM & ib & ",Fav Proj Bene Base," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavProjBeneBaseTest), CStr(ReadFIARelayINI.strFIAUnfavProjBeneBaseTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavProjBeneBaseTest
                Dim strTool As String = ReadFIARelayINI.strFIAUnfavProjBeneBaseTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAUnfavProjBeneBaseToolMM = ReadFIARelayINI.strFIAUnfavProjBeneBaseToolMM & ib & ",Unfav Proj Bene Base," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIASpecProjWDLimitTest), CStr(ReadFIARelayINI.strFIASpecProjWDLimitTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIASpecProjWDLimitTest
                Dim strTool As String = ReadFIARelayINI.strFIASpecProjWDLimitTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIASpecProjWDLimitToolMM = ReadFIARelayINI.strFIASpecProjWDLimitToolMM & ib & ",Spec Proj WD Limit," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAFavProjWDLimitTest), CStr(ReadFIARelayINI.strFIAFavProjWDLimitTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAFavProjWDLimitTest
                Dim strTool As String = ReadFIARelayINI.strFIAFavProjWDLimitTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAFavProjWDLimitToolMM = ReadFIARelayINI.strFIAFavProjWDLimitToolMM & ib & ",Fav Proj WD Limit," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavProjWDLimitTest), CStr(ReadFIARelayINI.strFIAUnfavProjWDLimitTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAUnfavProjWDLimitTest
                Dim strTool As String = ReadFIARelayINI.strFIAUnfavProjWDLimitTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAUnfavProjWDLimitToolMM = ReadFIARelayINI.strFIAUnfavProjWDLimitToolMM & ib & ",Unfav Proj WD Limit," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAMonthlyCapIndexCreditTest), CStr(ReadFIARelayINI.strFIAMonthlyCapIndexCreditTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAMonthlyCapIndexCreditTest
            Dim strTool As String = ReadFIARelayINI.strFIAMonthlyCapIndexCreditTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAMonthlyCapIndexCreditToolMM = ReadFIARelayINI.strFIAMonthlyCapIndexCreditToolMM & ib & ",Monthly Cap Index Credit," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualCapIndexCreditTest), CStr(ReadFIARelayINI.strFIAAnnualCapIndexCreditTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAAnnualCapIndexCreditTest
            Dim strTool As String = ReadFIARelayINI.strFIAAnnualCapIndexCreditTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAAnnualCapIndexCreditToolMM = ReadFIARelayINI.strFIAAnnualCapIndexCreditToolMM & ib & ",Annual Cap Index Credit," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAPerfTriggerIndexCreditTest), CStr(ReadFIARelayINI.strFIAPerfTriggerIndexCreditTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAPerfTriggerIndexCreditTest
            Dim strTool As String = ReadFIARelayINI.strFIAPerfTriggerIndexCreditTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAPerfTriggerIndexCreditToolMM = ReadFIARelayINI.strFIAPerfTriggerIndexCreditToolMM & ib & ",Perf Trigger Index Credit," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASevenYearIntRateTest), CStr(ReadFIARelayINI.strFIASevenYearIntRateTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASevenYearIntRateTest
            Dim strTool As String = ReadFIARelayINI.strFIASevenYearIntRateTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASevenYearIntRateToolMM = ReadFIARelayINI.strFIASevenYearIntRateToolMM & ib & ",7 Yr Int Rate," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIATenYearIntRateTest), CStr(ReadFIARelayINI.strFIATenYearIntRateTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIATenYearIntRateTest
            Dim strTool As String = ReadFIARelayINI.strFIATenYearIntRateTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIATenYearIntRateToolMM = ReadFIARelayINI.strFIATenYearIntRateToolMM & ib & ",10 Yr Int Rate," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASevenYearAccumValueTest), CStr(ReadFIARelayINI.strFIASevenYearAccumValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASevenYearAccumValueTest
            Dim strTool As String = ReadFIARelayINI.strFIASevenYearAccumValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASevenYearAccumValueToolMM = ReadFIARelayINI.strFIASevenYearAccumValueToolMM & ib & ",7 Year Accum Value," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIATenYearAccumValueTest), CStr(ReadFIARelayINI.strFIATenYearAccumValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIATenYearAccumValueTest
            Dim strTool As String = ReadFIARelayINI.strFIATenYearAccumValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIATenYearAccumValueToolMM = ReadFIARelayINI.strFIATenYearAccumValueToolMM & ib & ",10 Year Accum Value," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAMonthlyCapAccumValueTest), CStr(ReadFIARelayINI.strFIAMonthlyCapAccumValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAMonthlyCapAccumValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAMonthlyCapAccumValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAMonthlyCapAccumValueToolMM = ReadFIARelayINI.strFIAMonthlyCapAccumValueToolMM & ib & ",Montly Cap Accum Value," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAAnnualCapAccumValueTest), CStr(ReadFIARelayINI.strFIAAnnualCapAccumValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAAnnualCapAccumValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAAnnualCapAccumValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAAnnualCapAccumValueToolMM = ReadFIARelayINI.strFIAAnnualCapAccumValueToolMM & ib & ",Annual Cap Accum Value," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAPerfTriggerAccumValueTest), CStr(ReadFIARelayINI.strFIAPerfTriggerAccumValueTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAPerfTriggerAccumValueTest
            Dim strTool As String = ReadFIARelayINI.strFIAPerfTriggerAccumValueTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAPerfTriggerAccumValueToolMM = ReadFIARelayINI.strFIAPerfTriggerAccumValueToolMM & ib & ",Perf Trigger Accum Value," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If ReadFIARelayINI.strFIAINIINcomeRider = "Yes" Then
            If (String.Compare(CStr(clsReadFIAValues.strFIAContractValueNoWDTest), CStr(ReadFIARelayINI.strFIAContractValueNoWDTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAContractValueNoWDTest
                Dim strTool As String = ReadFIARelayINI.strFIAContractValueNoWDTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAContractValueNoWDToolMM = ReadFIARelayINI.strFIAContractValueNoWDToolMM & ib & ",CV No WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAGuarBeneBaseNoWDTest), CStr(ReadFIARelayINI.strFIAGuarBeneBaseNoWDTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAGuarBeneBaseNoWDTest
                Dim strTool As String = ReadFIARelayINI.strFIAGuarBeneBaseNoWDTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAGuarBeneBaseNoWDToolMM = ReadFIARelayINI.strFIAGuarBeneBaseNoWDToolMM & ib & ",Guar BeneBase No WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAProjBeneBaseNoWDTest), CStr(ReadFIARelayINI.strFIAProjBeneBaseNoWDTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAProjBeneBaseNoWDTest
                Dim strTool As String = ReadFIARelayINI.strFIAProjBeneBaseNoWDTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAProjBeneBaseNoWDToolMM = ReadFIARelayINI.strFIAProjBeneBaseNoWDToolMM & ib & ",Proj BeneBase No WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAGuarWDLimitNoWDTest), CStr(ReadFIARelayINI.strFIAGuarWDLimitNoWDTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAGuarWDLimitNoWDTest
                Dim strTool As String = ReadFIARelayINI.strFIAGuarWDLimitNoWDTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAGuarWDLimitNoWDToolMM = ReadFIARelayINI.strFIAGuarWDLimitNoWDToolMM & ib & ",Guar WD Limit No WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAProjWDLimitNoWDTest), CStr(ReadFIARelayINI.strFIAProjWDLimitNoWDTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAProjWDLimitNoWDTest
                Dim strTool As String = ReadFIARelayINI.strFIAProjWDLimitNoWDTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAProjWDLimitNoWDToolMM = ReadFIARelayINI.strFIAProjWDLimitNoWDToolMM & ib & ",Proj WD Limit No WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadFIAValues.strFIAGuarWDFactorTest), CStr(ReadFIARelayINI.strFIAGuarWDFactorTool), True)) <> 0 Then
                ReadFIARelayINI.bMisMatchTool = True

                Dim strTest As String = clsReadFIAValues.strFIAGuarWDFactorTest
                Dim strTool As String = ReadFIARelayINI.strFIAGuarWDFactorTool

                Dim SplitTest = Split(strTest, ",")
                Dim SplitTool = Split(strTool, ",")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitTool) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                    iShort = UBound(SplitTool)
                    ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                    For iAdd = iShort To UBound(SplitTool)
                        SplitTool(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitTool)
                    If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                        ReadFIARelayINI.strFIAGuarWDFactorToolMM = ReadFIARelayINI.strFIAGuarWDFactorToolMM & ib & ",Guar WD Factor," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If
        End If
        If (String.Compare(CStr(clsReadFIAValues.strFIASpecSPChangeTest), CStr(ReadFIARelayINI.strFIASpecSPChangeTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASpecSPChangeTest
            Dim strTool As String = ReadFIARelayINI.strFIASpecSPChangeTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASpecSPChangeToolMM = ReadFIARelayINI.strFIASpecSPChangeToolMM & ib & ",Spec SP Change," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAFavSPChangeTest), CStr(ReadFIARelayINI.strFIAFavSPChangeTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAFavSPChangeTest
            Dim strTool As String = ReadFIARelayINI.strFIAFavSPChangeTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAFavSPChangeToolMM = ReadFIARelayINI.strFIAFavSPChangeToolMM & ib & ",Fav SP Change," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavSPChangeTest), CStr(ReadFIARelayINI.strFIAUnfavSPChangeTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAUnfavSPChangeTest
            Dim strTool As String = ReadFIARelayINI.strFIAUnfavSPChangeTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAUnfavSPChangeToolMM = ReadFIARelayINI.strFIAUnfavSPChangeToolMM & ib & ",UnFav SP Change," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIASpecWDTest), CStr(ReadFIARelayINI.strFIASpecWDTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIASpecWDTest
            Dim strTool As String = ReadFIARelayINI.strFIASpecWDTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIASpecWDToolMM = ReadFIARelayINI.strFIASpecWDToolMM & ib & ",Spec WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAFavWDTest), CStr(ReadFIARelayINI.strFIAFavWDTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAFavWDTest
            Dim strTool As String = ReadFIARelayINI.strFIAFavWDTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAFavWDToolMM = ReadFIARelayINI.strFIAFavWDToolMM & ib & ",Fav WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        If (String.Compare(CStr(clsReadFIAValues.strFIAUnfavWDTest), CStr(ReadFIARelayINI.strFIAUnfavWDTool), True)) <> 0 Then
            ReadFIARelayINI.bMisMatchTool = True

            Dim strTest As String = clsReadFIAValues.strFIAUnfavWDTest
            Dim strTool As String = ReadFIARelayINI.strFIAUnfavWDTool

            Dim SplitTest = Split(strTest, ",")
            Dim SplitTool = Split(strTool, ",")

            'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
            If UBound(SplitTest) < UBound(SplitTool) Then
                iShort = UBound(SplitTest)
                ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitTool) - UBound(SplitTest))))
                For iAdd = iShort To UBound(SplitTest)
                    SplitTest(iAdd) = "0"
                Next
            ElseIf UBound(SplitTool) < UBound(SplitTest) Then
                iShort = UBound(SplitTool)
                ReDim Preserve SplitTool((UBound(SplitTool) + UBound(SplitTest) - UBound(SplitTool)))
                For iAdd = iShort To UBound(SplitTool)
                    SplitTool(iAdd) = "0"
                Next
            End If

            For iElement = 0 To UBound(SplitTool)
                If String.Compare(SplitTest(iElement), SplitTool(iElement), True) <> 0 Then
                    ReadFIARelayINI.strFIAUnfavWDToolMM = ReadFIARelayINI.strFIAUnfavWDToolMM & ib & ",UnFav WD," & iElement & "," & SplitTool(iElement) & "," & SplitTest(iElement) & "&"
                End If
            Next
        End If

        'if cases don't match, create the mismatch reports
        If ReadFIARelayINI.bMisMatchTool = True Then
            ReadFIARelayINI.bMismatchToolAtLeastOnce = True
            'SaveNewTool(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
            CreateFIATOOLMismatchReport(ib)
            strClientMisMatchListFIATool = strClientMisMatchListFIATool & "," & ib
            strsplitmmlistFIATool = Split(strClientMisMatchListFIATool, ",")
            'WriteMismatchList(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If
        'copy the Test.pdf to the test folder under the client folder, whether there are mismatches or not
        'If kclbClientList.GetItemChecked(ib - 1) Then
        '    CopyPDF(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        'End If
        'write the stats to the file
        WriteMatchStatusFIATool(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp, ReadFIARelayINI.bMisMatchTool, bNewBench:=False)


    End Sub
    Private Sub CompareValuesSPIA(ByVal ic As Integer, ByVal ib As Integer, ByVal strcomp() As String)

        Dim ix As Integer = 0
        Dim imax As Integer = 0
        Dim iAdd As Integer = 0
        Dim iShort As Integer = 0

        'set mismatch flag to false
        clsReadSPIAValues.bMisMatch = False

        'go through each value in relay.out to compare between bench and test, if mismatch, create strings for datagrid

        'below 20 are for warning and error messages
        If clsReadSPIAValues.bErrorBench = True And clsReadSPIAValues.bErrorTest = True Then
            If (String.Compare(clsReadSPIAValues.strMessage1Test, clsReadSPIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage1MM = ib & ",Message 1," & clsReadSPIAValues.strMessage1Bench & "," & clsReadSPIAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage2Test, clsReadSPIAValues.strMessage2Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage2MM = ib & ",Message 2," & clsReadSPIAValues.strMessage2Bench & "," & clsReadSPIAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage3Test, clsReadSPIAValues.strMessage3Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage3MM = ib & ",Message 3," & clsReadSPIAValues.strMessage3Bench & "," & clsReadSPIAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage4Test, clsReadSPIAValues.strMessage4Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage4MM = ib & ",Message 4," & clsReadSPIAValues.strMessage4Bench & "," & clsReadSPIAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage5Test, clsReadSPIAValues.strMessage5Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage5MM = ib & ",Message 5," & clsReadSPIAValues.strMessage5Bench & "," & clsReadSPIAValues.strMessage5Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage6Test, clsReadSPIAValues.strMessage6Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage6MM = ib & ",Message 6," & clsReadSPIAValues.strMessage6Bench & "," & clsReadSPIAValues.strMessage6Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage7Test, clsReadSPIAValues.strMessage7Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage7MM = ib & ",Message 7," & clsReadSPIAValues.strMessage7Bench & "," & clsReadSPIAValues.strMessage7Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage8Test, clsReadSPIAValues.strMessage8Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage8MM = ib & ",Message 8," & clsReadSPIAValues.strMessage8Bench & "," & clsReadSPIAValues.strMessage8Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage9Test, clsReadSPIAValues.strMessage9Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage9MM = ib & ",Message 9," & clsReadSPIAValues.strMessage9Bench & "," & clsReadSPIAValues.strMessage9Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage10Test, clsReadSPIAValues.strMessage10Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage10MM = ib & ",Message 10," & clsReadSPIAValues.strMessage10Bench & "," & clsReadSPIAValues.strMessage10Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage11Test, clsReadSPIAValues.strMessage11Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage11MM = ib & ",Message 11," & clsReadSPIAValues.strMessage11Bench & "," & clsReadSPIAValues.strMessage11Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage12Test, clsReadSPIAValues.strMessage12Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage12MM = ib & ",Message 12," & clsReadSPIAValues.strMessage12Bench & "," & clsReadSPIAValues.strMessage12Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage13Test, clsReadSPIAValues.strMessage13Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage13MM = ib & ",Message 13," & clsReadSPIAValues.strMessage13Bench & "," & clsReadSPIAValues.strMessage13Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage14Test, clsReadSPIAValues.strMessage14Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage14MM = ib & ",Message 14," & clsReadSPIAValues.strMessage14Bench & "," & clsReadSPIAValues.strMessage14Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage15Test, clsReadSPIAValues.strMessage15Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage15MM = ib & ",Message 15," & clsReadSPIAValues.strMessage15Bench & "," & clsReadSPIAValues.strMessage15Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage16Test, clsReadSPIAValues.strMessage16Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage16MM = ib & ",Message 16," & clsReadSPIAValues.strMessage16Bench & "," & clsReadSPIAValues.strMessage16Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage17Test, clsReadSPIAValues.strMessage17Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage17MM = ib & ",Message 17," & clsReadSPIAValues.strMessage17Bench & "," & clsReadSPIAValues.strMessage17Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage18Test, clsReadSPIAValues.strMessage18Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage18MM = ib & ",Message 18," & clsReadSPIAValues.strMessage18Bench & "," & clsReadSPIAValues.strMessage18Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage19Test, clsReadSPIAValues.strMessage19Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage19MM = ib & ",Message 19," & clsReadSPIAValues.strMessage19Bench & "," & clsReadSPIAValues.strMessage19Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage20Test, clsReadSPIAValues.strMessage20Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage20MM = ib & ",Message 20," & clsReadSPIAValues.strMessage20Bench & "," & clsReadSPIAValues.strMessage20Test & "&"
            End If

            'Set messages back to nothing
            clsReadSPIAValues.strMessage1Bench = ""
            clsReadSPIAValues.strMessage1Test = ""
            clsReadSPIAValues.strMessage2Bench = ""
            clsReadSPIAValues.strMessage2Test = ""
            clsReadSPIAValues.strMessage3Bench = ""
            clsReadSPIAValues.strMessage3Test = ""
            clsReadSPIAValues.strMessage4Bench = ""
            clsReadSPIAValues.strMessage4Test = ""
            clsReadSPIAValues.strMessage5Bench = ""
            clsReadSPIAValues.strMessage5Test = ""
            clsReadSPIAValues.strMessage6Bench = ""
            clsReadSPIAValues.strMessage6Test = ""
            clsReadSPIAValues.strMessage7Bench = ""
            clsReadSPIAValues.strMessage7Test = ""
            clsReadSPIAValues.strMessage8Bench = ""
            clsReadSPIAValues.strMessage8Test = ""
            clsReadSPIAValues.strMessage9Bench = ""
            clsReadSPIAValues.strMessage9Test = ""
            clsReadSPIAValues.strMessage10Bench = ""
            clsReadSPIAValues.strMessage10Test = ""
            clsReadSPIAValues.strMessage11Bench = ""
            clsReadSPIAValues.strMessage11Test = ""
            clsReadSPIAValues.strMessage12Bench = ""
            clsReadSPIAValues.strMessage12Test = ""
            clsReadSPIAValues.strMessage13Bench = ""
            clsReadSPIAValues.strMessage13Test = ""
            clsReadSPIAValues.strMessage14Bench = ""
            clsReadSPIAValues.strMessage14Test = ""
            clsReadSPIAValues.strMessage15Bench = ""
            clsReadSPIAValues.strMessage15Test = ""
            clsReadSPIAValues.strMessage16Bench = ""
            clsReadSPIAValues.strMessage16Test = ""
            clsReadSPIAValues.strMessage17Bench = ""
            clsReadSPIAValues.strMessage17Test = ""
            clsReadSPIAValues.strMessage18Bench = ""
            clsReadSPIAValues.strMessage18Test = ""
            clsReadSPIAValues.strMessage19Bench = ""
            clsReadSPIAValues.strMessage19Test = ""
            clsReadSPIAValues.strMessage20Bench = ""
            clsReadSPIAValues.strMessage20Test = ""

            'see if one runs and one doesnt
        ElseIf clsReadSPIAValues.bErrorBench = True And clsReadSPIAValues.bErrorTest = False Then
            clsReadSPIAValues.bMisMatch = True
            clsReadSPIAValues.strRunNoRunMM = ib & ",Test runs/Bench doesn't run,,&"

            If (String.Compare(clsReadSPIAValues.strMessage1Test, clsReadSPIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage1MM = ib & ",Message1," & clsReadSPIAValues.strMessage1Bench & "," & clsReadSPIAValues.strMessage1Test & "&"
            End If

            clsReadSPIAValues.strMessage1Test = ""
            clsReadSPIAValues.strMessage1Bench = ""

        ElseIf clsReadSPIAValues.bErrorBench = False And clsReadSPIAValues.bErrorTest = True Then
            clsReadSPIAValues.bMisMatch = True
            clsReadSPIAValues.strRunNoRunMM = ib & ",Bench runs/Test doesn't run,,&"

            If (String.Compare(clsReadSPIAValues.strMessage1Test, clsReadSPIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage1MM = ib & ",Message1," & clsReadSPIAValues.strMessage1Bench & "," & clsReadSPIAValues.strMessage1Test & "&"
            End If

            clsReadSPIAValues.strMessage1Test = ""
            clsReadSPIAValues.strMessage1Bench = ""
        Else
            'if both cases run, then compare...

            If (String.Compare(clsReadSPIAValues.strMessage1Test, clsReadSPIAValues.strMessage1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage1MM = ib & ",Message1," & clsReadSPIAValues.strMessage1Bench & "," & clsReadSPIAValues.strMessage1Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage2Test, clsReadSPIAValues.strMessage2Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage2MM = ib & ",Message2," & clsReadSPIAValues.strMessage2Bench & "," & clsReadSPIAValues.strMessage2Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage3Test, clsReadSPIAValues.strMessage3Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage3MM = ib & ",Message1," & clsReadSPIAValues.strMessage3Bench & "," & clsReadSPIAValues.strMessage3Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage4Test, clsReadSPIAValues.strMessage4Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage4MM = ib & ",Message4," & clsReadSPIAValues.strMessage4Bench & "," & clsReadSPIAValues.strMessage4Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strMessage5Test, clsReadSPIAValues.strMessage5Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage5MM = ib & ",Message5," & clsReadSPIAValues.strMessage5Bench & "," & clsReadSPIAValues.strMessage5Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage6Test, clsReadSPIAValues.strMessage6Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage6MM = ib & ",Message6," & clsReadSPIAValues.strMessage6Bench & "," & clsReadSPIAValues.strMessage6Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage7Test, clsReadSPIAValues.strMessage7Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage7MM = ib & ",Message7," & clsReadSPIAValues.strMessage7Bench & "," & clsReadSPIAValues.strMessage7Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage8Test, clsReadSPIAValues.strMessage8Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage8MM = ib & ",Message8," & clsReadSPIAValues.strMessage8Bench & "," & clsReadSPIAValues.strMessage8Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage9Test, clsReadSPIAValues.strMessage9Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage9MM = ib & ",Message9," & clsReadSPIAValues.strMessage9Bench & "," & clsReadSPIAValues.strMessage9Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage10Test, clsReadSPIAValues.strMessage10Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage10MM = ib & ",Message10," & clsReadSPIAValues.strMessage10Bench & "," & clsReadSPIAValues.strMessage10Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage11Test, clsReadSPIAValues.strMessage11Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage11MM = ib & ",Message11," & clsReadSPIAValues.strMessage11Bench & "," & clsReadSPIAValues.strMessage11Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage12Test, clsReadSPIAValues.strMessage12Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage12MM = ib & ",Message12," & clsReadSPIAValues.strMessage12Bench & "," & clsReadSPIAValues.strMessage12Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage13Test, clsReadSPIAValues.strMessage13Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage13MM = ib & ",Message13," & clsReadSPIAValues.strMessage13Bench & "," & clsReadSPIAValues.strMessage13Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage14Test, clsReadSPIAValues.strMessage14Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage14MM = ib & ",Message14," & clsReadSPIAValues.strMessage14Bench & "," & clsReadSPIAValues.strMessage14Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage15Test, clsReadSPIAValues.strMessage15Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage15MM = ib & ",Message15," & clsReadSPIAValues.strMessage15Bench & "," & clsReadSPIAValues.strMessage15Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage16Test, clsReadSPIAValues.strMessage16Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage16MM = ib & ",Message16," & clsReadSPIAValues.strMessage16Bench & "," & clsReadSPIAValues.strMessage16Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage17Test, clsReadSPIAValues.strMessage17Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage17MM = ib & ",Message17," & clsReadSPIAValues.strMessage17Bench & "," & clsReadSPIAValues.strMessage17Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage18Test, clsReadSPIAValues.strMessage18Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage18MM = ib & ",Message18," & clsReadSPIAValues.strMessage18Bench & "," & clsReadSPIAValues.strMessage18Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage19Test, clsReadSPIAValues.strMessage19Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage19MM = ib & ",Message19," & clsReadSPIAValues.strMessage19Bench & "," & clsReadSPIAValues.strMessage19Test & "&"
            End If
            If (String.Compare(clsReadSPIAValues.strMessage20Test, clsReadSPIAValues.strMessage20Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strMessage20MM = ib & ",Message20," & clsReadSPIAValues.strMessage20Bench & "," & clsReadSPIAValues.strMessage20Test & "&"
            End If

            'Set messages back to nothing
            clsReadSPIAValues.strMessage1Bench = ""
            clsReadSPIAValues.strMessage1Test = ""
            clsReadSPIAValues.strMessage2Bench = ""
            clsReadSPIAValues.strMessage2Test = ""
            clsReadSPIAValues.strMessage3Bench = ""
            clsReadSPIAValues.strMessage3Test = ""
            clsReadSPIAValues.strMessage4Bench = ""
            clsReadSPIAValues.strMessage4Test = ""
            clsReadSPIAValues.strMessage5Bench = ""
            clsReadSPIAValues.strMessage5Test = ""
            clsReadSPIAValues.strMessage6Bench = ""
            clsReadSPIAValues.strMessage6Test = ""
            clsReadSPIAValues.strMessage7Bench = ""
            clsReadSPIAValues.strMessage7Test = ""
            clsReadSPIAValues.strMessage8Bench = ""
            clsReadSPIAValues.strMessage8Test = ""
            clsReadSPIAValues.strMessage9Bench = ""
            clsReadSPIAValues.strMessage9Test = ""
            clsReadSPIAValues.strMessage10Bench = ""
            clsReadSPIAValues.strMessage10Test = ""
            clsReadSPIAValues.strMessage11Bench = ""
            clsReadSPIAValues.strMessage11Test = ""
            clsReadSPIAValues.strMessage12Bench = ""
            clsReadSPIAValues.strMessage12Test = ""
            clsReadSPIAValues.strMessage13Bench = ""
            clsReadSPIAValues.strMessage13Test = ""
            clsReadSPIAValues.strMessage14Bench = ""
            clsReadSPIAValues.strMessage14Test = ""
            clsReadSPIAValues.strMessage15Bench = ""
            clsReadSPIAValues.strMessage15Test = ""
            clsReadSPIAValues.strMessage16Bench = ""
            clsReadSPIAValues.strMessage16Test = ""
            clsReadSPIAValues.strMessage17Bench = ""
            clsReadSPIAValues.strMessage17Test = ""
            clsReadSPIAValues.strMessage18Bench = ""
            clsReadSPIAValues.strMessage18Test = ""
            clsReadSPIAValues.strMessage19Bench = ""
            clsReadSPIAValues.strMessage19Test = ""
            clsReadSPIAValues.strMessage20Bench = ""
            clsReadSPIAValues.strMessage20Test = ""


            If (String.Compare(clsReadSPIAValues.strSPIAStreamCountTest, clsReadSPIAValues.strSPIAStreamCountBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAstreamcountMM = ib & ",Stream Count," & clsReadSPIAValues.strSPIAStreamCountBench & "," & clsReadSPIAValues.strSPIAStreamCountTest & "&"
            End If


            If (String.Compare(clsReadSPIAValues.strSPIACompanyNameTest, clsReadSPIAValues.strSPIACompanyNameBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIACompanyNameMM = ib & ",Company Name," & clsReadSPIAValues.strSPIACompanyNameBench & "," & clsReadSPIAValues.strSPIACompanyNameTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAClient1Test, clsReadSPIAValues.strSPIAClient1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAClient1MM = ib & ",Client1 Name, " & clsReadSPIAValues.strSPIAClient1Bench & "," & clsReadSPIAValues.strSPIAClient1Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAAge1Test, clsReadSPIAValues.strSPIAAge1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAAge1MM = ib & ",Client1 Age," & clsReadSPIAValues.strSPIAAge1Bench & "," & clsReadSPIAValues.strSPIAAge1Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIASex1Test, clsReadSPIAValues.strSPIASex1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIASex1MM = ib & ",Client1 Sex," & clsReadSPIAValues.strSPIASex1Bench & "," & clsReadSPIAValues.strSPIASex1Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIADOB1Test, clsReadSPIAValues.strSPIADOB1Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIADOB1MM = ib & ",Client1 DOB," & clsReadSPIAValues.strSPIADOB1Bench & "," & clsReadSPIAValues.strSPIADOB1Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAClient2Test, clsReadSPIAValues.strSPIAClient2Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAClient2MM = ib & ",Client2 Name, " & clsReadSPIAValues.strSPIAClient2Bench & "," & clsReadSPIAValues.strSPIAClient2Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAAge2Test, clsReadSPIAValues.strSPIAAge2Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAAge2MM = ib & ",Client2 Age," & clsReadSPIAValues.strSPIAAge2Bench & "," & clsReadSPIAValues.strSPIAAge2Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIASex2Test, clsReadSPIAValues.strSPIASex2Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIASex2MM = ib & ",Client2 Sex," & clsReadSPIAValues.strSPIASex2Bench & "," & clsReadSPIAValues.strSPIASex2Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIADOB2Test, clsReadSPIAValues.strSPIADOB2Bench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIADOB2MM = ib & ",Client2 DOB," & clsReadSPIAValues.strSPIADOB2Bench & "," & clsReadSPIAValues.strSPIADOB2Test & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACompanyNameLongTest, clsReadSPIAValues.strSPIACompanyNameLongBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIACompanyNameLongMM = ib & ",Comp Name Long," & clsReadSPIAValues.strSPIACompanyNameLongBench & "," & clsReadSPIAValues.strSPIACompanyNameLongTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAProductNameLongTest, clsReadSPIAValues.strSPIAProductNameLongBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAProductNameLongMM = ib & ",Prod Name Long," & clsReadSPIAValues.strSPIAProductNameLongBench & "," & clsReadSPIAValues.strSPIAProductNameLongTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAProdNameTest, clsReadSPIAValues.strSPIAProdNameBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAProdNameMM = ib & ",Prod Name Short," & clsReadSPIAValues.strSPIAProdNameBench & "," & clsReadSPIAValues.strSPIAProdNameTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIARatePricingCodeTest, clsReadSPIAValues.strSPIARatePricingCodeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIARatePricingCodeMM = ib & ",Rate Pricing Code," & clsReadSPIAValues.strSPIARatePricingCodeBench & "," & clsReadSPIAValues.strSPIARatePricingCodeTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIASystemVersionTest, clsReadSPIAValues.strSPIASystemVersionBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIASystemVersionMM = ib & ",System Version," & clsReadSPIAValues.strSPIASystemVersionBench & "," & clsReadSPIAValues.strSPIASystemVersionTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAPayoutRateCodesTest, clsReadSPIAValues.strSPIAPayoutRateCodesBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPayoutRateCodesMM = ib & ",Payout Rate Codes," & clsReadSPIAValues.strSPIAPayoutRateCodesBench & "," & clsReadSPIAValues.strSPIAPayoutRateCodesTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAAgentTest, clsReadSPIAValues.strSPIAAgentBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAAgentMM = ib & ",Agent Name," & clsReadSPIAValues.strSPIAAgentBench & "," & clsReadSPIAValues.strSPIAAgentTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAHOApprovalAmtTest, clsReadSPIAValues.strSPIAHOApprovalAmtBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAHOApprovalAmtMM = ib & ",HO Premium Limit," & clsReadSPIAValues.strSPIAHOApprovalAmtBench & "," & clsReadSPIAValues.strSPIAHOApprovalAmtTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAPolicyFeeThresholdTest, clsReadSPIAValues.strSPIAPolicyFeeThresholdBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPolicyFeeThresholdMM = ib & ",Policy Fee Threshold," & clsReadSPIAValues.strSPIAPolicyFeeThresholdBench & "," & clsReadSPIAValues.strSPIAPolicyFeeThresholdTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAPolicyFeeTest, clsReadSPIAValues.strSPIAPolicyFeeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPolicyFeeMM = ib & ",Policy Fee Amt," & clsReadSPIAValues.strSPIAPolicyFeeBench & "," & clsReadSPIAValues.strSPIAPolicyFeeTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommissionLoadTest, clsReadSPIAValues.strSPIACommissionLoadBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIACommissionLoadMM = ib & ",Commission Load," & clsReadSPIAValues.strSPIACommissionLoadBench & "," & clsReadSPIAValues.strSPIACommissionLoadTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAGenerationCodeTest, clsReadSPIAValues.strSPIAGenerationCodeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAGenerationCodeMM = ib & ",Generational Code," & clsReadSPIAValues.strSPIAGenerationCodeBench & "," & clsReadSPIAValues.strSPIAGenerationCodeTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAGuarMortalityCodeTest, clsReadSPIAValues.strSPIAGuarMortalityCodeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAGuarMortalityCodeMM = ib & ",Mortality Code," & clsReadSPIAValues.strSPIAGuarMortalityCodeBench & "," & clsReadSPIAValues.strSPIAGuarMortalityCodeTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAAggregationAmtTest, clsReadSPIAValues.strSPIAAggregationAmtBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAAggregationAmtMM = ib & ",Aggregation Amount," & clsReadSPIAValues.strSPIAAggregationAmtBench & "," & clsReadSPIAValues.strSPIAAggregationAmtTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIARestrictionsEndorsementTest, clsReadSPIAValues.strSPIARestrictionsEndorsementBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIARestrictionsEndorsementMM = ib & ",Restrictions Endorsement?," & clsReadSPIAValues.strSPIARestrictionsEndorsementBench & "," & clsReadSPIAValues.strSPIARestrictionsEndorsementTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAGroupCodeTest, clsReadSPIAValues.strSPIAGroupCodeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAGroupCodeMM = ib & ",Group Code," & clsReadSPIAValues.strSPIAGroupCodeBench & "," & clsReadSPIAValues.strSPIAGroupCodeTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAPremiumTest, clsReadSPIAValues.strSPIAPremiumBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAPremiumTest
                Dim strBench As String = clsReadSPIAValues.strSPIAPremiumBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAPremiumMM(ix) = clsReadSPIAValues.strSPIAPremiumMM(ix) & ib & ",Premium " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAInitBenefitTest, clsReadSPIAValues.strSPIAInitBenefitBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAInitBenefitTest
                Dim strBench As String = clsReadSPIAValues.strSPIAInitBenefitBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAInitBenefitMM(ix) = clsReadSPIAValues.strSPIAInitBenefitMM(ix) & ib & ",Initial Benefit " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAPolicyFeeUsedTest, clsReadSPIAValues.strSPIAPolicyFeeUsedBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAPolicyFeeUsedTest
                Dim strBench As String = clsReadSPIAValues.strSPIAPolicyFeeUsedBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAPolicyFeeUsedMM(ix) = clsReadSPIAValues.strSPIAPolicyFeeUsedMM(ix) & ib & ",Policy Fee? " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAAdvanceEligibleTest, clsReadSPIAValues.strSPIAAdvanceEligibleBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAAdvanceEligibleTest
                Dim strBench As String = clsReadSPIAValues.strSPIAAdvanceEligibleBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAAdvanceEligibleMM(ix) = clsReadSPIAValues.strSPIAAdvanceEligibleMM(ix) & ib & ",Advance Eligible? " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAGuaranteedYrsTest, clsReadSPIAValues.strSPIAGuaranteedYrsBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAGuaranteedYrsTest
                Dim strBench As String = clsReadSPIAValues.strSPIAGuaranteedYrsBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAGuaranteedYrsMM(ix) = clsReadSPIAValues.strSPIAGuaranteedYrsMM(ix) & ib & ",Certain Years " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(clsReadSPIAValues.strSPIAGuaranteedMthsTest, clsReadSPIAValues.strSPIAGuaranteedMthsBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAGuaranteedMthsTest
                Dim strBench As String = clsReadSPIAValues.strSPIAGuaranteedMthsBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAGuaranteedMthsMM(ix) = clsReadSPIAValues.strSPIAGuaranteedMthsMM(ix) & ib & ",Certain Months " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAIRRTest, clsReadSPIAValues.strSPIAIRRBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAIRRTest
                Dim strBench As String = clsReadSPIAValues.strSPIAIRRBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAIRRMM(ix) = clsReadSPIAValues.strSPIAIRRMM(ix) & ib & ",IRR % " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If


            If (String.Compare(clsReadSPIAValues.strSPIAIncreasePctTest, clsReadSPIAValues.strSPIAIncreasePctBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAIncreasePctTest
                Dim strBench As String = clsReadSPIAValues.strSPIAIncreasePctBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAIncreasePctMM(ix) = clsReadSPIAValues.strSPIAIncreasePctMM(ix) & ib & ",Increase % " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIARateChangeDateTest, clsReadSPIAValues.strSPIARateChangeDateBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIARateChangeDateMM = ib & ",Rate Effective Date," & clsReadSPIAValues.strSPIARateChangeDateBench & "," & clsReadSPIAValues.strSPIARateChangeDateTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAAnnuityTypeTest, clsReadSPIAValues.strSPIAAnnuityTypeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAAnnuityTypeTest
                Dim strBench As String = clsReadSPIAValues.strSPIAAnnuityTypeBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAAnnuityTypeMM(ix) = clsReadSPIAValues.strSPIAAnnuityTypeMM(ix) & ib & ",Annuity Type " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIATaxFreeAmtTest, clsReadSPIAValues.strSPIATaxFreeAmtBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIATaxFreeAmtTest
                Dim strBench As String = clsReadSPIAValues.strSPIATaxFreeAmtBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIATaxFreeAmtMM(ix) = clsReadSPIAValues.strSPIATaxFreeAmtMM(ix) & ib & ",Tax Free Amt " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAExclusionRatioTest, clsReadSPIAValues.strSPIAExclusionRatioBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAExclusionRatioTest
                Dim strBench As String = clsReadSPIAValues.strSPIAExclusionRatioBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAExclusionRatioMM(ix) = clsReadSPIAValues.strSPIAExclusionRatioMM(ix) & ib & ",Exclusion Ratio " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIATotalCertPaymentsTest, clsReadSPIAValues.strSPIATotalCertPaymentsBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIATotalCertPaymentsTest
                Dim strBench As String = clsReadSPIAValues.strSPIATotalCertPaymentsBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIATotalCertPaymentsMM(ix) = clsReadSPIAValues.strSPIATotalCertPaymentsMM(ix) & ib & ",Total Cert Payments " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIATotalPaymentsLifeExpecTest, clsReadSPIAValues.strSPIATotalPaymentsLifeExpecBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIATotalPaymentsLifeExpecTest
                Dim strBench As String = clsReadSPIAValues.strSPIATotalPaymentsLifeExpecBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIATotalPaymentsLifeExpecMM(ix) = clsReadSPIAValues.strSPIATotalPaymentsLifeExpecMM(ix) & ib & ",Total Payments Life Expec " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIASurvivorPctTest, clsReadSPIAValues.strSPIASurvivorPctBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIASurvivorPctTest
                Dim strBench As String = clsReadSPIAValues.strSPIASurvivorPctBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIASurvivorPctMM(ix) = clsReadSPIAValues.strSPIASurvivorPctMM(ix) & ib & ",Survivor % " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAResidentStateTest, clsReadSPIAValues.strSPIAResidentStateBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAResidentStateMM = ib & ",Owner State," & clsReadSPIAValues.strSPIAResidentStateBench & "," & clsReadSPIAValues.strSPIAResidentStateTest & "&"
            End If

            If (String.Compare(clsReadSPIAValues.strSPIAApplicationStateTest, clsReadSPIAValues.strSPIAApplicationStateBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAApplicationStateMM = ib & ",Application State," & clsReadSPIAValues.strSPIAApplicationStateBench & "," & clsReadSPIAValues.strSPIAApplicationStateTest & "&"
            End If


            If (String.Compare(clsReadSPIAValues.strSPIAPremiumTaxTest, clsReadSPIAValues.strSPIAPremiumTaxBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPremiumTaxMM = ib & ",Premium Tax," & clsReadSPIAValues.strSPIAPremiumTaxBench & "," & clsReadSPIAValues.strSPIAPremiumTaxTest & "&"
            End If

            If CInt(clsReadSPIAValues.strSPIAStreamCountTest) = CInt(clsReadSPIAValues.strSPIAStreamCountBench) Then

                For ix = 1 To CInt(clsReadSPIAValues.strSPIAStreamCountTest)

                    If (String.Compare(clsReadSPIAValues.strSPIABeneSchedStartDateTest(ix), clsReadSPIAValues.strSPIABeneSchedStartDateBench(ix), True)) <> 0 Then
                        clsReadSPIAValues.bMisMatch = True

                        Dim strTest As String = clsReadSPIAValues.strSPIABeneSchedStartDateTest(ix)
                        Dim strBench As String = clsReadSPIAValues.strSPIABeneSchedStartDateBench(ix)

                        Dim SplitTest = Split(strTest, "�")
                        Dim SplitBench = Split(strBench, "�")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadSPIAValues.strSPIABeneSchedStartDateMM(ix) = clsReadSPIAValues.strSPIABeneSchedStartDateMM(ix) & ib & ",Bene Sched Start Date " & ix & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentTest(ix), clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentBench(ix), True)) <> 0 Then
                        clsReadSPIAValues.bMisMatch = True

                        Dim strTest As String = clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentTest(ix)
                        Dim strBench As String = clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentBench(ix)

                        Dim SplitTest = Split(strTest, "�")
                        Dim SplitBench = Split(strBench, "�")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentMM(ix) = clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentMM(ix) & ib & ",Bene Sched Ann Income Payment " & ix & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyTest(ix), clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyBench(ix), True)) <> 0 Then
                        clsReadSPIAValues.bMisMatch = True

                        Dim strTest As String = clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyTest(ix)
                        Dim strBench As String = clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyBench(ix)

                        Dim SplitTest = Split(strTest, "�")
                        Dim SplitBench = Split(strBench, "�")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If

                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyMM(ix) = clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyMM(ix) & ib & ",Bene Sched Ann Payment Living Only " & ix & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If

                    If (String.Compare(clsReadSPIAValues.strSPIABeneSchedCumPayoutTest(ix), clsReadSPIAValues.strSPIABeneSchedCumPayoutBench(ix), True)) <> 0 Then
                        clsReadSPIAValues.bMisMatch = True

                        Dim strTest As String = clsReadSPIAValues.strSPIABeneSchedCumPayoutTest(ix)
                        Dim strBench As String = clsReadSPIAValues.strSPIABeneSchedCumPayoutBench(ix)

                        Dim SplitTest = Split(strTest, "�")
                        Dim SplitBench = Split(strBench, "�")

                        'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                        If UBound(SplitTest) < UBound(SplitBench) Then
                            iShort = UBound(SplitTest)
                            ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                            For iAdd = iShort To UBound(SplitTest)
                                SplitTest(iAdd) = "0"
                            Next
                        ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                            iShort = UBound(SplitBench)
                            ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                            For iAdd = iShort To UBound(SplitBench)
                                SplitBench(iAdd) = "0"
                            Next
                        End If


                        For iElement = 0 To UBound(SplitBench)
                            If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                                clsReadSPIAValues.strSPIABeneSchedCumPayoutMM(ix) = clsReadSPIAValues.strSPIABeneSchedCumPayoutMM(ix) & ib & ",Bene Sched Cumulative Payout " & ix & "," & iElement & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                            End If
                        Next
                    End If
                Next ix
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIATaxStatusTest), CStr(clsReadSPIAValues.strSPIATaxStatusBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIATaxStatusMM = ib & ",Type of Funds," & clsReadSPIAValues.strSPIATaxStatusBench & "," & clsReadSPIAValues.strSPIATaxStatusTest & "&"
            Else

            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIACostBasisTest), CStr(clsReadSPIAValues.strSPIACostBasisBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIACostBasisMM = ib & ",Cost Basis," & clsReadSPIAValues.strSPIACostBasisBench & "," & clsReadSPIAValues.strSPIACostBasisTest & "&"
            Else

            End If

            If (String.Compare(clsReadSPIAValues.strSPIABeneReductionOptionTest, clsReadSPIAValues.strSPIABeneReductionOptionBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIABeneReductionOptionTest
                Dim strBench As String = clsReadSPIAValues.strSPIABeneReductionOptionBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIABeneReductionOptionMM = clsReadSPIAValues.strSPIABeneReductionOptionMM & ib & ",Benefit Reduction Opt " & "," & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAPurchaseDateTest), CStr(clsReadSPIAValues.strSPIAPurchaseDateBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPurchaseDateMM = ib & ",Purchase Date," & clsReadSPIAValues.strSPIAPurchaseDateBench & "," & clsReadSPIAValues.strSPIAPurchaseDateTest & "&"
            End If


            If (String.Compare(clsReadSPIAValues.strSPIAIncomeStartDateTest, clsReadSPIAValues.strSPIAIncomeStartDateBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAIncomeStartDateTest
                Dim strBench As String = clsReadSPIAValues.strSPIAIncomeStartDateBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAIncomeStartDateMM = clsReadSPIAValues.strSPIAIncomeStartDateMM & ib & ",Income Start Date " & "," & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIABenefitErrorTest, clsReadSPIAValues.strSPIABenefitErrorBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIABenefitErrorTest
                Dim strBench As String = clsReadSPIAValues.strSPIABenefitErrorBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIABenefitErrorMM(ix) = clsReadSPIAValues.strSPIABenefitErrorMM(ix) & ib & ",Stream Error Message " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If
            
            If (String.Compare(clsReadSPIAValues.strSPIAColaTest, clsReadSPIAValues.strSPIAColaBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIAColaTest
                Dim strBench As String = clsReadSPIAValues.strSPIAColaBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIAColaMM = clsReadSPIAValues.strSPIAColaMM & ib & ",COLA Description " & "," & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If
            
            If (String.Compare(clsReadSPIAValues.strSPIALifeExpectancyTest, clsReadSPIAValues.strSPIALifeExpectancyBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIALifeExpectancyTest
                Dim strBench As String = clsReadSPIAValues.strSPIALifeExpectancyBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIALifeExpectancyMM = clsReadSPIAValues.strSPIALifeExpectancyMM & ib & ",Life Expectancy " & "," & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If
            
            If (String.Compare(clsReadSPIAValues.strSPIACommutationAmtTest, clsReadSPIAValues.strSPIACommutationAmtBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationAmtTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationAmtBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationAmtMM = clsReadSPIAValues.strSPIACommutationAmtMM & ib & ",Commutation Amount " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationTypeTest, clsReadSPIAValues.strSPIACommutationTypeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationTypeTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationTypeBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationTypeMM = clsReadSPIAValues.strSPIACommutationTypeMM & ib & ",Commutation Type " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationTaxFreeAmtTest, clsReadSPIAValues.strSPIACommutationTaxFreeAmtBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationTaxFreeAmtTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationTaxFreeAmtBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationTaxFreeAmtMM = clsReadSPIAValues.strSPIACommutationTaxFreeAmtMM & ib & ",Commutation Tax Free Amt " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationDateTest, clsReadSPIAValues.strSPIACommutationDateBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationDateTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationDateBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationDateMM = clsReadSPIAValues.strSPIACommutationDateMM & ib & ",Commutation Date " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationBeneTypeTest, clsReadSPIAValues.strSPIACommutationBeneTypeBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationBeneTypeTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationBeneTypeBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationBeneTypeMM = clsReadSPIAValues.strSPIACommutationBeneTypeMM & ib & ",Commutation Bene Type " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationErrorTest, clsReadSPIAValues.strSPIACommutationErrorBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationErrorTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationErrorBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationErrorMM = clsReadSPIAValues.strSPIACommutationErrorMM & ib & ",Commutation Error Message " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationInputPctTest, clsReadSPIAValues.strSPIACommutationInputPctBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationInputPctTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationInputPctBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationInputPctMM = clsReadSPIAValues.strSPIACommutationInputPctMM & ib & ",Commutation Input % " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(clsReadSPIAValues.strSPIACommutationInputAmtTest, clsReadSPIAValues.strSPIACommutationInputAmtBench, True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True

                Dim strTest As String = clsReadSPIAValues.strSPIACommutationInputAmtTest
                Dim strBench As String = clsReadSPIAValues.strSPIACommutationInputAmtBench

                Dim SplitTest = Split(strTest, "�")
                Dim SplitBench = Split(strBench, "�")

                'If the arrays don't have the same # of elements, fill in the shorter one with "0's"
                If UBound(SplitTest) < UBound(SplitBench) Then
                    iShort = UBound(SplitTest)
                    ReDim Preserve SplitTest((UBound(SplitTest) + (UBound(SplitBench) - UBound(SplitTest))))
                    For iAdd = iShort To UBound(SplitTest)
                        SplitTest(iAdd) = "0"
                    Next
                ElseIf UBound(SplitBench) < UBound(SplitTest) Then
                    iShort = UBound(SplitBench)
                    ReDim Preserve SplitBench((UBound(SplitBench) + UBound(SplitTest) - UBound(SplitBench)))
                    For iAdd = iShort To UBound(SplitBench)
                        SplitBench(iAdd) = "0"
                    Next
                End If

                For iElement = 0 To UBound(SplitBench)
                    If String.Compare(SplitTest(iElement), SplitBench(iElement), True) <> 0 Then
                        clsReadSPIAValues.strSPIACommutationInputAmtMM = clsReadSPIAValues.strSPIACommutationInputAmtMM & ib & ",Commutation Input Amt " & "," & iElement + 1 & "," & SplitBench(iElement) & "," & SplitTest(iElement) & "&"
                    End If
                Next
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIABankTest), CStr(clsReadSPIAValues.strSPIABankBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIABankMM = ib & ",Sold in Bank?," & clsReadSPIAValues.strSPIABankBench & "," & clsReadSPIAValues.strSPIABankTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAChannelTest), CStr(clsReadSPIAValues.strSPIAChannelBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAChannelMM = ib & ",Channel," & clsReadSPIAValues.strSPIAChannelBench & "," & clsReadSPIAValues.strSPIAChannelTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAPayoutRateIDTest), CStr(clsReadSPIAValues.strSPIAPayoutRateIDBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPayoutRateIDMM = ib & ",Payout Rate ID," & clsReadSPIAValues.strSPIAPayoutRateIDBench & "," & clsReadSPIAValues.strSPIAPayoutRateIDTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAPricingCodeTest), CStr(clsReadSPIAValues.strSPIAPricingCodeBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPricingCodeMM = ib & ",Pricing Code," & clsReadSPIAValues.strSPIAPricingCodeBench & "," & clsReadSPIAValues.strSPIAPricingCodeTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAPricingCodeSuffixTest), CStr(clsReadSPIAValues.strSPIAPricingCodeSuffixBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPricingCodeSuffixMM = ib & ",Pricing Code Suffix," & clsReadSPIAValues.strSPIAPricingCodeSuffixBench & "," & clsReadSPIAValues.strSPIAPricingCodeSuffixTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAPayoutRateEffectiveDateTest), CStr(clsReadSPIAValues.strSPIAPayoutRateEffectiveDateBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPayoutRateEffectiveDateMM = ib & ",Payout Rate Effec Date," & clsReadSPIAValues.strSPIAPayoutRateEffectiveDateBench & "," & clsReadSPIAValues.strSPIAPayoutRateEffectiveDateTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAProofOfBirthListTest), CStr(clsReadSPIAValues.strSPIAProofOfBirthListBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAProofOfBirthListMM = ib & ",Proof of Birth List?," & clsReadSPIAValues.strSPIAProofOfBirthListBench & "," & clsReadSPIAValues.strSPIAProofOfBirthListTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIABeneSchedLengthTypeTest), CStr(clsReadSPIAValues.strSPIABeneSchedLengthTypeBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIABeneSchedLengthTypeMM = ib & ",Bene Sched Length Type," & clsReadSPIAValues.strSPIABeneSchedLengthTypeBench & "," & clsReadSPIAValues.strSPIABeneSchedLengthTypeTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIABeneSchedInputYearsTest), CStr(clsReadSPIAValues.strSPIABeneSchedInputYearsBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIABeneSchedInputYearsMM = ib & ",Bene Sched Input Years," & clsReadSPIAValues.strSPIABeneSchedInputYearsBench & "," & clsReadSPIAValues.strSPIABeneSchedInputYearsTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIABeneSchedYNTest), CStr(clsReadSPIAValues.strSPIABeneSchedYNBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIABeneSchedYNMM = ib & ",Bene Sched?," & clsReadSPIAValues.strSPIABeneSchedYNBench & "," & clsReadSPIAValues.strSPIABeneSchedYNTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAIRRYNTest), CStr(clsReadSPIAValues.strSPIAIRRYNBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAIRRYNMM = ib & ",IRR?," & clsReadSPIAValues.strSPIAIRRYNBench & "," & clsReadSPIAValues.strSPIAIRRYNTest & "&"
            End If

            If (String.Compare(CStr(clsReadSPIAValues.strSPIAPopulationTableTest), CStr(clsReadSPIAValues.strSPIAPopulationTableBench), True)) <> 0 Then
                clsReadSPIAValues.bMisMatch = True
                clsReadSPIAValues.strSPIAPopulationTableMM = ib & ",Population Table," & clsReadSPIAValues.strSPIAPopulationTableBench & "," & clsReadSPIAValues.strSPIAPopulationTableTest & "&"
            End If

        End If
        'if cases don't match, create the mismatch reports
        If clsReadSPIAValues.bMisMatch = True Then
            clsReadSPIAValues.bMismatchAtLeastOnce = True
            SaveNewBench(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
            CreatespiaMismatchReport(ib)
            strClientMisMatchList = strClientMisMatchList & "," & ib
            strsplitMMList = Split(strClientMisMatchList, ",")
            'WriteMismatchList(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If
        ''copy the Test.pdf to the test folder under the client folder, whether there are mismatches or not
        If kclbClientList.GetItemChecked(ib - 1) Then
            CopyPDF(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp)
        End If

        'write the stats to the file
        WriteMatchStatus(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strcomp, clsReadSPIAValues.bMisMatch, bNewBench:=False)



    End Sub

    Private Sub DeleteFiles(ByVal strcomp() As String, ByVal i As Short)

        'Delete any .emf or relay files before calcing each case

        Dim di As DirectoryInfo = New DirectoryInfo("C:\WinFlex6\" & strcomp(i))
        Dim emf As FileInfo() = di.GetFiles("*.emf")
        Dim ini As FileInfo() = di.GetFiles("*.ini")
        Dim out As FileInfo() = di.GetFiles("*.out")

        For Each finext In emf
            finext.Delete()
        Next

        For Each finext In ini
            If finext.Name = "Gnawin.ini" Then
            Else
                finext.Delete()
            End If

        Next

        For Each finext In out
            finext.Delete()
        Next
    End Sub
    Public Sub WriteMatchStatusFIATool(ByVal strpath As String, ByVal strcase As String, ByVal iclient As Integer, ByVal strcomp() As String, ByVal bMisMatch As Boolean, ByVal bNewBench As Boolean, Optional ByVal note As String = "")

        'write the run stats to the file

        Dim strComplPath As String = strpath & strcase & "\" & iclient & "\"

        If FileIO.FileSystem.FileExists(strComplPath & "\fiatoolstats" & iclient & ".txt") Then
            Dim sw As System.IO.StreamWriter
            Dim sr As System.IO.StreamReader
            sr = System.IO.File.OpenText(strComplPath & "\fiatoolstats" & iclient & ".txt")
            Dim MyContents As String = sr.ReadToEnd
            sr.Close()

            sw = System.IO.File.AppendText(strComplPath & "\fiatoolstats" & iclient & ".txt")
            sw.WriteLine(vbCrLf)

            If bNewBench = True Then
                'sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                'sw.WriteLine(Today & " " & TimeOfDay & " New Benchmark Created")
                'If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                '    sw.WriteLine("The age was saved")
                'End If
                'If note <> "" Then
                '    sw.WriteLine("Notes:  " & note)
                'End If
            ElseIf bMisMatch = True Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & " WinFlex did not match FIA Tool")
                sw.WriteLine(ReadFIARelayINI.strToolVersion)
                If gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            ElseIf bMisMatch = False Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & "  WinFlex matched FIA Tool")
                sw.WriteLine(ReadFIARelayINI.strToolVersion)
                If gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            End If
            

            sw.WriteLine(clsReadVAValues.strFIAEXE)
            sw.WriteLine(clsReadVAValues.strFIACPY)
            sw.WriteLine(clsReadVAValues.strFIAMDB)
            sw.Close()
            sr.Close()
        Else
            Dim sw As System.IO.StreamWriter
            sw = System.IO.File.CreateText(strComplPath & "\fiatoolstats" & iclient & ".txt")
            If bNewBench = True Then
                'sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                'sw.WriteLine(Today & " " & TimeOfDay & " New Benchmark Created")
                'If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                '    sw.WriteLine("The age was saved")
                'End If
                'If note <> "" Then
                '    sw.WriteLine("Notes:  " & note)
                'End If
            ElseIf bMisMatch = True Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & " WinFlex did not match FIA Tool")
                sw.WriteLine(ReadFIARelayINI.strToolVersion)
                If gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            ElseIf bMisMatch = False Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & "  WinFlex matched FIA Tool")
                sw.WriteLine(ReadFIARelayINI.strToolVersion)
                If gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            End If
            sw.WriteLine(clsReadVAValues.strFIAEXE)
                sw.WriteLine(clsReadVAValues.strFIACPY)
                sw.WriteLine(clsReadVAValues.strFIAMDB)
            sw.Flush()
            sw.Close()
        End If
    End Sub
    Public Sub WriteMatchStatus(ByVal strpath As String, ByVal strcase As String, ByVal iclient As Integer, ByVal strcomp() As String, ByVal bMisMatch As Boolean, ByVal bNewBench As Boolean, Optional ByVal note As String = "")

        'write the run stats to the file

        Dim strComplPath As String = strpath & strcase & "\" & iclient & "\"

        If FileIO.FileSystem.FileExists(strComplPath & "\runstatscase" & iclient & ".txt") Then
            Dim sw As System.IO.StreamWriter
            Dim sr As System.IO.StreamReader
            sr = System.IO.File.OpenText(strComplPath & "\runstatscase" & iclient & ".txt")
            Dim MyContents As String = sr.ReadToEnd
            sr.Close()

            sw = System.IO.File.AppendText(strComplPath & "\runstatscase" & iclient & ".txt")
            sw.WriteLine(vbCrLf)

            If bNewBench = True Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & " New Benchmark Created")
                If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
                If note <> "" Then
                    sw.WriteLine("Notes:  " & note)
                End If
            ElseIf bMisMatch = True Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & " Case Did Not Match")
                If gbSPIAEffectiveDate Then
                    sw.WriteLine("An Effective Rate Date of:  " & gstrSPIARateDate & " was used")
                End If
                If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            ElseIf bMisMatch = False Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & "  Case Matched")
                If gbSPIAEffectiveDate Then
                    sw.WriteLine("An Effective Rate Date of:  " & gstrSPIARateDate & " was used")
                End If
                If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            End If
            If gstrpathProduct = "VA" Then
                If gstrCompCode(iclient) = "GELA" Then
                    sw.WriteLine(clsReadVAValues.strWFProp)
                    sw.WriteLine(clsReadVAValues.strGLAICCPY)
                    sw.WriteLine(clsReadVAValues.strAnn1)
                    sw.WriteLine(clsReadVAValues.strAnn2)
                    sw.WriteLine(clsReadVAValues.strAnn3)
                    sw.WriteLine(clsReadVAValues.strAnn4)
                    sw.WriteLine(clsReadVAValues.strWFGELA)
                ElseIf gstrCompCode(iclient) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYCPY)
                    sw.WriteLine(clsReadVAValues.strGECLRIC)
                    sw.WriteLine(clsReadVAValues.strGECLVA1)
                    sw.WriteLine(clsReadVAValues.strGECLVA2)
                    sw.WriteLine(clsReadVAValues.strGECLVA3)
                    sw.WriteLine(clsReadVAValues.strGECLEXE)
                End If
            ElseIf gstrpathProduct = "SPIA" Then
                If gstrCompCode(iclient) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICSPIANNVER)
                ElseIf gstrCompCode(iclient) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNVER)
                ElseIf gstrCompCode(iclient) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIANNVER)

                End If
            ElseIf gstrpathProduct = "SPDA" Then
                If gstrCompCode(iclient) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(iclient) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(iclient) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDARATE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDACPY)
                End If
            ElseIf gstrpathProduct = "FIA" Then
                sw.WriteLine(clsreadvavalues.strfiaexe)
                sw.WriteLine(clsreadvavalues.strfiacpy)
                sw.WriteLine(clsreadvavalues.strfiamdb)

            End If
            sw.Close()
            sr.Close()
        Else
            Dim sw As System.IO.StreamWriter
            sw = System.IO.File.CreateText(strComplPath & "\runstatscase" & iclient & ".txt")
            If bNewBench = True Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & " New Benchmark Created")
                If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
                If note <> "" Then
                    sw.WriteLine("Notes:  " & note)
                End If
            ElseIf bMisMatch = True Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & " Case Did Not Match")
                If gbSPIAEffectiveDate Then
                    sw.WriteLine("An Effective Rate Date of:  " & gstrSPIARateDate & " was used")
                End If
                If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            ElseIf bMisMatch = False Then
                sw.WriteLine(strcase & " Client # " & iclient & "  Run by:  " & username)
                sw.WriteLine(Today & " " & TimeOfDay & "  Case Matched")
                If gbSPIAEffectiveDate Then
                    sw.WriteLine("An Effective Rate Date of:  " & gstrSPIARateDate & " was used")
                End If
                If gbVASaveAge(iclient) Or gbFIASaveAge(iclient) Then
                    sw.WriteLine("The age was saved")
                End If
            End If
            If gstrpathProduct = "VA" Then
                If gstrCompCode(iclient) = "GELA" Then
                    sw.WriteLine(clsReadVAValues.strWFProp)
                    sw.WriteLine(clsReadVAValues.strGLAICCPY)
                    sw.WriteLine(clsReadVAValues.strAnn1)
                    sw.WriteLine(clsReadVAValues.strAnn2)
                    sw.WriteLine(clsReadVAValues.strAnn3)
                    sw.WriteLine(clsReadVAValues.strAnn4)
                    sw.WriteLine(clsReadVAValues.strWFGELA)
                ElseIf gstrCompCode(iclient) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYCPY)
                    sw.WriteLine(clsReadVAValues.strGECLRIC)
                    sw.WriteLine(clsReadVAValues.strGECLVA1)
                    sw.WriteLine(clsReadVAValues.strGECLVA2)
                    sw.WriteLine(clsReadVAValues.strGECLVA3)
                    sw.WriteLine(clsReadVAValues.strGECLEXE)
                End If
            ElseIf gstrpathProduct = "SPIA" Then
                If gstrCompCode(iclient) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICSPIANNVER)
                ElseIf gstrCompCode(iclient) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNVER)
                ElseIf gstrCompCode(iclient) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIANNVER)
                End If
            ElseIf gstrpathProduct = "SPDA" Then
                If gstrCompCode(iclient) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(iclient) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(iclient) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDARATE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDACPY)
                End If
            ElseIf gstrpathProduct = "FIA" Then
                sw.WriteLine(clsreadvavalues.strfiaexe)
                sw.WriteLine(clsreadvavalues.strfiacpy)
                sw.WriteLine(clsreadvavalues.strfiamdb)
            End If
            sw.Flush()
            sw.Close()
        End If
    End Sub
    Public Sub WriteBenchMarkDate(ByVal strpath As String, ByVal strcase As String, ByVal iclient As Integer)

        'write the new benchmark date to a file to be read at run time

        Dim strComplPath As String = strpath & strcase & "\" & iclient & "\"
        Dim sw As System.IO.StreamWriter

        sw = System.IO.File.CreateText(strComplPath & "\lastbenchmarkdate" & iclient & ".txt")
        sw.WriteLine(Today & " " & TimeOfDay)
        sw.Flush()
        sw.Close()
    End Sub
    Public Sub WriteMismatchList(ByVal strpath As String, ByVal strcase As String)

    End Sub
    Public Sub DeleteNewBenchFolders()

        'When exiting program, delete any 'NEW' folders created during testing

        Dim FSO As New FileSystemObject
        Dim fld As Folder

        If FSO.FolderExists(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0)) Then
            fld = FSO.GetFolder(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0))
            For Each subdi As DirectoryInfo In New DirectoryInfo(gstrpath & "\" & gstrpathProduct & "\" & fld.Name).GetDirectories()
                If InStr(subdi.Name, "New") Then
                    Directory.Delete(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & subdi.Name, True)
                End If
            Next
        End If
    End Sub
    Private Sub SaveNewBench(ByVal strpath As String, ByVal strcase As String, ByVal iClient As Integer, ByVal strcomp() As String)

        'save the run files as a potentially new benchmark, and write the status to the file

        Dim di As DirectoryInfo = New DirectoryInfo("C:\WinFlex6\" & strcomp(iClient))
        Dim emf As FileInfo() = di.GetFiles("*.emf")
        Dim ini As FileInfo() = di.GetFiles("*.ini")
        Dim out As FileInfo() = di.GetFiles("*.out")
        Dim pdf As FileInfo() = di.GetFiles("Test.pdf")
        Dim finext As FileInfo
        Dim strComplPath As String = strpath & "\" & strcase & "\" & "New" & iClient & "\"

        Dim sw As StreamWriter

        'create a new numbered directory, i.e. "New1"
        Directory.CreateDirectory(strComplPath)

        'copy the illustration files to the new directory
        For Each finext In emf
            System.IO.File.Copy("C:\WinFlex6\" & strcomp(iClient) & "\" & finext.Name, strComplPath & finext.Name, True)
        Next

        For Each finext In pdf
            System.IO.File.Copy("C:\WinFlex6\" & strcomp(iClient) & "\" & finext.Name, strComplPath & finext.Name, True)
        Next

        For Each finext In ini
            If finext.Name = "Gnawin.ini" Then
            Else
                System.IO.File.Copy("C:\WinFlex6\" & strcomp(iClient) & "\" & finext.Name, strComplPath & finext.Name, True)
            End If
        Next

        For Each finext In out
            System.IO.File.Copy("C:\WinFlex6\" & strcomp(iClient) & "\" & finext.Name, strComplPath & finext.Name, True)
        Next

        sw = System.IO.File.CreateText(strComplPath & "\runstatscase" & iClient & ".txt")
        sw.WriteLine(strcase & " Client # " & iClient & "  Run by:  " & username)
        sw.WriteLine(Today & " " & TimeOfDay & " Case Did Not Match")

        If gstrpathProduct = "VA" Then
            If gstrCompCode(iClient) = "GELA" Then
                sw.WriteLine(clsReadVAValues.strWFProp)
                sw.WriteLine(clsReadVAValues.strGLAICCPY)
                sw.WriteLine(clsReadVAValues.strAnn1)
                sw.WriteLine(clsReadVAValues.strAnn2)
                sw.WriteLine(clsReadVAValues.strAnn3)
                sw.WriteLine(clsReadVAValues.strAnn4)
                sw.WriteLine(clsReadVAValues.strWFGELA)
            ElseIf gstrCompCode(iClient) = "GECL" Then
                sw.WriteLine(clsReadVAValues.strGLICNYCPY)
                sw.WriteLine(clsReadVAValues.strGECLRIC)
                sw.WriteLine(clsReadVAValues.strGECLVA1)
                sw.WriteLine(clsReadVAValues.strGECLVA2)
                sw.WriteLine(clsReadVAValues.strGECLVA3)
                sw.WriteLine(clsReadVAValues.strGECLEXE)
            End If
        ElseIf gstrpathProduct = "SPIA" Then
            If gstrCompCode(iClient) = "GECA" Then
                sw.WriteLine(clsReadVAValues.strGLICSPIAEXE)
                sw.WriteLine(clsReadVAValues.strGLICSPIAANNRATES)
                sw.WriteLine(clsReadVAValues.strGLICSPIAANNPROD)
                sw.WriteLine(clsReadVAValues.strGLICSPIAANNSYS)
                sw.WriteLine(clsReadVAValues.strGLICSPIACPY)
                sw.WriteLine(clsReadVAValues.strGLICSPIAANNSUPP)
                sw.WriteLine(clsReadVAValues.strGLICSPIANNVER)
            ElseIf gstrCompCode(iClient) = "GECL" Then
                sw.WriteLine(clsReadVAValues.strGLICNYSPIAEXE)
                sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNRATES)
                sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNPROD)
                sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSYS)
                sw.WriteLine(clsReadVAValues.strGLICNYSPIACPY)
                sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSUPP)
                sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNVER)
            ElseIf gstrCompCode(iClient) = "FCOL" Then
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAEXE)
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNRATES)
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNPROD)
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSYS)
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIACPY)
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSUPP)
                sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIANNVER)
            End If
        ElseIf gstrpathProduct = "SPDA" Then
            If gstrCompCode(iClient) = "GECA" Then
                sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
            ElseIf gstrCompCode(iClient) = "GECL" Then
                sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
            ElseIf gstrCompCode(iClient) = "FCOL" Then
                sw.WriteLine(clsReadVAValues.strFCOLSPDAEXE)
                sw.WriteLine(clsReadVAValues.strFCOLSPDARATE)
                sw.WriteLine(clsReadVAValues.strFCOLSPDAGNAWINE)
                sw.WriteLine(clsReadVAValues.strFCOLSPDACPY)
            End If
        ElseIf gstrpathProduct = "FIA" Then
            sw.WriteLine(clsreadvavalues.strfiaexe)
            sw.WriteLine(clsreadvavalues.strfiacpy)
            sw.WriteLine(clsreadvavalues.strfiamdb)
        End If
        sw.Close()
    End Sub
    Private Sub CopyPDF(ByVal strpath As String, ByVal strcase As String, ByVal iClient As Integer, ByVal strcomp() As String)

        'save the newly run Test.PDF to compare via Paloma

        Dim di As DirectoryInfo = New DirectoryInfo("C:\WinFlex6\" & strcomp(iClient))

        Dim strComplPath As String = strpath & strcase & "\" & iClient & "\Test\"

        'copy the pdf to the Test directory under the client folder

        'For Each finext In pdf
        If gbClientDoesntRun = False Then
            System.IO.File.Copy("C:\WinFlex6\" & strcomp(iClient) & "\Test.pdf", strComplPath & "Test.pdf", True)
        End If
        'Next

    End Sub
    Private Sub FillMismatchBox(ByVal i As Integer)
    End Sub
    Public Sub CreateVAMismatchReport(ByVal ib As Integer)

        'write the mismatch strings for each variable

        Dim ix As Integer = 0

        If gbVASaveAge(ib) Then
            If gstrVASaveAgeDOB1 <> "" Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & gstrVASaveAgeDOB1
                gstrDOB1New = ""
                gstrVASaveAgeDOB1 = ""
            End If
            If gstrVASaveAgeDOB2 <> "" Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & gstrVASaveAgeDOB2
                gstrDOB2New = ""
                gstrVASaveAgeDOB2 = ""
            End If
        End If

        If Len(clsReadVAValues.strRunNoRunMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRunNoRunMM
            clsReadVAValues.strRunNoRunMM = ""
        End If


        If Len(clsReadVAValues.strMessage1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMessage1MM
            clsReadVAValues.strMessage1MM = ""
        End If

        If Len(clsReadVAValues.strMessage2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMessage2MM
            clsReadVAValues.strMessage2MM = ""
        End If

        If Len(clsReadVAValues.strMessage3MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMessage3MM
            clsReadVAValues.strMessage3MM = ""
        End If

        If Len(clsReadVAValues.strMessage4MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMessage4MM
            clsReadVAValues.strMessage4MM = ""
        End If
        If Len(clsReadVAValues.strMessage5MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMessage5MM
            clsReadVAValues.strMessage5MM = ""
        End If
        If Len(clsReadVAValues.strMessage6MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMessage6MM
            clsReadVAValues.strMessage6MM = ""
        End If

        If Len(clsReadVAValues.strCompanyNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strCompanyNameMM
            clsReadVAValues.strCompanyNameMM = ""
        End If

        If Len(clsReadVAValues.strClient1NameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strClient1NameMM
            clsReadVAValues.strClient1NameMM = ""
        End If

        If Len(clsReadVAValues.strClient2NameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strClient2NameMM
            clsReadVAValues.strClient2NameMM = ""
        End If

        If Len(clsReadVAValues.strSex1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strSex1MM
            clsReadVAValues.strSex1MM = ""
        End If

        If Len(clsReadVAValues.strSex2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strSex2MM
            clsReadVAValues.strSex2MM = ""
        End If

        If Len(clsReadVAValues.strAge1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strAge1MM
            clsReadVAValues.strAge1MM = ""
        End If

        If Len(clsReadVAValues.strAge2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strAge2MM
            clsReadVAValues.strAge2MM = ""
        End If

        If Len(clsReadVAValues.strAgeOlderMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strAgeOlderMM
            clsReadVAValues.strAgeOlderMM = ""
        End If

        If Len(clsReadVAValues.strIRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIRateMM
            clsReadVAValues.strIRateMM = ""
        End If

        If Len(clsReadVAValues.strInitialDBMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInitialDBMM
            clsReadVAValues.strInitialDBMM = ""
        End If

        If Len(clsReadVAValues.strInitialDB2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInitialDB2MM
            clsReadVAValues.strInitialDB2MM = ""
        End If

        If Len(clsReadVAValues.strJointMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strJointMM
            clsReadVAValues.strJointMM = ""
        End If

        If Len(clsReadVAValues.strContractTypeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strContractTypeMM
            clsReadVAValues.strContractTypeMM = ""
        End If

        If Len(clsReadVAValues.strSurrChargeYrsMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strSurrChargeYrsMM
            clsReadVAValues.strSurrChargeYrsMM = ""
        End If

        If Len(clsReadVAValues.strHypoNorGMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHypoNorGMM
            clsReadVAValues.strHypoNorGMM = ""
        End If

        If Len(clsReadVAValues.strZeroNetMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strZeroNetMM
            clsReadVAValues.strZeroNetMM = ""
        End If

        If Len(clsReadVAValues.strHypoNetMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHypoNetMM
            clsReadVAValues.strHypoNetMM = ""
        End If

        If Len(clsReadVAValues.strHypoGrossMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHypoGrossMM
            clsReadVAValues.strHypoGrossMM = ""
        End If

        If Len(clsReadVAValues.strHypoGISRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHypoGISRateMM
            clsReadVAValues.strHypoGISRateMM = ""
        End If

        If Len(clsReadVAValues.strZeroGISRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strZeroGISRateMM
            clsReadVAValues.strZeroGISRateMM = ""
        End If

        If Len(clsReadVAValues.strExpenseVAMandEOnlyMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strExpenseVAMandEOnlyMM
            clsReadVAValues.strExpenseVAMandEOnlyMM = ""
        End If

        If Len(clsReadVAValues.strExpensesTotalBaseContractMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strExpensesTotalBaseContractMM
            clsReadVAValues.strExpensesTotalBaseContractMM = ""
        End If

        If Len(clsReadVAValues.strZeroGrowthRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strZeroGrowthRateMM
            clsReadVAValues.strZeroGrowthRateMM = ""
        End If

        If Len(clsReadVAValues.strExpensesAdminOnlyMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strExpensesAdminOnlyMM
            clsReadVAValues.strExpensesAdminOnlyMM = ""
        End If

        If Len(clsReadVAValues.strFundExpensesVAMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundExpensesVAMM
            clsReadVAValues.strFundExpensesVAMM = ""
        End If

        If Len(clsReadVAValues.strStartWDsYoungerBDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strStartWDsYoungerBDMM
            clsReadVAValues.strStartWDsYoungerBDMM = ""
        End If

        If Len(clsReadVAValues.strFundExpenseGISMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundExpenseGISMM
            clsReadVAValues.strFundExpenseGISMM = ""
        End If

        If Len(clsReadVAValues.strFundExpenseEffDateGISMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundExpenseEffDateGISMM
            clsReadVAValues.strFundExpenseEffDateGISMM = ""
        End If

        If Len(clsReadVAValues.strVADBBenefitRiderChargeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strVADBBenefitRiderChargeMM
            clsReadVAValues.strVADBBenefitRiderChargeMM = ""
        End If

        If Len(clsReadVAValues.strVAContractChargeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strVAContractChargeMM
            clsReadVAValues.strVAContractChargeMM = ""
        End If

        If Len(clsReadVAValues.strVAContractChargeWaiverLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strVAContractChargeWaiverLimitMM
            clsReadVAValues.strVAContractChargeWaiverLimitMM = ""
        End If

        If Len(clsReadVAValues.strFundExpenseEffDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundExpenseEffDateMM
            clsReadVAValues.strFundExpenseEffDateMM = ""
        End If

        If Len(clsReadVAValues.strEarlyAccessChargeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strEarlyAccessChargeMM
            clsReadVAValues.strEarlyAccessChargeMM = ""
        End If

        If Len(clsReadVAValues.strLivingBenefitRiderChargeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLivingBenefitRiderChargeMM
            clsReadVAValues.strLivingBenefitRiderChargeMM = ""
        End If

        If Len(clsReadVAValues.strInitialPremiumMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInitialPremiumMM
            clsReadVAValues.strInitialPremiumMM = ""
        End If

        If Len(clsReadVAValues.strPrintYearsMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strPrintYearsMM
            clsReadVAValues.strPrintYearsMM = ""
        End If

        If Len(clsReadVAValues.strDBTypeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBTypeMM
            clsReadVAValues.strDBTypeMM = ""
        End If

        If Len(clsReadVAValues.strFundCountMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundCountMM
            clsReadVAValues.strFundCountMM = ""
        End If

        If Len(clsReadVAValues.strInvestStratMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInvestStratMM
            clsReadVAValues.strInvestStratMM = ""
        End If

        If Len(clsReadVAValues.strIncomeStartAgeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIncomeStartAgeMM
            clsReadVAValues.strIncomeStartAgeMM = ""
        End If

        If Len(clsReadVAValues.strIncomeStartAgeJointMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIncomeStartAgeJointMM
            clsReadVAValues.strIncomeStartAgeJointMM = ""
        End If

        If Len(clsReadVAValues.strIncomeStartYearMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIncomeStartYearMM
            clsReadVAValues.strIncomeStartYearMM = ""
        End If

        If Len(clsReadVAValues.strIncomeStartMonthMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIncomeStartMonthMM
            clsReadVAValues.strIncomeStartMonthMM = ""
        End If

        If Len(clsReadVAValues.strYearsCertainMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strYearsCertainMM
            clsReadVAValues.strYearsCertainMM = ""
        End If

        If Len(clsReadVAValues.strGIAInitialMonthlyPayoutHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strGIAInitialMonthlyPayoutHypoMM
            clsReadVAValues.strGIAInitialMonthlyPayoutHypoMM = ""
        End If

        If Len(clsReadVAValues.strGIAInitialMonthlyPayoutZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strGIAInitialMonthlyPayoutZeroMM
            clsReadVAValues.strGIAInitialMonthlyPayoutZeroMM = ""
        End If

        If Len(clsReadVAValues.strGIASchedInstallmentMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strGIASchedInstallmentMM
            clsReadVAValues.strGIASchedInstallmentMM = ""
        End If

        If Len(clsReadVAValues.strAnnAmountZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strAnnAmountZeroMM
            clsReadVAValues.strAnnAmountZeroMM = ""
        End If

        If Len(clsReadVAValues.strAnnAmountHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strAnnAmountHypoMM
            clsReadVAValues.strAnnAmountHypoMM = ""
        End If

        If Len(clsReadVAValues.strInstallmentCountZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInstallmentCountZeroMM
            clsReadVAValues.strInstallmentCountZeroMM = ""
        End If

        If Len(clsReadVAValues.strInstallmentCountHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInstallmentCountHypoMM
            clsReadVAValues.strInstallmentCountHypoMM = ""
        End If

        If Len(clsReadVAValues.strPPDBChargeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strPPDBChargeMM
            clsReadVAValues.strPPDBChargeMM = ""
        End If

        If Len(clsReadVAValues.strLIPFactorFirstWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPFactorFirstWDMM
            clsReadVAValues.strLIPFactorFirstWDMM = ""
        End If

        If Len(clsReadVAValues.strLIPGuarWithdrawalMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPGuarWithdrawalMM
            clsReadVAValues.strLIPGuarWithdrawalMM = ""
        End If

        If Len(clsReadVAValues.strLIPWDStartYearMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPWDStartYearMM
            clsReadVAValues.strLIPWDStartYearMM = ""
        End If

        If Len(clsReadVAValues.strLIPWDStartMonthMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPWDStartMonthMM
            clsReadVAValues.strLIPWDStartMonthMM = ""
        End If

        If Len(clsReadVAValues.strLIPBenBaseFirstWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPBenBaseFirstWDMM
            clsReadVAValues.strLIPBenBaseFirstWDMM = ""
        End If

        If Len(clsReadVAValues.strTaxExcludableAmtZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTaxExcludableAmtZeroMM
            clsReadVAValues.strTaxExcludableAmtZeroMM = ""
        End If

        If Len(clsReadVAValues.strTaxExcludableAmtHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTaxExcludableAmtHypoMM
            clsReadVAValues.strTaxExcludableAmtHypoMM = ""
        End If

        If Len(clsReadVAValues.strTaxExcludableAmtHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTaxExcludableAmtHistMM
            clsReadVAValues.strTaxExcludableAmtHistMM = ""
        End If

        If Len(clsReadVAValues.strTaxBracketMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTaxBracketMM
            clsReadVAValues.strTaxBracketMM = ""
        End If

        If Len(clsReadVAValues.strTaxBasisMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTaxBasisMM
            clsReadVAValues.strTaxBasisMM = ""
        End If

        If Len(clsReadVAValues.strInvestmentMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strInvestmentMM
            clsReadVAValues.strInvestmentMM = ""
        End If

        If Len(clsReadVAValues.strBaseContractValueZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strBaseContractValueZeroMM
            clsReadVAValues.strBaseContractValueZeroMM = ""
        End If

        If Len(clsReadVAValues.strCombinedSurrValueZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strCombinedSurrValueZeroMM
            clsReadVAValues.strCombinedSurrValueZeroMM = ""
        End If

        If Len(clsReadVAValues.strCombinedSurrValueHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strCombinedSurrValueHypoMM
            clsReadVAValues.strCombinedSurrValueHypoMM = ""
        End If

        If Len(clsReadVAValues.strGISValueZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strGISValueZeroMM
            clsReadVAValues.strGISValueZeroMM = ""
        End If

        If Len(clsReadVAValues.strAnnualIncomeZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strAnnualIncomeZeroMM
            clsReadVAValues.strAnnualIncomeZeroMM = ""
        End If

        If Len(clsReadVAValues.strDeathBenefitZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDeathBenefitZeroMM
            clsReadVAValues.strDeathBenefitZeroMM = ""
        End If

        If Len(clsReadVAValues.strTransfertoGISZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTransfertoGISZeroMM
            clsReadVAValues.strTransfertoGISZeroMM = ""
        End If

        If Len(clsReadVAValues.strTotalContractValueZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strTotalContractValueZeroMM
            clsReadVAValues.strTotalContractValueZeroMM = ""
        End If

        If Len(clsReadVAValues.strSurrenderChargesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strSurrenderChargesMM
            clsReadVAValues.strSurrenderChargesMM = ""
        End If

        If Len(clsReadVAValues.strHypoAnnIncomeFloorMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHypoAnnIncomeFloorMM
            clsReadVAValues.strHypoAnnIncomeFloorMM = ""
        End If

        If Len(clsReadVAValues.strGIASumGtdAmtMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strGIASumGtdAmtMM
            clsReadVAValues.strGIASumGtdAmtMM = ""
        End If

        If Len(clsReadVAValues.strHistReturnForPeriodMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistReturnForPeriodMM
            clsReadVAValues.strHistReturnForPeriodMM = ""
        End If

        If Len(clsReadVAValues.strHistAnnIncomeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistAnnIncomeMM
            clsReadVAValues.strHistAnnIncomeMM = ""
        End If

        If Len(clsReadVAValues.strHistAccountGISMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistAccountGISMM
            clsReadVAValues.strHistAccountGISMM = ""
        End If

        If Len(clsReadVAValues.strIPRHistTotalContractValueCurrMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistTotalContractValueCurrMM
            clsReadVAValues.strIPRHistTotalContractValueCurrMM = ""
        End If

        If Len(clsReadVAValues.strIPRHistTotalSurrValueCurrMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistTotalSurrValueCurrMM
            clsReadVAValues.strIPRHistTotalSurrValueCurrMM = ""
        End If

        If Len(clsReadVAValues.strIPRHistTotalContractValueMaxMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistTotalContractValueMaxMM
            clsReadVAValues.strIPRHistTotalContractValueMaxMM = ""
        End If

        If Len(clsReadVAValues.strIPRHistTotalSurrValueMaxMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistTotalSurrValueMaxMM
            clsReadVAValues.strIPRHistTotalSurrValueMaxMM = ""
        End If

        If Len(clsReadVAValues.strHistDeathBenefitGIAMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistDeathBenefitGIAMM
            clsReadVAValues.strHistDeathBenefitGIAMM = ""
        End If

        If Len(clsReadVAValues.strPPBAHypoRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strPPBAHypoRateMM
            clsReadVAValues.strPPBAHypoRateMM = ""
        End If

        If Len(clsReadVAValues.strPPBAZeroRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strPPBAZeroRateMM
            clsReadVAValues.strPPBAZeroRateMM = ""
        End If

        If Len(clsReadVAValues.strBenefitBaseZeroRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strBenefitBaseZeroRateMM
            clsReadVAValues.strBenefitBaseZeroRateMM = ""
        End If

        If Len(clsReadVAValues.strBenefitBaseHypoRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strBenefitBaseHypoRateMM
            clsReadVAValues.strBenefitBaseHypoRateMM = ""
        End If

        If Len(clsReadVAValues.strRollupZeroRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRollupZeroRateMM
            clsReadVAValues.strRollupZeroRateMM = ""
        End If

        If Len(clsReadVAValues.strRollupHypoRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRollupHypoRateMM
            clsReadVAValues.strRollupHypoRateMM = ""
        End If

        If Len(clsReadVAValues.strLIPAnnIncomeZeroRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPAnnIncomeZeroRateMM
            clsReadVAValues.strLIPAnnIncomeZeroRateMM = ""
        End If

        If Len(clsReadVAValues.strLIPAnnIncomeHypoRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPAnnIncomeHypoRateMM
            clsReadVAValues.strLIPAnnIncomeHypoRateMM = ""
        End If

        If Len(clsReadVAValues.strLIPResetValueZeroRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPResetValueZeroRateMM
            clsReadVAValues.strLIPResetValueZeroRateMM = ""
        End If

        If Len(clsReadVAValues.strLIPContractValueHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strLIPContractValueHypoMM
            clsReadVAValues.strLIPContractValueHypoMM = ""
        End If

        If Len(clsReadVAValues.strBASEWithdrawalZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strBASEWithdrawalZeroMM
            clsReadVAValues.strBASEWithdrawalZeroMM = ""
        End If

        If Len(clsReadVAValues.strBASEWithdrawalHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strBASEWithdrawalHypoMM
            clsReadVAValues.strBASEWithdrawalHypoMM = ""
        End If

        If Len(clsReadVAValues.strEPRZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strEPRZeroMM
            clsReadVAValues.strEPRZeroMM = ""
        End If

        If Len(clsReadVAValues.strEPRHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strEPRHypoMM
            clsReadVAValues.strEPRHypoMM = ""
        End If

        If Len(clsReadVAValues.strEPRHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strEPRHistMM
            clsReadVAValues.strEPRHistMM = ""
        End If

        If Len(clsReadVAValues.strDBGIAZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBGIAZeroMM
            clsReadVAValues.strDBGIAZeroMM = ""
        End If

        If Len(clsReadVAValues.strDBGIAHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBGIAHypoMM
            clsReadVAValues.strDBGIAHypoMM = ""
        End If

        If Len(clsReadVAValues.strDBLIPZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBLIPZeroMM
            clsReadVAValues.strDBLIPZeroMM = ""
        End If

        If Len(clsReadVAValues.strDBLIPHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBLIPHypoMM
            clsReadVAValues.strDBLIPHypoMM = ""
        End If

        If Len(clsReadVAValues.strDBLIPHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBLIPHistMM
            clsReadVAValues.strDBLIPHistMM = ""
        End If

        If Len(clsReadVAValues.strDBComboZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBComboZeroMM
            clsReadVAValues.strDBComboZeroMM = ""
        End If

        If Len(clsReadVAValues.strDBComboHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBComboHypoMM
            clsReadVAValues.strDBComboHypoMM = ""
        End If

        If Len(clsReadVAValues.strDBComboHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBComboHistMM
            clsReadVAValues.strDBComboHistMM = ""
        End If

        If Len(clsReadVAValues.strDBASDBZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBASDBZeroMM
            clsReadVAValues.strDBASDBZeroMM = ""
        End If

        If Len(clsReadVAValues.strDBASDBHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBASDBHypoMM
            clsReadVAValues.strDBASDBHypoMM = ""
        End If

        If Len(clsReadVAValues.strDBASDBHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBASDBHistMM
            clsReadVAValues.strDBASDBHistMM = ""
        End If

        If Len(clsReadVAValues.strDBRollupZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBRollupZeroMM
            clsReadVAValues.strDBRollupZeroMM = ""
        End If

        If Len(clsReadVAValues.strDBRollupHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBRollupHypoMM
            clsReadVAValues.strDBRollupHypoMM = ""
        End If

        If Len(clsReadVAValues.strDBRollupHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBRollupHistMM
            clsReadVAValues.strDBRollupHistMM = ""
        End If

        If Len(clsReadVAValues.strDBStandardZeroMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBStandardZeroMM
            clsReadVAValues.strDBStandardZeroMM = ""
        End If

        If Len(clsReadVAValues.strDBStandardHypoMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBStandardHypoMM
            clsReadVAValues.strDBStandardHypoMM = ""
        End If

        If Len(clsReadVAValues.strDBStandardHistMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBStandardHistMM
            clsReadVAValues.strDBStandardHistMM = ""
        End If

        For ix = 1 To clsReadVAValues.strFundCountTest


            If Len(clsReadVAValues.strFundCodeMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundCodeMM(ix)
                clsReadVAValues.strFundCodeMM(ix) = ""
            End If

            If Len(clsReadVAValues.strFundPctMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundPctMM(ix)
                clsReadVAValues.strFundPctMM(ix) = ""
            End If

            If Len(clsReadVAValues.strFundNameMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strFundNameMM(ix)
                clsReadVAValues.strFundNameMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr1StdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr1StdMM(ix)
                clsReadVAValues.strReturnYr1StdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDatestdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDatestdMM(ix)
                clsReadVAValues.strReturnAdoptionDatestdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr5StdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr5StdMM(ix)
                clsReadVAValues.strReturnYr5StdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr10StdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr10StdMM(ix)
                clsReadVAValues.strReturnYr10StdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionStdMM(ix)
                clsReadVAValues.strReturnAdoptionStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDatestdIGAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDatestdIGAMM(ix)
                clsReadVAValues.strReturnAdoptionDatestdIGAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr1StdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr1StdGIAMM(ix)
                clsReadVAValues.strReturnYr1StdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr5StdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr5StdGIAMM(ix)
                clsReadVAValues.strReturnYr5StdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr10StdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr10StdGIAMM(ix)
                clsReadVAValues.strReturnYr10StdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionStdGIAMM(ix)
                clsReadVAValues.strReturnAdoptionStdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDatestdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDatestdGIAMM(ix)
                clsReadVAValues.strReturnAdoptionDatestdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr1NonStdSCMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr1NonStdSCMM(ix)
                clsReadVAValues.strReturnYr1NonStdSCMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr5NonStdSCMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr5NonStdSCMM(ix)
                clsReadVAValues.strReturnYr5NonStdSCMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr10NonStdSCMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr10NonStdSCMM(ix)
                clsReadVAValues.strReturnYr10NonStdSCMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionNonStdSCMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionNonStdSCMM(ix)
                clsReadVAValues.strReturnAdoptionNonStdSCMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr1NonStdSCGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr1NonStdSCGIAMM(ix)
                clsReadVAValues.strReturnYr1NonStdSCGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr5NonStdSCGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr5NonStdSCGIAMM(ix)
                clsReadVAValues.strReturnYr5NonStdSCGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr10NonStdSCGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr10NonStdSCGIAMM(ix)
                clsReadVAValues.strReturnYr10NonStdSCGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionNonStdSCGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionNonStdSCGIAMM(ix)
                clsReadVAValues.strReturnAdoptionNonStdSCGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr1NonStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr1NonStdMM(ix)
                clsReadVAValues.strReturnYr1NonStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr5NonStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr5NonStdMM(ix)
                clsReadVAValues.strReturnYr5NonStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr10NonStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr10NonStdMM(ix)
                clsReadVAValues.strReturnYr10NonStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionNonStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionNonStdMM(ix)
                clsReadVAValues.strReturnAdoptionNonStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr1NonStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr1NonStdGIAMM(ix)
                clsReadVAValues.strReturnYr1NonStdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr5NonStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr5NonStdGIAMM(ix)
                clsReadVAValues.strReturnYr5NonStdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnYr10NonStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnYr10NonStdGIAMM(ix)
                clsReadVAValues.strReturnYr10NonStdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionNonStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionNonStdGIAMM(ix)
                clsReadVAValues.strReturnAdoptionNonStdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix)
                clsReadVAValues.strReturnAdoptionDateNonStdSCMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnInceptionDateNonStdSCMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnInceptionDateNonStdSCMM(ix)
                clsReadVAValues.strReturnInceptionDateNonStdSCMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDateNonStdSCGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDateNonStdSCGIAMM(ix)
                clsReadVAValues.strReturnAdoptionDateNonStdSCGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnInceptionDateNonStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnInceptionDateNonStdMM(ix)
                clsReadVAValues.strReturnInceptionDateNonStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnInceptionDateNonStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnInceptionDateNonStdGIAMM(ix)
                clsReadVAValues.strReturnInceptionDateNonStdGIAMM(ix) = ""
            End If

            If Len(clsReadVAValues.strHistPeriodEndingMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistPeriodEndingMM(ix)
                clsReadVAValues.strHistPeriodEndingMM(ix) = ""
            End If

            If Len(clsReadVAValues.strHistCumulativeReturnMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistCumulativeReturnMM(ix)
                clsReadVAValues.strHistCumulativeReturnMM(ix) = ""
            End If

            If Len(clsReadVAValues.strHistAverageAnnReturnMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistAverageAnnReturnMM(ix)
                clsReadVAValues.strHistAverageAnnReturnMM(ix) = ""
            End If

            If Len(clsReadVAValues.strHistCumulativeReturnMaxMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistCumulativeReturnMaxMM(ix)
                clsReadVAValues.strHistCumulativeReturnMaxMM(ix) = ""
            End If

            If Len(clsReadVAValues.strHistAverageAnnReturnMaxMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strHistAverageAnnReturnMaxMM(ix)
                clsReadVAValues.strHistAverageAnnReturnMaxMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDateNonStdMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDateNonStdMM(ix)
                clsReadVAValues.strReturnAdoptionDateNonStdMM(ix) = ""
            End If

            If Len(clsReadVAValues.strReturnAdoptionDateNonStdGIAMM(ix)) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strReturnAdoptionDateNonStdGIAMM(ix)
                clsReadVAValues.strReturnAdoptionDateNonStdGIAMM(ix) = ""
            End If

            'IPR Hist DB Values
            If Len(clsReadVAValues.strDBIPRHistContractDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistContractDBCurrMM
                clsReadVAValues.strDBIPRHistContractDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistASDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistASDBCurrMM
                clsReadVAValues.strDBIPRHistASDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistBasicDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistBasicDBCurrMM
                clsReadVAValues.strDBIPRHistBasicDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistRollUpDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistRollUpDBCurrMM
                clsReadVAValues.strDBIPRHistRollUpDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistEPRDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistEPRDBCurrMM
                clsReadVAValues.strDBIPRHistEPRDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistPPDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistPPDBCurrMM
                clsReadVAValues.strDBIPRHistPPDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistContractDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistContractDBMaxMM
                clsReadVAValues.strDBIPRHistContractDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistASDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistASDBMaxMM
                clsReadVAValues.strDBIPRHistASDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistBasicDBmaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistBasicDBmaxMM
                clsReadVAValues.strDBIPRHistBasicDBmaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistRollUpDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistRollUpDBMaxMM
                clsReadVAValues.strDBIPRHistRollUpDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistEPRDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistEPRDBMaxMM
                clsReadVAValues.strDBIPRHistEPRDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHistPPDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHistPPDBMaxMM
                clsReadVAValues.strDBIPRHistPPDBMaxMM = ""
            End If

            'IPR WD Limits
            If Len(clsReadVAValues.strIPRWDLimitCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitCurrMM
                clsReadVAValues.strIPRWDLimitCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitGtdMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitGtdMM
                clsReadVAValues.strIPRWDLimitGtdMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitHistMM
                clsReadVAValues.strIPRWDLimitHistMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitGtdGCMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitGtdGCMM
                clsReadVAValues.strIPRWDLimitGtdGCMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitHistGCMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitHistGCMM
                clsReadVAValues.strIPRWDLimitHistGCMM = ""
            End If


            'IPR Hypo DB Values
            If Len(clsReadVAValues.strDBIPRHypoContractDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoContractDBCurrMM
                clsReadVAValues.strDBIPRHypoContractDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoASDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoASDBCurrMM
                clsReadVAValues.strDBIPRHypoASDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoBasicDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoBasicDBCurrMM
                clsReadVAValues.strDBIPRHypoBasicDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoRollUpDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoRollUpDBCurrMM
                clsReadVAValues.strDBIPRHypoRollUpDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoEPRDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoEPRDBCurrMM
                clsReadVAValues.strDBIPRHypoEPRDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoPPDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoPPDBCurrMM
                clsReadVAValues.strDBIPRHypoPPDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoContractDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoContractDBMaxMM
                clsReadVAValues.strDBIPRHypoContractDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoASDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoASDBMaxMM
                clsReadVAValues.strDBIPRHypoASDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoBasicDBmaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoBasicDBmaxMM
                clsReadVAValues.strDBIPRHypoBasicDBmaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoRollUpDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoRollUpDBMaxMM
                clsReadVAValues.strDBIPRHypoRollUpDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoEPRDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoEPRDBMaxMM
                clsReadVAValues.strDBIPRHypoEPRDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strDBIPRHypoPPDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strDBIPRHypoPPDBMaxMM
                clsReadVAValues.strDBIPRHypoPPDBMaxMM = ""
            End If

            'IPR WD Limits
            If Len(clsReadVAValues.strIPRWDLimitCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitCurrMM
                clsReadVAValues.strIPRWDLimitCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitGtdMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitGtdMM
                clsReadVAValues.strIPRWDLimitGtdMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitHistMM
                clsReadVAValues.strIPRWDLimitHistMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitGtdGCMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitGtdGCMM
                clsReadVAValues.strIPRWDLimitGtdGCMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDLimitHistGCMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDLimitHistGCMM
                clsReadVAValues.strIPRWDLimitHistGCMM = ""
            End If

            'RR One Return for Period
            If Len(clsReadVAValues.strRROneHistReturnForPeriodCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRROneHistReturnForPeriodCurrMM
                clsReadVAValues.strRROneHistReturnForPeriodCurrMM = ""
            End If

            If Len(clsReadVAValues.strRROneHistReturnForPeriodMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRROneHistReturnForPeriodMaxMM
                clsReadVAValues.strRROneHistReturnForPeriodMaxMM = ""
            End If

            'IPR Hypo Values
            If Len(clsReadVAValues.strIPRHypoPPBAMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoPPBAMaxMM
                clsReadVAValues.strIPRHypoPPBAMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoMaxAnnValueMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoMaxAnnValueMaxMM
                clsReadVAValues.strIPRHypoMaxAnnValueMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoRollupValueMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoRollupValueMaxMM
                clsReadVAValues.strIPRHypoRollupValueMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoBenefitBaseMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoBenefitBaseMaxMM
                clsReadVAValues.strIPRHypoBenefitBaseMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoDBMaxMM
                clsReadVAValues.strIPRHypoDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoPPBACurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoPPBACurrMM
                clsReadVAValues.strIPRHypoPPBACurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoMaxAnnValueCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoMaxAnnValueCurrMM
                clsReadVAValues.strIPRHypoMaxAnnValueCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoRollupValueCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoRollupValueCurrMM
                clsReadVAValues.strIPRHypoRollupValueCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoBenefitBaseCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoBenefitBaseCurrMM
                clsReadVAValues.strIPRHypoBenefitBaseCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoDBCurrMM
                clsReadVAValues.strIPRHypoDBCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoContractValueCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoContractValueCurrMM
                clsReadVAValues.strIPRHypoContractValueCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoSurrenderValueCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoSurrenderValueCurrMM
                clsReadVAValues.strIPRHypoSurrenderValueCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoContractValueMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoContractValueMaxMM
                clsReadVAValues.strIPRHypoContractValueMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHypoSurrenderValueMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHypoSurrenderValueMaxMM
                clsReadVAValues.strIPRHypoSurrenderValueMaxMM = ""
            End If

            'IPR Hist Values
            If Len(clsReadVAValues.strIPRHistPPBAMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistPPBAMaxMM
                clsReadVAValues.strIPRHistPPBAMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistMaxAnnValueMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistMaxAnnValueMaxMM
                clsReadVAValues.strIPRHistMaxAnnValueMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistRollupValueMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistRollupValueMaxMM
                clsReadVAValues.strIPRHistRollupValueMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistBenefitBaseMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistBenefitBaseMaxMM
                clsReadVAValues.strIPRHistBenefitBaseMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistDBMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistDBMaxMM
                clsReadVAValues.strIPRHistDBMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistPPBACurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistPPBACurrMM
                clsReadVAValues.strIPRHistPPBACurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistMaxAnnValueCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistMaxAnnValueCurrMM
                clsReadVAValues.strIPRHistMaxAnnValueCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistRollupValueCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistRollupValueCurrMM
                clsReadVAValues.strIPRHistRollupValueCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistBenefitBaseCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistBenefitBaseCurrMM
                clsReadVAValues.strIPRHistBenefitBaseCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRHistDBCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRHistDBCurrMM
                clsReadVAValues.strIPRHistDBCurrMM = ""
            End If

            'IPR WD Taken
            If Len(clsReadVAValues.strIPRWDTakenHypoCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDTakenHypoCurrMM
                clsReadVAValues.strIPRWDTakenHypoCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDTakenHypoMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDTakenHypoMaxMM
                clsReadVAValues.strIPRWDTakenHypoMaxMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDTakenHistCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDTakenHistCurrMM
                clsReadVAValues.strIPRWDTakenHistCurrMM = ""
            End If

            If Len(clsReadVAValues.strIPRWDTakenHistMaxMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strIPRWDTakenHistMaxMM
                clsReadVAValues.strIPRWDTakenHistMaxMM = ""
            End If

            'RR One Annual Income
            If Len(clsReadVAValues.strRROneAnnualIncomeZeroMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRROneAnnualIncomeZeroMM
                clsReadVAValues.strRROneAnnualIncomeZeroMM = ""
            End If

            If Len(clsReadVAValues.strRROneAnnualIncomeCurrMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRROneAnnualIncomeCurrMM
                clsReadVAValues.strRROneAnnualIncomeCurrMM = ""
            End If

            If Len(clsReadVAValues.strRROneAnnualIncomeHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strRROneAnnualIncomeHistMM
                clsReadVAValues.strRROneAnnualIncomeHistMM = ""
            End If

            'For MCC
            'MCC Gtd Payment Floor
            If Len(clsReadVAValues.strMCCGtdPaymentFloorZeroMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdPaymentFloorZeroMM
                clsReadVAValues.strMCCGtdPaymentFloorZeroMM = ""
            End If

            If Len(clsReadVAValues.strMCCGtdPaymentFloorHypoMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdPaymentFloorHypoMM
                clsReadVAValues.strMCCGtdPaymentFloorHypoMM = ""
            End If

            If Len(clsReadVAValues.strMCCGtdPaymentFloorHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdPaymentFloorHistMM
                clsReadVAValues.strMCCGtdPaymentFloorHistMM = ""
            End If

            'MCC Payment Floor Factor
            If Len(clsReadVAValues.strMCCGtdPaymentFloorFactorZeroMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdPaymentFloorFactorZeroMM
                clsReadVAValues.strMCCGtdPaymentFloorFactorZeroMM = ""
            End If

            If Len(clsReadVAValues.strMCCGtdPaymentFloorFactorHypoMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdPaymentFloorFactorHypoMM
                clsReadVAValues.strMCCGtdPaymentFloorFactorHypoMM = ""
            End If

            If Len(clsReadVAValues.strMCCGtdPaymentFloorFactorHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdPaymentFloorFactorHistMM
                clsReadVAValues.strMCCGtdPaymentFloorFactorHistMM = ""
            End If

            'MCC First Full Yr Income
            If Len(clsReadVAValues.strMCCFirstFullYrIncomeZeroMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCFirstFullYrIncomeZeroMM
                clsReadVAValues.strMCCFirstFullYrIncomeZeroMM = ""
            End If

            If Len(clsReadVAValues.strMCCFirstFullYrIncomeHypoMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCFirstFullYrIncomeHypoMM
                clsReadVAValues.strMCCFirstFullYrIncomeHypoMM = ""
            End If

            If Len(clsReadVAValues.strMCCFirstFullYrIncomeHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCFirstFullYrIncomeHistMM
                clsReadVAValues.strMCCFirstFullYrIncomeHistMM = ""
            End If

            'MCC Life with Period Certain Of
            If Len(clsReadVAValues.strMCCLifeWithPeriodCertainOfMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCLifeWithPeriodCertainOfMM
                clsReadVAValues.strMCCLifeWithPeriodCertainOfMM = ""
            End If

            'MCC Guarantee From Plan
            If Len(clsReadVAValues.strMCCGuaranteeFromPlanMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGuaranteeFromPlanMM
                clsReadVAValues.strMCCGuaranteeFromPlanMM = ""
            End If

            'MCC Desired Retirement Age
            If Len(clsReadVAValues.strMCCDesiredRetirementAgeMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCDesiredRetirementAgeMM
                clsReadVAValues.strMCCDesiredRetirementAgeMM = ""
            End If

            'MCC Gtd Income Payments at 0% (chart)
            If Len(clsReadVAValues.strMCCGtdIncomePayments0MM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCGtdIncomePayments0MM
                clsReadVAValues.strMCCGtdIncomePayments0MM = ""
            End If

            'MCC First Full Yr Income Hypo Net (chart)
            If Len(clsReadVAValues.strMCCFirstFullYrIncomeHypoNetMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCFirstFullYrIncomeHypoNetMM
                clsReadVAValues.strMCCFirstFullYrIncomeHypoNetMM = ""
            End If

            'MCC Total Gtd Income Payments at 0% (chart)
            If Len(clsReadVAValues.strMCCTotalGtdIncomePayments0MM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCTotalGtdIncomePayments0MM
                clsReadVAValues.strMCCTotalGtdIncomePayments0MM = ""
            End If

            'MCC Total Income Payments Hypo Net (chart)
            If Len(clsReadVAValues.strMCCTotalIncomePaymentsHypoNetMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCTotalGtdIncomePayments0MM
                clsReadVAValues.strMCCTotalIncomePaymentsHypoNetMM = ""
            End If

            'MCC Adjustment Account
            If Len(clsReadVAValues.strMCCAdjustmentAccountZeroMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCAdjustmentAccountZeroMM
                clsReadVAValues.strMCCAdjustmentAccountZeroMM = ""
            End If

            If Len(clsReadVAValues.strMCCAdjustmentAccountHypoMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCAdjustmentAccountHypoMM
                clsReadVAValues.strMCCAdjustmentAccountHypoMM = ""
            End If

            If Len(clsReadVAValues.strMCCAdjustmentAccountHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCAdjustmentAccountHistMM
                clsReadVAValues.strMCCAdjustmentAccountHistMM = ""
            End If

            'MCC Commutation Value
            If Len(clsReadVAValues.strMCCcommutationValueZeroMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCcommutationValueZeroMM
                clsReadVAValues.strMCCcommutationValueZeroMM = ""
            End If

            If Len(clsReadVAValues.strMCCCommutationValueHypoMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCCommutationValueHypoMM
                clsReadVAValues.strMCCCommutationValueHypoMM = ""
            End If

            If Len(clsReadVAValues.strMCCCommutationValueHistMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCCommutationValueHistMM
                clsReadVAValues.strMCCCommutationValueHistMM = ""
            End If

            'MCC Hist INCOME Return
            If Len(clsReadVAValues.strMCCHistIncomePeriodReturnMM) > 0 Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadVAValues.strMCCHistIncomePeriodReturnMM
                clsReadVAValues.strMCCHistIncomePeriodReturnMM = ""
            End If


        Next ix
    End Sub

    Public Sub CreateSPIAMismatchReport(ByVal ib As Integer)

        'write the mismatch strings for each variable

        Dim ix As Integer = 0


        If Len(clsReadSPIAValues.strRunNoRunMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strRunNoRunMM
            clsReadSPIAValues.strRunNoRunMM = ""
        End If


        If Len(clsReadSPIAValues.strMessage1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage1MM
            clsReadSPIAValues.strMessage1MM = ""
        End If

        If Len(clsReadSPIAValues.strMessage2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage2MM
            clsReadSPIAValues.strMessage2MM = ""
        End If

        If Len(clsReadSPIAValues.strMessage3MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage3MM
            clsReadSPIAValues.strMessage3MM = ""
        End If

        If Len(clsReadSPIAValues.strMessage4MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage4MM
            clsReadSPIAValues.strMessage4MM = ""
        End If

        If Len(clsReadSPIAValues.strMessage5MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage5MM
            clsReadSPIAValues.strMessage5MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage6MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage6MM
            clsReadSPIAValues.strMessage6MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage7MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage7MM
            clsReadSPIAValues.strMessage7MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage8MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage8MM
            clsReadSPIAValues.strMessage8MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage9MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage9MM
            clsReadSPIAValues.strMessage9MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage10MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage10MM
            clsReadSPIAValues.strMessage10MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage11MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage11MM
            clsReadSPIAValues.strMessage11MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage12MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage12MM
            clsReadSPIAValues.strMessage12MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage13MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage13MM
            clsReadSPIAValues.strMessage13MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage14MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage14MM
            clsReadSPIAValues.strMessage14MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage15MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage15MM
            clsReadSPIAValues.strMessage15MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage16MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage16MM
            clsReadSPIAValues.strMessage16MM = ""
        End If

        If Len(clsReadSPIAValues.strMessage17MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage17MM
            clsReadSPIAValues.strMessage17MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage18MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage18MM
            clsReadSPIAValues.strMessage18MM = ""
        End If

        If Len(clsReadSPIAValues.strMessage19MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage19MM
            clsReadSPIAValues.strMessage19MM = ""
        End If
        If Len(clsReadSPIAValues.strMessage20MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strMessage20MM
            clsReadSPIAValues.strMessage20MM = ""
        End If


        If Len(clsReadSPIAValues.strSPIAstreamcountMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAstreamcountMM
            clsReadSPIAValues.strSPIAstreamcountMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACompanyNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACompanyNameMM
            clsReadSPIAValues.strSPIACompanyNameMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAClient1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAClient1MM
            clsReadSPIAValues.strSPIAClient1MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAAge1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAAge1MM
            clsReadSPIAValues.strSPIAAge1MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIASex1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIASex1MM
            clsReadSPIAValues.strSPIASex1MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIADOB1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIADOB1MM
            clsReadSPIAValues.strSPIADOB1MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAClient2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAClient2MM
            clsReadSPIAValues.strSPIAClient2MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAAge2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAAge2MM
            clsReadSPIAValues.strSPIAAge2MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIASex2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIASex2MM
            clsReadSPIAValues.strSPIASex2MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIADOB2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIADOB2MM
            clsReadSPIAValues.strSPIADOB2MM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACompanyNameLongMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACompanyNameLongMM
            clsReadSPIAValues.strSPIACompanyNameLongMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAProductNameLongMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAProductNameLongMM
            clsReadSPIAValues.strSPIAProductNameLongMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAProdNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAProdNameMM
            clsReadSPIAValues.strSPIAProdNameMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIARatePricingCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIARatePricingCodeMM
            clsReadSPIAValues.strSPIARatePricingCodeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIASystemVersionMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIASystemVersionMM
            clsReadSPIAValues.strSPIASystemVersionMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPayoutRateCodesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPayoutRateCodesMM
            clsReadSPIAValues.strSPIAPayoutRateCodesMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAAgentMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAAgentMM
            clsReadSPIAValues.strSPIAAgentMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAHOApprovalAmtMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAHOApprovalAmtMM
            clsReadSPIAValues.strSPIAHOApprovalAmtMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPolicyFeeThresholdMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPolicyFeeThresholdMM
            clsReadSPIAValues.strSPIAPolicyFeeThresholdMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPolicyFeeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPolicyFeeMM
            clsReadSPIAValues.strSPIAPolicyFeeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommissionLoadMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommissionLoadMM
            clsReadSPIAValues.strSPIACommissionLoadMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAGenerationCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAGenerationCodeMM
            clsReadSPIAValues.strSPIAGenerationCodeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAGuarMortalityCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAGuarMortalityCodeMM
            clsReadSPIAValues.strSPIAGuarMortalityCodeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAAggregationAmtMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAAggregationAmtMM
            clsReadSPIAValues.strSPIAAggregationAmtMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIARestrictionsEndorsementMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIARestrictionsEndorsementMM
            clsReadSPIAValues.strSPIARestrictionsEndorsementMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAGroupCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAGroupCodeMM
            clsReadSPIAValues.strSPIAGroupCodeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAAdvanceEligibleMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAAdvanceEligibleMM(ix)
            clsReadSPIAValues.strSPIAAdvanceEligibleMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPremiumMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPremiumMM(ix)
            clsReadSPIAValues.strSPIAPremiumMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAInitBenefitMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAInitBenefitMM(ix)
            clsReadSPIAValues.strSPIAInitBenefitMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAGuaranteedYrsMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAGuaranteedYrsMM(ix)
            clsReadSPIAValues.strSPIAGuaranteedYrsMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAGuaranteedMthsMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAGuaranteedMthsMM(ix)
            clsReadSPIAValues.strSPIAGuaranteedMthsMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAIRRMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAIRRMM(ix)
            clsReadSPIAValues.strSPIAIRRMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAIncreasePctMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAIncreasePctMM(ix)
            clsReadSPIAValues.strSPIAIncreasePctMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIARateChangeDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIARateChangeDateMM
            clsReadSPIAValues.strSPIARateChangeDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAAnnuityTypeMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAAnnuityTypeMM(ix)
            clsReadSPIAValues.strSPIAAnnuityTypeMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIATaxFreeAmtMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIATaxFreeAmtMM(ix)
            clsReadSPIAValues.strSPIATaxFreeAmtMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAExclusionRatioMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAExclusionRatioMM(ix)
            clsReadSPIAValues.strSPIAExclusionRatioMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIATotalCertPaymentsMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIATotalCertPaymentsMM(ix)
            clsReadSPIAValues.strSPIATotalCertPaymentsMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIATotalPaymentsLifeExpecMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIATotalPaymentsLifeExpecMM(ix)
            clsReadSPIAValues.strSPIATotalPaymentsLifeExpecMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAResidentStateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAResidentStateMM
            clsReadSPIAValues.strSPIAResidentStateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAApplicationStateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAApplicationStateMM
            clsReadSPIAValues.strSPIAApplicationStateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPremiumTaxMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPremiumTaxMM
            clsReadSPIAValues.strSPIAPremiumTaxMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationInputAmtMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationInputAmtMM
            clsReadSPIAValues.strSPIACommutationInputAmtMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationInputPctMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationInputPctMM
            clsReadSPIAValues.strSPIACommutationInputPctMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationErrorMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationErrorMM
            clsReadSPIAValues.strSPIACommutationErrorMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationBeneTypeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationBeneTypeMM
            clsReadSPIAValues.strSPIACommutationBeneTypeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationDateMM
            clsReadSPIAValues.strSPIACommutationDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationTaxFreeAmtMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationTaxFreeAmtMM
            clsReadSPIAValues.strSPIACommutationTaxFreeAmtMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationTypeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationTypeMM
            clsReadSPIAValues.strSPIACommutationTypeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIACommutationAmtMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACommutationAmtMM
            clsReadSPIAValues.strSPIACommutationAmtMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIALifeExpectancyMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIALifeExpectancyMM
            clsReadSPIAValues.strSPIALifeExpectancyMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAColaMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAColaMM
            clsReadSPIAValues.strSPIAColaMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIABenefitErrorMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABenefitErrorMM(ix)
            clsReadSPIAValues.strSPIABenefitErrorMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIAIncomeStartDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAIncomeStartDateMM
            clsReadSPIAValues.strSPIAIncomeStartDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPurchaseDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPurchaseDateMM
            clsReadSPIAValues.strSPIAPurchaseDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAQuoteExpirationDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAQuoteExpirationDateMM
            clsReadSPIAValues.strSPIAQuoteExpirationDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAQuoteDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAQuoteDateMM
            clsReadSPIAValues.strSPIAQuoteDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIABeneReductionOptionMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneReductionOptionMM
            clsReadSPIAValues.strSPIABeneReductionOptionMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIASurvivorPctMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIASurvivorPctMM(ix)
            clsReadSPIAValues.strSPIASurvivorPctMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIACostBasisMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIACostBasisMM
            clsReadSPIAValues.strSPIACostBasisMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIATaxStatusMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIATaxStatusMM
            clsReadSPIAValues.strSPIATaxStatusMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPolicyFeeUsedMM(ix)) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPolicyFeeUsedMM(ix)
            clsReadSPIAValues.strSPIAPolicyFeeUsedMM(ix) = ""
        End If

        If Len(clsReadSPIAValues.strSPIABankMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABankMM
            clsReadSPIAValues.strSPIABankMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAChannelMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAChannelMM
            clsReadSPIAValues.strSPIAChannelMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPayoutRateIDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPayoutRateIDMM
            clsReadSPIAValues.strSPIAPayoutRateIDMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPricingCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPricingCodeMM
            clsReadSPIAValues.strSPIAPricingCodeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPricingCodeSuffixMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPricingCodeSuffixMM
            clsReadSPIAValues.strSPIAPricingCodeSuffixMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPayoutRateEffectiveDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPayoutRateEffectiveDateMM
            clsReadSPIAValues.strSPIAPayoutRateEffectiveDateMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAProofOfBirthListMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAProofOfBirthListMM
            clsReadSPIAValues.strSPIAProofOfBirthListMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIABeneSchedLengthTypeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedLengthTypeMM
            clsReadSPIAValues.strSPIABeneSchedLengthTypeMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIABeneSchedInputYearsMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedInputYearsMM
            clsReadSPIAValues.strSPIABeneSchedInputYearsMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIABeneSchedYNMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedYNMM
            clsReadSPIAValues.strSPIABeneSchedYNMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAIRRYNMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAIRRYNMM
            clsReadSPIAValues.strSPIAIRRYNMM = ""
        End If

        If Len(clsReadSPIAValues.strSPIAPopulationTableMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIAPopulationTableMM
            clsReadSPIAValues.strSPIAPopulationTableMM = ""
        End If

        If CInt(clsReadSPIAValues.strSPIAStreamCountTest) = CInt(clsReadSPIAValues.strSPIAStreamCountBench) Then

            For ix = 1 To CInt(clsReadSPIAValues.strSPIAStreamCountTest)

                If Len(clsReadSPIAValues.strSPIABeneSchedStartDateMM(ix)) > 0 Then
                    strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedStartDateMM(ix)
                    clsReadSPIAValues.strSPIABeneSchedStartDateMM(ix) = ""
                End If

                If Len(clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentMM(ix)) > 0 Then
                    strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentMM(ix)
                    clsReadSPIAValues.strSPIABeneSchedAnnIncomePaymentMM(ix) = ""
                End If

                If Len(clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyMM(ix)) > 0 Then
                    strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyMM(ix)
                    clsReadSPIAValues.strSPIABeneSchedAnnPayLivingOnlyMM(ix) = ""
                End If

                If Len(clsReadSPIAValues.strSPIABeneSchedCumPayoutMM(ix)) > 0 Then
                    strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPIAValues.strSPIABeneSchedCumPayoutMM(ix)
                    clsReadSPIAValues.strSPIABeneSchedCumPayoutMM(ix) = ""
                End If

            Next ix
        End If
    End Sub
    Public Sub CreateSPDAMismatchReport(ByVal ib As Integer)

        'write the mismatch strings for each variable

        Dim ix As Integer = 0


        If Len(clsReadSPDAValues.strRunNoRunMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strRunNoRunMM
            clsReadSPDAValues.strRunNoRunMM = ""
        End If


        If Len(clsReadSPDAValues.strMessage1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strMessage1MM
            clsReadSPDAValues.strMessage1MM = ""
        End If

        If Len(clsReadSPDAValues.strMessage2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strMessage2MM
            clsReadSPDAValues.strMessage2MM = ""
        End If

        If Len(clsReadSPDAValues.strMessage3MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strMessage3MM
            clsReadSPDAValues.strMessage3MM = ""
        End If

        If Len(clsReadSPDAValues.strMessage4MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strMessage4MM
            clsReadSPDAValues.strMessage4MM = ""
        End If

        If Len(clsReadSPDAValues.strMessage5MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strMessage5MM
            clsReadSPDAValues.strMessage5MM = ""
        End If
        If Len(clsReadSPDAValues.strMessage6MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strMessage6MM
            clsReadSPDAValues.strMessage6MM = ""
        End If


        If Len(clsReadSPDAValues.strSPDAAnnualPremiumMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAnnualPremiumMM
            clsReadSPDAValues.strSPDAAnnualPremiumMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDACashBenefitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDACashBenefitMM
            clsReadSPDAValues.strSPDACashBenefitMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDADeathBenefitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDADeathBenefitMM
            clsReadSPDAValues.strSPDADeathBenefitMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAClient1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAClient1MM
            clsReadSPDAValues.strSPDAClient1MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAge1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAge1MM
            clsReadSPDAValues.strSPDAAge1MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDASex1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDASex1MM
            clsReadSPDAValues.strSPDASex1MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDADOB1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDADOB1MM
            clsReadSPDAValues.strSPDADOB1MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAClient2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAClient2MM
            clsReadSPDAValues.strSPDAClient2MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAge2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAge2MM
            clsReadSPDAValues.strSPDAAge2MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDASex2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDASex2MM
            clsReadSPDAValues.strSPDASex2MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDADOB2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDADOB2MM
            clsReadSPDAValues.strSPDADOB2MM = ""
        End If

        If Len(clsReadSPDAValues.strSPDACompanyNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDACompanyNameMM
            clsReadSPDAValues.strSPDACompanyNameMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAProdNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAProdNameMM
            clsReadSPDAValues.strSPDAProdNameMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAStateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAStateMM
            clsReadSPDAValues.strSPDAStateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDATaxStatusMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDATaxStatusMM
            clsReadSPDAValues.strSPDATaxStatusMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAPremiumBonusMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAPremiumBonusMM
            clsReadSPDAValues.strSPDAPremiumBonusMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAgentMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAgentMM
            clsReadSPDAValues.strSPDAAgentMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAPremiumTaxRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAPremiumTaxRateMM
            clsReadSPDAValues.strSPDAPremiumTaxRateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDABailoutRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDABailoutRateMM
            clsReadSPDAValues.strSPDABailoutRateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAPayoutRateCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAPayoutRateCodeMM
            clsReadSPDAValues.strSPDAPayoutRateCodeMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAPolicyFormMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAPolicyFormMM
            clsReadSPDAValues.strSPDAPolicyFormMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDASurrChargeYrsMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDASurrChargeYrsMM
            clsReadSPDAValues.strSPDASurrChargeYrsMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAChannelCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAChannelCodeMM
            clsReadSPDAValues.strSPDAChannelCodeMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAMaxIssueAgeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAMaxIssueAgeMM
            clsReadSPDAValues.strSPDAMaxIssueAgeMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAMinIntRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAMinIntRateMM
            clsReadSPDAValues.strSPDAMinIntRateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDADeclaredRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDADeclaredRateMM
            clsReadSPDAValues.strSPDADeclaredRateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDARederminationYearMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDARederminationYearMM
            clsReadSPDAValues.strSPDARederminationYearMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAGuarPeriodMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAGuarPeriodMM
            clsReadSPDAValues.strSPDAGuarPeriodMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDATaxBracketMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDATaxBracketMM
            clsReadSPDAValues.strSPDATaxBracketMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAProjectedRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAProjectedRateMM
            clsReadSPDAValues.strSPDAProjectedRateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAnnuitizationAgeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAnnuitizationAgeMM
            clsReadSPDAValues.strSPDAAnnuitizationAgeMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAWithdrawalTypeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAWithdrawalTypeMM
            clsReadSPDAValues.strSPDAWithdrawalTypeMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAWithdrawalFrequencyMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAWithdrawalFrequencyMM
            clsReadSPDAValues.strSPDAWithdrawalFrequencyMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAutoInterestStartYearMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAutoInterestStartYearMM
            clsReadSPDAValues.strSPDAAutoInterestStartYearMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAutoInterestStopYearMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAutoInterestStopYearMM
            clsReadSPDAValues.strSPDAAutoInterestStopYearMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDASurrenderChargesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDASurrenderChargesMM
            clsReadSPDAValues.strSPDASurrenderChargesMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAHypoInterestRatesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAHypoInterestRatesMM
            clsReadSPDAValues.strSPDAHypoInterestRatesMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAGuarInterestRatesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAGuarInterestRatesMM
            clsReadSPDAValues.strSPDAGuarInterestRatesMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAHypoPartialWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAHypoPartialWDMM
            clsReadSPDAValues.strSPDAHypoPartialWDMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAGuarPartialWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAGuarPartialWDMM
            clsReadSPDAValues.strSPDAGuarPartialWDMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAWDPercentMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAWDPercentMM
            clsReadSPDAValues.strSPDAWDPercentMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAAnnualWDAmountMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAAnnualWDAmountMM
            clsReadSPDAValues.strSPDAAnnualWDAmountMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAHypoCVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAHypoCVMM
            clsReadSPDAValues.strSPDAHypoCVMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAGuarCVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAGuarCVMM
            clsReadSPDAValues.strSPDAGuarCVMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAHypoSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAHypoSVMM
            clsReadSPDAValues.strSPDAHypoSVMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAGuarSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAGuarSVMM
            clsReadSPDAValues.strSPDAGuarSVMM = ""
        End If

        'Added for MVA

        If Len(clsReadSPDAValues.strSPDAMGSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAMGSVMM
            clsReadSPDAValues.strSPDAMGSVMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAMGIRMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAMGIRMM
            clsReadSPDAValues.strSPDAMGIRMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDANonForfRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDANonForfRateMM
            clsReadSPDAValues.strSPDANonForfRateMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDAJumboRatesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDAJumboRatesMM
            clsReadSPDAValues.strSPDAJumboRatesMM = ""
        End If

        If Len(clsReadSPDAValues.strSPDARenewalSurrChargesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadSPDAValues.strSPDARenewalSurrChargesMM
            clsReadSPDAValues.strSPDARenewalSurrChargesMM = ""
        End If


    End Sub
    Public Sub CreateFIAMismatchReport(ByVal ib As Integer)

        'write the mismatch strings for each variable

        Dim ix As Integer = 0

        If gbFIASaveAge(ib) Then
            If gstrFIASaveAgeDOB1 <> "" Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & gstrFIASaveAgeDOB1
                gstrDOB1New = ""
                gstrFIASaveAgeDOB1 = ""
            End If
            If gstrFIASaveAgeDOB2 <> "" Then
                strClientXMisMatch(ib) = strClientXMisMatch(ib) & gstrFIASaveAgeDOB2
                gstrDOB2New = ""
                gstrFIASaveAgeDOB2 = ""
            End If
        End If

        If Len(clsReadFIAValues.strRunNoRunMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strRunNoRunMM
            clsReadFIAValues.strRunNoRunMM = ""
        End If


        If Len(clsReadFIAValues.strMessage1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strMessage1MM
            clsReadFIAValues.strMessage1MM = ""
        End If

        If Len(clsReadFIAValues.strMessage2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strMessage2MM
            clsReadFIAValues.strMessage2MM = ""
        End If

        If Len(clsReadFIAValues.strMessage3MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strMessage3MM
            clsReadFIAValues.strMessage3MM = ""
        End If

        If Len(clsReadFIAValues.strMessage4MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strMessage4MM
            clsReadFIAValues.strMessage4MM = ""
        End If

        If Len(clsReadFIAValues.strMessage5MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strMessage5MM
            clsReadFIAValues.strMessage5MM = ""
        End If
        If Len(clsReadFIAValues.strMessage6MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strMessage6MM
            clsReadFIAValues.strMessage6MM = ""
        End If


        If Len(clsReadFIAValues.strFIAAnnualPremiumMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnualPremiumMM
            clsReadFIAValues.strFIAAnnualPremiumMM = ""
        End If

        If Len(clsReadFIAValues.strFIAClient1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAClient1MM
            clsReadFIAValues.strFIAClient1MM = ""
        End If

        If Len(clsReadFIAValues.strFIAAge1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAge1MM
            clsReadFIAValues.strFIAAge1MM = ""
        End If

        If Len(clsReadFIAValues.strFIASex1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASex1MM
            clsReadFIAValues.strFIASex1MM = ""
        End If

        If Len(clsReadFIAValues.strFIADOB1MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIADOB1MM
            clsReadFIAValues.strFIADOB1MM = ""
        End If

        If Len(clsReadFIAValues.strFIAClient2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAClient2MM
            clsReadFIAValues.strFIAClient2MM = ""
        End If

        If Len(clsReadFIAValues.strFIAAge2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAge2MM
            clsReadFIAValues.strFIAAge2MM = ""
        End If

        If Len(clsReadFIAValues.strFIASex2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASex2MM
            clsReadFIAValues.strFIASex2MM = ""
        End If

        If Len(clsReadFIAValues.strFIADOB2MM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIADOB2MM
            clsReadFIAValues.strFIADOB2MM = ""
        End If

        If Len(clsReadFIAValues.strFIACompanyNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIACompanyNameMM
            clsReadFIAValues.strFIACompanyNameMM = ""
        End If

        If Len(clsReadFIAValues.strFIAProdNameMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAProdNameMM
            clsReadFIAValues.strFIAProdNameMM = ""
        End If

        If Len(clsReadFIAValues.strFIAStateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAStateMM
            clsReadFIAValues.strFIAStateMM = ""
        End If

        If Len(clsReadFIAValues.strFIATaxStatusMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIATaxStatusMM
            clsReadFIAValues.strFIATaxStatusMM = ""
        End If

        If Len(clsReadFIAValues.strFIAPremiumTaxRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPremiumTaxRateMM
            clsReadFIAValues.strFIAPremiumTaxRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAPolicyFormMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPolicyFormMM
            clsReadFIAValues.strFIAPolicyFormMM = ""
        End If

        If Len(clsReadFIAValues.strFIASurrChargeYrsMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASurrChargeYrsMM
            clsReadFIAValues.strFIASurrChargeYrsMM = ""
        End If

        If Len(clsReadFIAValues.strFIAChannelCodeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAChannelCodeMM
            clsReadFIAValues.strFIAChannelCodeMM = ""
        End If

        If Len(clsReadFIAValues.strFIAWDFrequencyMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAWDFrequencyMM
            clsReadFIAValues.strFIAWDFrequencyMM = ""
        End If

        If Len(clsReadFIAValues.strFIASurrenderChargesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASurrenderChargesMM
            clsReadFIAValues.strFIASurrenderChargesMM = ""
        End If

        If Len(clsReadFIAValues.strFIAWDPercentMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAWDPercentMM
            clsReadFIAValues.strFIAWDPercentMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnualWDAmountMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnualWDAmountMM
            clsReadFIAValues.strFIAAnnualWDAmountMM = ""
        End If


        If Len(clsReadFIAValues.strFIAPremiumEnhancementMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPremiumEnhancementMM
            clsReadFIAValues.strFIAPremiumEnhancementMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAgeAtFirstWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAgeAtFirstWDMM
            clsReadFIAValues.strFIAAgeAtFirstWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIARiderChargeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIARiderChargeMM
            clsReadFIAValues.strFIARiderChargeMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnWDLimitGuarMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnWDLimitGuarMM
            clsReadFIAValues.strFIAAnnWDLimitGuarMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnWDLimitProjMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnWDLimitProjMM
            clsReadFIAValues.strFIAAnnWDLimitProjMM = ""
        End If

        If Len(clsReadFIAValues.strFIAOneYearFixedAllocMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAOneYearFixedAllocMM
            clsReadFIAValues.strFIAOneYearFixedAllocMM = ""
        End If

        If Len(clsReadFIAValues.strFIASevenYearFixedAllocMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASevenYearFixedAllocMM
            clsReadFIAValues.strFIASevenYearFixedAllocMM = ""
        End If

        If Len(clsReadFIAValues.strFIATenYearFixedAllocMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIATenYearFixedAllocMM
            clsReadFIAValues.strFIATenYearFixedAllocMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnCapAllocMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnCapAllocMM
            clsReadFIAValues.strFIAAnnCapAllocMM = ""
        End If

        If Len(clsReadFIAValues.strFIAMonCapAllocMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAMonCapAllocMM
            clsReadFIAValues.strFIAMonCapAllocMM = ""
        End If

        If Len(clsReadFIAValues.strFIAPerfTrigAllocMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPerfTrigAllocMM
            clsReadFIAValues.strFIAPerfTrigAllocMM = ""
        End If

        If Len(clsReadFIAValues.strFIAOneYearFixedInitialRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAOneYearFixedInitialRateMM
            clsReadFIAValues.strFIAOneYearFixedInitialRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIASevenYearFixedInitialRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASevenYearFixedInitialRateMM
            clsReadFIAValues.strFIASevenYearFixedInitialRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIATenYearFixedInitialRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIATenYearFixedInitialRateMM
            clsReadFIAValues.strFIATenYearFixedInitialRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnualCapCapMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnualCapCapMM
            clsReadFIAValues.strFIAAnnualCapCapMM = ""
        End If

        If Len(clsReadFIAValues.strFIAMonthlyCapCapMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAMonthlyCapCapMM
            clsReadFIAValues.strFIAMonthlyCapCapMM = ""
        End If

        If Len(clsReadFIAValues.strFIARiderRollupRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIARiderRollupRateMM
            clsReadFIAValues.strFIARiderRollupRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAInitialBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAInitialBeneBaseMM
            clsReadFIAValues.strFIAInitialBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIABailoutAnnualCapMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIABailoutAnnualCapMM
            clsReadFIAValues.strFIABailoutAnnualCapMM = ""
        End If

        If Len(clsReadFIAValues.strFIAperfTrigSpecifiedRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAperfTrigSpecifiedRateMM
            clsReadFIAValues.strFIAperfTrigSpecifiedRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAYearsToPrintMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAYearsToPrintMM
            clsReadFIAValues.strFIAYearsToPrintMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecPeriodStartDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecPeriodStartDateMM
            clsReadFIAValues.strFIASpecPeriodStartDateMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecPeriodEndDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecPeriodEndDateMM
            clsReadFIAValues.strFIASpecPeriodEndDateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavPeriodStartDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavPeriodStartDateMM
            clsReadFIAValues.strFIAFavPeriodStartDateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavPeriodEndDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavPeriodEndDateMM
            clsReadFIAValues.strFIAFavPeriodEndDateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnFavPeriodStartDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnFavPeriodStartDateMM
            clsReadFIAValues.strFIAUnFavPeriodStartDateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnFavPeriodEndDateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnFavPeriodEndDateMM
            clsReadFIAValues.strFIAUnFavPeriodEndDateMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecWDMM
            clsReadFIAValues.strFIASpecWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavWDMM
            clsReadFIAValues.strFIAFavWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavWDMM
            clsReadFIAValues.strFIAUnfavWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecSPChangeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecSPChangeMM
            clsReadFIAValues.strFIASpecSPChangeMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecAnnCreditRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecAnnCreditRateMM
            clsReadFIAValues.strFIASpecAnnCreditRateMM = ""
        End If


        If Len(clsReadFIAValues.strFIAAnnCreditRateNoWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnCreditRateNoWDMM
            clsReadFIAValues.strFIAAnnCreditRateNoWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavAnnCreditRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavAnnCreditRateMM
            clsReadFIAValues.strFIAFavAnnCreditRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavAnnCreditRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavAnnCreditRateMM
            clsReadFIAValues.strFIAUnfavAnnCreditRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecContractValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecContractValueMM
            clsReadFIAValues.strFIASpecContractValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAGMCVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAGMCVMM
            clsReadFIAValues.strFIAGMCVMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavContractValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavContractValueMM
            clsReadFIAValues.strFIAFavContractValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavContractValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavContractValueMM
            clsReadFIAValues.strFIAUnfavContractValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavSurrenderValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavSurrenderValueMM
            clsReadFIAValues.strFIAUnfavSurrenderValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecSurrenderValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecSurrenderValueMM
            clsReadFIAValues.strFIASpecSurrenderValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavSurrenderValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavSurrenderValueMM
            clsReadFIAValues.strFIAFavSurrenderValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavSurrenderValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavSurrenderValueMM
            clsReadFIAValues.strFIAUnfavSurrenderValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecProjBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecProjBeneBaseMM
            clsReadFIAValues.strFIASpecProjBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecGuarBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecGuarBeneBaseMM
            clsReadFIAValues.strFIASpecGuarBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecMGSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecMGSVMM
            clsReadFIAValues.strFIASpecMGSVMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavMGSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavMGSVMM
            clsReadFIAValues.strFIAFavMGSVMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavMGSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavMGSVMM
            clsReadFIAValues.strFIAUnfavMGSVMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecProjWDLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecProjWDLimitMM
            clsReadFIAValues.strFIASpecProjWDLimitMM = ""
        End If

        If Len(clsReadFIAValues.strFIASevenYearIntRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASevenYearIntRateMM
            clsReadFIAValues.strFIASevenYearIntRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIATenYearIntRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIATenYearIntRateMM
            clsReadFIAValues.strFIATenYearIntRateMM = ""
        End If

        If Len(clsReadFIAValues.strFIAMonthlyCapIndexCreditMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAMonthlyCapIndexCreditMM
            clsReadFIAValues.strFIAMonthlyCapIndexCreditMM = ""
        End If

        If Len(clsReadFIAValues.strFIASpecGuarWDLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASpecGuarWDLimitMM
            clsReadFIAValues.strFIASpecGuarWDLimitMM = ""
        End If

        If Len(clsReadFIAValues.strFIASevenYearAccumValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIASevenYearAccumValueMM
            clsReadFIAValues.strFIASevenYearAccumValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIATenYearAccumValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIATenYearAccumValueMM
            clsReadFIAValues.strFIATenYearAccumValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAMonthlyCapAccumValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAMonthlyCapAccumValueMM
            clsReadFIAValues.strFIAMonthlyCapAccumValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnualCapAccumValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnualCapAccumValueMM
            clsReadFIAValues.strFIAAnnualCapAccumValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAPerfTriggerAccumValueMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPerfTriggerAccumValueMM
            clsReadFIAValues.strFIAPerfTriggerAccumValueMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnualCapIndexCreditMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnualCapIndexCreditMM
            clsReadFIAValues.strFIAAnnualCapIndexCreditMM = ""
        End If

        If Len(clsReadFIAValues.strFIAPerfTriggerIndexCreditMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPerfTriggerIndexCreditMM
            clsReadFIAValues.strFIAPerfTriggerIndexCreditMM = ""
        End If

        If Len(clsReadFIAValues.strFIAContractValueNoWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAContractValueNoWDMM
            clsReadFIAValues.strFIAContractValueNoWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAGuarBeneBaseNoWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAGuarBeneBaseNoWDMM
            clsReadFIAValues.strFIAGuarBeneBaseNoWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAProjBeneBAseNoWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAProjBeneBAseNoWDMM
            clsReadFIAValues.strFIAProjBeneBAseNoWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAProjWDLimitNoWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAProjWDLimitNoWDMM
            clsReadFIAValues.strFIAProjWDLimitNoWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAGuarWDLimitNoWDMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAGuarWDLimitNoWDMM
            clsReadFIAValues.strFIAGuarWDLimitNoWDMM = ""
        End If

        If Len(clsReadFIAValues.strFIAGuarWDFactorMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAGuarWDFactorMM
            clsReadFIAValues.strFIAGuarWDFactorMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavSPChangeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavSPChangeMM
            clsReadFIAValues.strFIAFavSPChangeMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnFavSPChangeMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnFavSPChangeMM
            clsReadFIAValues.strFIAUnFavSPChangeMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavProjBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavProjBeneBaseMM
            clsReadFIAValues.strFIAFavProjBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavProjBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavProjBeneBaseMM
            clsReadFIAValues.strFIAUnfavProjBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavGuarBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavGuarBeneBaseMM
            clsReadFIAValues.strFIAFavGuarBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavGuarBeneBaseMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavGuarBeneBaseMM
            clsReadFIAValues.strFIAUnfavGuarBeneBaseMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavProjWDLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavProjWDLimitMM
            clsReadFIAValues.strFIAFavProjWDLimitMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFavGuarWDLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFavGuarWDLimitMM
            clsReadFIAValues.strFIAFavGuarWDLimitMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavGuarWDLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavGuarWDLimitMM
            clsReadFIAValues.strFIAUnfavGuarWDLimitMM = ""
        End If

        If Len(clsReadFIAValues.strFIAUnfavProjWDLimitMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAUnfavProjWDLimitMM
            clsReadFIAValues.strFIAUnfavProjWDLimitMM = ""
        End If

        If Len(clsReadFIAValues.strFIAFixedJumboMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAFixedJumboMM
            clsReadFIAValues.strFIAFixedJumboMM = ""
        End If

        If Len(clsReadFIAValues.strFIAAnnCapJumboMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAAnnCapJumboMM
            clsReadFIAValues.strFIAAnnCapJumboMM = ""
        End If

        If Len(clsReadFIAValues.strFIAMonthlyCapJumboMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAMonthlyCapJumboMM
            clsReadFIAValues.strFIAMonthlyCapJumboMM = ""
        End If

        If Len(clsReadFIAValues.strFIAPerfTriggerJumboMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAPerfTriggerJumboMM
            clsReadFIAValues.strFIAPerfTriggerJumboMM = ""
        End If

        If Len(clsReadFIAValues.strFIAWDFrequencyMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAWDFrequencyMM
            clsReadFIAValues.strFIAWDFrequencyMM = ""
        End If

        If Len(clsReadFIAValues.strFIARenewalSurrChargesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIARenewalSurrChargesMM
            clsReadFIAValues.strFIARenewalSurrChargesMM = ""
        End If

        If Len(clsReadFIAValues.strFIANonForfIntRateMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIANonForfIntRateMM
            clsReadFIAValues.strFIANonForfIntRateMM = ""
        End If





        If Len(clsReadFIAValues.strFIAMGSVMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIAMGSVMM
            clsReadFIAValues.strFIAMGSVMM = ""
        End If


        If Len(clsReadFIAValues.strFIARenewalSurrChargesMM) > 0 Then
            strClientXMisMatch(ib) = strClientXMisMatch(ib) & clsReadFIAValues.strFIARenewalSurrChargesMM
            clsReadFIAValues.strFIARenewalSurrChargesMM = ""
        End If


    End Sub
    Public Sub CreateFIATOOLMismatchReport(ByVal ib As Integer)

        'write the mismatch strings for each variable

        Dim ix As Integer = 0



        If Len(ReadFIARelayINI.strFIAGMCVToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAGMCVToolMM
            ReadFIARelayINI.strFIAGMCVToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecWDToolMM
            ReadFIARelayINI.strFIASpecWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavWDToolMM
            ReadFIARelayINI.strFIAFavWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavWDToolMM
            ReadFIARelayINI.strFIAUnfavWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecSPChangeToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecSPChangeToolMM
            ReadFIARelayINI.strFIASpecSPChangeToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecAnnCreditRateToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecAnnCreditRateToolMM
            ReadFIARelayINI.strFIASpecAnnCreditRateToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavAnnCreditRateToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavAnnCreditRateToolMM
            ReadFIARelayINI.strFIAFavAnnCreditRateToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavAnnCreditRateToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavAnnCreditRateToolMM
            ReadFIARelayINI.strFIAUnfavAnnCreditRateToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecContractValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecContractValueToolMM
            ReadFIARelayINI.strFIASpecContractValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavContractValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavContractValueToolMM
            ReadFIARelayINI.strFIAFavContractValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavContractValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavContractValueToolMM
            ReadFIARelayINI.strFIAUnfavContractValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM
            ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecSurrenderValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecSurrenderValueToolMM
            ReadFIARelayINI.strFIASpecSurrenderValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavSurrenderValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavSurrenderValueToolMM
            ReadFIARelayINI.strFIAFavSurrenderValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM
            ReadFIARelayINI.strFIAUnfavSurrenderValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecProjBeneBaseToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecProjBeneBaseToolMM
            ReadFIARelayINI.strFIASpecProjBeneBaseToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecMGSVToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecMGSVToolMM
            ReadFIARelayINI.strFIASpecMGSVToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavMGSVToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavMGSVToolMM
            ReadFIARelayINI.strFIAFavMGSVToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavMGSVToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavMGSVToolMM
            ReadFIARelayINI.strFIAUnfavMGSVToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASpecProjWDLimitToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASpecProjWDLimitToolMM
            ReadFIARelayINI.strFIASpecProjWDLimitToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASevenYearIntRateToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASevenYearIntRateToolMM
            ReadFIARelayINI.strFIASevenYearIntRateToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIATenYearIntRateToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIATenYearIntRateToolMM
            ReadFIARelayINI.strFIATenYearIntRateToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAMonthlyCapIndexCreditToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAMonthlyCapIndexCreditToolMM
            ReadFIARelayINI.strFIAMonthlyCapIndexCreditToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIASevenYearAccumValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIASevenYearAccumValueToolMM
            ReadFIARelayINI.strFIASevenYearAccumValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIATenYearAccumValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIATenYearAccumValueToolMM
            ReadFIARelayINI.strFIATenYearAccumValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAMonthlyCapAccumValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAMonthlyCapAccumValueToolMM
            ReadFIARelayINI.strFIAMonthlyCapAccumValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAAnnualCapAccumValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAAnnualCapAccumValueToolMM
            ReadFIARelayINI.strFIAAnnualCapAccumValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAPerfTriggerAccumValueToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAPerfTriggerAccumValueToolMM
            ReadFIARelayINI.strFIAPerfTriggerAccumValueToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAAnnualCapIndexCreditToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAAnnualCapIndexCreditToolMM
            ReadFIARelayINI.strFIAAnnualCapIndexCreditToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAPerfTriggerIndexCreditToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAPerfTriggerIndexCreditToolMM
            ReadFIARelayINI.strFIAPerfTriggerIndexCreditToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAContractValueNoWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAContractValueNoWDToolMM
            ReadFIARelayINI.strFIAContractValueNoWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAGuarBeneBaseNoWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAGuarBeneBaseNoWDToolMM
            ReadFIARelayINI.strFIAGuarBeneBaseNoWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAProjBeneBaseNoWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAProjBeneBaseNoWDToolMM
            ReadFIARelayINI.strFIAProjBeneBaseNoWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAProjWDLimitNoWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAProjWDLimitNoWDToolMM
            ReadFIARelayINI.strFIAProjWDLimitNoWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAGuarWDLimitNoWDToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAGuarWDLimitNoWDToolMM
            ReadFIARelayINI.strFIAGuarWDLimitNoWDToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAGuarWDFactorToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAGuarWDFactorToolMM
            ReadFIARelayINI.strFIAGuarWDFactorToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavSPChangeToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavSPChangeToolMM
            ReadFIARelayINI.strFIAFavSPChangeToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavSPChangeToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavSPChangeToolMM
            ReadFIARelayINI.strFIAUnfavSPChangeToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAFavProjBeneBaseToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavProjBeneBaseToolMM
            ReadFIARelayINI.strFIAFavProjBeneBaseToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavProjBeneBaseToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavProjBeneBaseToolMM
            ReadFIARelayINI.strFIAUnfavProjBeneBaseToolMM = ""
        End If


        If Len(ReadFIARelayINI.strFIAFavProjWDLimitToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAFavProjWDLimitToolMM
            ReadFIARelayINI.strFIAFavProjWDLimitToolMM = ""
        End If

        If Len(ReadFIARelayINI.strFIAUnfavProjWDLimitToolMM) > 0 Then
            strClientXMisMatchFIATool(ib) = strClientXMisMatchFIATool(ib) & ReadFIARelayINI.strFIAUnfavProjWDLimitToolMM
            ReadFIARelayINI.strFIAUnfavProjWDLimitToolMM = ""
        End If

    End Sub
    Private Sub FileExists(ByVal strPath As String)

        If FileIO.FileSystem.FileExists(strPath) Then
        Else
            MsgBox(strPath & " is missing!")
            End
        End If
    End Sub
    Private Sub RunExe(ByVal strcomp() As String, ByVal strengine() As String, ByVal i As Short)

        'once the relay.ini is copied over, run the calc engine

        If strcomp(i) = "GELADATA" Then
            'GLAIC VA
            Process.Start("C:\WinFlex6\GELADATA\WFPROP.exe").WaitForExit()
        ElseIf strcomp(i) = "GECLDATA" Then
            'GLICNY VA
            If gstrpathProduct = "VA" Then
                Process.Start("C:\WinFlex6\GECLDATA\GECLRIC.EXE").WaitForExit()
                'GLICNY SPIA
            ElseIf gstrpathProduct = "SPIA" Then
                Process.Start("C:\WinFlex6\GECLDATA\WINANN.EXE").WaitForExit()
                'GLICNY SPDA
            ElseIf gstrpathProduct = "SPDA" Then
                Process.Start("C:\WinFlex6\GECLDATA\GEDefAnn.exe").WaitForExit()
            End If
        ElseIf strcomp(i) = "GECADATA" Then
            'GLIC SPIA
            If gstrpathProduct = "SPIA" Then
                Process.Start("C:\WinFlex6\GECADATA\WINANN.EXE").WaitForExit()
                'GLIC SPDA
            ElseIf gstrpathProduct = "SPDA" Then
                Process.Start("C:\WinFlex6\GECADATA\GEDefAnn.exe").WaitForExit()
            End If
        ElseIf strcomp(i) = "FCOLDATA" Then
            'GLAICFixed SPIA
            If gstrpathProduct = "SPIA" Then
                Process.Start("C:\WinFlex6\FCOLDATA\WINANN.EXE").WaitForExit()
            ElseIf gstrpathProduct = "SPDA" Then
                Process.Start("C:\WinFlex6\FCOLDATA\GEDefAnn.exe").WaitForExit()
            ElseIf gstrpathProduct = "FIA" Then
                Process.Start("C:\WinFlex6\FCOLDATA\FIA.exe").WaitForExit()
            End If
        End If
        Call FileExists("C:\WinFlex6\" & strcomp(i) & "\Relay.out")
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("C:\Program Files\Genworth Financial\Code Mover 3.0\Code Mover 3.0.exe").WaitForExit()
    End Sub
    Private Sub CopyRelayToC(ByVal strcase() As String, ByVal i As Short, ByVal strcomp() As String)
        'copy the relay.ini file from the benchmark folders to the company folders
        System.IO.File.Copy(gstrpath & "\" & gstrpathProduct & "\" & strcase(0) & "\" & i & "\relay.ini", "C:\WinFlex6\" & strcomp(i) & "\relay.ini", True)


        If gstrpathProduct = "SPDA" Then
            ModifyRelaySPDARateDate("C:\WinFlex6\" & strcomp(i) & "\relay.ini", gstrSPDARateDate)
        ElseIf gstrpathProduct = "SPIA" Then
            'If gbSPIAEffectiveDate = True Then
            ModifyRelaySPIARateDate("C:\WinFlex6\" & strcomp(i) & "\relay.ini", gstrSPIARateDate)
            'End If
        ElseIf gstrpathProduct = "VA" Then

            'Insert code to modify VA to save age here: 

            If gbVASaveAgeChecked = True And gbClientDoesntRun = False Then
                CheckVARelayForAge("C:\WinFlex6\" & strcomp(i) & "\relay.ini", Today, i)
            Else
                gbVASaveAge(i) = False
            End If
            gbClientDoesntRun = False

            'Insert code to modify VA to use previous historical numbers:

            If gbVAHistoricalDate And gbClientDoesntRun = False Then
                ModifyRelayVAHistDate("C:\WinFlex6\" & strcomp(i) & "\relay.ini", gstrVAHistDate)
            End If

            'save FIA age
        ElseIf gstrpathProduct = "FIA" Then
            If gbFIASaveAgeChecked = True And gbClientDoesntRun = False Then
                checkfiarelayforage("C:\WinFlex6\" & strcomp(i) & "\relay.ini", Today, i)
            Else
                gbFIASaveAge(i) = False
            End If
            gbClientDoesntRun = False
        End If
    End Sub
    Public Function FolderCount(ByVal PathName As String) As Long

        Dim FSO As New FileSystemObject
        Dim fld As Folder

        'count the number of numbered folders in order to determine the # of clients in the case
        If FSO.FolderExists(PathName) Then
            fld = FSO.GetFolder(PathName)

            For Each subdi As DirectoryInfo In New DirectoryInfo(PathName).GetDirectories()
                If InStr(subdi.Name, "New") Then
                    Directory.Delete(PathName & "\" & subdi.Name, True)
                End If
            Next
            FolderCount = fld.SubFolders.Count

            Return FolderCount
        End If
    End Function

    Public Function FillCaseBox(ByVal pathname As String) As Long

        Dim fso As New FileSystemObject
        Dim fld As Folder

        'fill the list box with cases per product line
        If fso.FolderExists(pathname) Then
            fld = fso.GetFolder(pathname)
            For Each subdi As DirectoryInfo In New DirectoryInfo(pathname).GetDirectories()
                klbCases.Items.Add(subdi)

            Next

        End If

    End Function
    Public Function FillClientList(ByVal pathname As String) As Long

        Dim fso As New FileSystemObject
        Dim fld As Folder

        'clear the list first
        kclbClientList.Items.Clear()  'new client list box
        kclbClientList.Enabled = False
        'kdgvRegressionStatus.Enabled = False
        'fill the list box with clients per case
        If fso.FolderExists(pathname) Then
            fld = fso.GetFolder(pathname)
            For Each subdi As DirectoryInfo In New DirectoryInfo(pathname).GetDirectories()
                If InStr(subdi.Name, "New") Then
                Else
                    Dim s As String = subdi.Name
                    If s.Length = 1 Then
                        s = s.Insert(0, "0")
                    Else
                    End If
                    kclbClientList.Items.Add(s)
                End If
            Next
        End If

        kclbClientList.Sorted = True
        kclbClientList.Enabled = True

    End Function
    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub kbmovecode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbMoveCode.Click

        'run codemover
        Process.Start("\\ricfiles\HQ Depts\HQ Depts 6\IT\WFPROP\codemover4.0\Code Mover 4.0.exe").WaitForExit()

    End Sub
    Public Sub kbrun_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbRun.Click

        KryptonLabel7.Text = ""

        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next

        kbPalomaOnly.Enabled = False

        bPalomaOnly = False
        Run()

        If gstrpathProduct = "FIA" Then
            ReadFIARelayINI.bNoToolRun = False
        Else

            kbPaloma.Enabled = True
        End If
    End Sub
    Private Sub Run()
        'this is the mechanism that runs the illustrations  in order to compare the values from relay.out

        Dim ix As Integer  'General counter
        Dim ic As Integer  'number of cases selected...this is just going to be 1 for now at least
        Dim ib As Integer 'Benchmark folders
        Dim i As Integer
        Dim iFolders As Integer
        Dim iComp As Integer
        Dim strEngine(100) As String
        Dim strComp(100) As String
        Dim im As Integer
        'Dim iClients As Integer '# of clients chosen to run
        Dim icl As Integer
        Dim iCount As Integer
        Dim iSave As Integer

        'set mismatches back to none
        clsReadVAValues.bMismatchAtLeastOnce = False
        clsReadVAValues.bErrorBench = False
        clsReadVAValues.bErrorTest = False
        clsReadVAValues.bMisMatch = False

        clsReadSPIAValues.bMismatchAtLeastOnce = False
        clsReadSPIAValues.bErrorBench = False
        clsReadSPIAValues.bErrorTest = False
        clsReadSPIAValues.bMisMatch = False

        clsReadSPDAValues.bMismatchAtLeastOnce = False
        clsReadSPDAValues.bErrorBench = False
        clsReadSPDAValues.bErrorTest = False
        clsReadSPDAValues.bMisMatch = False

        clsReadFIAValues.bMismatchAtLeastOnce = False
        clsReadFIAValues.bErrorBench = False
        clsReadFIAValues.bErrorTest = False
        clsReadFIAValues.bMisMatch = False

        ReadFIARelayINI.bMismatchToolAtLeastOnce = False
        ReadFIARelayINI.bMisMatchTool = False
       


        'set error messages back to empty
        clsReadVAValues.strMessage1Bench = ""
        clsReadVAValues.strMessage1Test = ""
        clsReadVAValues.strMessage2Bench = ""
        clsReadVAValues.strMessage2Test = ""
        clsReadVAValues.strMessage3Bench = ""
        clsReadVAValues.strMessage3Test = ""
        clsReadVAValues.strMessage4Bench = ""
        clsReadVAValues.strMessage4Test = ""
        clsReadVAValues.strMessage5Bench = ""
        clsReadVAValues.strMessage5Test = ""
        clsReadVAValues.strMessage6Bench = ""
        clsReadVAValues.strMessage6Test = ""


        clsReadSPIAValues.strMessage1Bench = ""
        clsReadSPIAValues.strMessage1Test = ""
        clsReadSPIAValues.strMessage2Bench = ""
        clsReadSPIAValues.strMessage2Test = ""
        clsReadSPIAValues.strMessage3Bench = ""
        clsReadSPIAValues.strMessage3Test = ""
        clsReadSPIAValues.strMessage4Bench = ""
        clsReadSPIAValues.strMessage4Test = ""
        clsReadSPIAValues.strMessage5Bench = ""
        clsReadSPIAValues.strMessage5Test = ""
        clsReadSPIAValues.strMessage6Bench = ""
        clsReadSPIAValues.strMessage6Test = ""
        clsReadSPIAValues.strMessage7Bench = ""
        clsReadSPIAValues.strMessage7Test = ""
        clsReadSPIAValues.strMessage8Bench = ""
        clsReadSPIAValues.strMessage8Test = ""
        clsReadSPIAValues.strMessage9Bench = ""
        clsReadSPIAValues.strMessage9Test = ""
        clsReadSPIAValues.strMessage10Bench = ""
        clsReadSPIAValues.strMessage10Test = ""
        clsReadSPIAValues.strMessage11Bench = ""
        clsReadSPIAValues.strMessage11Test = ""
        clsReadSPIAValues.strMessage12Bench = ""
        clsReadSPIAValues.strMessage12Test = ""
        clsReadSPIAValues.strMessage13Bench = ""
        clsReadSPIAValues.strMessage13Test = ""
        clsReadSPIAValues.strMessage14Bench = ""
        clsReadSPIAValues.strMessage14Test = ""
        clsReadSPIAValues.strMessage15Bench = ""
        clsReadSPIAValues.strMessage15Test = ""
        clsReadSPIAValues.strMessage16Bench = ""
        clsReadSPIAValues.strMessage16Test = ""
        clsReadSPIAValues.strMessage17Bench = ""
        clsReadSPIAValues.strMessage17Test = ""
        clsReadSPIAValues.strMessage18Bench = ""
        clsReadSPIAValues.strMessage18Test = ""
        clsReadSPIAValues.strMessage19Bench = ""
        clsReadSPIAValues.strMessage19Test = ""
        clsReadSPIAValues.strMessage20Bench = ""
        clsReadSPIAValues.strMessage20Test = ""

        clsReadSPDAValues.strMessage1Bench = ""
        clsReadSPDAValues.strMessage1Test = ""
        clsReadSPDAValues.strMessage2Bench = ""
        clsReadSPDAValues.strMessage2Test = ""
        clsReadSPDAValues.strMessage3Bench = ""
        clsReadSPDAValues.strMessage3Test = ""
        clsReadSPDAValues.strMessage4Bench = ""
        clsReadSPDAValues.strMessage4Test = ""
        clsReadSPDAValues.strMessage5Bench = ""
        clsReadSPDAValues.strMessage5Test = ""
        clsReadSPDAValues.strMessage6Bench = ""
        clsReadSPDAValues.strMessage6Test = ""

        clsReadFIAValues.strMessage1Bench = ""
        clsReadFIAValues.strMessage1Test = ""
        clsReadFIAValues.strMessage2Bench = ""
        clsReadFIAValues.strMessage2Test = ""
        clsReadFIAValues.strMessage3Bench = ""
        clsReadFIAValues.strMessage3Test = ""
        clsReadFIAValues.strMessage4Bench = ""
        clsReadFIAValues.strMessage4Test = ""
        clsReadFIAValues.strMessage5Bench = ""
        clsReadFIAValues.strMessage5Test = ""
        clsReadFIAValues.strMessage6Bench = ""
        clsReadFIAValues.strMessage6Test = ""

       

        'Set this to false so statuses can be written properly
        gbSPIAEffectiveDate = False

        'initialize the progress bar
        progress.SmoothProgressBar1.Value = 0

        'need to select a case before clicking on the run button
        If klbCases.SelectedIndex = -1 Then
            MsgBox("Please select a case")
            Return
        End If

        'dont run a new case or rerun the same case until any mismatch files are either benchmarked or deleted
        If klbMismatchedClients.Enabled = True Then
            MessageBox.Show("There were mismatches on this case.  Before running this case again or running a different case, either delete these mismatched files or save them as new benchmarks.", "New Files", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
            DeleteTestFilesNoExit()
        End If

        'determine the # of clients in the case by counting the numbered folders
        iFolders = FolderCount(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0))

        'Set this to false so statuses can be written properly
        For iSave = 1 To iFolders
            gbVASaveAge(iSave) = False
            gbFIASaveAge(iSave) = False
        Next

        'determine the # of clients selected to run
        For icl = 0 To kclbClientList.Items.Count - 1
            If kclbClientList.GetItemCheckState(icl) Then
                iCount = iCount + 1

            End If
        Next

        If iCount = 0 Then
            MsgBox("Please select at least one client to run.", MsgBoxStyle.Critical)
            Return
        End If

        'Modify VA to use previous Historical numbers

        If klbAnnTypes.SelectedItem = "VA" Then
            gbVAPrevHistoricalCancel = False
            Dim form8 As New Form8
            form8.ShowDialog()
            If gbvaprevhistoricalcancel = True Then
                Return
            Else
                gbVAHistoricalDate = form8.krbVAHistoricalEffective.Checked
                If gbvahistoricaldate Then
                    gstrVAHistDate = form8.kdtpVAHistoricalDate.Value.Date
                End If
            End If

        End If

        'Modify VA to save age

        If klbAnnTypes.SelectedItem = "VA" Then
            gbVASaveAgeCancel = False
            Dim form7 As New Form7
            form7.ShowDialog()
            If gbVASaveAgeCancel = True Then
                Return
            Else
                gbVASaveAgeChecked = form7.krbFIASaveAgeYes.Checked
            End If
        End If

        'Modify FIA to save age

        If klbAnnTypes.SelectedItem = "FIA" Then
            gbFIASaveAgeCancel = False
            Dim form7 As New Form7
            form7.ShowDialog()
            If gbFIASaveAgeCancel = True Then
                Return
            Else
                gbFIASaveAgeChecked = form7.krbFIASaveAgeYes.Checked
            End If
        End If

        'Modify SPIA Relay.ini for rates

        If klbAnnTypes.SelectedItem = "SPIA" Then
            gbRateCancel = False
            Dim form6 As New Form6
            'use showdialog to make it modal, wait for this form to exit before moving on
            form6.ShowDialog()
            'gbSPIAEffectiveDate = Form6.krbSPIAEffective.Checked
            If gbRateCancel = True Then
                Return
            Else
                gbSPIAEffectiveDate = form6.krbSPIAEffective.Checked
                If gbSPIAEffectiveDate Then
                    gstrSPIARateDate = form6.kdtpSPIARateDate.Value.Date
                End If
            End If
        End If


        'Modify SPDA Relay.ini for rates

        If klbAnnTypes.SelectedItem = "SPDA" Then
            gbRateCancel = False
            Dim form9 As New Form9
            'use showdialog to make it modal, wait for this form to exit before moving on
            form9.ShowDialog()
            If gbRateCancel = True Then
                Return
            Else
                gbSPDAEffectiveDate = form9.krbSPDAEffective.Checked
                If gbSPDAEffectiveDate Then
                    gstrSPDARateDate = form9.kdtpSPDARateDate.Value.Date
                End If
            End If
        End If

        'set the values on the SPIA Rate form back to defaults
        Form6.krbSPIACurrent.Checked = True
        Form6.krbSPIAEffective.Checked = False
        Form6.kdtpSPIARateDate.Visible = False

        'set the values on the SPDA Rate form back to defaults
        Form9.krbSPDACurrent.Checked = True
        Form9.krbSPDAEffective.Checked = False
        Form9.kdtpSPDARateDate.Visible = False

        'set the values on the VA Save Age form back to defaults
        form7.krbFIASaveAgeYes.Checked = True
        form7.krbFIASaveAgeNo.Checked = False

        'set the values on the VA Historical Date back to defaults
        Form8.krbVAHistoricalCurrent.Checked = True
        Form8.krbVAHistoricalEffective.Checked = False
        Form8.kdtpVAHistoricalDate.Visible = False


        'if no benchmarks already exist
        If iFolders = 0 Then
            MsgBox("There are no existing benchmarks for this case.  Please create them, and try again.", MsgBoxStyle.Critical)
            Return
        End If

        System.Threading.Thread.Sleep(1000)


        progress.Show()

        'Clear and reset the controls on the form
        klbMismatchedClients.Items.Clear()
        klbMismatchedClients.Enabled = False
        KryptonDataGridView1.RowCount = 0

        klbMismatchedClientsFIATool.Items.Clear()
        klbMismatchedClientsFIATool.Enabled = False
        kdgvFIATool.RowCount = 0

        kbNewBench.Enabled = False
        kbViewIllustration.Visible = False
        kbViewIllustration.Enabled = False
        KryptonBorderEdge4.Visible = False
        KryptonBorderEdge4.Enabled = False
        kbViewIllustration.Text = ""

       

        'determine the selected case
        For ix = 0 To klbCases.Items.Count - 1
            If klbCases.GetSelected(ix) Then
                gstrCase(ic) = klbCases.GetItemText(klbCases.Items(ix))
                ic = ic + 1
            End If
        Next ix


        'determine time and date stamps for all of the relevant files 
        TimeStamps.GetTimeStamps()

        ReDim Preserve gstrCase(ic)   'Re-Dimension gstrCase

        i = UBound(gstrCase)   'i = the total # of cases selected

        'set progress bar minimum to 0
        progress.SmoothProgressBar1.Minimum = 0

        'set progress bar maximum to the total number of client in the case
        progress.SmoothProgressBar1.Maximum = iCount 'iFolders


        'for each client, determine the company
        For iComp = 1 To iFolders
            Dim Source = New IniConfigSource(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & iComp & "\relay.ini")
            Dim config As IConfig

            'set up NINI to read the company and engine info from relay.ini
            config = Source.Configs("Identifiers")
            strEngine(iComp) = config.Get("pol_EngineName")
            gstrCompCode(iComp) = config.Get("comp_Code")

            'determine the company based on the engine name
            If strEngine(iComp) = "WFPROP.EXE" Then
                strComp(iComp) = "GELADATA"
            ElseIf strEngine(iComp) = "GECLRIC.EXE" Then
                strComp(iComp) = "GECLDATA"
            ElseIf strEngine(iComp) = "WINANN.EXE" Then
                If gstrCompCode(iComp) = "GECA" Then
                    strComp(iComp) = "GECADATA"
                ElseIf gstrCompCode(iComp) = "GECL" Then
                    strComp(iComp) = "GECLDATA"
                ElseIf gstrCompCode(iComp) = "FCOL" Then
                    strComp(iComp) = "FCOLDATA"
                End If
            ElseIf strEngine(iComp) = "GEDefAnn.exe" Then
                If gstrCompCode(iComp) = "GECL" Then
                    strComp(iComp) = "GECLDATA"
                ElseIf gstrCompCode(iComp) = "GECA" Then
                    strComp(iComp) = "GECADATA"
                ElseIf gstrCompCode(iComp) = "FCOL" Then
                    strComp(iComp) = "FCOLDATA"
                End If
            ElseIf strEngine(iComp) = "FIA.EXE" Then
                strComp(iComp) = "FCOLDATA"
            End If
        Next

        ReDim Preserve strComp(iComp - 1)
        ReDim Preserve strEngine(iComp - 1)


        'set up loop to compare values for each client
        For ib = 1 To iFolders

            'determine the # of clients selected to run

            If kclbClientList.GetItemCheckState(ib - 1) Then

                'initialize the list of mismatches
                strClientXMisMatch(ib) = ""
                strClientXMisMatchFIATool(ib) = ""
                If gstrpathProduct = "SPIA" Then
                    clsReadSPIAValues.bErrorBench = False
                    clsReadSPIAValues.bErrorTest = False
                ElseIf gstrpathProduct = "VA" Then
                    clsReadVAValues.bErrorBench = False
                    clsReadVAValues.bErrorTest = False
                ElseIf gstrpathProduct = "SPDA" Then
                    clsReadSPDAValues.bErrorBench = False
                    clsReadSPDAValues.bErrorTest = False
                ElseIf gstrpathProduct = "FIA" Then
                    clsReadFIAValues.bErrorBench = False
                    clsReadFIAValues.bErrorTest = False
                End If

                'read the exisiting bench values from the relay.out file
                If gstrpathProduct = "SPIA" Then
                    ReadSPIABench.ReadSPIABenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out") 'this will work for only 1 case for now
                ElseIf gstrpathProduct = "VA" Then
                    ReadBench.ReadBenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out")
                ElseIf gstrpathProduct = "SPDA" Then
                    ReadSPDABench.ReadSPDABenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out")
                ElseIf gstrpathProduct = "FIA" Then
                    ReadFIABench.ReadFIABenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out")
                End If

                'delete any existing illustration and relay files from the company directory before calcing the case

                DeleteFiles(strComp, ib)

                'Write to the Relay.ini before copying to C, to produce a .pdf for Paloma
                ModifyRelayPaloma(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0), ib)

                'copy the relay.ini file from the existing benchmark files to the company directory
                CopyRelayToC(gstrCase, ib, strComp)


                'run the calc engine using the relay.ini files
                RunExe(strComp, strEngine, ib)

                'read the test values just created
                If gstrpathProduct = "SPIA" Then
                    ReadSPIATest.ReadSPIATestValues("C:\WinFlex6\" & strComp(ib) & "\Relay.out")
                ElseIf gstrpathProduct = "VA" Then
                    ReadTest.ReadTestValues("C:\WinFlex6\" & strComp(ib) & "\Relay.out")
                ElseIf gstrpathProduct = "SPDA" Then
                    ReadSPDATest.ReadSPDATestValues("C:\WinFlex6\" & strComp(ib) & "\Relay.out")
                ElseIf gstrpathProduct = "FIA" Then
                    ReadFIATest.ReadFIATestValues("C:\WinFlex6\" & strComp(ib) & "\Relay.out")
                End If


                'FOR FIA TOOL
                If gstrpathProduct = "FIA" Then
                    If kcbFIATool.Checked = True Then
                        ''clear out the array of FIA TOOL mismatches
                        'Array.Clear(strsplitmmlistFIATool, 0, strsplitmmlistFIATool.Length)
                        'Array.Resize(strsplitmmlistFIATool, 0)
                        FIATool.ReadFIAInputs("\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ib) & "\relay.ini", "C:\WinFlex6\" & strComp(ib) & "\Relay.out", ib)
                        If ib = 1 Then
                            KryptonLabel7.Text = KryptonLabel7.Text & "Current WinFlex Run vs. FIA Tool (Version:  " & ReadFIARelayINI.strToolVersion & ")"
                        End If
                    End If
                End If

                'End If


                'compare the new test values against the existing benchmark values
                If gstrpathProduct = "VA" Then
                    CompareValuesVA(ic, ib, strComp)
                ElseIf gstrpathProduct = "SPIA" Then
                    CompareValuesSPIA(ic, ib, strComp)
                ElseIf gstrpathProduct = "SPDA" Then
                    CompareValuesSPDA(ic, ib, strComp)
                ElseIf gstrpathProduct = "FIA" Then
                    CompareValuesFIA(ic, ib, strComp)
                    If kcbFIATool.Checked = True And ReadFIARelayINI.bNoToolRun = False Then
                        CompareValuesFIATOOL(ic, ib, strComp)
                    End If
                End If


                'advance the progress bar for each client compared

                progress.SmoothProgressBar1.Value += 1

                System.Windows.Forms.Application.DoEvents()
            End If


            'set client doesn't run back to false, so VA Save Age works properly
            gbClientDoesntRun = False


        Next

        'make sure a case was selected
        If klbCases.SelectedIndex = -1 Then
        Else
            progress.Close()

        End If

        'if the cases does not match, enable the necessary controls on the form

        If clsReadVAValues.bMismatchAtLeastOnce = True Or clsReadSPIAValues.bMismatchAtLeastOnce = True Or clsReadSPDAValues.bMismatchAtLeastOnce = True Or clsReadFIAValues.bMismatchAtLeastOnce = True Then

            'kbDisplayMismatches.Enabled = True
            If gstrpathProduct = "SPIA" And gbSPIAEffectiveDate = True Then
                kbNewBench.Enabled = False
            ElseIf gstrpathProduct = "VA" And gbVAHistoricalDate = True Then
                kbNewBench.Enabled = False
            Else
                kbNewBench.Enabled = True
            End If

            klbMismatchedClients.Enabled = True
            kbDeleteFiles.Enabled = True

            'fill the mismatched clients listbox
            klbMismatchedClients.Items.Clear() 'clear first?
            For im = 1 To UBound(strsplitMMList) 'iFolders
                If Len(strsplitMMList(im)) > 0 Then
                    klbMismatchedClients.Items.Add(strsplitMMList(im))
                End If
            Next

            MsgBox("MISMATCHES EXIST IN THIS CASE.  The stats  for " & gstrCase(0) & " have been written to:  " & gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & " to a file named runstatscase(client#).txt in each client folder.", MsgBoxStyle.Critical)

        ElseIf klbMismatchedClients.Items.Count = 0 Then

            klbMismatchedClients.Enabled = False
            kbViewIllustration.Visible = False
            kbViewIllustration.Enabled = False
            KryptonBorderEdge4.Visible = False
            KryptonBorderEdge4.Enabled = False
            kbDeleteFiles.Enabled = False

            MsgBox("ALL CLIENTS IN THIS CASE MATCH THE BENCHMARK.  The stats  for " & gstrCase(0) & " have been written to:  " & gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & " to a file named runstatscase(client#).txt in each client folder.", MsgBoxStyle.Information, "The Regressionator")
        End If

        If kcbFIATool.Checked = True Then
            If ReadFIARelayINI.bMismatchToolAtLeastOnce = True Then 'FIA tool mismatches

                klbMismatchedClientsFIATool.Enabled = True
                kbClearToolMM.Enabled = True

                'fill the mismatched clients listbox for FIA TOOL Cases
                klbMismatchedClientsFIATool.Items.Clear() 'clear first?
                For im = 1 To UBound(strsplitMMListFIATool) 'iFolders
                    If Len(strsplitMMListfiatool(im)) > 0 Then
                        klbMismatchedClientsFIATool.Items.Add(strsplitMMListFIAtool(im))
                    End If
                Next

                MsgBox("MISMATCHES EXIST IN THIS CASE BETWEEN WINFLEX AND FIA TOOL.  Any cases that do not run in WinFlex (test) have not been run in the tool.  The stats  for " & gstrCase(0) & " have been written to:  " & gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & " to a file named runstatscase(client#).txt in each client folder.", MsgBoxStyle.Critical)

            ElseIf klbMismatchedClientsFIATool.Items.Count = 0 Then

                klbMismatchedClientsFIATool.Enabled = False
                'kbViewIllustration.Visible = False
                'kbViewIllustration.Enabled = False
                'KryptonBorderEdge4.Visible = False
                'KryptonBorderEdge4.Enabled = False
                'kbDeleteFiles.Enabled = False

                MsgBox("ALL CLIENTS IN THIS CASE MATCH THE BETWEEN WINFLEX AND FIA TOOL.  Any cases that do not run in WinFlex (test) have not been run in the tool.  The stats  for " & gstrCase(0) & " have been written to:  " & gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & " to a file named runstatscase(client#).txt in each client folder.", MsgBoxStyle.Information, "The Regressionator")
            End If
        End If


        Dim iskip As Integer

        If ReadFIARelayINI.bFIAToolSkippedAtLeastOnce = True Then
            klbFIAToolSkippedList.Enabled = True
            kbClearToolMM.Enabled = True

            'fill the skipped clients listbox for FIA TOOL Cases
            klbFIAToolSkippedList.Items.Clear() 'clear first?
            For iskip = 1 To UBound(ReadFIARelayINI.strsplitFIAToolSkipList) 'iFolders
                If Len(ReadFIARelayINI.strsplitFIAToolSkipList(iskip)) > 0 Then
                    klbFIAToolSkippedList.Items.Add(ReadFIARelayINI.strsplitFIAToolSkipList(iskip))
                End If
            Next
        End If


        'reset the progress bar
        progress.SmoothProgressBar1.Value = 0

       


        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next
    End Sub
    Private Sub RunPalomaOnly()

        'this is the mechanism that runs the illustrations  in order to compare the values from relay.out

        Dim ix As Integer  'General counter
        Dim ic As Integer  'number of cases selected...this is just going to be 1 for now at least
        Dim ib As Integer 'Benchmark folders
        Dim i As Integer
        Dim iFolders As Integer
        Dim iComp As Integer
        Dim strEngine(100) As String
        Dim strComp(100) As String
        Dim icl As Integer
        Dim iCount As Integer
        Dim iSave As Integer

        'set mismatches back to none
        clsReadVAValues.bMismatchAtLeastOnce = False
        clsReadVAValues.bErrorBench = False
        clsReadVAValues.bErrorTest = False
        clsReadVAValues.bMisMatch = False

        clsReadSPIAValues.bMismatchAtLeastOnce = False
        clsReadSPIAValues.bErrorBench = False
        clsReadSPIAValues.bErrorTest = False
        clsReadSPIAValues.bMisMatch = False

        clsReadSPDAValues.bMismatchAtLeastOnce = False
        clsReadSPDAValues.bErrorBench = False
        clsReadSPDAValues.bErrorTest = False
        clsReadSPDAValues.bMisMatch = False

        'set error messages back to empty
        clsReadVAValues.strMessage1Bench = ""
        clsReadVAValues.strMessage1Test = ""
        clsReadVAValues.strMessage2Bench = ""
        clsReadVAValues.strMessage2Test = ""
        clsReadVAValues.strMessage3Bench = ""
        clsReadVAValues.strMessage3Test = ""
        clsReadVAValues.strMessage4Bench = ""
        clsReadVAValues.strMessage4Test = ""
        clsReadVAValues.strMessage5Bench = ""
        clsReadVAValues.strMessage5Test = ""
        clsReadVAValues.strMessage6Bench = ""
        clsReadVAValues.strMessage6Test = ""


        clsReadSPIAValues.strMessage1Bench = ""
        clsReadSPIAValues.strMessage1Test = ""
        clsReadSPIAValues.strMessage2Bench = ""
        clsReadSPIAValues.strMessage2Test = ""
        clsReadSPIAValues.strMessage3Bench = ""
        clsReadSPIAValues.strMessage3Test = ""
        clsReadSPIAValues.strMessage4Bench = ""
        clsReadSPIAValues.strMessage4Test = ""
        clsReadSPIAValues.strMessage5Bench = ""
        clsReadSPIAValues.strMessage5Test = ""
        clsReadSPIAValues.strMessage6Bench = ""
        clsReadSPIAValues.strMessage6Test = ""
        clsReadSPIAValues.strMessage7Bench = ""
        clsReadSPIAValues.strMessage7Test = ""
        clsReadSPIAValues.strMessage8Bench = ""
        clsReadSPIAValues.strMessage8Test = ""
        clsReadSPIAValues.strMessage9Bench = ""
        clsReadSPIAValues.strMessage9Test = ""
        clsReadSPIAValues.strMessage10Bench = ""
        clsReadSPIAValues.strMessage10Test = ""
        clsReadSPIAValues.strMessage11Bench = ""
        clsReadSPIAValues.strMessage11Test = ""
        clsReadSPIAValues.strMessage12Bench = ""
        clsReadSPIAValues.strMessage12Test = ""
        clsReadSPIAValues.strMessage13Bench = ""
        clsReadSPIAValues.strMessage13Test = ""
        clsReadSPIAValues.strMessage14Bench = ""
        clsReadSPIAValues.strMessage14Test = ""
        clsReadSPIAValues.strMessage15Bench = ""
        clsReadSPIAValues.strMessage15Test = ""
        clsReadSPIAValues.strMessage16Bench = ""
        clsReadSPIAValues.strMessage16Test = ""
        clsReadSPIAValues.strMessage17Bench = ""
        clsReadSPIAValues.strMessage17Test = ""
        clsReadSPIAValues.strMessage18Bench = ""
        clsReadSPIAValues.strMessage18Test = ""
        clsReadSPIAValues.strMessage19Bench = ""
        clsReadSPIAValues.strMessage19Test = ""
        clsReadSPIAValues.strMessage20Bench = ""
        clsReadSPIAValues.strMessage20Test = ""

        clsReadSPDAValues.strMessage1Bench = ""
        clsReadSPDAValues.strMessage1Test = ""
        clsReadSPDAValues.strMessage2Bench = ""
        clsReadSPDAValues.strMessage2Test = ""
        clsReadSPDAValues.strMessage3Bench = ""
        clsReadSPDAValues.strMessage3Test = ""
        clsReadSPDAValues.strMessage4Bench = ""
        clsReadSPDAValues.strMessage4Test = ""
        clsReadSPDAValues.strMessage5Bench = ""
        clsReadSPDAValues.strMessage5Test = ""
        clsReadSPDAValues.strMessage6Bench = ""
        clsReadSPDAValues.strMessage6Test = ""

        clsReadFIAValues.strMessage1Bench = ""
        clsReadFIAValues.strMessage1Test = ""
        clsReadFIAValues.strMessage2Bench = ""
        clsReadFIAValues.strMessage2Test = ""
        clsReadFIAValues.strMessage3Bench = ""
        clsReadFIAValues.strMessage3Test = ""
        clsReadFIAValues.strMessage4Bench = ""
        clsReadFIAValues.strMessage4Test = ""
        clsReadFIAValues.strMessage5Bench = ""
        clsReadFIAValues.strMessage5Test = ""
        clsReadFIAValues.strMessage6Bench = ""
        clsReadFIAValues.strMessage6Test = ""

        'Set this to false so statuses can be written properly
        gbSPIAEffectiveDate = False

        'initialize the progress bar
        progress.SmoothProgressBar1.Value = 0

        'need to select a case before clicking on the run button
        If klbCases.SelectedIndex = -1 Then
            MsgBox("Please select a case")
            Return
        End If

        'dont run a new case or rerun the same case until any mismatch files are either benchmarked or deleted
        If klbMismatchedClients.Enabled = True Then
            MessageBox.Show("There were mismatches on this case.  Before running this case again or running a different case, either delete these mismatched files or save them as new benchmarks.", "New Files", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
            DeleteTestFilesNoExit()
        End If

        'determine the # of clients in the case by counting the numbered folders
        iFolders = FolderCount(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0))

        'Set this to false so statuses can be written properly
        For iSave = 1 To iFolders
            gbVASaveAge(iSave) = False
        Next

        'determine the # of clients selected to run
        For icl = 0 To kclbClientList.Items.Count - 1
            If kclbClientList.GetItemCheckState(icl) Then
                iCount = iCount + 1

            End If
        Next

        If iCount = 0 Then
            MsgBox("Please select at least one client to run.", MsgBoxStyle.Critical)
            Return
        End If

        'Modify SPIA Relay.ini for rates

        If klbAnnTypes.SelectedItem = "SPIA" Then
            gbRateCancel = False
            Dim form6 As New Form6
            'use showdialog to make it modal, wait for this form to exit before moving on
            form6.ShowDialog()

            If gbRateCancel = True Then
                Return
            Else

                gbSPIAEffectiveDate = form6.krbSPIAEffective.Checked
                If gbSPIAEffectiveDate Then
                    gstrSPIARateDate = form6.kdtpSPIARateDate.Value.Date
                End If
            End If
        End If

        'Modify SPDA Relay.ini for rates

        If klbAnnTypes.SelectedItem = "SPDA" Then
            gbRateCancel = False
            Dim form9 As New Form9
            'use showdialog to make it modal, wait for this form to exit before moving on
            form9.ShowDialog()
            If gbRateCancel = True Then
                Return
            Else
                gbSPDAEffectiveDate = form9.krbSPDAEffective.Checked
                If gbSPDAEffectiveDate Then
                    gstrSPDARateDate = form9.kdtpSPDARateDate.Value.Date
                End If
            End If
        End If

        'Modify VA to use previous Historical numbers

        If klbAnnTypes.SelectedItem = "VA" Then
            gbVAPrevHistoricalCancel = False
            Dim form8 As New Form8
            form8.ShowDialog()
            If gbVAPrevHistoricalCancel = True Then
                Return
            Else
                gbVAHistoricalDate = form8.krbVAHistoricalEffective.Checked
                If gbVAHistoricalDate Then
                    gstrVAHistDate = form8.kdtpVAHistoricalDate.Value.Date
                End If
            End If

        End If

        'Modify VA to save age

        If klbAnnTypes.SelectedItem = "VA" Then
            gbVASaveAgeCancel = False
            Dim form7 As New Form7
            form7.ShowDialog()
            If gbVASaveAgeCancel = True Then
                Return
            Else
                gbVASaveAgeChecked = form7.krbFIASaveAgeYes.Checked
            End If
        End If

        'set the values on the SPIA Rate form back to defaults
        Form6.krbSPIACurrent.Checked = True
        Form6.krbSPIAEffective.Checked = False
        Form6.kdtpSPIARateDate.Visible = False


        'set the values on the SPDA Rate form back to defaults
        Form9.krbSPDACurrent.Checked = True
        Form9.krbSPDAEffective.Checked = False
        Form9.kdtpSPDARateDate.Visible = False

        'set the values on the VA Save Age form back to defaults
        Form7.krbFIASaveAgeYes.Checked = True
        Form7.krbFIASaveAgeNo.Checked = False

        'set the values on the VA Historical Date back to defaults
        Form8.krbVAHistoricalCurrent.Checked = True
        Form8.krbVAHistoricalEffective.Checked = False
        Form8.kdtpVAHistoricalDate.Visible = False

        'if no benchmarks already exist
        If iFolders = 0 Then
            MsgBox("There are no existing benchmarks for this case.  Please create them, and try again.", MsgBoxStyle.Critical)
            Return
        End If

        progress.Show()

        'Clear and reset the controls on the form
        klbMismatchedClients.Items.Clear()
        klbMismatchedClients.Enabled = False
        KryptonDataGridView1.RowCount = 0
        kbNewBench.Enabled = False
        kbViewIllustration.Visible = False
        kbViewIllustration.Enabled = False
        KryptonBorderEdge4.Visible = False
        KryptonBorderEdge4.Enabled = False
        kbViewIllustration.Text = ""

        'determine the selected case
        For ix = 0 To klbCases.Items.Count - 1
            If klbCases.GetSelected(ix) Then
                gstrCase(ic) = klbCases.GetItemText(klbCases.Items(ix))
                ic = ic + 1
            End If
        Next ix


        'determine time and date stamps for all of the relevant files 
        TimeStamps.GetTimeStamps()

        ReDim Preserve gstrCase(ic)   'Re-Dimension gstrCase

        i = UBound(gstrCase)   'i = the total # of cases selected


        'set progress bar minimum to 0
        progress.SmoothProgressBar1.Minimum = 0

        'set progress bar maximum to the total number of client in the case
        progress.SmoothProgressBar1.Maximum = iCount 'iFolders


        'for each client, determine the company
        For iComp = 1 To iFolders
            Dim Source = New IniConfigSource(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & iComp & "\relay.ini")
            Dim config As IConfig

            'set up NINI to read the company and engine info from relay.ini
            config = Source.Configs("Identifiers")
            strEngine(iComp) = config.Get("pol_EngineName")
            gstrCompCode(iComp) = config.Get("comp_Code")

            'determine the company based on the engine name
            If strEngine(iComp) = "WFPROP.EXE" Then
                strComp(iComp) = "GELADATA"
            ElseIf strEngine(iComp) = "GECLRIC.EXE" Then
                strComp(iComp) = "GECLDATA"
            ElseIf strEngine(iComp) = "WINANN.EXE" Then
                If gstrCompCode(iComp) = "GECA" Then
                    strComp(iComp) = "GECADATA"
                ElseIf gstrCompCode(iComp) = "GECL" Then
                    strComp(iComp) = "GECLDATA"
                ElseIf gstrCompCode(iComp) = "FCOL" Then
                    strComp(iComp) = "FCOLDATA"
                End If
            ElseIf strEngine(iComp) = "GEDefAnn.exe" Then
                If gstrCompCode(iComp) = "GECL" Then
                    strComp(iComp) = "GECLDATA"
                ElseIf gstrCompCode(iComp) = "GECA" Then
                    strComp(iComp) = "GECADATA"
                ElseIf gstrCompCode(iComp) = "FCOL" Then
                    strComp(iComp) = "FCOLDATA"
                End If
            End If
        Next

        ReDim Preserve strComp(iComp - 1)
        ReDim Preserve strEngine(iComp - 1)


        'set up loop to compare values for each client
        For ib = 1 To iFolders

            'determine the # of clients selected to run

            If kclbClientList.GetItemCheckState(ib - 1) Then


                'initialzie the listof mismatches
                strClientXMisMatch(ib) = ""
                If gstrpathProduct = "SPIA" Then
                    clsReadSPIAValues.bErrorBench = False
                    clsReadSPIAValues.bErrorTest = False
                ElseIf gstrpathProduct = "VA" Then
                    clsReadVAValues.bErrorBench = False
                    clsReadVAValues.bErrorTest = False
                ElseIf gstrpathProduct = "SPDA" Then
                    clsReadSPDAValues.bErrorBench = False
                    clsReadSPDAValues.bErrorTest = False
                ElseIf gstrpathProduct = "FIA" Then
                    clsReadFIAValues.bErrorBench = False
                    clsReadFIAValues.bErrorTest = False
                End If

                'read the exisiting bench values from the relay.out file
                If gstrpathProduct = "SPIA" Then
                    ReadSPIABench.ReadSPIABenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out") 'this will work for only 1 case for now
                ElseIf gstrpathProduct = "VA" Then
                    ReadBench.ReadBenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out")
                ElseIf gstrpathProduct = "SPDA" Then
                    ReadSPDABench.ReadSPDABenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out")
                ElseIf gstrpathProduct = "FIA" Then
                    ReadFIABench.ReadFIABenchValues(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & ib & "\Relay.out")
                End If


                'Write to the Relay.ini before copying to C, to produce a .pdf for Paloma
                ModifyRelayPaloma(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0), ib)

                'copy the relay.ini file from the existing benchmark files to the company directory
                CopyRelayToC(gstrCase, ib, strComp)

                'run the calc engine using the relay.ini files
                RunExe(strComp, strEngine, ib)

                'advance the progress bar for each client compared

                progress.SmoothProgressBar1.Value += 1

                System.Windows.Forms.Application.DoEvents()
            Else
            End If

            'copy test.pdf to folders, even if not comparing
            If kclbClientList.GetItemChecked(ib - 1) Then
                'if no test pdf, dont try to copy :)
                If FileIO.FileSystem.FileExists("\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ib) & "\Test\Test.pdf") Then
                    CopyPDF(gstrpath & "\" & gstrpathProduct & "\", gstrCase(ic - 1), ib, strComp)
                End If
            End If

            'set client doesn't run back to false, so VA Save Age works properly
            gbClientDoesntRun = False

        Next


        'make sure a case was selected
        If klbCases.SelectedIndex = -1 Then
        Else
            progress.Close()

            'MsgBox("The run is complete.  The stats  for " & gstrCase(0) & " have been written to:  " & gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & " to a file named runstatscase(client#).txt in each client folder.", MsgBoxStyle.Information)
            'MsgBox("Run is complete.", MsgBoxStyle.Information, "The Regressionator")
        End If

        'if the cases does not match, enable the necessary controls on the form

        'reset the progress bar
        progress.SmoothProgressBar1.Value = 0
        ShowPaloma(iFolders)
    End Sub
    Private Sub ModifyRelayPaloma(ByVal PathName As String, ByVal ib As Integer)
        'set the correct path below
        PathName = PathName & "\" & ib & "\Relay.ini"
        Dim Fs As FileStream = New FileStream(PathName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String

        str = sr.ReadToEnd()

        If InStr(str, "Paloma.pdf=Y") Then
        Else
            str = str & "Paloma.pdf=Y"
            str = str & ControlChars.NewLine
        End If

        Fs.Position = 0
        sw.Write(str)

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()
    End Sub
    Private Sub ModifyRelaySPIARateDate(ByVal pathname As String, ByVal strdate As String)

        'set the correct path below
        Dim Fs As FileStream = New FileStream(pathname, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String
        Dim strSplits() As String
        Dim ix As Integer

        str = sr.ReadToEnd()
        ix = 0

        strSplits = str.Split(ControlChars.NewLine)
        For ix = 0 To strSplits.Length - 1

            If gbSPIAEffectiveDate = True Then
                If InStr(strSplits(ix), "FCOL.PayoutRI") Then
                    str = str.Replace(strSplits(ix).TrimStart, "FCOL.PayoutRI=Effective Date")
                    Fs.SetLength(str.Length)
                End If
            Else
                If InStr(strSplits(ix), "FCOL.PayoutRI") Then
                    str = str.Replace(strSplits(ix).TrimStart, "FCOL.PayoutRI=Current Rate")
                    Fs.SetLength(str.Length)
                End If

                If InStr(strSplits(ix), "FCOL.SPIAEffectDate") Then
                    str = str.Replace(strSplits(ix).TrimStart, "                             ")
                End If

            End If

        Next ix

        If gbSPIAEffectiveDate = True Then
            str = str & ControlChars.NewLine & "FCOL.SPIAEffectDate=" & gstrSPIARateDate
        End If

        Fs.Position = 0
        sw.Write(str)

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()

    End Sub
    Private Sub ModifyRelaySPDARateDate(ByVal pathname As String, ByVal strdate As String)

        'set the correct path below
        Dim Fs As FileStream = New FileStream(pathname, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String


        str = sr.ReadToEnd()

        If gbSPDAEffectiveDate = True Then
            str = str & ControlChars.NewLine & "NewHypoDate=" & gstrSPDARateDate
            str = str & ControlChars.NewLine & "NewGuarDate=" & gstrSPDARateDate
        End If

        Fs.Position = 0
        sw.Write(str)

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()

    End Sub
    Private Sub ModifyRelayVAHistDate(ByVal pathname As String, ByVal strdate As String)

        'set the correct path below
        Dim Fs As FileStream = New FileStream(pathname, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String
        'Dim strSplits() As String
        'Dim ix As Integer

        str = sr.ReadToEnd()

        If gbVAHistoricalDate = True Then
            str = str & ControlChars.NewLine & "PrevHistoricalDate=" & gstrVAHistDate
        End If

        Fs.Position = 0
        sw.Write(str)

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()

    End Sub
    Private Sub ModifyRelayVASaveAge(ByVal pathname As String, ByVal iAge1 As Integer, ByVal strDOB1 As String, ByVal iclient As Integer, ByVal iage2 As Integer, ByVal strDOB2 As String)

        'set the correct path below

        Dim Fs As FileStream = New FileStream(pathname, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String
        Dim strSplits() As String
        Dim ix As Integer

        str = sr.ReadToEnd()
        ix = 0

        strSplits = str.Split(ControlChars.NewLine)
        For ix = 0 To strSplits.Length - 1

            'Write the modified DOB(s) to the Relay.ini on the local drive

            If InStr(strSplits(ix), "Insured.DateOfBirth") Then
                str = str.Replace(strSplits(ix).TrimStart, "Insured.DateOfBirth=" & strDOB1)
                Fs.SetLength(str.Length)
            End If

            If InStr(strSplits(ix), "Insured2.DateOfBirth") Then
                str = str.Replace(strSplits(ix).TrimStart, "Insured2.DateOfBirth=" & strDOB2)
                Fs.SetLength(str.Length)
            End If

        Next ix

        Fs.Position = 0
        sw.Write(str)

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()

    End Sub
    Private Sub CheckVARelayForAge(ByVal pathname As String, ByVal strdate As String, ByVal iclient As Integer)

        'set the correct path below
        Dim Fs As FileStream = New FileStream(pathname, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String
        Dim strSplits() As String
        Dim ix As Integer

        Dim strAge1 As String = ""
        Dim strDOB1 As String = ""
        Dim strAge2 As String = ""
        Dim strDOB2 As String = ""
        Dim iAge1 As Integer
        Dim iAge2 As Integer = 0
        Dim dtDOB1 As Date
        Dim dtDOB2 As Date

        str = sr.ReadToEnd()
        ix = 0

        strSplits = str.Split(ControlChars.NewLine)
        For ix = 0 To strSplits.Length - 1

            'Read the DOB(s) from the Relay.ini on the local drive

            If InStr(strSplits(ix), "Insured.DateOfBirth") Then
                strDOB1 = strSplits(ix)
            End If

            If InStr(strSplits(ix), "Insured2.DateOfBirth") Then
                strDOB2 = strSplits(ix)
            End If

        Next ix

        'pull the Ages from the Relay.out file, as the Relay.ini file does not get updated on any runs or Benchmarks
        strAge1 = clsReadVAValues.strAge1Bench
        strAge2 = clsReadVAValues.strAge2Bench

        'Convert Strings to Integers, call to check ages...

        strDOB1 = strDOB1.Remove(0, 21)
        dtDOB1 = CDate(strDOB1)
        iAge1 = CInt(strAge1)
        strDOB1 = Format(dtDOB1, "MM/dd/yyyy")

        If strDOB2 <> "" Then
            strDOB2 = strDOB2.Remove(0, 22)
            dtDOB2 = CDate(strDOB2)
            iAge2 = CInt(strAge2)
            strDOB2 = Format(dtDOB2, "MM/dd/yyyy")
            CalculateAgeLastJoint(pathname, strDOB1, iAge1, iclient, strDOB2, iAge2)
        Else
            CalculateAgeLastSingle(pathname, strDOB1, iAge1, iclient)
        End If

        Fs.Position = 0

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()

    End Sub

    Private Sub CheckFIARelayForAge(ByVal pathname As String, ByVal strdate As String, ByVal iclient As Integer)

        'set the correct path below
        Dim Fs As FileStream = New FileStream(pathname, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As New StreamWriter(Fs, Encoding.Default)
        Dim sr As New StreamReader(Fs, Encoding.Default)
        Dim str As String
        Dim strSplits() As String
        Dim ix As Integer

        Dim strAge1 As String = ""
        Dim strDOB1 As String = ""
        Dim strAge2 As String = ""
        Dim strDOB2 As String = ""
        Dim iAge1 As Integer
        Dim iAge2 As Integer = 0
        Dim dtDOB1 As Date
        Dim dtDOB2 As Date

        str = sr.ReadToEnd()
        ix = 0

        strSplits = str.Split(ControlChars.NewLine)
        For ix = 0 To strSplits.Length - 1

            'Read the DOB(s) from the Relay.ini on the local drive

            If InStr(strSplits(ix), "Insured.DateOfBirth") Then
                strDOB1 = strSplits(ix)
            End If

            If InStr(strSplits(ix), "Insured2.DateOfBirth") Then
                strDOB2 = strSplits(ix)
            End If

        Next ix

        'pull the Ages from the Relay.out file, as the Relay.ini file does not get updated on any runs or Benchmarks
        strAge1 = clsReadFIAValues.strFIAAge1Bench
        strAge2 = clsReadFIAValues.strFIAAge2Bench

        'Convert Strings to Integers, call to check ages...

        strDOB1 = strDOB1.Remove(0, 21)
        dtDOB1 = CDate(strDOB1)
        iAge1 = CInt(strAge1)
        strDOB1 = Format(dtDOB1, "MM/dd/yyyy")

        If strDOB2 <> "" Then
            strDOB2 = strDOB2.Remove(0, 22)
            dtDOB2 = CDate(strDOB2)
            iAge2 = CInt(strAge2)
            strDOB2 = Format(dtDOB2, "MM/dd/yyyy")
            CalculateAgeLastJoint(pathname, strDOB1, iAge1, iclient, strDOB2, iAge2)
        Else
            CalculateAgeLastSingle(pathname, strDOB1, iAge1, iclient)
        End If

        Fs.Position = 0

        Fs.Flush()
        sw.Flush()
        sr.Close()
        Fs.Close()

    End Sub
    Private Sub KryptonButton3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub KryptonDataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles KryptonDataGridView1.CellContentClick

    End Sub
  
    Public Sub Kbdisplaymismatches_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim strSep As String = ""

        'Try
        '    'display the values that do not match on the selected case
        '    If klbMismatchedClients.SelectedIndex = -1 Then
        '        MsgBox("Please select a client in order to view the mismatches.", MsgBoxStyle.Critical)
        '    End If
        'Catch ex As Exception
        'End Try
        'parse the string to fill the data grid
        Dim splittest = Split(strClientXMisMatch(klbMismatchedClients.SelectedItem), "&")
        Dim splittestrow(splittest.Count, 5)
        Dim strmmlistnew As String = ""

        Dim strMMList As String = ""
        Dim ix As Integer = 0
        Dim iy As Integer = 0

        'initialize the data grid
        KryptonDataGridView1.RowCount = 0

        For ix = 0 To splittest.Count - 2
            For iy = 0 To 4
                If InStr(splittest(ix), "�") Then
                    splittestrow(ix, iy) = Split(splittest(ix), "�")
                Else
                    splittestrow(ix, iy) = Split(splittest(ix), ",")

                End If

            Next
        Next

        'Dim LastBench As String = ReadBenchMarkDate(gstrpath, gstrCase(0), klbMismatchedClients.SelectedItem)

        'fill the data grid with the values
        For i As Integer = 0 To splittest.Count - 2
            If InStr(splittest(i), "�") Then
                strSep = "�"
            Else
                strSep = ","
            End If
            Dim splittestgrid = Split(splittest(i), strSep)

            Dim item As New DataGridViewRow
            item.CreateCells(KryptonDataGridView1)
            With item
                .Cells(0).Value = splittestgrid(0)
                .Cells(1).Value = splittestgrid(1)
                If splittestgrid.Count = 4 Then
                    .Cells(2).Value = ""
                    .Cells(3).Value = splittestgrid(2)
                    .Cells(4).Value = splittestgrid(3)
                Else
                    .Cells(2).Value = splittestgrid(2)
                    .Cells(3).Value = splittestgrid(3)
                    .Cells(4).Value = splittestgrid(4)
                End If

            End With
            KryptonDataGridView1.Rows.Add(item)
        Next



        If klbMismatchedClients.SelectedIndex <> -1 And KryptonDataGridView1.RowCount > 0 Then
            'enable the button to view the illustration pages
            kbViewIllustration.Visible = True
            kbViewIllustration.Enabled = True
            KryptonBorderEdge4.Visible = True
            KryptonBorderEdge4.Enabled = True

            'set text on button corresponding to the selected case
            kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClients.SelectedItem

            'enable the button to save the mismatch results
            kbSaveMismatchResults.Enabled = True
        End If
    End Sub
    Public Sub DisplayMismatches()
        Dim strSep As String = ""
        Dim iback As Integer

        'Try
        '    'display the values that do not match on the selected case
        '    If klbMismatchedClients.SelectedIndex = -1 Then
        '        MsgBox("Please select a client in order to view the mismatches.", MsgBoxStyle.Critical)
        '    End If
        'Catch ex As Exception
        'End Try

        'parse the string to fill the data grid
        Dim splittest = Split(strClientXMisMatch(klbMismatchedClients.SelectedItem), "&")
        Dim splittestrow(splittest.Count, 5)
        Dim strmmlistnew As String = ""

        Dim strMMList As String = ""
        Dim ix As Integer = 0
        Dim iy As Integer = 0

        'initialize the data grid
        KryptonDataGridView1.RowCount = 0

        For ix = 0 To splittest.Count - 2
            For iy = 0 To 4
                If InStr(splittest(ix), "�") Then
                    If InStr(splittest(ix), "WARNING") Or InStr(splittest(ix), "ERROR") Then
                        splittestrow(ix, iy) = Split(splittest(ix), ",")
                    Else
                        splittestrow(ix, iy) = Split(splittest(ix), "�")
                    End If
                Else
                    splittestrow(ix, iy) = Split(splittest(ix), ",")

                End If
            Next
        Next

        'fill the data grid with the values

        For i As Integer = 0 To splittest.Count - 2

            If InStr(splittest(i), "�") Then
                If InStr(splittest(i), "WARNING") Or InStr(splittest(i), "ERROR") Then
                    strSep = ","
                Else
                    strSep = "�"
                End If
            Else
                strSep = ","
            End If
            Dim splittestgrid = Split(splittest(i), strSep)

            Dim item As New DataGridViewRow
            item.CreateCells(KryptonDataGridView1)
            With item
                .Cells(0).Value = splittestgrid(0)
                .Cells(1).Value = splittestgrid(1)
                If splittestgrid.Count = 4 Then
                    .Cells(2).Value = ""
                    .Cells(3).Value = splittestgrid(2)
                    .Cells(4).Value = splittestgrid(3)
                Else
                    .Cells(2).Value = splittestgrid(2)
                    .Cells(3).Value = splittestgrid(3)
                    .Cells(4).Value = splittestgrid(4)
                End If
            End With

            'color the text of mismatches Red if using SPIA Effective Date instead of Current Date
            For iy = 0 To 4
                If gstrpathProduct = "SPIA" And gbSPIAEffectiveDate = True Then
                    item.Cells(iy).Style.ForeColor = Color.Red
                ElseIf gstrpathProduct = "SPDA" And gbSPDAEffectiveDate = True Then
                    item.Cells(iy).Style.ForeColor = Color.Red
                ElseIf gstrpathProduct = "VA" And gbVASaveAge(item.Cells(0).Value) = True Then
                    item.Cells(iy).Style.ForeColor = Color.Green
                ElseIf gstrpathProduct = "FIA" And gbFIASaveAge(item.Cells(0).Value) = True Then
                    item.Cells(iy).Style.ForeColor = Color.Green
                End If
            Next



            'color the background of the item name depending on hypo/hist...
            If gstrpathProduct = "VA" Then
                If InStr(splittestgrid(1), "Hypo") Then
                    For iback = 0 To 4
                        item.Cells(iback).Style.BackColor = Color.LightSteelBlue
                    Next
                ElseIf InStr(splittestgrid(1), "Hist") Or InStr(splittestgrid(1), "Fund") Then
                    For iback = 0 To 4
                        item.Cells(iback).Style.BackColor = Color.Gold
                    Next
                End If
            End If

            KryptonDataGridView1.Rows.Add(item)
        Next



        If klbMismatchedClients.SelectedIndex <> -1 And KryptonDataGridView1.RowCount > 0 Then
            'enable the button to view the illustration pages
            kbViewIllustration.Visible = True
            kbViewIllustration.Enabled = True
            KryptonBorderEdge4.Visible = True
            KryptonBorderEdge4.Enabled = True

            'set text on button corresponding to the selected case
            kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClients.SelectedItem

            'enable the button to save the mismatch results
            kbSaveMismatchResults.Enabled = True
        End If
    End Sub
    Public Sub DisplayMismatchesFIATool()
        Dim strSep As String = ""
        'Dim iback As Integer

        'Try
        '    'display the values that do not match on the selected case
        '    If klbMismatchedClientsFIATool.SelectedIndex = -1 Then
        '        MsgBox("Please select a client in order to view the mismatches.", MsgBoxStyle.Critical)
        '    End If
        'Catch ex As Exception
        'End Try

        'parse the string to fill the data grid
        Dim splittest = Split(strClientXMisMatchFIATool(klbMismatchedClientsFIATool.SelectedItem), "&")
        Dim splittestrow(splittest.Count, 5)
        Dim strmmlistnew As String = ""

        Dim strMMList As String = ""
        Dim ix As Integer = 0
        Dim iy As Integer = 0

        'initialize the data grid
        kdgvFIATool.RowCount = 0

        For ix = 0 To splittest.Count - 2
            For iy = 0 To 4
                If InStr(splittest(ix), "�") Then
                    If InStr(splittest(ix), "WARNING") Or InStr(splittest(ix), "ERROR") Then
                        splittestrow(ix, iy) = Split(splittest(ix), ",")
                    Else
                        splittestrow(ix, iy) = Split(splittest(ix), "�")
                    End If
                Else
                    splittestrow(ix, iy) = Split(splittest(ix), ",")

                End If
            Next
        Next

        'fill the data grid with the values

        For i As Integer = 0 To splittest.Count - 2

            If InStr(splittest(i), "�") Then
                If InStr(splittest(i), "WARNING") Or InStr(splittest(i), "ERROR") Then
                    strSep = ","
                Else
                    strSep = "�"
                End If
            Else
                strSep = ","
            End If
            Dim splittestgrid = Split(splittest(i), strSep)

            Dim item As New DataGridViewRow
            item.CreateCells(kdgvFIATool)
            With item
                .Cells(0).Value = splittestgrid(0)
                .Cells(1).Value = splittestgrid(1)
                If splittestgrid.Count = 4 Then
                    .Cells(2).Value = ""
                    .Cells(3).Value = splittestgrid(2)
                    .Cells(4).Value = splittestgrid(3)
                Else
                    .Cells(2).Value = splittestgrid(2)
                    .Cells(3).Value = splittestgrid(3)
                    .Cells(4).Value = splittestgrid(4)
                End If
            End With

            'color the text of mismatches Red if using SPIA Effective Date instead of Current Date
            'For iy = 0 To 4
            '    If gstrpathProduct = "SPIA" And gbSPIAEffectiveDate = True Then
            '        item.Cells(iy).Style.ForeColor = Color.Red
            '    ElseIf gstrpathProduct = "SPDA" And gbSPDAEffectiveDate = True Then
            '        item.Cells(iy).Style.ForeColor = Color.Red
            '    ElseIf gstrpathProduct = "VA" And gbVASaveAge(item.Cells(0).Value) = True Then
            '        item.Cells(iy).Style.ForeColor = Color.Green
            '    ElseIf gstrpathProduct = "FIA" And gbFIASaveAge(item.Cells(0).Value) = True Then
            '        item.Cells(iy).Style.ForeColor = Color.Green
            '    End If
            'Next



            'color the background of the item name depending on hypo/hist...
            'If gstrpathProduct = "VA" Then
            '    If InStr(splittestgrid(1), "Hypo") Then
            '        For iback = 0 To 4
            '            item.Cells(iback).Style.BackColor = Color.LightSteelBlue
            '        Next
            '    ElseIf InStr(splittestgrid(1), "Hist") Or InStr(splittestgrid(1), "Fund") Then
            '        For iback = 0 To 4
            '            item.Cells(iback).Style.BackColor = Color.Gold
            '        Next
            '    End If
            'End If

            kdgvFIATool.Rows.Add(item)
        Next



        If klbMismatchedClientsFIATool.SelectedIndex <> -1 And kdgvFIATool.RowCount > 0 Then
            'enable the button to view the illustration pages
            'kbViewIllustration.Visible = True
            'kbViewIllustration.Enabled = True
            'KryptonBorderEdge4.Visible = True
            'KryptonBorderEdge4.Enabled = True

            'set text on button corresponding to the selected case
            'kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClients.SelectedItem

            'enable the button to save the mismatch results
            'kbSaveMismatchResults.Enabled = True
        End If
    End Sub
    Public Function ReadBenchMarkDate(ByVal strpath As String, ByVal strcase As String, ByVal iclient As Integer)

        'read the last benchmark date

        Dim strComplPath As String = strpath & "\" & gstrpathProduct & "\" & strcase & "\" & iclient & "\"
        Dim LastBench As String

        If FileIO.FileSystem.FileExists(strComplPath & "\lastbenchmarkdate" & iclient & ".txt") Then

            Dim sr As System.IO.StreamReader
            sr = System.IO.File.OpenText(strComplPath & "\lastbenchmarkdate" & iclient & ".txt")
            LastBench = sr.ReadToEnd
            If LastBench = "" Then LastBench = "None Yet"
            sr.Close()

        Else
            LastBench = "None"

        End If
        Return lastBench
    End Function
    Private Sub kbnewbench_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbNewBench.Click

        Dim cCtrl As Control
        For Each cCtrl In Me.Controls
            If cCtrl.TabIndex <> 20 Then
                cCtrl.Enabled = False
            End If
        Next

        'open the form to create new benchmarks
        Form4.Show()

        'fill the checkedlistbox with the cases that do not match
        For ix = 0 To klbMismatchedClients.Items.Count - 1
            Form4.kclbNewBench.Items.Add(klbMismatchedClients.Items(ix))
        Next

        'clear the mismatch list box
        strClientMisMatchList = ""

        'clear out the array of mismatches
        Array.Clear(strsplitMMList, 0, strsplitMMList.Length)
        Array.Resize(strsplitMMList, 0)
        'If kcbFIATool.Checked = True Then
        '    Array.Clear(strsplitmmlistFIATool, 0, strsplitmmlistFIATool.Length)
        '    Array.Resize(strsplitmmlistFIATool, 0)
        'End If

        ''clear and reset the controls on the form

        klbMismatchedClients.Enabled = False
        kbNewBench.Enabled = False
        kbViewIllustration.Text = ""
        kbViewIllustration.Visible = False
        kbViewIllustration.Enabled = False
        KryptonBorderEdge4.Visible = False
        KryptonBorderEdge4.Enabled = False

    End Sub

    Private Sub kbexit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbExit.Click

        'delete the new bench files and exit the program
        If klbMismatchedClients.Enabled = True Then
            Dim result As DialogResult = MessageBox.Show("There were mismatches on this case, exiting now without creating new benchmarks will delete the new results.  Click YES to exit anyway, click NO to create new benchmarks first", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = System.Windows.Forms.DialogResult.Yes Then
                DeleteNewBenchFolders()
                Me.Close()
                'Me.Dispose()


            ElseIf result = System.Windows.Forms.DialogResult.No Then
            End If
        Else
            If Form3.Enabled = True Then Form3.Close()
            If Form4.Enabled = True Then Form4.Close()

            Me.Close()
            'Me.Dispose()


        End If
        System.Windows.Forms.Application.DoEvents()

        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next


    End Sub

    Private Sub DeleteTestFilesNoExit()

        'Delete the new bench files without exiting the program
        Dim result As DialogResult = MessageBox.Show("There were mismatches on this case, running a new case now without creating new benchmarks will delete the new results.  Click YES to run it anyway, click NO to create new benchmarks first", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = System.Windows.Forms.DialogResult.Yes Then
            DeleteNewBenchFolders()
            Return
        ElseIf result = System.Windows.Forms.DialogResult.No Then
        End If

    End Sub
    Private Sub kbviewillustration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbViewIllustration.Click


        'set up the bench and test directories in order to view the illustration pages
        Dim diTest As DirectoryInfo = New DirectoryInfo(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(klbCases.SelectedValue) & "\new" & klbMismatchedClients.SelectedItem)
        Dim diBench As DirectoryInfo = New DirectoryInfo(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(klbCases.SelectedValue) & "\" & klbMismatchedClients.SelectedItem)

        'get the .emf pages of the illustrations
        Dim emfTest As FileInfo() = diTest.GetFiles("*.emf")
        Dim emfBench As FileInfo() = diBench.GetFiles("*.emf")

        'initialize the # of pages, etc
        Dim iTestPages As Integer = 0
        Dim iBenchPages As Integer = 0
        Dim iMore As Integer = 0
        Dim iboxorig As Integer = 0

        'disable main form when viewing illustrations
        Dim cCtrl As Control
        For Each cCtrl In Me.Controls
            If cCtrl.TabIndex <> 20 Then
                cCtrl.Enabled = False
            End If
        Next

        'show the view illustration form
        Form3.Show()

        'clear out and fill the list box with page #'s
        Form3.klbPages.Items.Clear()

        'get # of pages and add a "0" for pages below 10 so will sort properly
        For Each finext In emfTest
            iTestPages = iTestPages + 1
            Dim s As String = finext.Name
            If s.Length = 9 Then
                s = s.Substring(0, 5)
                s = s.Insert(4, "0")
            Else
                s = s.Substring(0, 6)
            End If
            Form3.klbPages.Items.Add(s)
        Next

        'get # of pages and add a "0" for pages below 10 so will sort properly
        For Each finext In emfBench
            iBenchPages = iBenchPages + 1
            Dim s1 As String = finext.Name
            If s1.Length = 9 Then
                s1 = s1.Substring(0, 5)
                s1 = s1.Insert(4, "0")
            Else
                s1 = s1.Substring(0, 6)
            End If
        Next


        If iTestPages = 0 Or iBenchPages = 0 Then
            'just show empty form if no test or bench illustration
        Else
           
            kbViewIllustration.Visible = True
            kbViewIllustration.Enabled = True
            'if bench and test illustration do not have the same # of pages, make them equal and then put 
            'up message that pages are not available
            If iTestPages <> iBenchPages Then
                If iBenchPages > iTestPages Then
                    iTestPages = iBenchPages
                ElseIf iTestPages > iBenchPages Then
                    iBenchPages = iTestPages
                End If

                Dim s2 As String = ""
                If iBenchPages > Form3.klbPages.Items.Count Then
                    iboxorig = Form3.klbPages.Items.Count
                    iMore = iBenchPages - Form3.klbPages.Items.Count
                    For ix = 1 To iMore
                        s2 = "page" & iboxorig + ix
                        If s2.Length = 9 Then
                            s2 = s2.Substring(0, 5)
                            s2 = s2.Insert(4, "0")
                        Else
                            s2 = s2.Substring(0, 6)
                        End If
                        Form3.klbPages.Items.Add(s2)

                    Next
                End If
            End If
        End If
    End Sub
    Private Sub KryptonLabel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles KryptonLabel1.Paint
    End Sub
    Private Sub klbanntypes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles klbAnnTypes.SelectedIndexChanged

        kcbFIATool.Checked = False
        kcbFIATool.Enabled = False
        klbCases.SelectedIndex = -1
        'Load the list box with the 4 Ann types
        klbCases.Items.Clear()
        klbCases.SelectedIndex = -1
        gstrpathProduct = klbAnnTypes.SelectedItem
        FillCaseBox("\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & gstrpathProduct)
        kclbClientList.Items.Clear()  'new client list box
        kbOpenCaseFolder.Enabled = True
        kdgvRegressionStatus.Rows.Clear()
        kbPaloma.Enabled = False
        kbClearAll.Enabled = False
        kbSelectAll.Enabled = False
        If gstrpathProduct = "FIA" Then
            kcbFIATool.Enabled = True
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Kbdeletefiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbDeleteFiles.Click

        'delete the new records created
        MsgBox("This will delete the mismatched test cases just run.", MsgBoxStyle.Exclamation)

        DeleteNewBenchFolders()

        MsgBox("The mismatched test cases have been deleted.", MsgBoxStyle.Information)

        System.Windows.Forms.Application.DoEvents()

        'clear the mismatch list box
        strClientMisMatchList = ""

        'clear out the array of mismatches
        Array.Clear(strsplitMMList, 0, strsplitMMList.Length)
        Array.Resize(strsplitMMList, 0)
        'If kcbFIATool.Checked = True Then
        '    Array.Clear(strsplitmmlistFIATool, 0, strsplitmmlistFIATool.Length)
        '    Array.Resize(strsplitmmlistFIATool, 0)
        'End If

        'set mismatches back to none
        clsReadVAValues.bMismatchAtLeastOnce = False
        clsReadVAValues.bErrorBench = False
        clsReadVAValues.bErrorTest = False
        clsReadSPIAValues.bMismatchAtLeastOnce = False
        clsReadSPIAValues.bErrorBench = False
        clsReadSPIAValues.bErrorTest = False
        clsReadSPDAValues.bMismatchAtLeastOnce = False
        clsReadSPDAValues.bErrorBench = False
        clsReadSPDAValues.bErrorTest = False
        clsReadFIAValues.bMismatchAtLeastOnce = False
        clsReadFIAValues.bErrorBench = False
        clsReadFIAValues.bErrorTest = False

        'clear and reset the controls on the form
        klbMismatchedClients.Items.Clear()
        klbMismatchedClients.Enabled = False
        KryptonDataGridView1.RowCount = 0
        kbNewBench.Enabled = False
        kbViewIllustration.Visible = False
        kbViewIllustration.Enabled = False
        KryptonBorderEdge4.Visible = False
        KryptonBorderEdge4.Enabled = False
        kbSaveMismatchResults.Enabled = False
        kbDeleteFiles.Enabled = False

        'For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '    proc.Kill()
        'Next
    End Sub
    Private Sub klbMismatchedClients_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles klbMismatchedClients.SelectedIndexChanged
        If kbViewIllustration.Enabled = True And kbViewIllustration.Visible = True Then
            kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClients.SelectedItem
            'klbMismatchedClients.SelectedIndex = -1
        End If

        DisplayMismatches()

    End Sub
    
   
    Private Sub klbCases_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles klbCases.Click

    End Sub


    Private Sub klbCases_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles klbCases.SelectedIndexChanged

        kbPaloma.Enabled = False

        FillRegStatus()

    End Sub
    Private Sub FillRegStatus()
        'fill the DataGrid with ALL of the clients SELECTED to run in the case, whether they match or not

        kdgvRegressionStatus.Rows.Clear()

        kbSelectAll.Enabled = False
        kbClearAll.Enabled = False
        For ix = 0 To kclbClientList.Items.Count - 1
            Dim item As New DataGridViewRow
            item.CreateCells(kdgvRegressionStatus)
            'If kclbClientList.GetItemCheckState(ix) Then
            Dim strReport As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\Control\Report.pdf" & ""
            Dim strBench As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\Control\bench.txt" & ""
            Dim strLastRegCompare As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\runstatscase" & (ix + 1) & ".txt" & ""
            Dim strLastRegBench As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\lastbenchmarkdate" & (ix + 1) & ".txt" & ""

            Dim objFileInfoLastCompare As New FileInfo(strReport)
            Dim objfileinfoLastBench As New FileInfo(strBench)
            Dim objFileInfoLastCompareReg As New FileInfo(strLastRegCompare)
            Dim objFileInfoLastBenchReg As New FileInfo(strLastRegBench)
            Dim LastPalomaComp As DateTime = objFileInfoLastCompare.LastWriteTime
            Dim LastPalomaBench As DateTime = objfileinfoLastBench.LastWriteTime
            Dim LastRegComp As DateTime = objFileInfoLastCompareReg.LastWriteTime
            Dim LastRegBench As DateTime = objFileInfoLastBenchReg.LastWriteTime

            If FileIO.FileSystem.FileExists("\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\Test\Test.pdf") Then


                item.Cells(0).Value = CStr(ix + 1)
                If LastRegComp = #12/31/1600 7:00:00 PM# Then
                    item.Cells(1).Value = "None Yet"
                Else
                    item.Cells(1).Value = LastRegComp
                End If
                If LastRegBench = #12/31/1600 7:00:00 PM# Then
                    item.Cells(2).Value = "None Yet"
                Else
                    item.Cells(2).Value = LastRegBench
                End If
                If LastPalomaComp = #12/31/1600 7:00:00 PM# Then
                    item.Cells(3).Value = "None Yet"
                Else
                    item.Cells(3).Value = LastPalomaComp
                End If
                If LastPalomaBench = #12/31/1600 7:00:00 PM# Then
                    item.Cells(4).Value = ("None Yet")
                Else
                    item.Cells(4).Value = LastPalomaBench
                End If

                kdgvRegressionStatus.Rows.Add(item)

            Else
                item.Cells(0).Value = CStr(ix + 1)
                If LastRegComp = #12/31/1600 7:00:00 PM# Then
                    item.Cells(1).Value = "None Yet"
                Else
                    item.Cells(1).Value = LastRegComp
                End If
                If LastRegBench = #12/31/1600 7:00:00 PM# Then
                    item.Cells(2).Value = "None Yet"
                Else
                    item.Cells(2).Value = LastRegBench
                End If
                item.Cells(3).Value = "--"
                item.Cells(4).Value = "--"

                kdgvRegressionStatus.Rows.Add(item)
            End If

        Next
        kdgvRegressionStatus.Enabled = True
        kbSelectAll.Enabled = True
        kbClearAll.Enabled = True
    End Sub
    Private Sub kbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbSaveMismatchResults.Click

        'Write the datagridview results to a file in the numbered benchmark directory

        'parse the string to fill the data grid
        Dim splittest = Split(strClientXMisMatch(klbMismatchedClients.SelectedItem), "&")

        'Dim splittestrow(100, 5)
        Dim ix As Integer = 0
        Dim iy As Integer = 0
        Dim strComplPath As String = gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\" & klbMismatchedClients.SelectedItem & "\"
        Dim str As String = ""
        Dim stritem As String = ""
        Dim message As String
        Dim title As String
        Dim value As String

        message = "Add notes for Client # " & klbMismatchedClients.SelectedItem
        title = "The Regressionator:  Notes"
        value = InputBox(message, title)

        If FileIO.FileSystem.FileExists(strComplPath & "mismatchdetails" & klbMismatchedClients.SelectedItem & ".txt") Then
            Dim sw As System.IO.StreamWriter
            Dim sr As System.IO.StreamReader
            sr = System.IO.File.OpenText(strComplPath & "mismatchdetails" & klbMismatchedClients.SelectedItem & ".txt")
            Dim MyContents As String = sr.ReadToEnd
            sr.Close()

            sw = System.IO.File.AppendText(strComplPath & "mismatchdetails" & klbMismatchedClients.SelectedItem & ".txt")
            sw.WriteLine(vbCrLf)
            sw.WriteLine(Today & " " & TimeOfDay & "  " & gstrCase(0) & " Client # " & klbMismatchedClients.SelectedItem & "  Run by:  " & username)

            'write the time stamps for the relevant files
            If gstrpathProduct = "VA" Then
                If gstrCompCode(klbMismatchedClients.SelectedItem) = "GELA" Then
                    sw.WriteLine(clsReadVAValues.strWFProp)
                    sw.WriteLine(clsReadVAValues.strGLAICCPY)
                    sw.WriteLine(clsReadVAValues.strAnn1)
                    sw.WriteLine(clsReadVAValues.strAnn2)
                    sw.WriteLine(clsReadVAValues.strAnn3)
                    sw.WriteLine(clsReadVAValues.strAnn4)
                    sw.WriteLine(clsReadVAValues.strWFGELA)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYCPY)
                    sw.WriteLine(clsReadVAValues.strGECLRIC)
                    sw.WriteLine(clsReadVAValues.strGECLVA1)
                    sw.WriteLine(clsReadVAValues.strGECLVA2)
                    sw.WriteLine(clsReadVAValues.strGECLVA3)
                    sw.WriteLine(clsReadVAValues.strGECLEXE)
                End If
            ElseIf gstrpathProduct = "SPIA" Then
                If gstrCompCode(klbMismatchedClients.SelectedItem) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICSPIANNVER)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNVER)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIANNVER)
                End If
            ElseIf gstrpathProduct = "SPDA" Then
                If gstrCompCode(klbMismatchedClients.SelectedItem) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDARATE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDACPY)
                End If
            ElseIf gstrpathProduct = "FIA" Then
                sw.WriteLine(clsreadvavalues.strfiaexe)
                sw.WriteLine(clsreadvavalues.strfiamdb)
                sw.WriteLine(clsreadvavalues.strfiacpy)

            End If
            sw.WriteLine(vbCrLf)

            'write the mismatches to the file
            For i As Integer = 0 To splittest.Count - 2
                Dim splittestgrid = Split(splittest(i), ",")

                If splittestgrid.Count = 4 Then
                    str = "Variable/Element:" & ControlChars.Tab & splittestgrid(1) & "/ N/A"
                    sw.WriteLine(str)
                    str = "Bench Value/TestValue:" & ControlChars.Tab & splittestgrid(2) & "/" & splittestgrid(3)
                    sw.WriteLine(str)
                    sw.WriteLine("________")
                Else
                    str = "Variable/Element:" & ControlChars.Tab & splittestgrid(1) & "/" & splittestgrid(2)
                    sw.WriteLine(str)
                    str = "Bench Value/Test Value:" & ControlChars.Tab & splittestgrid(3) & "/" & splittestgrid(4)
                    sw.WriteLine(str)
                    sw.WriteLine("________")
                End If
            Next
            If value <> "" Then
                sw.WriteLine("Notes:  " & value)
            End If
            sw.Close()
            sr.Close()
        Else
            Dim sw As System.IO.StreamWriter
            sw = System.IO.File.CreateText(strComplPath & "mismatchdetails" & klbMismatchedClients.SelectedItem & ".txt")
            sw.WriteLine(Today & " " & TimeOfDay & "  " & gstrCase(0) & " Client # " & klbMismatchedClients.SelectedItem & "  Run by:  " & username)

            If gstrpathProduct = "VA" Then
                If gstrCompCode(klbMismatchedClients.SelectedItem) = "GELA" Then
                    sw.WriteLine(clsReadVAValues.strWFProp)
                    sw.WriteLine(clsReadVAValues.strGLAICCPY)
                    sw.WriteLine(clsReadVAValues.strAnn1)
                    sw.WriteLine(clsReadVAValues.strAnn2)
                    sw.WriteLine(clsReadVAValues.strAnn3)
                    sw.WriteLine(clsReadVAValues.strAnn4)
                    sw.WriteLine(clsReadVAValues.strWFGELA)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYCPY)
                    sw.WriteLine(clsReadVAValues.strGECLRIC)
                    sw.WriteLine(clsReadVAValues.strGECLVA1)
                    sw.WriteLine(clsReadVAValues.strGECLVA2)
                    sw.WriteLine(clsReadVAValues.strGECLVA3)
                    sw.WriteLine(clsReadVAValues.strGECLEXE)
                End If
            ElseIf gstrpathProduct = "SPIA" Then
                If gstrCompCode(klbMismatchedClients.SelectedItem) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICSPIANNVER)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLICNYSPIAANNVER)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAEXE)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNRATES)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNPROD)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSYS)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIACPY)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIAANNSUPP)
                    sw.WriteLine(clsReadVAValues.strGLAICFIXEDSPIANNVER)
                End If
            ElseIf gstrpathProduct = "SPDA" Then
                If gstrCompCode(klbMismatchedClients.SelectedItem) = "GECA" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "GECL" Then
                    sw.WriteLine(clsReadVAValues.strGLICSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDARATE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strGLICSPDACPY)
                ElseIf gstrCompCode(klbMismatchedClients.SelectedItem) = "FCOL" Then
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAEXE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDARATE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDAGNAWINE)
                    sw.WriteLine(clsReadVAValues.strFCOLSPDACPY)
                ElseIf gstrpathProduct = "FIA" Then
                    sw.WriteLine(clsreadvavalues.strfiaexe)
                    sw.WriteLine(clsreadvavalues.strfiacpy)
                    sw.WriteLine(clsreadvavalues.strfiamdb)
                End If
            End If
            sw.WriteLine(vbCrLf)
            'write the mismatches to the file
            For i As Integer = 0 To splittest.Count - 2
                Dim splittestgrid = Split(splittest(i), ",")

                If splittestgrid.Count = 4 Then
                    str = "Variable/Element:" & ControlChars.Tab & splittestgrid(1) & "/ N/A"
                    sw.WriteLine(str)
                    str = "Bench Value/TestValue:" & ControlChars.Tab & splittestgrid(2) & "/" & splittestgrid(3)
                    sw.WriteLine(str)
                    sw.WriteLine("________")
                Else
                    str = "Variable/Element:" & ControlChars.Tab & splittestgrid(1) & "/" & splittestgrid(2)
                    sw.WriteLine(str)
                    str = "Bench Value/Test Value:" & ControlChars.Tab & splittestgrid(3) & "/" & splittestgrid(4)
                    sw.WriteLine(str)
                    sw.WriteLine("________")
                End If
            Next
            If value <> "" Then
                sw.WriteLine("Notes:  " & value)
            End If
            sw.Flush()
            sw.Close()

        End If

        MsgBox("The list of mismatches for " & gstrCase(0) & " Client # " & klbMismatchedClients.SelectedItem & " have been written to:  " & strComplPath & "mismatchdetails" & klbMismatchedClients.SelectedItem & ".txt", MsgBoxStyle.Information)
        Return
        'End If
    End Sub
    Private Sub VersionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox.Show()
    End Sub

    Private Sub KryptonButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbOpenCaseFolder.Click
        Dim strComplPath As String = gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0) & "\"
        Process.Start("explorer.exe", strComplPath)
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        InputBox("Input notes")
    End Sub
    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("C:\WinFlex6\GECADATA\gedefann.EXE").WaitForExit()
    End Sub

    Private Sub kbSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbSelectAll.Click
        'Select all
        For ix = 0 To kclbClientList.Items.Count - 1
            kclbClientList.SetItemChecked(ix, True)
        Next
    End Sub

    Private Sub kbClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbClearAll.Click
        'Unselect all
        For ix = 0 To kclbClientList.Items.Count - 1
            kclbClientList.SetItemChecked(ix, False)
        Next
    End Sub

    Private Sub klbCases_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles klbCases.SelectedValueChanged

        kbRun.Enabled = True
        If gstrpathProduct = "FIA" Then
        Else
            kbPalomaOnly.Enabled = True
        End If
        For ix = 0 To klbCases.Items.Count - 1
            If klbCases.GetSelected(ix) Then
                gstrCase(0) = klbCases.GetItemText(klbCases.Items(ix))
            End If
        Next ix
        kclbClientList.Items.Clear()  'new client list box
        FillClientList("\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & gstrpathProduct & "\" & gstrCase(0))

    End Sub
    Private Sub klbMismatchedClients_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles klbMismatchedClients.SelectedValueChanged


        If kbViewIllustration.Enabled = True And kbViewIllustration.Visible = True Then
            kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClients.SelectedItem
        End If

        DisplayMismatches()

    End Sub
    Private Sub klbMismatchedClientsfiatool_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles klbMismatchedClientsFIATool.SelectedValueChanged


        'If kbViewIllustration.Enabled = True And kbViewIllustration.Visible = True Then
        '    kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClientsFIATool.SelectedItem
        'End If

        DisplayMismatchesFIATool()

    End Sub
    Private Sub kclbClientList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kclbClientList.SelectedIndexChanged

    End Sub

    Private Sub KryptonRadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Select all
        For ix = 0 To kclbClientList.Items.Count - 1
            kclbClientList.SetItemChecked(ix, True)
        Next
    End Sub

    Private Sub KryptonRadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Unselect all
        For ix = 0 To kclbClientList.Items.Count - 1
            kclbClientList.SetItemChecked(ix, False)
        Next
    End Sub

    Private Sub klBenchMarkDate_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub kbPaloma_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbPaloma.Click

        Dim iFolders As Integer

        iFolders = FolderCount(gstrpath & "\" & gstrpathProduct & "\" & gstrCase(0))

        If gstrpathProduct = "FIA" Then
            kbPaloma.Enabled = False
        Else

            Dim icount As Integer = 0
            For ix = 0 To kclbClientList.Items.Count - 1
                If kclbClientList.GetItemCheckState(ix) Then
                    icount = icount + 1
                End If
            Next
            If icount = 0 Then
                MsgBox("Please select at least one client to run Paloma Compare", MsgBoxStyle.Critical)
            Else
                ShowPaloma(icount) 'iFolders)
            End If
        End If
    End Sub
    Private Sub ShowPaloma(ByVal ifolders As Integer)
        Dim cCtrl As Control
        For Each cCtrl In Me.Controls
            If cCtrl.TabIndex <> 20 Then
                cCtrl.Enabled = False
            End If
        Next

        'open the form to create new benchmarks
        Form5.Show()

        'fill the DataGrid with ALL of the clients SELECTED to run in the case, whether they match or not

        For ix = 0 To kclbClientList.Items.Count - 1
            Dim item As New DataGridViewRow
            item.CreateCells(Form5.kdgvPalomaStatus)
            If kclbClientList.GetItemCheckState(ix) Then
                Dim strReport As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\Control\Report.pdf" & ""
                Dim strBench As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\Control\bench.txt" & ""
                Dim strLastRegCompare As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\runstatscase" & (ix + 1) & ".txt" & ""
                Dim strLastRegBench As String = "\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\lastbenchmarkdate" & (ix + 1) & ".txt" & ""

                Dim objFileInfoLastCompare As New FileInfo(strReport)
                Dim objfileinfoLastBench As New FileInfo(strBench)
                Dim objFileInfoLastCompareReg As New FileInfo(strLastRegCompare)
                Dim objFileInfoLastBenchReg As New FileInfo(strLastRegBench)
                Dim DateModified As DateTime = objFileInfoLastCompare.LastWriteTime
                Dim LastBench As DateTime = objfileinfoLastBench.LastWriteTime
                Dim LastRegComp As DateTime = objFileInfoLastCompareReg.LastWriteTime
                Dim LastRegBench As DateTime = objFileInfoLastBenchReg.LastWriteTime

                If FileIO.FileSystem.FileExists("\\ricfiles\RI&I\ProductDevelopment\Illustrations Shared Drive\BUGS\Regression\" & Regression.RegressionMain.gstrpathProduct & "\" & Regression.RegressionMain.gstrCase(0) & "\" & (ix + 1) & "\Test\Test.pdf") Then

                    item.Cells(1).Value = CStr(ix + 1)
                    item.Cells(2).Value = "Uncompared"
                    If DateModified = #12/31/1600 7:00:00 PM# Then
                        item.Cells(3).Value = "None Yet"
                    Else
                        item.Cells(3).Value = DateModified
                    End If
                    If LastBench = #12/31/1600 7:00:00 PM# Then
                        item.Cells(4).Value = "None Yet"
                    Else
                        item.Cells(4).Value = LastBench
                    End If
                    If LastRegComp = #12/31/1600 7:00:00 PM# Then
                        item.Cells(5).Value = "None Yet"
                    Else
                        item.Cells(5).Value = LastRegComp
                    End If
                    If LastRegBench = #12/31/1600 7:00:00 PM# Then
                        item.Cells(6).Value = "None Yet"
                    Else
                        item.Cells(6).Value = LastRegBench
                    End If

                    Form5.kdgvPalomaStatus.Rows.Add(item)

                Else
                    item.Cells(1).Value = CStr(ix + 1)
                    item.Cells(2).Value = "--"
                    item.Cells(3).Value = "--"
                    item.Cells(4).Value = "--"
                    If LastRegComp = #12/31/1600 7:00:00 PM# Then
                        item.Cells(5).Value = "None Yet"
                    Else
                        item.Cells(5).Value = LastRegComp
                    End If
                    If LastRegBench = #12/31/1600 7:00:00 PM# Then
                        item.Cells(6).Value = "None Yet"
                    Else
                        item.Cells(6).Value = LastRegBench
                    End If
                    Form5.kdgvPalomaStatus.Rows.Add(item)
                End If

            End If
        Next
        
        Form5.ifolders = ifolders

    End Sub

    Private Sub kbPalomaOnly_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kbPalomaOnly.Click

        kbRun.Enabled = False
        bPalomaOnly = True
        RunPalomaOnly()
    End Sub
    Private Sub CalculateAgeLastSingle(ByVal pathname As String, ByVal strDOB1 As String, ByVal iAge1 As Integer, ByVal iclient As Integer)

        'Single Annuitant

        Dim iStartYear As Integer
        Dim iStartMonth As Integer
        Dim iStartDay As Integer
        Dim iBirthYear1 As Integer
        Dim iBirthMonth1 As Integer
        Dim iBirthDay1 As Integer
        Dim iAgeNew1 As Integer
        Dim strDOBNew1 As String
        Dim dtDOB1 As Date
        Dim strDOBNew2 As String = ""
        Dim iAgeNew2 As Integer

        Dim strToday As String

        'Calculate the age(s) based on today's date

        strToday = Format(Today, "MM/dd/yyyy")

        'Today
        iStartYear = Val(Mid$(strToday, 7, 4))
        iStartMonth = Val(Mid$(strToday, 1, 2))
        iStartDay = Val(Mid$(strToday, 4, 2))

        'Birthday
        iBirthYear1 = Val(Mid$(strDOB1, 7, 4))
        iBirthMonth1 = Val(Mid$(strDOB1, 1, 2))
        iBirthDay1 = Val(Mid$(strDOB1, 4, 2))

        If iStartMonth > iBirthMonth1 Then
            iAgeNew1 = iStartYear - iBirthYear1
        ElseIf iStartMonth = iBirthMonth1 And iStartDay >= iBirthDay1 Then
            iAgeNew1 = iStartYear - iBirthYear1
        Else
            iAgeNew1 = iStartYear - iBirthYear1 - 1
        End If

        'Compare the age(s) as of today with the age from the last Benchmark

        If iAgeNew1 > iAge1 Then
            gbVASaveAge(iclient) = True
            gbFIASaveAge(iclient) = True
            dtDOB1 = CDate(strDOB1)
            gstrDOB1Original = CStr(dtDOB1)
            strDOBNew1 = dtDOB1.AddYears(1)
            gstrDOB1New = strDOBNew1
            ModifyRelayVASaveAge(pathname, iAgeNew1, strDOBNew1, iclient, iAgeNew2, strDOBNew2)
        Else
            gbVASaveAge(iclient) = False
            gbFIASaveAge(iclient) = False
        End If

    End Sub
    Private Sub CalculateAgeLastJoint(ByVal pathname As String, ByVal strDOB1 As String, ByVal iAge1 As Integer, ByVal iclient As Integer, ByVal strDOB2 As String, ByVal iAge2 As Integer)

        'Joint Annuitants

        Dim iStartYear As Integer
        Dim iStartMonth As Integer
        Dim iStartDay As Integer
        Dim iBirthYear1 As Integer
        Dim iBirthMonth1 As Integer
        Dim iBirthDay1 As Integer
        Dim iBirthYear2 As Integer
        Dim iBirthMonth2 As Integer
        Dim iBirthDay2 As Integer
        Dim iAgeNew1 As Integer
        Dim iAgeNew2 As Integer
        Dim strDOBNew1 As String = ""
        Dim strDOBNew2 As String = ""
        Dim dtDOB1 As Date
        Dim dtDOB2 As Date
        Dim strToday As String

        'Calculate the age(s) based on today's date

        strToday = Format(Today, "MM/dd/yyyy")

        'Today
        iStartYear = Val(Mid$(strToday, 7, 4))
        iStartMonth = Val(Mid$(strToday, 1, 2))
        iStartDay = Val(Mid$(strToday, 4, 2))

        'Birthday
        iBirthYear1 = Val(Mid$(strDOB1, 7, 4))
        iBirthMonth1 = Val(Mid$(strDOB1, 1, 2))
        iBirthDay1 = Val(Mid$(strDOB1, 4, 2))

        iBirthYear2 = Val(Mid$(strDOB2, 7, 4))
        iBirthMonth2 = Val(Mid$(strDOB2, 1, 2))
        iBirthDay2 = Val(Mid$(strDOB2, 4, 2))

        If iStartMonth > iBirthMonth1 Then
            iAgeNew1 = iStartYear - iBirthYear1
        ElseIf iStartMonth = iBirthMonth1 And iStartDay >= iBirthDay1 Then
            iAgeNew1 = iStartYear - iBirthYear1
        Else
            iAgeNew1 = iStartYear - iBirthYear1 - 1
        End If

        If iStartMonth > iBirthMonth2 Then
            iAgeNew2 = iStartYear - iBirthYear2
        ElseIf iStartMonth = iBirthMonth2 And iStartDay >= iBirthDay2 Then
            iAgeNew2 = iStartYear - iBirthYear2
        Else
            iAgeNew2 = iStartYear - iBirthYear2 - 1
        End If

        'Compare the age(s) as of today with the age from the last Benchmark

        If iAgeNew1 > iAge1 Or iAgeNew2 > iAge2 Then
            gbVASaveAge(iclient) = True
            gbFIASaveAge(iclient) = True
            If iAgeNew1 > iAge1 Then
                dtDOB1 = CDate(strDOB1)
                gstrDOB1Original = CStr(dtDOB1)
                strDOBNew1 = dtDOB1.AddYears(1)
                gstrDOB1New = strDOBNew1
            Else
                dtDOB1 = CDate(strDOB1)
                strDOBNew1 = dtDOB1
            End If
            If iAgeNew2 > iAge2 Then
                dtDOB2 = CDate(strDOB2)
                gstrDOB2Original = CStr(dtDOB2)
                strDOBNew2 = dtDOB2.AddYears(1)
                gstrDOB2New = strDOBNew2
            Else
                dtDOB2 = CDate(strDOB2)
                strDOBNew2 = dtDOB2
            End If
            ModifyRelayVASaveAge(pathname, iAgeNew1, strDOBNew1, iclient, iAgeNew2, strDOBNew2)
        Else
            gbVASaveAge(iclient) = False
            gbFIASaveAge(iclient) = False
        End If

    End Sub

    Private Sub KryptonLabel5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles KryptonLabel5.Paint

    End Sub

    Private Sub kdgvRegressionStatus_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles kdgvRegressionStatus.CellContentClick

    End Sub

    Private Sub kbNewBench_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles kbNewBench.Disposed

    End Sub

    'Private Sub Button1_Click_3(sender As System.Object, e As System.EventArgs) Handles Button1.Click
    '    FIATool.FillToolValues()


    'End Sub

    Private Sub klbMismatchedClientsFIATool_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles klbMismatchedClientsFIATool.SelectedIndexChanged



        DisplayMismatchesFIATool()

        'Dim strSep As String = ""

        ''Try
        ''    'display the values that do not match on the selected case
        ''    If klbMismatchedClientsFIATool.SelectedIndex = -1 Then
        ''        MsgBox("Please select a client in order to view the mismatches.", MsgBoxStyle.Critical)
        ''    End If
        ''Catch ex As Exception
        ''End Try
        ''parse the string to fill the data grid
        'Dim splittest = Split(strClientXMisMatchFIATool(klbMismatchedClientsFIATool.SelectedItem), "&")
        'Dim splittestrow(splittest.Count, 5)
        'Dim strmmlistnew As String = ""

        'Dim strMMList As String = ""
        'Dim ix As Integer = 0
        'Dim iy As Integer = 0

        ''initialize the data grid
        'kdgvFIATool.RowCount = 0

        'For ix = 0 To splittest.Count - 2
        '    For iy = 0 To 4
        '        If InStr(splittest(ix), "�") Then
        '            splittestrow(ix, iy) = Split(splittest(ix), "�")
        '        Else
        '            splittestrow(ix, iy) = Split(splittest(ix), ",")

        '        End If

        '    Next
        'Next


        ''fill the data grid with the values
        'For i As Integer = 0 To splittest.Count - 2
        '    If InStr(splittest(i), "�") Then
        '        strSep = "�"
        '    Else
        '        strSep = ","
        '    End If
        '    Dim splittestgrid = Split(splittest(i), strSep)

        '    Dim item As New DataGridViewRow
        '    item.CreateCells(kdgvFIATool)
        '    With item
        '        .Cells(0).Value = splittestgrid(0)
        '        .Cells(1).Value = splittestgrid(1)
        '        If splittestgrid.Count = 4 Then
        '            .Cells(2).Value = ""
        '            .Cells(3).Value = splittestgrid(2)
        '            .Cells(4).Value = splittestgrid(3)
        '        Else
        '            .Cells(2).Value = splittestgrid(2)
        '            .Cells(3).Value = splittestgrid(3)
        '            .Cells(4).Value = splittestgrid(4)
        '        End If

        '    End With
        '    kdgvFIATool.Rows.Add(item)
        'Next

        'If klbMismatchedClientsFIATool.SelectedIndex <> -1 And kdgvFIATool.RowCount > 0 Then
        '    'enable the button to view the illustration pages
        '    'kbViewIllustration.Visible = True
        '    'kbViewIllustration.Enabled = True
        '    'KryptonBorderEdge4.Visible = True
        '    'KryptonBorderEdge4.Enabled = True

        '    'set text on button corresponding to the selected case
        '    'kbViewIllustration.Text = "View Illustration for Client # " & klbMismatchedClients.SelectedItem

        '    'enable the button to save the mismatch results
        '    kbSaveMismatchResults.Enabled = True
        'End If
    End Sub

    Private Sub Button1_Click_3(sender As System.Object, e As System.EventArgs)

       





    End Sub

    Private Sub kbClearToolMM_Click(sender As System.Object, e As System.EventArgs) Handles kbClearToolMM.Click

        'clear out the array of FIA TOOL mismatches and skips

        Array.Clear(strsplitmmlistFIATool, 0, strsplitmmlistFIATool.Length)
        Array.Resize(strsplitmmlistFIATool, 0)

        Array.Clear(ReadFIARelayINI.strsplitFIAToolSkipList, 0, ReadFIARelayINI.strsplitFIAToolSkipList.Length)
        Array.Resize(ReadFIARelayINI.strsplitFIAToolSkipList, 0)

        'clear the mismatch list box
        strClientMisMatchListFIATool = ""
        ReadFIARelayINI.strFIAToolSkipList = ""

        'set mismatches back to none

        ReadFIARelayINI.bMismatchToolAtLeastOnce = False
        ReadFIARelayINI.bMisMatchTool = False

        ReadFIARelayINI.bFIAToolSkippedAtLeastOnce = False


        'clear and reset the controls on the form
        klbMismatchedClientsFIATool.Items.Clear()
        klbMismatchedClientsFIATool.Enabled = False
        kdgvFIATool.RowCount = 0
        kbClearToolMM.Enabled = False
        ReadFIARelayINI.bNoToolRun = False

        klbFIAToolSkippedList.Items.Clear()
        klbFIAToolSkippedList.Enabled = False


    End Sub

    Private Sub KryptonLabel8_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles KryptonLabel8.Paint

    End Sub
End Class