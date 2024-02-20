Imports Nini.Config
Imports System.IO
Imports System.Text

Public Class ReadFIARelayINI

    Public Shared strToolVersion As String
    Public Shared ID = System.Security.Principal.WindowsIdentity.GetCurrent
    Public Shared username = ID.Name.Substring(ID.Name.IndexOf("\") + 1)

    Public Shared b10Y As Boolean = False
    Public Shared bNoToolRun As Boolean = False

    Public Shared objExcel As New Microsoft.Office.Interop.Excel.Application
    Public Shared Books As Microsoft.Office.Interop.Excel.Workbooks

    Public Shared Book As Microsoft.Office.Interop.Excel.Workbook = objExcel.Workbooks.Open("C:\Documents and Settings\" & username & "\Desktop\FIATool.xlsm")


    Public Shared Sheet As Microsoft.Office.Interop.Excel.Worksheet

    Public Shared strFIAToolSkipList As String = ""
    Public Shared strsplitFIAToolSkipList(100) As String


    Public Shared strYearsToPrint As String
    Public Shared strParse(3) As String

    'variables for relay.ini 
    Public Shared Identifiers As IConfig
    Public Shared Engine As IConfig

    Public Shared strFIAINIINcomeRider As String 'B37
    Public Shared strFIAINIJointYN As String 'B15
    Public Shared strFIAINIStartIncomeMonth As String 'B22
    Public Shared strFIAINIStartIncomePeriodYN As String ' need this?
    Public Shared strFIAINIStartIncomeYear As String 'B23
    Public Shared strFIAINIWithdrawalFreq As String 'B41, B47
    Public Shared strFIAINIIntCreditStrategyTotal As String 'need this?
    Public Shared strFIAINI7YearFxRate As String 'B56
    Public Shared strFIAINI10YearFxRate As String 'B56
    Public Shared strFIAINIAnnualCap As String 'B57
    Public Shared strFIAINIMonthlyCap As String 'B58
    Public Shared strFIAINIPerformanceTrigger As String 'B59
    Public Shared strFIAINIInsuredDateOfBirth As String 'B9. B10. B11
    Public Shared strFIAINIInsuredSex As String 'B14
    Public Shared strFIAINIPremium As String 'B5
    Public Shared strFIAINIPrintYears As String 'B66
    Public Shared strFIAINISolveFor As String 'B4
    Public Shared strFIAINIInsured2DateOfBirth As String 'B16, B17, B18
    Public Shared strFIAINIInsured2Sex As String 'B21
    Public Shared strFIAINIProdName As String 'B2
    Public Shared strFIAINIIncomePayment As String 'B5
    Public Shared strFIAINIWDPct As String 'B45
    Public Shared strFIAINIWDType As String 'B44
    Public Shared strFIAINIWithdrawalAmt As String 'B46...B48, B49
    Public Shared strFIAINIInsuredState As String 'based on state, look up non forf rates, etc
    Public Shared iFIAINIInsuredAge As Integer
    Public Shared iFIAINIInsured2Age As Integer

    'Tool Values

    'variables for Tool values read from FIA Tool

    Public Shared strFIASpecSPChangeTool As String
    Public Shared strFIASpecWDTool As String
    Public Shared strFIASpecAnnCreditRateTool As String
    Public Shared strFIASpecContractValueTool As String
    Public Shared strFIASpecSurrenderValueTool As String
    Public Shared strFIASpecMGSVTool As String
    Public Shared strFIASpecProjBeneBaseTool As String
    Public Shared strFIASpecProjWDLimitTool As String
    Public Shared strFIASevenYearIntRateTool As String
    Public Shared strFIATenYearIntRateTool As String
    Public Shared strFIAMonthlyCapIndexCreditTool As String
    Public Shared strFIAAnnualCapIndexCreditTool As String
    Public Shared strFIAPerfTriggerIndexCreditTool As String
    Public Shared strFIASevenYearAccumValueTool As String
    Public Shared strFIATenYearAccumValueTool As String
    Public Shared strFIAMonthlyCapAccumValueTool As String
    Public Shared strFIAAnnualCapAccumValueTool As String
    Public Shared strFIAPerfTriggerAccumValueTool As String
    Public Shared strFIAContractValueNoWDTool As String
    Public Shared strFIAGuarWDFactorTool As String
    Public Shared strFIAGuarBeneBaseNoWDTool As String
    Public Shared strFIAGuarWDLimitNoWDTool As String
    Public Shared strFIAProjBeneBaseNoWDTool As String
    Public Shared strFIAProjWDLimitNoWDTool As String
    Public Shared strFIAFavSPChangeTool As String
    Public Shared strFIAUnfavSPChangeTool As String
    Public Shared strFIAFavWDTool As String
    Public Shared strFIAUnfavWDTool As String
    Public Shared strFIAFavAnnCreditRateTool As String
    Public Shared strFIAUnfavAnnCreditRateTool As String
    Public Shared strFIAFavContractValueTool As String
    Public Shared strFIAUnfavContractValueTool As String
    Public Shared strFIAFavSurrenderValueTool As String
    Public Shared strFIAUnfavSurrenderValueTool As String
    Public Shared strFIAFavMGSVTool As String
    Public Shared strFIAUnfavMGSVTool As String
    Public Shared strFIAFavProjBeneBaseTool As String
    Public Shared strFIAFavProjWDLimitTool As String
    Public Shared strFIAUnfavProjBeneBaseTool As String
    Public Shared strFIAUnfavProjWDLimitTool As String
    Public Shared strFIAGMCVTool As String

    'FOR TOOL MISMATCHES

    Public Shared bMisMatchTool As Boolean
    Public Shared bMismatchToolAtLeastOnce As Boolean

    Public Shared strFIASpecSPChangeToolMM As String
    Public Shared strFIASpecWDToolMM As String
    Public Shared strFIASpecAnnCreditRateToolMM As String
    Public Shared strFIASpecContractValueToolMM As String
    Public Shared strFIASpecSurrenderValueToolMM As String
    Public Shared strFIASpecMGSVToolMM As String
    Public Shared strFIASpecProjBeneBaseToolMM As String
    Public Shared strFIASpecProjWDLimitToolMM As String
    Public Shared strFIASevenYearIntRateToolMM As String
    Public Shared strFIATenYearIntRateToolMM As String
    Public Shared strFIAMonthlyCapIndexCreditToolMM As String
    Public Shared strFIAAnnualCapIndexCreditToolMM As String
    Public Shared strFIAPerfTriggerIndexCreditToolMM As String
    Public Shared strFIASevenYearAccumValueToolMM As String
    Public Shared strFIATenYearAccumValueToolMM As String
    Public Shared strFIAMonthlyCapAccumValueToolMM As String
    Public Shared strFIAAnnualCapAccumValueToolMM As String
    Public Shared strFIAPerfTriggerAccumValueToolMM As String
    Public Shared strFIAContractValueNoWDToolMM As String
    Public Shared strFIAGuarWDFactorToolMM As String
    Public Shared strFIAGuarBeneBaseNoWDToolMM As String
    Public Shared strFIAGuarWDLimitNoWDToolMM As String
    Public Shared strFIAProjBeneBaseNoWDToolMM As String
    Public Shared strFIAProjWDLimitNoWDToolMM As String
    Public Shared strFIAFavSPChangeToolMM As String
    Public Shared strFIAUnfavSPChangeToolMM As String
    Public Shared strFIAFavWDToolMM As String
    Public Shared strFIAUnfavWDToolMM As String
    Public Shared strFIAFavAnnCreditRateToolMM As String
    Public Shared strFIAUnfavAnnCreditRateToolMM As String
    Public Shared strFIAFavContractValueToolMM As String
    Public Shared strFIAUnfavContractValueToolMM As String
    Public Shared strFIAFavSurrenderValueToolMM As String
    Public Shared strFIAUnfavSurrenderValueToolMM As String
    Public Shared strFIAFavMGSVToolMM As String
    Public Shared strFIAUnfavMGSVToolMM As String
    Public Shared strFIAFavProjBeneBaseToolMM As String
    Public Shared strFIAFavProjWDLimitToolMM As String
    Public Shared strFIAUnfavProjBeneBaseToolMM As String
    Public Shared strFIAUnfavProjWDLimitToolMM As String
    Public Shared strFIAGMCVToolMM As String

    Public Shared bFIAToolSkippedAtLeastOnce As Boolean
   

    Public Sub ReadFIAInputs(ByVal strToolPath As String, ByVal strRelayOutPath As String, ib As Integer)

        'read the values from the relay.ini file

        Dim RelayINI = New IniConfigSource(strToolPath)
        Dim RelayOut = New IniConfigSource(strRelayOutPath)


        Dim icomma As Integer = 0
        Dim ir As Integer
        Dim strSplits() As String
        Dim isplit As Integer
        'Dim strStrategyTotal As String
        'Dim bNoToolRun As Boolean = False
        Dim strRelayOutMessage As String
        Dim iLength As Integer
        Dim ixsep As Integer
       


        'Dim attrib As System.IO.FileAttributes = System.IO.File.GetAttributes("C:\Documents and Settings\" & username & "\Desktop\FIATool.xlsm")
        'Dim myfileinfo As FileInfo = New FileInfo("C:\Documents and Settings\" & username & "\Desktop\FIATool.xlsm")

        'Dim fileinformation As String = myfileinfo.CreationTime

        'MsgBox(fileinformation)



        'variables for relay.ini
        Identifiers = RelayINI.Configs("Identifiers")

        'For Relay.out
        Engine = RelayOut.Configs("Engine")

        strRelayOutMessage = Engine.Get("Message_1")

        If InStr(strRelayOutMessage, "ERROR") Then
            bNoToolRun = True
        Else
            bNoToolRun = False
        End If


        strFIAINIProdName = Identifiers.Get("pol_ProductCode")
        If strFIAINIProdName = "FIA7L" Then
            strFIAINIProdName = "FIA 7 Yr"
        ElseIf strFIAINIProdName = "FIA10L" Then
            strFIAINIProdName = "FIA 10 Yr"
        End If
        strFIAINIJointYN = Identifiers.Get("Annuity.JointYN")
        If strFIAINIJointYN = "Y" Then
            strFIAINIJointYN = "Yes"
        End If
        strFIAINIINcomeRider = Identifiers.Get("Annuity.IncomeRider")
        If strFIAINIINcomeRider = "Y" Then
            strFIAINIINcomeRider = "Yes"
        ElseIf strFIAINIINcomeRider = "N" Then
            strFIAINIINcomeRider = "No"

        End If
        strFIAINIStartIncomeMonth = Identifiers.Get("Annuity.StartIncomeMonth")
        If strFIAINIStartIncomeMonth = "PA" Then
            strFIAINIStartIncomeMonth = "Contract Anniversary"
        ElseIf strFIAINIStartIncomeMonth = "YB" Then
            strFIAINIStartIncomeMonth = "Birth Month"
        End If
        strFIAINIStartIncomePeriodYN = Identifiers.Get("Annuity.StartIncomePeriod.YN")
        strFIAINIStartIncomeYear = Identifiers.Get("Annuity.StartIncomeYear")
        strFIAINIWithdrawalFreq = Identifiers.Get("Annuity.WithdrawalFreq")
        If strFIAINIWithdrawalFreq = "A" Then
            strFIAINIWithdrawalFreq = "Annual"
        ElseIf strFIAINIWithdrawalFreq = "M" Then
            strFIAINIWithdrawalFreq = "Monthly"
        End If
        strFIAINIIntCreditStrategyTotal = Identifiers.Get("FCOL.IntCreditStrategy.Total")
        strFIAINI7YearFxRate = Identifiers.Get("FUND.7YrFxRate")
        strFIAINI10YearFxRate = Identifiers.Get("FUND.10YrFxRate")
        strFIAINIAnnualCap = Identifiers.Get("FUND.AnnualCap")
        strFIAINIMonthlyCap = Identifiers.Get("FUND.MonthlyCap")
        strFIAINIPerformanceTrigger = Identifiers.Get("FUND.PerformanceTrigger")
        strFIAINIInsuredDateOfBirth = Identifiers.Get("Insured.DateOfBirth")
        iFIAINIInsuredAge = CInt(Identifiers.Get("Insured.Age"))
        strFIAINIInsuredSex = Identifiers.Get("Insured.Sex")
        strFIAINIPremium = Identifiers.Get("Policy.Premium")
        If strFIAINIPremium = Nothing Then
        Else

            strFIAINIPremium = strFIAINIPremium.Substring(6)
        End If
        strFIAINIPrintYears = Identifiers.Get("Policy.PrintYears")
        If strFIAINIPrintYears = "10Y" Then
            b10Y = True
        Else
            b10Y = False
        End If
        strFIAINISolveFor = Identifiers.Get("Policy.SolveFor")
        If strFIAINISolveFor = "IP" Then
            strFIAINISolveFor = "Income Payment"
        ElseIf strFIAINISolveFor = "PA" Then
            strFIAINISolveFor = "Single Premium Payment"
        End If
        strFIAINIInsured2DateOfBirth = Identifiers.Get("Insured2.DateOfBirth")
        iFIAINIInsured2Age = CInt(Identifiers.Get("Insured2.Age"))
        strFIAINIInsured2Sex = Identifiers.Get("Insured2.Sex")
        strFIAINIIncomePayment = Identifiers.Get("Annuity.IncomePayment")

        If strFIAINIINcomeRider = "No" And bNoToolRun = False Then
            strFIAINIWDType = Identifiers.Get("Annuity.WDType")
            If strFIAINIWDType = "1" Then
                strFIAINIWDType = "Percentage"
                strFIAINIWDPct = Identifiers.Get("Annuity.WDPct")
                For iLength = 1 To strFIAINIWDPct.Length 'if multiple amounts, dont run tool
                    If Mid$(strFIAINIWDPct, iLength, 1) = "�" Then
                        ixsep = ixsep + 1
                    End If
                Next
                If ixsep > 2 Then
                    bNoToolRun = True
                Else
                    bNoToolRun = False
                    strSplits = strFIAINIWDPct.Split("�") '("ü")
                    For isplit = 0 To strSplits.Length - 1
                        strParse(isplit) = strSplits(isplit)
                    Next
                End If
            ElseIf strFIAINIWDType = "3" Then
                strFIAINIWDType = "Amount"
                strFIAINIWithdrawalAmt = Identifiers.Get("Annuity.WithdrawalAmt")
                For iLength = 1 To strFIAINIWithdrawalAmt.Length 'if multiple amounts, dont run tool
                    If Mid$(strFIAINIWithdrawalAmt, iLength, 1) = "�" Then
                        ixsep = ixsep + 1
                    End If
                Next

                If ixsep > 2 Then
                    bNoToolRun = True
                Else
                    bNoToolRun = False
                    strSplits = strFIAINIWithdrawalAmt.Split("�") '("ü")
                    For isplit = 0 To strSplits.Length - 1
                        strParse(isplit) = strSplits(isplit)
                    Next
                End If

            End If

        End If

        strFIAINIInsuredState = Identifiers.Get("Insured.State")

        ''The below checks are for cases when WinFlex test doesnt run, to stop the tool from running

        'strStrategyTotal = Identifiers.Get("FCOL.IntCreditStrategy.Total")
        'If strStrategyTotal = "100" Then
        '    'bNoToolRun = False
        'Else
        '    bNoToolRun = True
        'End If

        'If strFIAINIProdName = "FIA 7 Yr" Then
        '    If strFIAINIInsuredState = "OK" Then
        '        If iFIAINIInsuredAge > 80 Or (strFIAINIJointYN = "Yes" And iFIAINIInsured2Age > 80) Then
        '            bNoToolRun = True
        '        Else
        '            'bNoToolRun = False
        '        End If
        '    Else
        '        If iFIAINIInsuredAge > 85 Or (strFIAINIJointYN = "Yes" And iFIAINIInsured2Age > 85) Then
        '            bNoToolRun = True
        '        Else
        '            'bNoToolRun = False
        '        End If
        '    End If
        'Else
        '    If strFIAINIProdName = "FIA 10 Yr" Then
        '        If iFIAINIInsuredAge > 80 Or (strFIAINIJointYN = "Yes" And iFIAINIInsured2Age > 80) Then
        '            bNoToolRun = True
        '        Else
        '            'bNoToolRun = False
        '        End If
        '    End If
        'End If

        'If strFIAINIINcomeRider = "Yes" Then
        '    If iFIAINIInsuredAge < 55 Or iFIAINIInsuredAge > 80 Or (strFIAINIJointYN = "Yes" And (iFIAINIInsured2Age < 55 Or iFIAINIInsured2Age > 80)) Then
        '        bNoToolRun = True
        '    Else
        '        'bNoToolRun = False
        '    End If
        'End If

        If bNoToolRun = True Then
            bFIAToolSkippedAtLeastOnce = True
            strFIAToolSkipList = strFIAToolSkipList & "," & ib
            strsplitFIAToolSkipList = Split(strFIAToolSkipList, ",")


        Else
            If strFIAINIINcomeRider = "No" Then
                For ir = 1 To 3
                    FillFIAToolValues(ir)
                Next
            Else
                For ir = 1 To 4
                    FillFIAToolValues(ir)

                Next
            End If
        End If

    End Sub
    Public Sub New()

    End Sub
    Public Sub FillFIAToolValues(ir As Integer)

        Dim strYearsToPrint As String

        'If ir = 1 And Book.Count = 0 Then


        '    Try
        '        Book = objExcel.Workbooks.Open("C:\Documents and Settings\" & username & "\Desktop\FIATool.xlsm")
        '    Catch ex As System.Runtime.InteropServices.COMException
        '        MsgBox(ex.Message)

        '        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
        '            proc.Kill()
        '        Next

        '        End


        '    Finally

        '    End Try
        'End If

        Sheet = DirectCast(Book.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        DirectCast(Book.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet).Activate()


        strToolVersion = Sheet.Cells.Range("C2").Value

        'Write values to cells

        If b10Y Then
            strYearsToPrint = strFIAINIStartIncomeYear + 10
        Else
            If CInt(strFIAINIPrintYears) <= CInt(strFIAINIStartIncomeYear) Then
                strYearsToPrint = strFIAINIStartIncomeYear + 1
            Else
                strYearsToPrint = strFIAINIPrintYears
            End If
        End If

        Sheet.Cells.Range("B2").Value = strFIAINIProdName
        Sheet.Cells.Range("B4").Value = strFIAINISolveFor
        If strFIAINISolveFor = "Income Payment" Then
            Sheet.Cells.Range("B5").Value = strFIAINIPremium
        Else

            Sheet.Cells.Range("B5").Value = strFIAINIIncomePayment
        End If
        If ir = 1 Then
            Sheet.Cells.Range("B6").Value = DatePart("yyyy", Now)
            Sheet.Cells.Range("B7").Value = DatePart("m", Now)
            Sheet.Cells.Range("B8").Value = DatePart("d", Now)
            Sheet.Cells.Range("B9").Value = DatePart("m", strFIAINIInsuredDateOfBirth)
            Sheet.Cells.Range("B10").Value = DatePart("d", strFIAINIInsuredDateOfBirth)
            Sheet.Cells.Range("B11").Value = DatePart("yyyy", strFIAINIInsuredDateOfBirth)
            Sheet.Cells.Range("B14").Value = strFIAINIInsuredSex

            If strFIAINIJointYN = "Yes" Then
                Sheet.Cells.Range("B15").Value = strFIAINIJointYN
                Sheet.Cells.Range("B16").Value = DatePart("m", strFIAINIInsured2DateOfBirth)
                Sheet.Cells.Range("B17").Value = DatePart("d", strFIAINIInsured2DateOfBirth)
                Sheet.Cells.Range("B18").Value = DatePart("yyyy", strFIAINIInsured2DateOfBirth)
                Sheet.Cells.Range("B21").Value = strFIAINIInsured2Sex
            Else
                Sheet.Cells.Range("B15").Value = "No"
            End If


            'Non Forf/mgsv by state:
            If strFIAINIProdName = "FIA 10 Yr" Then
                If strFIAINIInsuredState = "NJ" Or strFIAINIInsuredState = "OK" Or strFIAINIInsuredState = "UT" Or strFIAINIInsuredState = "TX" Then
                    Sheet.Cells.Range("B33").Value = 0.015
                    Sheet.Cells.Range("B34").Value = 0.015
                ElseIf strFIAINIInsuredState = "IA" Or strFIAINIInsuredState = "CA" Then
                    Sheet.Cells.Range("B33").Value = 0.0135
                    Sheet.Cells.Range("B34").Value = 0.0135

                Else
                    Sheet.Cells.Range("B33").Value = 0.01
                    Sheet.Cells.Range("B34").Value = 0.01
                End If
            ElseIf strFIAINIProdName = "FIA 7 Yr" Then
                If strFIAINIInsuredState = "IA" Or strFIAINIInsuredState = "CA" Then
                    Sheet.Cells.Range("B33").Value = 0.012
                    Sheet.Cells.Range("B34").Value = 0.012
                Else
                    Sheet.Cells.Range("B33").Value = 0.01
                    Sheet.Cells.Range("B34").Value = 0.01
                End If
            End If

            If strFIAINIInsuredState = "NJ" Then
                Sheet.Cells.Range("B35").Value = 0.9
                Sheet.Cells.Range("B36").Value = 0.9
            Else
                Sheet.Cells.Range("B35").Value = 0.875
                Sheet.Cells.Range("B36").Value = 0.875

            End If
        End If


        If strFIAINIINcomeRider = "Yes" Then
            Sheet.Cells.Range("B22").Value = strFIAINIStartIncomeMonth
            If ir = 4 Then
                Sheet.Cells.Range("B23").Value = strYearsToPrint + 1
            Else
                If strFIAINIStartIncomePeriodYN = "N" Then
                    Sheet.Cells.Range("B23").Value = strYearsToPrint + 1
                Else
                    Sheet.Cells.Range("B23").Value = strFIAINIStartIncomeYear
                End If
            End If
            Sheet.Cells.Range("B39").Value = "Yes"
            Sheet.Cells.Range("B43").Value = strFIAINIWithdrawalFreq
        Else
            Sheet.Cells.Range("B4").Value = "Income Payment"
            Sheet.Cells.Range("B5").Value = strFIAINIPremium
            Sheet.Cells.Range("B39").Value = "No"
            Sheet.Cells.Range("B46").Value = strFIAINIWDType
            If strFIAINIWDType = "Percentage" Then
                Sheet.Cells.Range("B47").Value = strParse(2) / 100
                Sheet.Cells.Range("B49").Value = strFIAINIWithdrawalFreq
                Sheet.Cells.Range("B50").Value = strParse(0)
                Sheet.Cells.Range("B51").Value = strParse(1)
            ElseIf strFIAINIWDType = "Amount" Then
                Sheet.Cells.Range("B48").Value = strParse(2)
                Sheet.Cells.Range("B49").Value = strFIAINIWithdrawalFreq
                Sheet.Cells.Range("B50").Value = strParse(0)
                Sheet.Cells.Range("B51").Value = strParse(1)
            Else
                'Dummy values
                Sheet.Cells.Range("B46").Value = "Amount"
                Sheet.Cells.Range("B47").Value = "0"
                Sheet.Cells.Range("B48").Value = "0"
                Sheet.Cells.Range("B49").Value = "Annual"
                Sheet.Cells.Range("B50").Value = "0"
                Sheet.Cells.Range("B51").Value = "0"
            End If
        End If

        If ir = 1 Then
            If strFIAINIProdName = "FIA 7 Yr" Then
                Sheet.Cells.Range("B58").Value = strFIAINI7YearFxRate / 100
            Else
                Sheet.Cells.Range("B58").Value = strFIAINI10YearFxRate / 100
            End If
            Sheet.Cells.Range("B59").Value = strFIAINIAnnualCap / 100
            Sheet.Cells.Range("B60").Value = strFIAINIMonthlyCap / 100
            Sheet.Cells.Range("B61").Value = strFIAINIPerformanceTrigger / 100
        End If

        If ir = 1 Then
            If b10Y Then
                Sheet.Cells.Range("B67").Value = "Ten Year Income View"
            Else
                Sheet.Cells.Range("B67").Value = "Specified Years to Print"
                Sheet.Cells.Range("B68").Value = strYearsToPrint
            End If

            ReadFIAToolValues(ir, CInt(strYearsToPrint))
        ElseIf ir = 2 Then
            Sheet.Cells.Range("B67").Value = "Favorable"

            ReadFIAToolValues(ir, CInt(strYearsToPrint))
        ElseIf ir = 3 Then
            Sheet.Cells.Range("B67").Value = "Unfavorable"
            ReadFIAToolValues(ir, CInt(strYearsToPrint))

        ElseIf ir = 4 Then
            Sheet.Cells.Range("B67").Value = "Specified Years to Print"
            Sheet.Cells.Range("B68").Value = strYearsToPrint

            If strFIAINISolveFor = "Single Premium Payment" Then
                Sheet.Cells.Range("B23").Value = strFIAINIStartIncomeYear 'set start year back to original
                Sheet.Cells.Range("B5").Value = FormatCurrency((Sheet.Cells.Range("C63").Value), 2, , TriState.False) 'put in premium
                Sheet.Cells.Range("B4").Value = "Income Payment"
                Sheet.Cells.Range("B23").Value = strYearsToPrint + 1 'set it back for no wds
            End If

            ReadFIAToolValues(ir, CInt(strYearsToPrint))

        End If

    End Sub

    Public Sub ReadFIAToolValues(ir As Integer, iyearstoprint As Integer)

        Dim ix As Integer
        Dim iYears As Integer


        Dim strFIASpecSPChangeTemp As String = ""
        Dim strFIASpecWDTemp As String = ""
        Dim strFIASpecAnnCreditRateTemp As String = ""
        Dim strFIASpecContractValueTemp As String = ""
        Dim strFIASpecSurrenderValueTemp As String = ""
        Dim strFIASpecMGSVTemp As String = ""
        Dim strFIASpecProjBeneBaseTemp As String = ""
        Dim strFIASpecProjWDLimitTemp As String = ""
        Dim strFIASevenYearIntRateTemp As String = ""
        Dim strFIATenYearIntRateTemp As String = ""
        Dim strFIAMonthlyCapIndexCreditTemp As String = ""
        Dim strFIAAnnualCapIndexCreditTemp As String = ""
        Dim strFIAPerfTriggerIndexCreditTemp As String = ""
        Dim strFIASevenYearAccumValueTemp As String = ""
        Dim strFIATenYearAccumValueTemp As String = ""
        Dim strFIAMonthlyCapAccumValueTemp As String = ""
        Dim strFIAAnnualCapAccumValueTemp As String = ""
        Dim strFIAPerfTriggerAccumValueTemp As String = ""
        Dim strFIAContractValueNoWDTemp As String = ""
        Dim strFIAGuarWDFactorTemp As String = ""
        Dim strFIAGuarBeneBaseNoWDTemp As String = ""
        Dim strFIAGuarWDLimitNoWDTemp As String = ""
        Dim strFIAProjBeneBaseNoWDTemp As String = ""
        Dim strFIAProjWDLimitNoWDTemp As String = ""
        Dim strFIAFavSPChangeTemp As String = ""
        Dim strFIAUnfavSPChangeTemp As String = ""
        Dim strFIAFavWDTemp As String = ""
        Dim strFIAUnfavWDTemp As String = ""
        Dim strFIAFavAnnCreditRateTemp As String = ""
        Dim strFIAUnfavAnnCreditRateTemp As String = ""
        Dim strFIAFavContractValueTemp As String = ""
        Dim strFIAUnfavContractValueTemp As String = ""
        Dim strFIAFavSurrenderValueTemp As String = ""
        Dim strFIAUnfavSurrenderValueTemp As String = ""
        Dim strFIAFavMGSVTemp As String = ""
        Dim strFIAUnfavMGSVTemp As String = ""
        Dim strFIAFavProjBeneBaseTemp As String = ""
        Dim strFIAFavProjWDLimitTemp As String = ""
        Dim strFIAUnfavProjBeneBaseTemp As String = ""
        Dim strFIAUnfavProjWDLimitTemp As String = ""
        Dim strFIAGMCVTemp As String = ""


        Sheet = DirectCast(Book.Worksheets(4), Microsoft.Office.Interop.Excel.Worksheet)
        DirectCast(Book.Sheets(4), Microsoft.Office.Interop.Excel.Worksheet).Activate()



        If b10Y Then
            iYears = CInt(strFIAINIStartIncomeYear) + 10
        Else
            iYears = iyearstoprint
        End If

        Select Case (ir)
            Case 2
                iYears = 10
            Case 3
                iYears = 10
        End Select


        If ir = 1 Then
            If strFIAINIProdName = "FIA 7 Yr" Then
                strFIAGMCVTemp = FormatCurrency((Sheet.Cells.Range("C" & 42).Value), 0, , , TriState.False)
            End If
        End If

        For ix = 1 To iYears
            If ir = 1 Then 'Specified
                strFIASpecWDTemp = strFIASpecWDTemp & FormatCurrency((Sheet.Cells.Range("C" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIASpecMGSVTemp = strFIASpecMGSVTemp & FormatCurrency((Sheet.Cells.Range("D" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIASpecSPChangeTemp = strFIASpecSPChangeTemp & FormatPercent(Sheet.Cells.Range("E" & ix + 4).Value) & ","
                strFIASpecAnnCreditRateTemp = strFIASpecAnnCreditRateTemp & FormatPercent(Sheet.Cells.Range("F" & ix + 4).Value) & ","
                strFIASpecContractValueTemp = strFIASpecContractValueTemp & FormatCurrency((Sheet.Cells.Range("G" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIASpecSurrenderValueTemp = strFIASpecSurrenderValueTemp & FormatCurrency((Sheet.Cells.Range("H" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIASpecProjBeneBaseTemp = strFIASpecProjBeneBaseTemp & FormatCurrency((Sheet.Cells.Range("I" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIASpecProjWDLimitTemp = strFIASpecProjWDLimitTemp & FormatCurrency((Sheet.Cells.Range("K" & ix + 4).Value), 0, , , TriState.False) & ","

                If strFIAINIProdName = "FIA 7 Yr" Then
                    If strFIAINI7YearFxRate > 0 Then
                        strFIASevenYearIntRateTemp = strFIASevenYearIntRateTemp & FormatPercent(Sheet.Cells.Range("Q" & ix + 4).Value) & ","
                        strFIASevenYearAccumValueTemp = strFIASevenYearAccumValueTemp & FormatCurrency((Sheet.Cells.Range("R" & ix + 4).Value), 0, , , TriState.False) & ","
                    Else
                        strFIASevenYearIntRateTemp = strFIASevenYearIntRateTemp & "NA,"
                        strFIASevenYearAccumValueTemp = strFIASevenYearAccumValueTemp & "NA,"
                    End If
                    strFIATenYearIntRateTemp = strFIATenYearIntRateTemp & "NA,"
                    strFIATenYearAccumValueTemp = strFIATenYearAccumValueTemp & "NA,"
                ElseIf strFIAINIProdName = "FIA 10 Yr" Then
                    If strFIAINI10YearFxRate > 0 Then
                        strFIATenYearIntRateTemp = strFIATenYearIntRateTemp & FormatPercent(Sheet.Cells.Range("Q" & ix + 4).Value) & ","
                        strFIATenYearAccumValueTemp = strFIATenYearAccumValueTemp & FormatCurrency((Sheet.Cells.Range("R" & ix + 4).Value), 0, , , TriState.False) & ","
                    Else
                        strFIATenYearIntRateTemp = strFIATenYearIntRateTemp & "NA,"
                        strFIATenYearAccumValueTemp = strFIATenYearAccumValueTemp & "NA,"
                    End If
                    strFIASevenYearIntRateTemp = strFIASevenYearIntRateTemp & "NA,"
                    strFIASevenYearAccumValueTemp = strFIASevenYearAccumValueTemp & "NA,"
                End If

                If strFIAINIAnnualCap > 0 Then
                    strFIAAnnualCapIndexCreditTemp = strFIAAnnualCapIndexCreditTemp & FormatPercent(Sheet.Cells.Range("S" & ix + 4).Value) & ","
                    strFIAAnnualCapAccumValueTemp = strFIAAnnualCapAccumValueTemp & FormatCurrency((Sheet.Cells.Range("T" & ix + 4).Value), 0, , , TriState.False) & ","
                Else
                    strFIAAnnualCapIndexCreditTemp = strFIAAnnualCapIndexCreditTemp & "NA,"
                    strFIAAnnualCapAccumValueTemp = strFIAAnnualCapAccumValueTemp & "NA,"
                End If

                If strFIAINIMonthlyCap > 0 Then
                    strFIAMonthlyCapIndexCreditTemp = strFIAMonthlyCapIndexCreditTemp & FormatPercent(Sheet.Cells.Range("U" & ix + 4).Value) & ","
                    strFIAMonthlyCapAccumValueTemp = strFIAMonthlyCapAccumValueTemp & FormatCurrency((Sheet.Cells.Range("V" & ix + 4).Value), 0, , , TriState.False) & ","
                Else
                    strFIAMonthlyCapIndexCreditTemp = strFIAMonthlyCapIndexCreditTemp & "NA,"
                    strFIAMonthlyCapAccumValueTemp = strFIAMonthlyCapAccumValueTemp & "NA,"
                End If

                If strFIAINIPerformanceTrigger > 0 Then
                    strFIAPerfTriggerIndexCreditTemp = strFIAPerfTriggerIndexCreditTemp & FormatPercent(Sheet.Cells.Range("W" & ix + 4).Value) & ","
                    strFIAPerfTriggerAccumValueTemp = strFIAPerfTriggerAccumValueTemp & FormatCurrency((Sheet.Cells.Range("X" & ix + 4).Value), 0, , , TriState.False) & ","
                Else
                    strFIAPerfTriggerIndexCreditTemp = strFIAPerfTriggerIndexCreditTemp & "NA,"
                    strFIAPerfTriggerAccumValueTemp = strFIAPerfTriggerAccumValueTemp & "NA,"
                End If

            ElseIf ir = 2 Then 'Favorable 
                strFIAFavWDTemp = strFIAFavWDTemp & FormatCurrency((Sheet.Cells.Range("C" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAFavMGSVTemp = strFIAFavMGSVTemp & FormatCurrency((Sheet.Cells.Range("D" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAFavSPChangeTemp = strFIAFavSPChangeTemp & FormatPercent(Sheet.Cells.Range("E" & ix + 4).Value) & ","
                strFIAFavAnnCreditRateTemp = strFIAFavAnnCreditRateTemp & FormatPercent(Sheet.Cells.Range("F" & ix + 4).Value) & ","
                strFIAFavContractValueTemp = strFIAFavContractValueTemp & FormatCurrency((Sheet.Cells.Range("G" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAFavSurrenderValueTemp = strFIAFavSurrenderValueTemp & FormatCurrency((Sheet.Cells.Range("H" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAFavProjBeneBaseTemp = strFIAFavProjBeneBaseTemp & FormatCurrency((Sheet.Cells.Range("I" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAFavProjWDLimitTemp = strFIAFavProjWDLimitTemp & FormatCurrency((Sheet.Cells.Range("K" & ix + 4).Value), 0, , , TriState.False) & ","

            ElseIf ir = 3 Then 'Unfavorable
                strFIAUnfavWDTemp = strFIAUnfavWDTemp & FormatCurrency((Sheet.Cells.Range("C" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAUnfavMGSVTemp = strFIAUnfavMGSVTemp & FormatCurrency((Sheet.Cells.Range("D" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAUnfavSPChangeTemp = strFIAUnfavSPChangeTemp & FormatPercent(Sheet.Cells.Range("E" & ix + 4).Value) & ","
                strFIAUnfavAnnCreditRateTemp = strFIAUnfavAnnCreditRateTemp & FormatPercent(Sheet.Cells.Range("F" & ix + 4).Value) & ","
                strFIAUnfavContractValueTemp = strFIAUnfavContractValueTemp & FormatCurrency((Sheet.Cells.Range("G" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAUnfavSurrenderValueTemp = strFIAUnfavSurrenderValueTemp & FormatCurrency((Sheet.Cells.Range("H" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAUnfavProjBeneBaseTemp = strFIAUnfavProjBeneBaseTemp & FormatCurrency((Sheet.Cells.Range("I" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAUnfavProjWDLimitTemp = strFIAUnfavProjWDLimitTemp & FormatCurrency((Sheet.Cells.Range("K" & ix + 4).Value), 0, , , TriState.False) & ","

            ElseIf ir = 4 Then 'Income Rider
                strFIAContractValueNoWDTemp = strFIAContractValueNoWDTemp & FormatCurrency((Sheet.Cells.Range("AE" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAGuarBeneBaseNoWDTemp = strFIAGuarBeneBaseNoWDTemp & FormatCurrency((Sheet.Cells.Range("AG" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAProjBeneBaseNoWDTemp = strFIAProjBeneBaseNoWDTemp & FormatCurrency((Sheet.Cells.Range("AH" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAGuarWDFactorTemp = strFIAGuarWDFactorTemp & FormatPercent(Sheet.Cells.Range("AI" & ix + 4).Value) & ","
                strFIAGuarWDLimitNoWDTemp = strFIAGuarWDLimitNoWDTemp & FormatCurrency((Sheet.Cells.Range("AJ" & ix + 4).Value), 0, , , TriState.False) & ","
                strFIAProjWDLimitNoWDTemp = strFIAProjWDLimitNoWDTemp & FormatCurrency((Sheet.Cells.Range("AK" & ix + 4).Value), 0, , , TriState.False) & ","

            End If

        Next

        If ir = 1 Then
            strFIASpecSPChangeTool = strFIASpecSPChangeTemp
            strFIASpecAnnCreditRateTool = strFIASpecAnnCreditRateTemp
            strFIASpecContractValueTool = strFIASpecContractValueTemp
            strFIASpecSurrenderValueTool = strFIASpecSurrenderValueTemp
            strFIASpecMGSVTool = strFIASpecMGSVTemp
            strFIASpecProjBeneBaseTool = strFIASpecProjBeneBaseTemp
            strFIASpecProjWDLimitTool = strFIASpecProjWDLimitTemp
            strFIASevenYearIntRateTool = strFIASevenYearIntRateTemp
            strFIATenYearIntRateTool = strFIATenYearIntRateTemp
            strFIAMonthlyCapIndexCreditTool = strFIAMonthlyCapIndexCreditTemp
            strFIAAnnualCapIndexCreditTool = strFIAAnnualCapIndexCreditTemp
            strFIAPerfTriggerIndexCreditTool = strFIAPerfTriggerIndexCreditTemp
            strFIASevenYearAccumValueTool = strFIASevenYearAccumValueTemp
            strFIATenYearAccumValueTool = strFIATenYearAccumValueTemp
            strFIAMonthlyCapAccumValueTool = strFIAMonthlyCapAccumValueTemp
            strFIAAnnualCapAccumValueTool = strFIAAnnualCapAccumValueTemp
            strFIAPerfTriggerAccumValueTool = strFIAPerfTriggerAccumValueTemp
            strFIASpecWDTool = strFIASpecWDTemp
            strFIAGMCVTool = strFIAGMCVTemp

        ElseIf ir = 2 Then
            strFIAFavContractValueTool = strFIAFavContractValueTemp
            strFIAFavSurrenderValueTool = strFIAFavSurrenderValueTemp
            strFIAFavMGSVTool = strFIAFavMGSVTemp
            strFIAFavProjBeneBaseTool = strFIAFavProjBeneBaseTemp
            strFIAFavProjWDLimitTool = strFIAFavProjWDLimitTemp
            strFIAFavSPChangeTool = strFIAFavSPChangeTemp
            strFIAFavWDTool = strFIAFavWDTemp
            strFIAFavAnnCreditRateTool = strFIAFavAnnCreditRateTemp
        ElseIf ir = 3 Then
            strFIAUnfavAnnCreditRateTool = strFIAUnfavAnnCreditRateTemp
            strFIAUnfavContractValueTool = strFIAUnfavContractValueTemp
            strFIAUnfavSurrenderValueTool = strFIAUnfavSurrenderValueTemp
            strFIAUnfavProjBeneBaseTool = strFIAUnfavProjBeneBaseTemp
            strFIAUnfavProjWDLimitTool = strFIAUnfavProjWDLimitTemp
            strFIAUnfavMGSVTool = strFIAUnfavMGSVTemp
            strFIAUnfavSPChangeTool = strFIAUnfavSPChangeTemp
            strFIAUnfavWDTool = strFIAUnfavWDTemp
        ElseIf ir = 4 Then
            strFIAGuarBeneBaseNoWDTool = strFIAGuarBeneBaseNoWDTemp
            strFIAGuarWDLimitNoWDTool = strFIAGuarWDLimitNoWDTemp
            strFIAProjBeneBaseNoWDTool = strFIAProjBeneBaseNoWDTemp
            strFIAProjWDLimitNoWDTool = strFIAProjWDLimitNoWDTemp
            strFIAContractValueNoWDTool = strFIAContractValueNoWDTemp
            strFIAGuarWDFactorTool = strFIAGuarWDFactorTemp
        End If
    End Sub

End Class




