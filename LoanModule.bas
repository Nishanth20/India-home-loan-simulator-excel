Attribute VB_Name = "LoanModule"
Option Explicit

Private Const INDIAN_FORMAT As Boolean = True

Private Const INPUT_SHEET As String = "Loan_Input"
Private Const SCHEDULE_SHEET As String = "Loan_Schedule"
Private Const RATE_SHEET As String = "Rate_Changes"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const COMPARE_SHEET As String = "Scenario_Compare"
Private Const ANNUAL_SUMMARY_SHEET As String = "Annual_Summary"

Private Const MODE_REDUCE_EMI As String = "REDUCE_EMI"
Private Const MODE_REDUCE_TENURE As String = "REDUCE_TENURE"

Private Const BODY_FONT As String = "Calibri"
Private Const BODY_FONT_SIZE As Long = 11
Private Const TITLE_FONT_SIZE As Long = 14
Private Const SHEET_HEADER_FONT_SIZE As Long = 18

Private Const COLOR_HEADER As Long = 2446373          ' RGB(37,68,65)
Private Const COLOR_EDITABLE As Long = 13498367       ' RGB(255,247,204)
Private Const COLOR_COMPUTED As Long = 15988979       ' RGB(243,248,252)
Private Const COLOR_SUMMARY As Long = 15135974        ' RGB(230,244,241)
Private Const COLOR_WHITE As Long = 16777215          ' RGB(255,255,255)
Private Const COLOR_GREEN As Long = 13434828          ' RGB(236,245,233)
Private Const COLOR_AMBER As Long = 13495551          ' RGB(255,241,204)
Private Const COLOR_RED As Long = 13421823            ' RGB(255,204,204)

Private Const INPUT_COL_LABEL As Long = 1
Private Const INPUT_COL_VALUE As Long = 2
Private Const INPUT_ROW_LOAN As Long = 2
Private Const INPUT_ROW_RATE As Long = 3
Private Const INPUT_ROW_TENURE As Long = 4
Private Const INPUT_ROW_MODE As Long = 5
Private Const INPUT_ROW_START_DATE As Long = 6

Private Const INPUT_LUMP_MONTH_COL As Long = 4
Private Const INPUT_LUMP_AMOUNT_COL As Long = 5
Private Const INPUT_EXTRA_MONTH_COL As Long = 7
Private Const INPUT_EXTRA_COUNT_COL As Long = 8
Private Const INPUT_TABLE_START_ROW As Long = 3

Private Const SUMMARY_TOP_ROW As Long = 3
Private Const SUMMARY_BOTTOM_ROW As Long = 11
Private Const HEADER_ROW As Long = 20
Private Const FIRST_DATA_ROW As Long = 21

Private Const COL_MONTH_NO As Long = 1
Private Const COL_MONTH_LABEL As Long = 2
Private Const COL_OPENING_BAL As Long = 3
Private Const COL_RATE As Long = 4
Private Const COL_EMI As Long = 5
Private Const COL_INTEREST As Long = 6
Private Const COL_PRINCIPAL As Long = 7
Private Const COL_LUMP_SUM As Long = 8
Private Const COL_EXTRA_EMI_COUNT As Long = 9
Private Const COL_EXTRA_EMI_AMOUNT As Long = 10
Private Const COL_PREPAY_APPLIED As Long = 11
Private Const COL_CLOSING_BAL As Long = 12
Private Const COL_REMAINING_MONTHS As Long = 13
Private Const COL_NEXT_EMI As Long = 14
Private Const COL_TOTAL_EMIS_YEAR As Long = 15
Private Const COL_SEC24B As Long = 16
Private Const COL_80C As Long = 17

Private Const TAX_SUMMARY_GAP As Long = 3
Private Const ANNUAL_HEADER_ROW As Long = 3
Private Const ANNUAL_DATA_ROW As Long = 4

Private Type LoanSummary
    TotalInterest As Double
    TotalPayment As Double
    MonthsUsed As Long
    TotalPrepayment As Double
    FinalEmi As Double
    Tax24bTotal As Double
    Tax80CTotal As Double
End Type

Public Sub SetupLoanCalculator()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    BuildInputSheet
    BuildRateSheet
    BuildScheduleShell
    BuildDashboardShell
    BuildCompareShell
    BuildAnnualSummaryShell
    Application.ScreenUpdating = True
    Application.StatusBar = "Loan calculator setup completed."
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "SetupLoanCalculator", Err.Description
End Sub

Public Sub GenerateLoanSchedule()
    On Error GoTo ErrHandler

    Dim principal As Double
    Dim annualRate As Double
    Dim tenureYears As Double
    Dim modeText As String
    Dim startDateValue As Date
    Dim lumpDict As Object
    Dim extraDict As Object
    Dim rateDict As Object
    Dim stressRateDict As Object
    Dim actual As LoanSummary
    Dim baseline As LoanSummary
    Dim planReduceEmi As LoanSummary
    Dim planReduceTenure As LoanSummary
    Dim rateStress As LoanSummary

    principal = GetPositiveDouble(ReadInputValue(INPUT_ROW_LOAN, INPUT_COL_VALUE), "Loan Amount")
    annualRate = GetNonNegativeDouble(ReadInputValue(INPUT_ROW_RATE, INPUT_COL_VALUE), "Annual Interest Rate")
    tenureYears = GetPositiveDouble(ReadInputValue(INPUT_ROW_TENURE, INPUT_COL_VALUE), "Tenure (Years)")
    modeText = NormalizeMode(CStr(ReadInputValue(INPUT_ROW_MODE, INPUT_COL_VALUE)))
    startDateValue = GetValidatedStartDate(ReadInputValue(INPUT_ROW_START_DATE, INPUT_COL_VALUE))

    WriteInputValue INPUT_ROW_MODE, INPUT_COL_VALUE, modeText
    WriteInputValue INPUT_ROW_START_DATE, INPUT_COL_VALUE, startDateValue

    Set lumpDict = CreateObject("Scripting.Dictionary")
    Set extraDict = CreateObject("Scripting.Dictionary")
    Set rateDict = CreateObject("Scripting.Dictionary")
    Set stressRateDict = CreateObject("Scripting.Dictionary")

    LoadAmountTable GetOrCreateSheet(INPUT_SHEET), lumpDict, INPUT_LUMP_MONTH_COL, INPUT_LUMP_AMOUNT_COL, INPUT_TABLE_START_ROW
    LoadCountTable GetOrCreateSheet(INPUT_SHEET), extraDict, INPUT_EXTRA_MONTH_COL, INPUT_EXTRA_COUNT_COL, INPUT_TABLE_START_ROW
    LoadRateChangeTable GetOrCreateSheet(RATE_SHEET), rateDict
    Set stressRateDict = BuildRateStressDictionary(rateDict)

    baseline = SimulateLoan(principal, annualRate, tenureYears, MODE_REDUCE_TENURE, startDateValue, _
                            CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"), rateDict, _
                            Nothing, False)

    actual = BuildLoanSchedule(principal, annualRate, tenureYears, modeText, startDateValue, lumpDict, extraDict, rateDict, baseline)

    planReduceEmi = SimulateLoan(principal, annualRate, tenureYears, MODE_REDUCE_EMI, startDateValue, _
                                 lumpDict, extraDict, rateDict, Nothing, False)
    planReduceTenure = SimulateLoan(principal, annualRate, tenureYears, MODE_REDUCE_TENURE, startDateValue, _
                                    lumpDict, extraDict, rateDict, Nothing, False)
    rateStress = SimulateLoan(principal, annualRate + 2#, tenureYears, MODE_REDUCE_TENURE, startDateValue, _
                              lumpDict, extraDict, stressRateDict, Nothing, False)

    BuildDashboard principal, annualRate, tenureYears, actual, baseline
    BuildScenarioCompare baseline, planReduceEmi, planReduceTenure, rateStress
    RefreshAnnualSummary

    Application.StatusBar = "Loan schedule generated successfully."
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "GenerateLoanSchedule", Err.Description
End Sub

Public Sub RecalculateSchedule()
    On Error GoTo ErrHandler

    Dim principal As Double
    Dim annualRate As Double
    Dim tenureYears As Double
    Dim modeText As String
    Dim startDateValue As Date
    Dim lumpDict As Object
    Dim extraDict As Object
    Dim rateDict As Object
    Dim stressRateDict As Object
    Dim actual As LoanSummary
    Dim baseline As LoanSummary
    Dim planReduceEmi As LoanSummary
    Dim planReduceTenure As LoanSummary
    Dim rateStress As LoanSummary
    Dim scheduleWs As Worksheet

    Set scheduleWs = GetOrCreateSheet(SCHEDULE_SHEET)

    principal = GetPositiveDouble(scheduleWs.Cells(SUMMARY_TOP_ROW, 2).Value, "Loan Amount")
    annualRate = GetNonNegativeDouble(scheduleWs.Cells(SUMMARY_TOP_ROW + 1, 2).Value, "Annual Interest Rate")
    tenureYears = GetPositiveDouble(scheduleWs.Cells(SUMMARY_TOP_ROW + 2, 2).Value, "Tenure (Years)")
    modeText = NormalizeMode(CStr(scheduleWs.Cells(SUMMARY_TOP_ROW + 1, 5).Value))
    startDateValue = GetValidatedStartDate(ReadInputValue(INPUT_ROW_START_DATE, INPUT_COL_VALUE))

    Set lumpDict = CreateObject("Scripting.Dictionary")
    Set extraDict = CreateObject("Scripting.Dictionary")
    Set rateDict = CreateObject("Scripting.Dictionary")
    Set stressRateDict = CreateObject("Scripting.Dictionary")

    LoadManualEntriesFromSchedule scheduleWs, lumpDict, extraDict
    LoadRateChangeTable GetOrCreateSheet(RATE_SHEET), rateDict
    Set stressRateDict = BuildRateStressDictionary(rateDict)

    baseline = SimulateLoan(principal, annualRate, tenureYears, MODE_REDUCE_TENURE, startDateValue, _
                            CreateObject("Scripting.Dictionary"), CreateObject("Scripting.Dictionary"), rateDict, _
                            Nothing, False)

    actual = BuildLoanSchedule(principal, annualRate, tenureYears, modeText, startDateValue, lumpDict, extraDict, rateDict, baseline)

    planReduceEmi = SimulateLoan(principal, annualRate, tenureYears, MODE_REDUCE_EMI, startDateValue, _
                                 lumpDict, extraDict, rateDict, Nothing, False)
    planReduceTenure = SimulateLoan(principal, annualRate, tenureYears, MODE_REDUCE_TENURE, startDateValue, _
                                    lumpDict, extraDict, rateDict, Nothing, False)
    rateStress = SimulateLoan(principal, annualRate + 2#, tenureYears, MODE_REDUCE_TENURE, startDateValue, _
                              lumpDict, extraDict, stressRateDict, Nothing, False)

    BuildDashboard principal, annualRate, tenureYears, actual, baseline
    BuildScenarioCompare baseline, planReduceEmi, planReduceTenure, rateStress
    RefreshAnnualSummary

    Application.StatusBar = "Loan schedule recalculated successfully."
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "RecalculateSchedule", Err.Description
End Sub

Public Sub RefreshAnnualSummary()
    On Error GoTo ErrHandler

    BuildAnnualSummaryFromSchedule
    Application.StatusBar = "Annual summary refreshed successfully."
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "RefreshAnnualSummary", Err.Description
End Sub

Private Function BuildLoanSchedule(ByVal principal As Double, ByVal annualRate As Double, ByVal tenureYears As Double, _
                                   ByVal modeText As String, ByVal startDateValue As Date, ByVal lumpDict As Object, _
                                   ByVal extraDict As Object, ByVal rateDict As Object, ByRef baseline As LoanSummary) As LoanSummary
    Dim scheduleWs As Worksheet
    Dim actual As LoanSummary
    Dim totalMonths As Long
    Dim originalEmi As Double
    Dim lastDataRow As Long
    Dim taxSummaryStartRow As Long

    Set scheduleWs = GetOrCreateSheet(SCHEDULE_SHEET)
    ClearWorksheet scheduleWs
    ApplyBaseSheetStyle scheduleWs

    totalMonths = CLng(Application.WorksheetFunction.RoundUp(tenureYears * 12, 0))
    originalEmi = Round(CalcEmi(principal, annualRate / 1200#, totalMonths), 2)

    WriteScheduleTitle scheduleWs
    WriteScheduleSummaryBlock scheduleWs, principal, annualRate, tenureYears, totalMonths, originalEmi, modeText, baseline
    WriteScheduleInstructions scheduleWs
    WriteScheduleHeaders scheduleWs

    actual = SimulateLoan(principal, annualRate, tenureYears, modeText, startDateValue, lumpDict, extraDict, rateDict, scheduleWs, True)
    lastDataRow = FIRST_DATA_ROW + actual.MonthsUsed - 1

    WriteScheduleActualSummary scheduleWs, principal, actual, baseline
    FormatScheduleDataArea scheduleWs, principal, lastDataRow
    CreateLoanChart scheduleWs, FIRST_DATA_ROW, lastDataRow
    AddOrReplaceButton scheduleWs, "RecalculateScheduleButton", "Recalculate Schedule", "LoanModule.RecalculateSchedule", _
                       scheduleWs.Cells(13, 9).Left, scheduleWs.Cells(13, 9).Top, 220, 34

    taxSummaryStartRow = lastDataRow + TAX_SUMMARY_GAP
    WriteScheduleTaxSummary scheduleWs, FIRST_DATA_ROW, lastDataRow, taxSummaryStartRow

    BuildLoanSchedule = actual
End Function

Private Function SimulateLoan(ByVal principal As Double, ByVal initialAnnualRate As Double, ByVal tenureYears As Double, _
                              ByVal modeText As String, ByVal startDateValue As Date, ByVal lumpDict As Object, _
                              ByVal extraDict As Object, ByVal rateDict As Object, ByVal targetWs As Worksheet, _
                              ByVal writeRows As Boolean) As LoanSummary
    Dim totalMonths As Long
    Dim monthNo As Long
    Dim loanYearNo As Long
    Dim rowOut As Long
    Dim openingBal As Double
    Dim closingBal As Double
    Dim balance As Double
    Dim annualRate As Double
    Dim monthlyRate As Double
    Dim originalEmi As Double
    Dim currentEmi As Double
    Dim emiPaid As Double
    Dim interestAmt As Double
    Dim principalAmt As Double
    Dim lumpAmt As Double
    Dim extraCount As Long
    Dim extraAmt As Double
    Dim prepaymentApplied As Double
    Dim allowedPrepayment As Double
    Dim remainingMonths As Long
    Dim nextEmi As Double
    Dim yearExtraDict As Object
    Dim totalEmisInLoanYear As Long
    Dim monthLabel As String
    Dim annual24bUsed As Double
    Dim annual80CUsed As Double
    Dim sec24bEligible As Double
    Dim sec80cEligible As Double
    Dim monthly24bCap As Double
    Dim monthly80cCap As Double

    totalMonths = CLng(Application.WorksheetFunction.RoundUp(tenureYears * 12, 0))
    annualRate = initialAnnualRate
    monthlyRate = annualRate / 1200#
    originalEmi = Round(CalcEmi(principal, monthlyRate, totalMonths), 2)
    currentEmi = originalEmi
    balance = principal
    rowOut = FIRST_DATA_ROW - 1
    monthly24bCap = 200000# / 12#
    monthly80cCap = 150000# / 12#

    Set yearExtraDict = BuildYearExtraDictionary(extraDict)

    Do While balance > 0.005 And monthNo < totalMonths + 600
        monthNo = monthNo + 1
        loanYearNo = ((monthNo - 1) \ 12) + 1

        If ((monthNo - 1) Mod 12) = 0 Then
            annual24bUsed = 0
            annual80CUsed = 0
        End If

        If rateDict.Exists(CStr(monthNo)) Then
            annualRate = CDbl(rateDict(CStr(monthNo)))
            monthlyRate = annualRate / 1200#
            If modeText = MODE_REDUCE_EMI Then
                remainingMonths = totalMonths - monthNo + 1
                If remainingMonths > 0 And balance > 0.005 Then
                    currentEmi = Round(CalcEmi(balance, monthlyRate, remainingMonths), 2)
                End If
            End If
        End If

        openingBal = balance
        interestAmt = Round(openingBal * monthlyRate, 2)

        If modeText = MODE_REDUCE_TENURE Then
            emiPaid = originalEmi
        Else
            emiPaid = currentEmi
        End If

        If emiPaid <= interestAmt And balance > 0.005 Then
            Err.Raise vbObjectError + 900, "SimulateLoan", "EMI is too low to cover interest in month " & CStr(monthNo) & "."
        End If

        principalAmt = Round(emiPaid - interestAmt, 2)
        If principalAmt > balance Then
            principalAmt = balance
            emiPaid = Round(interestAmt + principalAmt, 2)
        End If

        lumpAmt = GetDictionaryAmount(lumpDict, monthNo)
        extraCount = GetDictionaryCount(extraDict, monthNo)
        extraAmt = Round(extraCount * emiPaid, 2)
        allowedPrepayment = Round(balance - principalAmt, 2)
        If allowedPrepayment < 0 Then allowedPrepayment = 0

        prepaymentApplied = Round(Application.Min(allowedPrepayment, lumpAmt + extraAmt), 2)
        If prepaymentApplied < 0 Then prepaymentApplied = 0

        closingBal = Round(balance - principalAmt - prepaymentApplied, 10)
        If closingBal < 0.005 Then closingBal = 0

        remainingMonths = Application.Max(0, totalMonths - monthNo)
        If closingBal <= 0.005 Then
            nextEmi = 0
        ElseIf modeText = MODE_REDUCE_TENURE Then
            nextEmi = originalEmi
        ElseIf remainingMonths > 0 Then
            nextEmi = Round(CalcEmi(closingBal, monthlyRate, remainingMonths), 2)
        Else
            nextEmi = 0
        End If

        totalEmisInLoanYear = 12 + GetLoanYearExtraCount(yearExtraDict, loanYearNo)
        sec24bEligible = GetEligibleTaxValue(interestAmt, monthly24bCap, 200000#, annual24bUsed)
        sec80cEligible = GetEligibleTaxValue(principalAmt, monthly80cCap, 150000#, annual80CUsed)
        annual24bUsed = annual24bUsed + sec24bEligible
        annual80CUsed = annual80CUsed + sec80cEligible

        monthLabel = GetMonthLabel(startDateValue, monthNo)

        If writeRows Then
            rowOut = rowOut + 1
            WriteScheduleRow targetWs, rowOut, monthNo, monthLabel, openingBal, annualRate, emiPaid, interestAmt, principalAmt, _
                             lumpAmt, extraCount, extraAmt, prepaymentApplied, closingBal, remainingMonths, nextEmi, _
                             totalEmisInLoanYear, sec24bEligible, sec80cEligible
        End If

        SimulateLoan.TotalInterest = SimulateLoan.TotalInterest + interestAmt
        SimulateLoan.TotalPayment = SimulateLoan.TotalPayment + emiPaid + prepaymentApplied
        SimulateLoan.TotalPrepayment = SimulateLoan.TotalPrepayment + prepaymentApplied
        SimulateLoan.Tax24bTotal = SimulateLoan.Tax24bTotal + sec24bEligible
        SimulateLoan.Tax80CTotal = SimulateLoan.Tax80CTotal + sec80cEligible
        SimulateLoan.MonthsUsed = monthNo
        SimulateLoan.FinalEmi = nextEmi

        balance = closingBal
        currentEmi = nextEmi
    Loop
End Function

Private Sub BuildInputSheet()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(INPUT_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws

    WriteSheetTitle ws, 1, 1, "HOME LOAN CONTROL PANEL"

    WriteLabelValue ws, INPUT_ROW_LOAN, "Loan Amount", 4500000#
    WriteLabelValue ws, INPUT_ROW_RATE, "Annual Interest Rate (%)", 9#
    WriteLabelValue ws, INPUT_ROW_TENURE, "Tenure (Years)", 20#
    WriteLabelValue ws, INPUT_ROW_MODE, "Adjustment Mode", MODE_REDUCE_EMI
    WriteLabelValue ws, INPUT_ROW_START_DATE, "Loan Start Date", Date

    HighlightRange ws.Range(ws.Cells(INPUT_ROW_LOAN, INPUT_COL_VALUE), ws.Cells(INPUT_ROW_START_DATE, INPUT_COL_VALUE)), COLOR_EDITABLE
    ApplyModeValidation ws.Cells(INPUT_ROW_MODE, INPUT_COL_VALUE)
    ws.Cells(INPUT_ROW_START_DATE, INPUT_COL_VALUE).NumberFormat = "dd-mmm-yyyy"

    ws.Cells(1, 4).Value = "Lump Sum Prepayments"
    ws.Cells(2, 4).Value = "Month No"
    ws.Cells(2, 5).Value = "Amount"
    ws.Cells(1, 7).Value = "Extra EMI Plan"
    ws.Cells(2, 7).Value = "Month No"
    ws.Cells(2, 8).Value = "Extra EMI Count"
    ws.Cells(1, 10).Value = "Notes"
    ws.Cells(2, 10).Value = "B5 accepts REDUCE_EMI or REDUCE_TENURE"
    ws.Cells(3, 10).Value = "B6 drives MMM-YYYY labels in Loan_Schedule"
    ws.Cells(4, 10).Value = "Use D:E for month-wise lump sum entries"
    ws.Cells(5, 10).Value = "Use G:H for month-wise extra EMI count"
    ws.Cells(6, 10).Value = "Rate_Changes sheet supports floating rate overrides"

    HighlightRange ws.Range(ws.Cells(1, 4), ws.Cells(2, 5)), COLOR_SUMMARY
    HighlightRange ws.Range(ws.Cells(1, 7), ws.Cells(2, 8)), COLOR_SUMMARY
    HighlightRange ws.Range(ws.Cells(1, 10), ws.Cells(6, 10)), COLOR_COMPUTED

    ApplyThinBorders ws.Range(ws.Cells(1, 4), ws.Cells(2, 5))
    ApplyThinBorders ws.Range(ws.Cells(1, 7), ws.Cells(2, 8))
    ApplyThinBorders ws.Range(ws.Cells(1, 10), ws.Cells(6, 10))

    AddOrReplaceButton ws, "GenerateScheduleButton", "Generate Loan Schedule", "LoanModule.GenerateLoanSchedule", _
                       ws.Cells(9, 4).Left, ws.Cells(9, 4).Top, 220, 34

    ApplyCurrencyColumnsInput ws
    AutoFitStandardColumns ws
End Sub

Private Sub BuildRateSheet()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(RATE_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws

    WriteSheetTitle ws, 1, 1, "FLOATING RATE CHANGE TABLE"

    ws.Cells(3, 1).Value = "Month No"
    ws.Cells(3, 2).Value = "New Annual Rate (%)"
    ws.Cells(2, 4).Value = "Add overrides only when lender resets rate."
    ws.Cells(3, 4).Value = "Example: month 25 -> 9.25"
    ws.Cells(4, 4).Value = "Example: month 48 -> 8.75"
    ws.Cells(5, 4).Value = "Rate Stress Preview: Scenario_Compare auto-simulates +2% stress."

    HighlightRange ws.Range(ws.Cells(3, 1), ws.Cells(3, 2)), COLOR_SUMMARY
    HighlightRange ws.Range(ws.Cells(2, 4), ws.Cells(5, 4)), COLOR_COMPUTED
    ApplyThinBorders ws.Range(ws.Cells(3, 1), ws.Cells(3, 2))
    ApplyThinBorders ws.Range(ws.Cells(2, 4), ws.Cells(5, 4))

    AutoFitStandardColumns ws
End Sub

Private Sub BuildScheduleShell()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(SCHEDULE_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws
    WriteSheetTitle ws, 1, 1, "LOAN SCHEDULE"
    WriteScheduleHeaders ws
    AddOrReplaceButton ws, "RecalculateScheduleButton", "Recalculate Schedule", "LoanModule.RecalculateSchedule", _
                       ws.Cells(13, 9).Left, ws.Cells(13, 9).Top, 220, 34
End Sub

Private Sub BuildDashboardShell()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(DASHBOARD_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws
    WriteSheetTitle ws, 1, 1, "DASHBOARD"
End Sub

Private Sub BuildCompareShell()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(COMPARE_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws
    WriteSheetTitle ws, 1, 1, "SCENARIO COMPARE"
End Sub

Private Sub BuildAnnualSummaryShell()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(ANNUAL_SUMMARY_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws
    WriteSheetTitle ws, 1, 1, "ANNUAL SUMMARY"
    AddOrReplaceButton ws, "RefreshAnnualSummaryButton", "Refresh Annual Summary", "LoanModule.RefreshAnnualSummary", _
                       ws.Cells(1, 8).Left, ws.Cells(1, 8).Top, 220, 34
End Sub

Private Sub BuildDashboard(ByVal principal As Double, ByVal annualRate As Double, ByVal tenureYears As Double, _
                           ByRef actual As LoanSummary, ByRef baseline As LoanSummary)
    Dim ws As Worksheet
    Dim yearsSaved As Double
    Dim interestSaved As Double
    Dim totalTaxBenefit As Double
    Dim recommendation As String

    Set ws = GetOrCreateSheet(DASHBOARD_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws
    WriteSheetTitle ws, 1, 1, "DASHBOARD"

    yearsSaved = Application.Max(0, (baseline.MonthsUsed - actual.MonthsUsed) / 12#)
    interestSaved = baseline.TotalInterest - actual.TotalInterest
    totalTaxBenefit = actual.Tax24bTotal + actual.Tax80CTotal
    recommendation = BuildRecommendationText(actual, baseline)

    PaintKpiCard ws, 3, 1, "Loan Amount", principal, True
    PaintKpiCard ws, 3, 4, "Rate %", annualRate, False
    PaintKpiCard ws, 3, 7, "Tenure Years", tenureYears, False
    PaintKpiCard ws, 3, 10, "Payoff Months", actual.MonthsUsed, False

    PaintKpiCard ws, 7, 1, "Total Interest", actual.TotalInterest, True
    PaintKpiCard ws, 7, 4, "Total Prepayment", actual.TotalPrepayment, True
    PaintKpiCard ws, 7, 7, "Years Saved", yearsSaved, False
    PaintKpiCard ws, 7, 10, "Interest Saved", interestSaved, True

    PaintKpiCard ws, 11, 1, "Estimated 24b + 80C Total Benefit", totalTaxBenefit, True

    ws.Cells(11, 4).Value = "Recommendation"
    ws.Cells(12, 4).Value = recommendation
    ws.Cells(11, 4).Font.Bold = True
    ws.Range(ws.Cells(11, 4), ws.Cells(13, 11)).Interior.Color = COLOR_COMPUTED
    ApplyThinBorders ws.Range(ws.Cells(11, 4), ws.Cells(13, 11))
    ws.Range(ws.Cells(11, 4), ws.Cells(13, 11)).WrapText = True

    AutoFitStandardColumns ws
End Sub

Private Sub BuildScenarioCompare(ByRef baseLoan As LoanSummary, ByRef yourReduceEmi As LoanSummary, _
                                 ByRef yourReduceTenure As LoanSummary, ByRef rateStress As LoanSummary)
    Dim ws As Worksheet
    Dim minInterest As Double

    Set ws = GetOrCreateSheet(COMPARE_SHEET)
    ClearWorksheet ws
    ApplyBaseSheetStyle ws
    WriteSheetTitle ws, 1, 1, "SCENARIO COMPARE"

    ws.Cells(3, 1).Value = "Scenario"
    ws.Cells(3, 2).Value = "Payoff Months"
    ws.Cells(3, 3).Value = "Years"
    ws.Cells(3, 4).Value = "Total Interest"
    ws.Cells(3, 5).Value = "Total Prepayment"
    ws.Cells(3, 6).Value = "Interest Saved vs Base"
    ws.Cells(3, 7).Value = "Recommended?"
    ws.Cells(3, 8).Value = "Quant Advisory Notes"
    HighlightRange ws.Range(ws.Cells(3, 1), ws.Cells(3, 8)), COLOR_SUMMARY
    ApplyThinBorders ws.Range(ws.Cells(3, 1), ws.Cells(7, 8))

    WriteScenarioRow ws, 4, "Base Loan", baseLoan, baseLoan, "Benchmark path, no prepayments."
    WriteScenarioRow ws, 5, "Your Plan (REDUCE_EMI)", yourReduceEmi, baseLoan, "Cash-flow friendly, EMI relief focus."
    WriteScenarioRow ws, 6, "Your Plan (REDUCE_TENURE)", yourReduceTenure, baseLoan, "Interest-cutting mode, closes faster."
    WriteScenarioRow ws, 7, "Rate Stress +2%", rateStress, baseLoan, "Stress test for floating-rate pressure."

    minInterest = Application.Min(baseLoan.TotalInterest, yourReduceEmi.TotalInterest, yourReduceTenure.TotalInterest, rateStress.TotalInterest)
    HighlightBestScenario ws, 4, 7, minInterest
    ws.Range(ws.Cells(4, 2), ws.Cells(7, 2)).NumberFormat = "0"
    ws.Range(ws.Cells(4, 3), ws.Cells(7, 3)).NumberFormat = "0.00"
    ApplyCurrencyFormat ws.Range(ws.Cells(4, 4), ws.Cells(7, 6))

    AutoFitStandardColumns ws
End Sub

Private Sub BuildAnnualSummaryFromSchedule()
    Dim scheduleWs As Worksheet
    Dim annualWs As Worksheet
    Dim lastDataRow As Long
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim currentYear As Long
    Dim loanYear As Long
    Dim startLabel As String
    Dim endLabel As String
    Dim emiTotal As Double
    Dim interestTotal As Double
    Dim principalTotal As Double
    Dim prepayTotal As Double
    Dim closingBal As Double
    Dim tax24bTotal As Double
    Dim tax80cTotal As Double

    Set scheduleWs = GetOrCreateSheet(SCHEDULE_SHEET)
    Set annualWs = GetOrCreateSheet(ANNUAL_SUMMARY_SHEET)

    lastDataRow = GetLastScheduleDataRow(scheduleWs)
    If lastDataRow < FIRST_DATA_ROW Then
        Err.Raise vbObjectError + 910, "BuildAnnualSummaryFromSchedule", "Loan_Schedule does not contain amortization data."
    End If

    ClearWorksheet annualWs
    ApplyBaseSheetStyle annualWs
    WriteSheetTitle annualWs, 1, 1, "ANNUAL SUMMARY"
    AddOrReplaceButton annualWs, "RefreshAnnualSummaryButton", "Refresh Annual Summary", "LoanModule.RefreshAnnualSummary", _
                       annualWs.Cells(1, 8).Left, annualWs.Cells(1, 8).Top, 220, 34

    annualWs.Cells(ANNUAL_HEADER_ROW, 1).Value = "Year"
    annualWs.Cells(ANNUAL_HEADER_ROW, 2).Value = "Start Month"
    annualWs.Cells(ANNUAL_HEADER_ROW, 3).Value = "End Month"
    annualWs.Cells(ANNUAL_HEADER_ROW, 4).Value = "EMI Total"
    annualWs.Cells(ANNUAL_HEADER_ROW, 5).Value = "Interest Total"
    annualWs.Cells(ANNUAL_HEADER_ROW, 6).Value = "Principal Total"
    annualWs.Cells(ANNUAL_HEADER_ROW, 7).Value = "Prepayment Total"
    annualWs.Cells(ANNUAL_HEADER_ROW, 8).Value = "Closing Balance"
    annualWs.Cells(ANNUAL_HEADER_ROW, 9).Value = "24b Benefit"
    annualWs.Cells(ANNUAL_HEADER_ROW, 10).Value = "80C Benefit"
    HighlightRange annualWs.Range(annualWs.Cells(ANNUAL_HEADER_ROW, 1), annualWs.Cells(ANNUAL_HEADER_ROW, 10)), COLOR_SUMMARY

    outputRow = ANNUAL_DATA_ROW
    currentYear = 0

    For rowIndex = FIRST_DATA_ROW To lastDataRow
        loanYear = ((CLng(scheduleWs.Cells(rowIndex, COL_MONTH_NO).Value) - 1) \ 12) + 1

        If currentYear = 0 Then
            currentYear = loanYear
            startLabel = CStr(scheduleWs.Cells(rowIndex, COL_MONTH_LABEL).Value)
        End If

        If loanYear <> currentYear Then
            WriteAnnualSummaryRow annualWs, outputRow, currentYear, startLabel, endLabel, emiTotal, interestTotal, principalTotal, _
                                  prepayTotal, closingBal, tax24bTotal, tax80cTotal
            outputRow = outputRow + 1
            currentYear = loanYear
            startLabel = CStr(scheduleWs.Cells(rowIndex, COL_MONTH_LABEL).Value)
            emiTotal = 0
            interestTotal = 0
            principalTotal = 0
            prepayTotal = 0
            tax24bTotal = 0
            tax80cTotal = 0
        End If

        endLabel = CStr(scheduleWs.Cells(rowIndex, COL_MONTH_LABEL).Value)
        emiTotal = emiTotal + CDbl(scheduleWs.Cells(rowIndex, COL_EMI).Value)
        interestTotal = interestTotal + CDbl(scheduleWs.Cells(rowIndex, COL_INTEREST).Value)
        principalTotal = principalTotal + CDbl(scheduleWs.Cells(rowIndex, COL_PRINCIPAL).Value)
        prepayTotal = prepayTotal + CDbl(scheduleWs.Cells(rowIndex, COL_PREPAY_APPLIED).Value)
        closingBal = CDbl(scheduleWs.Cells(rowIndex, COL_CLOSING_BAL).Value)
        tax24bTotal = tax24bTotal + CDbl(scheduleWs.Cells(rowIndex, COL_SEC24B).Value)
        tax80cTotal = tax80cTotal + CDbl(scheduleWs.Cells(rowIndex, COL_80C).Value)
    Next rowIndex

    WriteAnnualSummaryRow annualWs, outputRow, currentYear, startLabel, endLabel, emiTotal, interestTotal, principalTotal, _
                          prepayTotal, closingBal, tax24bTotal, tax80cTotal

    ApplyThinBorders annualWs.Range(annualWs.Cells(ANNUAL_HEADER_ROW, 1), annualWs.Cells(outputRow, 10))
    ApplyCurrencyFormat annualWs.Range(annualWs.Cells(ANNUAL_DATA_ROW, 4), annualWs.Cells(outputRow, 10))
    AutoFitStandardColumns annualWs
End Sub

Private Sub WriteScheduleTitle(ByVal ws As Worksheet)
    WriteSheetTitle ws, 1, 1, "HOME LOAN SCHEDULE"
End Sub

Private Sub WriteScheduleSummaryBlock(ByVal ws As Worksheet, ByVal principal As Double, ByVal annualRate As Double, _
                                      ByVal tenureYears As Double, ByVal totalMonths As Long, ByVal originalEmi As Double, _
                                      ByVal modeText As String, ByRef baseline As LoanSummary)
    ws.Cells(3, 1).Value = "Loan Amount"
    ws.Cells(3, 2).Value = principal
    ws.Cells(4, 1).Value = "Interest Rate (%)"
    ws.Cells(4, 2).Value = annualRate
    ws.Cells(5, 1).Value = "Tenure (Years)"
    ws.Cells(5, 2).Value = tenureYears
    ws.Cells(6, 1).Value = "Tenure (Months)"
    ws.Cells(6, 2).Value = totalMonths

    ws.Cells(3, 4).Value = "Original EMI"
    ws.Cells(3, 5).Value = originalEmi
    ws.Cells(4, 4).Value = "Mode"
    ws.Cells(4, 5).Value = modeText
    ws.Cells(5, 4).Value = "Baseline Total Interest"
    ws.Cells(5, 5).Value = baseline.TotalInterest
    ws.Cells(6, 4).Value = "Baseline Total Payment"
    ws.Cells(6, 5).Value = baseline.TotalPayment

    ApplyModeValidation ws.Cells(4, 5)
    HighlightRange ws.Range(ws.Cells(3, 1), ws.Cells(6, 5)), COLOR_SUMMARY
    ApplyThinBorders ws.Range(ws.Cells(3, 1), ws.Cells(6, 5))
    ApplyCurrencyFormat ws.Range(ws.Cells(3, 2), ws.Cells(3, 2))
    ws.Cells(4, 2).NumberFormat = "0.00"
    ws.Cells(5, 2).NumberFormat = "0.00"
    ws.Cells(6, 2).NumberFormat = "0"
    ApplyCurrencyFormat ws.Range(ws.Cells(3, 5), ws.Cells(3, 5))
    ws.Cells(4, 5).NumberFormat = "@"
    ApplyCurrencyFormat ws.Range(ws.Cells(5, 5), ws.Cells(6, 5))
End Sub

Private Sub WriteScheduleActualSummary(ByVal ws As Worksheet, ByVal principal As Double, ByRef actual As LoanSummary, ByRef baseline As LoanSummary)
    Dim yearsSaved As Double
    Dim interestSaved As Double

    yearsSaved = Application.Max(0, (baseline.MonthsUsed - actual.MonthsUsed) / 12#)
    interestSaved = baseline.TotalInterest - actual.TotalInterest

    ws.Cells(8, 1).Value = "Actual Total Interest"
    ws.Cells(8, 2).Value = actual.TotalInterest
    ws.Cells(9, 1).Value = "Actual Total Payment"
    ws.Cells(9, 2).Value = actual.TotalPayment
    ws.Cells(10, 1).Value = "Principal Repaid"
    ws.Cells(10, 2).Value = principal
    ws.Cells(11, 1).Value = "Total Prepayment Applied"
    ws.Cells(11, 2).Value = actual.TotalPrepayment

    ws.Cells(8, 4).Value = "Payoff Months"
    ws.Cells(8, 5).Value = actual.MonthsUsed
    ws.Cells(9, 4).Value = "Actual Payoff (Years)"
    ws.Cells(9, 5).Value = Round(actual.MonthsUsed / 12#, 2)
    ws.Cells(10, 4).Value = "Years Saved"
    ws.Cells(10, 5).Value = Round(yearsSaved, 2)
    ws.Cells(11, 4).Value = "Interest Saved"
    ws.Cells(11, 5).Value = Round(interestSaved, 2)

    HighlightRange ws.Range(ws.Cells(8, 1), ws.Cells(11, 5)), COLOR_SUMMARY
    ApplyThinBorders ws.Range(ws.Cells(8, 1), ws.Cells(11, 5))
    ApplyCurrencyFormat ws.Range(ws.Cells(8, 2), ws.Cells(11, 2))
    ws.Cells(8, 5).NumberFormat = "0"
    ws.Cells(9, 5).NumberFormat = "0.00"
    ws.Cells(10, 5).NumberFormat = "0.00"
    ApplyCurrencyFormat ws.Range(ws.Cells(11, 5), ws.Cells(11, 5))
End Sub

Private Sub WriteScheduleInstructions(ByVal ws As Worksheet)
    ws.Cells(3, 7).Value = "Controls"
    ws.Cells(4, 7).Value = "Edit H for month-wise lump sum"
    ws.Cells(5, 7).Value = "Edit I for extra EMI count"
    ws.Cells(6, 7).Value = "Use Rate_Changes for floating-rate resets"
    ws.Cells(7, 7).Value = "Yellow cells are editable"

    HighlightRange ws.Range(ws.Cells(3, 7), ws.Cells(7, 10)), COLOR_COMPUTED
    ApplyThinBorders ws.Range(ws.Cells(3, 7), ws.Cells(7, 10))
End Sub

Private Sub WriteScheduleHeaders(ByVal ws As Worksheet)
    ws.Cells(HEADER_ROW, COL_MONTH_NO).Value = "Month No"
    ws.Cells(HEADER_ROW, COL_MONTH_LABEL).Value = "MMM-YYYY"
    ws.Cells(HEADER_ROW, COL_OPENING_BAL).Value = "Opening Balance"
    ws.Cells(HEADER_ROW, COL_RATE).Value = "Rate %"
    ws.Cells(HEADER_ROW, COL_EMI).Value = "EMI Paid"
    ws.Cells(HEADER_ROW, COL_INTEREST).Value = "Interest"
    ws.Cells(HEADER_ROW, COL_PRINCIPAL).Value = "Principal"
    ws.Cells(HEADER_ROW, COL_LUMP_SUM).Value = "Lump Sum"
    ws.Cells(HEADER_ROW, COL_EXTRA_EMI_COUNT).Value = "Extra EMI Count"
    ws.Cells(HEADER_ROW, COL_EXTRA_EMI_AMOUNT).Value = "Extra EMI Amount"
    ws.Cells(HEADER_ROW, COL_PREPAY_APPLIED).Value = "Prepayment Applied"
    ws.Cells(HEADER_ROW, COL_CLOSING_BAL).Value = "Closing Balance"
    ws.Cells(HEADER_ROW, COL_REMAINING_MONTHS).Value = "Remaining Months"
    ws.Cells(HEADER_ROW, COL_NEXT_EMI).Value = "Next EMI"
    ws.Cells(HEADER_ROW, COL_TOTAL_EMIS_YEAR).Value = "Total EMIs In Loan Year"
    ws.Cells(HEADER_ROW, COL_SEC24B).Value = "Sec 24b Eligible Interest"
    ws.Cells(HEADER_ROW, COL_80C).Value = "80C Principal Deduction"

    HighlightHeaderRow ws.Range(ws.Cells(HEADER_ROW, COL_MONTH_NO), ws.Cells(HEADER_ROW, COL_80C))
End Sub

Private Sub WriteScheduleRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal monthNo As Long, ByVal monthLabel As String, _
                             ByVal openingBal As Double, ByVal annualRate As Double, ByVal emiPaid As Double, _
                             ByVal interestAmt As Double, ByVal principalAmt As Double, ByVal lumpAmt As Double, _
                             ByVal extraCount As Long, ByVal extraAmt As Double, ByVal prepaymentApplied As Double, _
                             ByVal closingBal As Double, ByVal remainingMonths As Long, ByVal nextEmi As Double, _
                             ByVal totalEmisInLoanYear As Long, ByVal sec24bEligible As Double, ByVal sec80cEligible As Double)
    ws.Cells(rowNo, COL_MONTH_NO).Value = monthNo
    ws.Cells(rowNo, COL_MONTH_LABEL).NumberFormat = "@"
    ws.Cells(rowNo, COL_MONTH_LABEL).Value = monthLabel
    ws.Cells(rowNo, COL_OPENING_BAL).Value = openingBal
    ws.Cells(rowNo, COL_RATE).Value = annualRate
    ws.Cells(rowNo, COL_EMI).Value = emiPaid
    ws.Cells(rowNo, COL_INTEREST).Value = interestAmt
    ws.Cells(rowNo, COL_PRINCIPAL).Value = principalAmt
    ws.Cells(rowNo, COL_LUMP_SUM).Value = lumpAmt
    ws.Cells(rowNo, COL_EXTRA_EMI_COUNT).Value = extraCount
    ws.Cells(rowNo, COL_EXTRA_EMI_AMOUNT).Value = extraAmt
    ws.Cells(rowNo, COL_PREPAY_APPLIED).Value = prepaymentApplied
    ws.Cells(rowNo, COL_CLOSING_BAL).Value = closingBal
    ws.Cells(rowNo, COL_REMAINING_MONTHS).Value = remainingMonths
    ws.Cells(rowNo, COL_NEXT_EMI).Value = nextEmi
    ws.Cells(rowNo, COL_TOTAL_EMIS_YEAR).Value = totalEmisInLoanYear
    ws.Cells(rowNo, COL_SEC24B).Value = sec24bEligible
    ws.Cells(rowNo, COL_80C).Value = sec80cEligible
End Sub

Private Sub FormatScheduleDataArea(ByVal ws As Worksheet, ByVal originalPrincipal As Double, ByVal lastDataRow As Long)
    Dim fullDataRange As Range
    Dim editableRange As Range
    Dim computedRange As Range

    Set fullDataRange = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_MONTH_NO), ws.Cells(lastDataRow, COL_80C))
    Set editableRange = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_LUMP_SUM), ws.Cells(lastDataRow, COL_EXTRA_EMI_COUNT))
    Set computedRange = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_EXTRA_EMI_AMOUNT), ws.Cells(lastDataRow, COL_80C))

    ApplyThinBorders fullDataRange
    HighlightRange editableRange, COLOR_EDITABLE
    HighlightRange computedRange, COLOR_COMPUTED

    ws.Range(ws.Cells(HEADER_ROW, COL_MONTH_NO), ws.Cells(lastDataRow, COL_80C)).AutoFilter
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_OPENING_BAL), ws.Cells(lastDataRow, COL_OPENING_BAL))
    ws.Range(ws.Cells(FIRST_DATA_ROW, COL_RATE), ws.Cells(lastDataRow, COL_RATE)).NumberFormat = "0.00"
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_EMI), ws.Cells(lastDataRow, COL_EMI))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_INTEREST), ws.Cells(lastDataRow, COL_PRINCIPAL))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_LUMP_SUM), ws.Cells(lastDataRow, COL_LUMP_SUM))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_EXTRA_EMI_AMOUNT), ws.Cells(lastDataRow, COL_EXTRA_EMI_AMOUNT))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_PREPAY_APPLIED), ws.Cells(lastDataRow, COL_PREPAY_APPLIED))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_CLOSING_BAL), ws.Cells(lastDataRow, COL_CLOSING_BAL))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_NEXT_EMI), ws.Cells(lastDataRow, COL_NEXT_EMI))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_SEC24B), ws.Cells(lastDataRow, COL_SEC24B))
    ApplyCurrencyFormat ws.Range(ws.Cells(FIRST_DATA_ROW, COL_80C), ws.Cells(lastDataRow, COL_80C))
    ws.Range(ws.Cells(FIRST_DATA_ROW, COL_MONTH_NO), ws.Cells(lastDataRow, COL_MONTH_NO)).NumberFormat = "0"
    ws.Range(ws.Cells(FIRST_DATA_ROW, COL_MONTH_LABEL), ws.Cells(lastDataRow, COL_MONTH_LABEL)).NumberFormat = "@"
    ws.Range(ws.Cells(FIRST_DATA_ROW, COL_EXTRA_EMI_COUNT), ws.Cells(lastDataRow, COL_EXTRA_EMI_COUNT)).NumberFormat = "0"
    ws.Range(ws.Cells(FIRST_DATA_ROW, COL_REMAINING_MONTHS), ws.Cells(lastDataRow, COL_REMAINING_MONTHS)).NumberFormat = "0"
    ws.Range(ws.Cells(FIRST_DATA_ROW, COL_TOTAL_EMIS_YEAR), ws.Cells(lastDataRow, COL_TOTAL_EMIS_YEAR)).NumberFormat = "0"

    ApplyClosingBalanceConditionalFormatting ws, originalPrincipal, lastDataRow
    AutoFitStandardColumns ws
End Sub

Private Sub WriteScheduleTaxSummary(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal startRow As Long)
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim currentYear As Long
    Dim loanYear As Long
    Dim totalInterest As Double
    Dim totalSec24b As Double
    Dim totalPrincipal As Double
    Dim total80C As Double

    ws.Cells(startRow, 1).Value = "Yearly Tax Summary"
    ws.Cells(startRow + 1, 1).Value = "Year"
    ws.Cells(startRow + 1, 2).Value = "Total Interest Paid"
    ws.Cells(startRow + 1, 3).Value = "Sec 24b Benefit"
    ws.Cells(startRow + 1, 4).Value = "Principal Paid"
    ws.Cells(startRow + 1, 5).Value = "80C Benefit"

    HighlightRange ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 1, 5)), COLOR_SUMMARY

    outputRow = startRow + 2

    For rowIndex = firstRow To lastRow
        loanYear = ((CLng(ws.Cells(rowIndex, COL_MONTH_NO).Value) - 1) \ 12) + 1

        If currentYear = 0 Then currentYear = loanYear

        If loanYear <> currentYear Then
            WriteTaxSummaryRow ws, outputRow, currentYear, totalInterest, totalSec24b, totalPrincipal, total80C
            outputRow = outputRow + 1
            currentYear = loanYear
            totalInterest = 0
            totalSec24b = 0
            totalPrincipal = 0
            total80C = 0
        End If

        totalInterest = totalInterest + CDbl(ws.Cells(rowIndex, COL_INTEREST).Value)
        totalSec24b = totalSec24b + CDbl(ws.Cells(rowIndex, COL_SEC24B).Value)
        totalPrincipal = totalPrincipal + CDbl(ws.Cells(rowIndex, COL_PRINCIPAL).Value)
        total80C = total80C + CDbl(ws.Cells(rowIndex, COL_80C).Value)
    Next rowIndex

    WriteTaxSummaryRow ws, outputRow, currentYear, totalInterest, totalSec24b, totalPrincipal, total80C
    ApplyThinBorders ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(outputRow, 5))
    ApplyCurrencyFormat ws.Range(ws.Cells(startRow + 2, 2), ws.Cells(outputRow, 5))
End Sub

Private Sub WriteTaxSummaryRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal loanYear As Long, ByVal totalInterest As Double, _
                               ByVal totalSec24b As Double, ByVal totalPrincipal As Double, ByVal total80C As Double)
    ws.Cells(rowNo, 1).Value = loanYear
    ws.Cells(rowNo, 2).Value = totalInterest
    ws.Cells(rowNo, 3).Value = totalSec24b
    ws.Cells(rowNo, 4).Value = totalPrincipal
    ws.Cells(rowNo, 5).Value = total80C
End Sub

Private Sub WriteAnnualSummaryRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal loanYear As Long, ByVal startMonth As String, _
                                  ByVal endMonth As String, ByVal emiTotal As Double, ByVal interestTotal As Double, _
                                  ByVal principalTotal As Double, ByVal prepayTotal As Double, ByVal closingBal As Double, _
                                  ByVal tax24bTotal As Double, ByVal tax80cTotal As Double)
    ws.Cells(rowNo, 1).Value = loanYear
    ws.Cells(rowNo, 2).NumberFormat = "@"
    ws.Cells(rowNo, 2).Value = startMonth
    ws.Cells(rowNo, 3).NumberFormat = "@"
    ws.Cells(rowNo, 3).Value = endMonth
    ws.Cells(rowNo, 4).Value = emiTotal
    ws.Cells(rowNo, 5).Value = interestTotal
    ws.Cells(rowNo, 6).Value = principalTotal
    ws.Cells(rowNo, 7).Value = prepayTotal
    ws.Cells(rowNo, 8).Value = closingBal
    ws.Cells(rowNo, 9).Value = tax24bTotal
    ws.Cells(rowNo, 10).Value = tax80cTotal
End Sub

Private Sub WriteScenarioRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal scenarioName As String, _
                             ByRef scenarioResult As LoanSummary, ByRef baseLoan As LoanSummary, ByVal noteText As String)
    Dim interestSaved As Double
    Dim recommendedText As String

    interestSaved = baseLoan.TotalInterest - scenarioResult.TotalInterest
    recommendedText = GetScenarioRecommendation(scenarioName, scenarioResult, baseLoan)

    ws.Cells(rowNo, 1).Value = scenarioName
    ws.Cells(rowNo, 2).Value = scenarioResult.MonthsUsed
    ws.Cells(rowNo, 3).Value = Round(scenarioResult.MonthsUsed / 12#, 2)
    ws.Cells(rowNo, 4).Value = scenarioResult.TotalInterest
    ws.Cells(rowNo, 5).Value = scenarioResult.TotalPrepayment
    ws.Cells(rowNo, 6).Value = interestSaved
    ws.Cells(rowNo, 7).Value = recommendedText
    ws.Cells(rowNo, 8).Value = noteText
End Sub

Private Sub HighlightBestScenario(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal minInterest As Double)
    Dim rowNo As Long

    For rowNo = startRow To endRow
        If CDbl(ws.Cells(rowNo, 4).Value) = minInterest Then
            ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 8)).Font.Bold = True
            ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 8)).Interior.Color = COLOR_GREEN
            Exit For
        End If
    Next rowNo
End Sub

Private Function BuildRecommendationText(ByRef actual As LoanSummary, ByRef baseline As LoanSummary) As String
    Dim yearsSaved As Double

    yearsSaved = Application.Max(0, (baseline.MonthsUsed - actual.MonthsUsed) / 12#)

    If yearsSaved >= 3 Then
        BuildRecommendationText = "This strategy is strongly tenure-compressive. You are materially reducing interest drag and loan duration."
    ElseIf actual.TotalPrepayment > 0 And yearsSaved < 1 Then
        BuildRecommendationText = "Your current path mainly reduces EMI pressure. If faster closure matters more, switch to REDUCE_TENURE."
    ElseIf actual.TotalPrepayment = 0 Then
        BuildRecommendationText = "This is a pure amortization path. Even small early prepayments can improve lifetime interest outcomes."
    Else
        BuildRecommendationText = "Balanced path: some interest saving and some cash-flow relief. Fine-tune annual bonuses and rate stress assumptions."
    End If
End Function

Private Function GetScenarioRecommendation(ByVal scenarioName As String, ByRef scenarioResult As LoanSummary, ByRef baseLoan As LoanSummary) As String
    If scenarioResult.TotalInterest = baseLoan.TotalInterest Then
        GetScenarioRecommendation = "Base"
    ElseIf scenarioResult.TotalInterest < baseLoan.TotalInterest Then
        GetScenarioRecommendation = "Yes"
    Else
        GetScenarioRecommendation = "No"
    End If
End Function

Private Sub PaintKpiCard(ByVal ws As Worksheet, ByVal topRow As Long, ByVal leftCol As Long, ByVal labelText As String, _
                         ByVal metricValue As Double, ByVal isCurrency As Boolean)
    Dim target As Range

    Set target = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + 2, leftCol + 1))
    HighlightRange target, COLOR_SUMMARY
    ApplyThinBorders target
    target.VerticalAlignment = xlCenter
    target.HorizontalAlignment = xlCenter

    ws.Cells(topRow, leftCol).Value = labelText
    ws.Cells(topRow, leftCol).Font.Bold = True
    ws.Cells(topRow + 1, leftCol).Value = metricValue
    ws.Cells(topRow + 1, leftCol).Font.Size = TITLE_FONT_SIZE

    If isCurrency Then
        ApplyCurrencyFormat ws.Range(ws.Cells(topRow + 1, leftCol), ws.Cells(topRow + 1, leftCol))
    Else
        ws.Cells(topRow + 1, leftCol).NumberFormat = "#,##0.00"
    End If
End Sub

Private Sub CreateLoanChart(ByVal ws As Worksheet, ByVal firstDataRow As Long, ByVal lastDataRow As Long)
    Dim chartObject As ChartObject

    DeleteAllCharts ws

    Set chartObject = ws.ChartObjects.Add(Left:=ws.Cells(2, 18).Left, Top:=ws.Cells(2, 18).Top, Width:=540, Height:=300)
    chartObject.Chart.ChartType = xlLine
    chartObject.Chart.HasTitle = True
    chartObject.Chart.ChartTitle.Text = "Interest vs Principal — Month by Month"

    chartObject.Chart.SeriesCollection.NewSeries
    chartObject.Chart.SeriesCollection(1).Name = "Interest"
    chartObject.Chart.SeriesCollection(1).Values = ws.Range(ws.Cells(firstDataRow, COL_INTEREST), ws.Cells(lastDataRow, COL_INTEREST))
    chartObject.Chart.SeriesCollection(1).XValues = ws.Range(ws.Cells(firstDataRow, COL_MONTH_NO), ws.Cells(lastDataRow, COL_MONTH_NO))
    chartObject.Chart.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(203, 103, 30)

    chartObject.Chart.SeriesCollection.NewSeries
    chartObject.Chart.SeriesCollection(2).Name = "Principal"
    chartObject.Chart.SeriesCollection(2).Values = ws.Range(ws.Cells(firstDataRow, COL_PRINCIPAL), ws.Cells(lastDataRow, COL_PRINCIPAL))
    chartObject.Chart.SeriesCollection(2).XValues = ws.Range(ws.Cells(firstDataRow, COL_MONTH_NO), ws.Cells(lastDataRow, COL_MONTH_NO))
    chartObject.Chart.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(46, 117, 181)

    chartObject.Chart.Legend.Position = xlLegendPositionBottom
End Sub

Private Sub ApplyClosingBalanceConditionalFormatting(ByVal ws As Worksheet, ByVal originalPrincipal As Double, ByVal lastDataRow As Long)
    Dim target As Range

    Set target = ws.Range(ws.Cells(FIRST_DATA_ROW, COL_CLOSING_BAL), ws.Cells(lastDataRow, COL_CLOSING_BAL))
    target.FormatConditions.Delete

    target.FormatConditions.Add Type:=xlExpression, Formula1:="=" & ws.Cells(FIRST_DATA_ROW, COL_CLOSING_BAL).Address(False, False) & "<" & Replace(CStr(originalPrincipal * 0.25), ",", "")
    target.FormatConditions(target.FormatConditions.Count).Interior.Color = COLOR_GREEN

    target.FormatConditions.Add Type:=xlExpression, Formula1:="=AND(" & ws.Cells(FIRST_DATA_ROW, COL_CLOSING_BAL).Address(False, False) & ">=" & Replace(CStr(originalPrincipal * 0.25), ",", "") & "," & ws.Cells(FIRST_DATA_ROW, COL_CLOSING_BAL).Address(False, False) & "<" & Replace(CStr(originalPrincipal * 0.75), ",", "") & ")"
    target.FormatConditions(target.FormatConditions.Count).Interior.Color = COLOR_AMBER

    target.FormatConditions.Add Type:=xlExpression, Formula1:="=" & ws.Cells(FIRST_DATA_ROW, COL_CLOSING_BAL).Address(False, False) & ">=" & Replace(CStr(originalPrincipal * 0.75), ",", "")
    target.FormatConditions(target.FormatConditions.Count).Interior.Color = COLOR_RED
End Sub

Private Function ReadInputValue(ByVal rowNo As Long, ByVal colNo As Long) As Variant
    ReadInputValue = GetOrCreateSheet(INPUT_SHEET).Cells(rowNo, colNo).Value
End Function

Private Sub WriteInputValue(ByVal rowNo As Long, ByVal colNo As Long, ByVal valueToWrite As Variant)
    GetOrCreateSheet(INPUT_SHEET).Cells(rowNo, colNo).Value = valueToWrite
End Sub

Private Function GetValidatedStartDate(ByVal rawValue As Variant) As Date
    If IsDate(rawValue) Then
        GetValidatedStartDate = CDate(rawValue)
    Else
        GetValidatedStartDate = Date
    End If
End Function

Private Function GetMonthLabel(ByVal startDateValue As Date, ByVal monthNo As Long) As String
    If IsDate(startDateValue) Then
        GetMonthLabel = Format(DateAdd("m", monthNo - 1, startDateValue), "MMM-YYYY")
    Else
        GetMonthLabel = "Month " & CStr(monthNo)
    End If
End Function

Private Function GetPositiveDouble(ByVal rawValue As Variant, ByVal fieldName As String) As Double
    If Not IsNumeric(rawValue) Then
        Err.Raise vbObjectError + 1001, "GetPositiveDouble", fieldName & " must be numeric."
    End If
    If CDbl(rawValue) <= 0 Then
        Err.Raise vbObjectError + 1002, "GetPositiveDouble", fieldName & " must be greater than zero."
    End If
    GetPositiveDouble = CDbl(rawValue)
End Function

Private Function GetNonNegativeDouble(ByVal rawValue As Variant, ByVal fieldName As String) As Double
    If Not IsNumeric(rawValue) Then
        Err.Raise vbObjectError + 1003, "GetNonNegativeDouble", fieldName & " must be numeric."
    End If
    If CDbl(rawValue) < 0 Then
        Err.Raise vbObjectError + 1004, "GetNonNegativeDouble", fieldName & " cannot be negative."
    End If
    GetNonNegativeDouble = CDbl(rawValue)
End Function

Private Function NormalizeMode(ByVal rawMode As String) As String
    Dim modeText As String

    modeText = UCase$(Trim$(rawMode))
    If modeText <> MODE_REDUCE_EMI And modeText <> MODE_REDUCE_TENURE Then
        NormalizeMode = MODE_REDUCE_EMI
    Else
        NormalizeMode = modeText
    End If
End Function

Private Function CalcEmi(ByVal principal As Double, ByVal monthlyRate As Double, ByVal totalMonths As Long) As Double
    If totalMonths <= 0 Then
        CalcEmi = 0
    ElseIf monthlyRate = 0 Then
        CalcEmi = principal / totalMonths
    Else
        CalcEmi = principal * monthlyRate * ((1 + monthlyRate) ^ totalMonths) / (((1 + monthlyRate) ^ totalMonths) - 1)
    End If
End Function

Private Function GetEligibleTaxValue(ByVal rawAmount As Double, ByVal monthlyCap As Double, ByVal annualCap As Double, ByVal annualUsed As Double) As Double
    Dim remainingCap As Double
    Dim cappedValue As Double

    remainingCap = annualCap - annualUsed
    If remainingCap < 0 Then remainingCap = 0

    cappedValue = Application.Min(rawAmount, monthlyCap)
    GetEligibleTaxValue = Application.Min(cappedValue, remainingCap)
End Function

Private Sub LoadAmountTable(ByVal ws As Worksheet, ByVal dict As Object, ByVal keyCol As Long, ByVal valueCol As Long, ByVal startRow As Long)
    Dim lastRow As Long
    Dim rowNo As Long
    Dim monthNo As Long
    Dim amountValue As Double

    lastRow = GetLastUsedRow(ws, keyCol, valueCol)
    If lastRow < startRow Then Exit Sub

    For rowNo = startRow To lastRow
        If IsNumeric(ws.Cells(rowNo, keyCol).Value) And IsNumeric(ws.Cells(rowNo, valueCol).Value) Then
            monthNo = CLng(ws.Cells(rowNo, keyCol).Value)
            amountValue = CDbl(ws.Cells(rowNo, valueCol).Value)
            If monthNo > 0 And amountValue > 0 Then
                dict(CStr(monthNo)) = amountValue
            End If
        End If
    Next rowNo
End Sub

Private Sub LoadCountTable(ByVal ws As Worksheet, ByVal dict As Object, ByVal keyCol As Long, ByVal valueCol As Long, ByVal startRow As Long)
    Dim lastRow As Long
    Dim rowNo As Long
    Dim monthNo As Long
    Dim countValue As Long

    lastRow = GetLastUsedRow(ws, keyCol, valueCol)
    If lastRow < startRow Then Exit Sub

    For rowNo = startRow To lastRow
        If IsNumeric(ws.Cells(rowNo, keyCol).Value) And IsNumeric(ws.Cells(rowNo, valueCol).Value) Then
            monthNo = CLng(ws.Cells(rowNo, keyCol).Value)
            countValue = CLng(ws.Cells(rowNo, valueCol).Value)
            If monthNo > 0 And countValue > 0 Then
                dict(CStr(monthNo)) = countValue
            End If
        End If
    Next rowNo
End Sub

Private Sub LoadRateChangeTable(ByVal ws As Worksheet, ByVal dict As Object)
    Dim lastRow As Long
    Dim rowNo As Long
    Dim monthNo As Long
    Dim newRate As Double

    lastRow = GetLastUsedRow(ws, 1, 2)
    If lastRow < 4 Then Exit Sub

    For rowNo = 4 To lastRow
        If IsNumeric(ws.Cells(rowNo, 1).Value) And IsNumeric(ws.Cells(rowNo, 2).Value) Then
            monthNo = CLng(ws.Cells(rowNo, 1).Value)
            newRate = CDbl(ws.Cells(rowNo, 2).Value)
            If monthNo > 0 And newRate >= 0 Then
                dict(CStr(monthNo)) = newRate
            End If
        End If
    Next rowNo
End Sub

Private Sub LoadManualEntriesFromSchedule(ByVal ws As Worksheet, ByVal lumpDict As Object, ByVal extraDict As Object)
    Dim lastRow As Long
    Dim rowNo As Long
    Dim monthNo As Long

    lastRow = GetLastScheduleDataRow(ws)
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For rowNo = FIRST_DATA_ROW To lastRow
        If IsNumeric(ws.Cells(rowNo, COL_MONTH_NO).Value) Then
            monthNo = CLng(ws.Cells(rowNo, COL_MONTH_NO).Value)

            If IsNumeric(ws.Cells(rowNo, COL_LUMP_SUM).Value) Then
                If CDbl(ws.Cells(rowNo, COL_LUMP_SUM).Value) > 0 Then
                    lumpDict(CStr(monthNo)) = CDbl(ws.Cells(rowNo, COL_LUMP_SUM).Value)
                End If
            End If

            If IsNumeric(ws.Cells(rowNo, COL_EXTRA_EMI_COUNT).Value) Then
                If CLng(ws.Cells(rowNo, COL_EXTRA_EMI_COUNT).Value) > 0 Then
                    extraDict(CStr(monthNo)) = CLng(ws.Cells(rowNo, COL_EXTRA_EMI_COUNT).Value)
                End If
            End If
        End If
    Next rowNo
End Sub

Private Function BuildYearExtraDictionary(ByVal extraDict As Object) As Object
    Dim result As Object
    Dim dictKey As Variant
    Dim monthNo As Long
    Dim loanYearNo As Long

    Set result = CreateObject("Scripting.Dictionary")

    For Each dictKey In extraDict.Keys
        monthNo = CLng(dictKey)
        loanYearNo = ((monthNo - 1) \ 12) + 1
        If result.Exists(CStr(loanYearNo)) Then
            result(CStr(loanYearNo)) = CLng(result(CStr(loanYearNo))) + CLng(extraDict(dictKey))
        Else
            result(CStr(loanYearNo)) = CLng(extraDict(dictKey))
        End If
    Next dictKey

    Set BuildYearExtraDictionary = result
End Function

Private Function BuildRateStressDictionary(ByVal sourceDict As Object) As Object
    Dim result As Object
    Dim dictKey As Variant

    Set result = CreateObject("Scripting.Dictionary")

    For Each dictKey In sourceDict.Keys
        result(CStr(dictKey)) = CDbl(sourceDict(dictKey)) + 2#
    Next dictKey

    Set BuildRateStressDictionary = result
End Function

Private Function GetLoanYearExtraCount(ByVal yearExtraDict As Object, ByVal loanYearNo As Long) As Long
    If yearExtraDict.Exists(CStr(loanYearNo)) Then
        GetLoanYearExtraCount = CLng(yearExtraDict(CStr(loanYearNo)))
    Else
        GetLoanYearExtraCount = 0
    End If
End Function

Private Function GetDictionaryAmount(ByVal dict As Object, ByVal monthNo As Long) As Double
    If dict.Exists(CStr(monthNo)) Then
        GetDictionaryAmount = CDbl(dict(CStr(monthNo)))
    Else
        GetDictionaryAmount = 0
    End If
End Function

Private Function GetDictionaryCount(ByVal dict As Object, ByVal monthNo As Long) As Long
    If dict.Exists(CStr(monthNo)) Then
        GetDictionaryCount = CLng(dict(CStr(monthNo)))
    Else
        GetDictionaryCount = 0
    End If
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet, ByVal colA As Long, ByVal colB As Long) As Long
    Dim lastA As Long
    Dim lastB As Long

    lastA = ws.Cells(ws.Rows.Count, colA).End(xlUp).Row
    lastB = ws.Cells(ws.Rows.Count, colB).End(xlUp).Row
    If lastA > lastB Then
        GetLastUsedRow = lastA
    Else
        GetLastUsedRow = lastB
    End If
End Function

Private Function GetLastScheduleDataRow(ByVal ws As Worksheet) As Long
    Dim rowNo As Long

    rowNo = FIRST_DATA_ROW
    Do While rowNo <= ws.Rows.Count
        If Not IsNumeric(ws.Cells(rowNo, COL_MONTH_NO).Value) Then Exit Do
        If CLng(ws.Cells(rowNo, COL_MONTH_NO).Value) <= 0 Then Exit Do
        rowNo = rowNo + 1
    Loop

    GetLastScheduleDataRow = rowNo - 1
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set GetOrCreateSheet = ws
            Exit Function
        End If
    Next ws

    Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    GetOrCreateSheet.Name = sheetName
End Function

Private Sub ClearWorksheet(ByVal ws As Worksheet)
    DeleteAllCharts ws
    DeleteAllShapes ws
    ws.Cells.Clear
End Sub

Private Sub DeleteAllCharts(ByVal ws As Worksheet)
    Dim chartObject As ChartObject

    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject
End Sub

Private Sub DeleteAllShapes(ByVal ws As Worksheet)
    Dim shp As Shape

    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
End Sub

Private Sub AddOrReplaceButton(ByVal ws As Worksheet, ByVal buttonName As String, ByVal captionText As String, _
                               ByVal onActionText As String, ByVal leftPos As Double, ByVal topPos As Double, _
                               ByVal buttonWidth As Double, ByVal buttonHeight As Double)
    Dim shp As Shape
    Dim formButton As Button

    On Error Resume Next
    For Each shp In ws.Shapes
        If shp.Name = buttonName Then shp.Delete
    Next shp
    On Error GoTo 0

    Set formButton = ws.Buttons.Add(leftPos, topPos, buttonWidth, buttonHeight)
    formButton.Name = buttonName
    formButton.Characters.Text = captionText
    formButton.OnAction = onActionText
    formButton.Font.Bold = True
End Sub

Private Sub ApplyBaseSheetStyle(ByVal ws As Worksheet)
    ws.Cells.Font.Name = BODY_FONT
    ws.Cells.Font.Size = BODY_FONT_SIZE
    ws.Cells.Interior.Color = COLOR_WHITE
End Sub

Private Sub WriteSheetTitle(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal colNo As Long, ByVal titleText As String)
    ws.Cells(rowNo, colNo).Value = titleText
    ws.Cells(rowNo, colNo).Font.Name = BODY_FONT
    ws.Cells(rowNo, colNo).Font.Size = SHEET_HEADER_FONT_SIZE
    ws.Cells(rowNo, colNo).Font.Bold = True
    ws.Range(ws.Cells(rowNo, colNo), ws.Cells(rowNo, colNo + 25)).Interior.Color = COLOR_HEADER
    ws.Range(ws.Cells(rowNo, colNo), ws.Cells(rowNo, colNo + 25)).Font.Color = COLOR_WHITE
End Sub

Private Sub HighlightHeaderRow(ByVal target As Range)
    target.Interior.Color = COLOR_HEADER
    target.Font.Color = COLOR_WHITE
    target.Font.Bold = True
    ApplyThinBorders target
End Sub

Private Sub HighlightRange(ByVal target As Range, ByVal fillColor As Long)
    target.Interior.Color = fillColor
End Sub

Private Sub ApplyThinBorders(ByVal target As Range)
    target.Borders.LineStyle = xlContinuous
    target.Borders.Weight = xlThin
End Sub

Private Sub WriteLabelValue(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal labelText As String, ByVal valueText As Variant)
    ws.Cells(rowNo, INPUT_COL_LABEL).Value = labelText
    ws.Cells(rowNo, INPUT_COL_LABEL).Font.Bold = True
    ws.Cells(rowNo, INPUT_COL_VALUE).Value = valueText
End Sub

Private Sub ApplyModeValidation(ByVal targetCell As Range)
    targetCell.Validation.Delete
    targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                              Formula1:=MODE_REDUCE_EMI & "," & MODE_REDUCE_TENURE
End Sub

Private Sub ApplyCurrencyColumnsInput(ByVal ws As Worksheet)
    ApplyCurrencyFormat ws.Range(ws.Cells(INPUT_ROW_LOAN, INPUT_COL_VALUE), ws.Cells(INPUT_ROW_LOAN, INPUT_COL_VALUE))
End Sub

Private Sub ApplyCurrencyFormat(ByVal target As Range)
    If INDIAN_FORMAT Then
        target.NumberFormat = GetIndianCurrencyFormat()
    Else
        target.NumberFormat = "#,##0.00"
    End If
End Sub

Private Function GetIndianCurrencyFormat() As String
    GetIndianCurrencyFormat = ChrW$(8377) & " #,##,##,##0.00"
End Function

Private Sub AutoFitStandardColumns(ByVal ws As Worksheet)
    ws.Columns.AutoFit
End Sub
