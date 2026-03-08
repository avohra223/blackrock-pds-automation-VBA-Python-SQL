Attribute VB_Name = "modValidationEngine"
'=============================================================================
' Module: modValidationEngine
' Purpose: Core eligibility validation engine using array-based batch processing
' Author: Akhil Vohra
' Notes: Reads loan data into arrays for performance, validates against
'        configurable criteria, outputs results with per-criterion pass/fail
'=============================================================================
Option Explicit

Public Type LoanRecord
    LoanID As String
    Borrower As String
    Country As String
    Sector As String
    NACECode As String
    SMEClass As String
    Revenue As Variant
    Employees As Variant
    LoanAmtLocal As Variant
    Currency As String
    LoanAmtEUR As Variant
    InterestRate As Variant
    MaturityDate As Variant
    Purpose As String
    ExistingGuarantee As Variant
    CollateralValue As Variant
    OriginationDate As Variant
    TaxID As String
    Status As String
    RowNum As Long
End Type

Public Type EligibilityCriteria
    MinLoanEUR As Double
    MaxLoanEUR As Double
    MaxRevenue As Double
    MaxEmployees As Long
    EligibleSMEClasses As String
    MinMaturityYears As Double
    MaxMaturityYears As Double
    MaxInterestRate As Double
    MaxGuarantee As Double
    MinOriginationDate As Date
    EligibleStatuses As String
End Type

Public Type ValidationResult
    LoanID As String
    Borrower As String
    Country As String
    LoanEUR As Double
    PassLoanSize As String
    PassRevenue As String
    PassEmployees As String
    PassMaturity As String
    PassInterestRate As String
    PassGuarantee As String
    PassOrigination As String
    OverallResult As String
    FailureReasons As String
End Type

'---------------------------------------------------------------------
' RunFullValidation: Main entry point for the eligibility check
'---------------------------------------------------------------------
Public Sub RunFullValidation()
    Dim loans() As LoanRecord
    Dim criteria As EligibilityCriteria
    Dim results() As ValidationResult
    Dim integrityIssues As Long
    Dim startTime As Double
    
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Step 1: Load criteria
    criteria = LoadCriteria()
    
    ' Step 2: Load loan data into arrays
    loans = LoadLoanData()
    If UBound(loans) < 1 Then
        MsgBox "No loan records found in the Loan Portfolio sheet.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Step 3: Run data integrity checks
    integrityIssues = RunIntegrityChecks(loans)
    
    ' Step 4: Run eligibility validation (array-based)
    results = ValidateLoans(loans, criteria)
    
    ' Step 5: Write results to Validation Results sheet
    WriteValidationResults results
    
    ' Step 6: Run concentration analysis
    RunConcentrationAnalysis loans, criteria
    
    ' Step 7: Update dashboard
    UpdateDashboard results, integrityIssues
    
    ' Step 8: Apply conditional formatting
    ApplyConditionalFormatting
    
    ' Step 9: Log to audit trail
    LogAuditTrail UBound(loans), results, integrityIssues, Timer - startTime
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Eligibility check complete." & vbCrLf & _
           "Records processed: " & UBound(loans) & vbCrLf & _
           "Time elapsed: " & Format(Timer - startTime, "0.00") & "s", _
           vbInformation, "Validation Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error in validation: " & Err.Description, vbCritical
    
CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------
' LoadLoanData: Reads Loan Portfolio sheet into typed array
'---------------------------------------------------------------------
Private Function LoadLoanData() As LoanRecord()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim loans() As LoanRecord
    Dim dataArr As Variant
    
    Set ws = ThisWorkbook.Sheets("Loan Portfolio")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        ReDim loans(0)
        LoadLoanData = loans
        Exit Function
    End If
    
    ' Read entire range into variant array (fast batch read)
    dataArr = ws.Range("A2:S" & lastRow).Value
    
    ReDim loans(1 To UBound(dataArr, 1))
    
    For i = 1 To UBound(dataArr, 1)
        With loans(i)
            .LoanID = CStr(dataArr(i, 1) & "")
            .Borrower = CStr(dataArr(i, 2) & "")
            .Country = CStr(dataArr(i, 3) & "")
            .Sector = CStr(dataArr(i, 4) & "")
            .NACECode = CStr(dataArr(i, 5) & "")
            .SMEClass = CStr(dataArr(i, 6) & "")
            .Revenue = dataArr(i, 7)
            .Employees = dataArr(i, 8)
            .LoanAmtLocal = dataArr(i, 9)
            .Currency = CStr(dataArr(i, 10) & "")
            .LoanAmtEUR = dataArr(i, 11)
            .InterestRate = dataArr(i, 12)
            .MaturityDate = dataArr(i, 13)
            .Purpose = CStr(dataArr(i, 14) & "")
            .ExistingGuarantee = dataArr(i, 15)
            .CollateralValue = dataArr(i, 16)
            .OriginationDate = dataArr(i, 17)
            .TaxID = CStr(dataArr(i, 18) & "")
            .Status = CStr(dataArr(i, 19) & "")
            .RowNum = i + 1
        End With
    Next i
    
    LoadLoanData = loans
End Function

'---------------------------------------------------------------------
' LoadCriteria: Reads Eligibility Criteria sheet into typed struct
'---------------------------------------------------------------------
Private Function LoadCriteria() As EligibilityCriteria
    Dim ws As Worksheet
    Dim c As EligibilityCriteria
    
    Set ws = ThisWorkbook.Sheets("Eligibility Criteria")
    
    ' Read from criteria table (rows 5-15, columns C-D for min/max)
    c.MinLoanEUR = SafeDbl(ws.Cells(5, 3).Value, 0)
    c.MaxLoanEUR = SafeDbl(ws.Cells(5, 4).Value, 999999999)
    c.MaxRevenue = SafeDbl(ws.Cells(6, 4).Value, 999999999)
    c.MaxEmployees = CLng(SafeDbl(ws.Cells(7, 4).Value, 99999))
    c.EligibleSMEClasses = "Micro,Small,Medium"
    c.MinMaturityYears = SafeDbl(ws.Cells(11, 3).Value, 0)
    c.MaxMaturityYears = SafeDbl(ws.Cells(11, 4).Value, 99)
    c.MaxInterestRate = SafeDbl(ws.Cells(12, 4).Value, 100)
    c.MaxGuarantee = SafeDbl(ws.Cells(13, 4).Value, 100)
    c.MinOriginationDate = DateSerial(2024, 1, 1)
    c.EligibleStatuses = "Active"
    
    LoadCriteria = c
End Function

'---------------------------------------------------------------------
' ValidateLoans: Array-based batch validation of all loans
'---------------------------------------------------------------------
Private Function ValidateLoans(ByRef loans() As LoanRecord, _
                                ByRef criteria As EligibilityCriteria) As ValidationResult()
    Dim results() As ValidationResult
    Dim i As Long
    Dim failCount As Long
    Dim reasons As String
    Dim matYears As Double
    
    ReDim results(1 To UBound(loans))
    
    For i = 1 To UBound(loans)
        failCount = 0
        reasons = ""
        
        With results(i)
            .LoanID = loans(i).LoanID
            .Borrower = loans(i).Borrower
            .Country = loans(i).Country
            
            ' Get EUR amount
            If IsNumeric(loans(i).LoanAmtEUR) Then
                .LoanEUR = CDbl(loans(i).LoanAmtEUR)
            Else
                .LoanEUR = 0
            End If
            
            ' CHECK 1: Loan Size
            If .LoanEUR >= criteria.MinLoanEUR And .LoanEUR <= criteria.MaxLoanEUR Then
                .PassLoanSize = "PASS"
            Else
                .PassLoanSize = "FAIL"
                failCount = failCount + 1
                If .LoanEUR < criteria.MinLoanEUR Then
                    reasons = reasons & "Loan below minimum EUR " & Format(criteria.MinLoanEUR, "#,##0") & "; "
                Else
                    reasons = reasons & "Loan exceeds maximum EUR " & Format(criteria.MaxLoanEUR, "#,##0") & "; "
                End If
            End If
            
            ' CHECK 2: Revenue Cap
            If IsNumeric(loans(i).Revenue) Then
                If CDbl(loans(i).Revenue) <= criteria.MaxRevenue Then
                    .PassRevenue = "PASS"
                Else
                    .PassRevenue = "FAIL"
                    failCount = failCount + 1
                    reasons = reasons & "Revenue exceeds SME cap; "
                End If
            Else
                .PassRevenue = "N/A"
                reasons = reasons & "Revenue data missing; "
            End If
            
            ' CHECK 3: Employee Limit
            If IsNumeric(loans(i).Employees) Then
                If CLng(loans(i).Employees) <= criteria.MaxEmployees Then
                    .PassEmployees = "PASS"
                Else
                    .PassEmployees = "FAIL"
                    failCount = failCount + 1
                    reasons = reasons & "Employees exceed " & criteria.MaxEmployees & "; "
                End If
            Else
                .PassEmployees = "N/A"
                reasons = reasons & "Employee data invalid; "
            End If
            
            ' CHECK 4: Maturity
            If IsDate(loans(i).MaturityDate) And IsDate(loans(i).OriginationDate) Then
                matYears = (CDate(loans(i).MaturityDate) - CDate(loans(i).OriginationDate)) / 365.25
                If matYears >= criteria.MinMaturityYears And matYears <= criteria.MaxMaturityYears Then
                    .PassMaturity = "PASS"
                Else
                    .PassMaturity = "FAIL"
                    failCount = failCount + 1
                    reasons = reasons & "Maturity " & Format(matYears, "0.0") & "y outside range; "
                End If
            Else
                .PassMaturity = "N/A"
                reasons = reasons & "Date data invalid; "
            End If
            
            ' CHECK 5: Interest Rate
            If IsNumeric(loans(i).InterestRate) Then
                Dim rateVal As Double
                rateVal = CDbl(loans(i).InterestRate)
                ' Handle both decimal (0.05) and percentage (5) formats
                If rateVal < 1 Then rateVal = rateVal * 100
                If rateVal <= criteria.MaxInterestRate Then
                    .PassInterestRate = "PASS"
                Else
                    .PassInterestRate = "FAIL"
                    failCount = failCount + 1
                    reasons = reasons & "Interest rate " & Format(rateVal, "0.00") & "% exceeds cap; "
                End If
            Else
                .PassInterestRate = "N/A"
            End If
            
            ' CHECK 6: Existing Guarantee
            If IsNumeric(loans(i).ExistingGuarantee) Then
                Dim guarVal As Double
                guarVal = CDbl(loans(i).ExistingGuarantee)
                If guarVal < 1 Then guarVal = guarVal * 100
                If guarVal <= criteria.MaxGuarantee Then
                    .PassGuarantee = "PASS"
                Else
                    .PassGuarantee = "FAIL"
                    failCount = failCount + 1
                    reasons = reasons & "Existing guarantee " & Format(guarVal, "0") & "% exceeds cap; "
                End If
            Else
                .PassGuarantee = "PASS"
            End If
            
            ' CHECK 7: Origination Date
            If IsDate(loans(i).OriginationDate) Then
                If CDate(loans(i).OriginationDate) >= criteria.MinOriginationDate Then
                    .PassOrigination = "PASS"
                Else
                    .PassOrigination = "FAIL"
                    failCount = failCount + 1
                    reasons = reasons & "Originated before " & Format(criteria.MinOriginationDate, "DD/MM/YYYY") & "; "
                End If
            Else
                .PassOrigination = "N/A"
                reasons = reasons & "Origination date invalid; "
            End If
            
            ' OVERALL RESULT
            If failCount = 0 And InStr(reasons, "missing") = 0 And InStr(reasons, "invalid") = 0 Then
                .OverallResult = "ELIGIBLE"
            Else
                .OverallResult = "INELIGIBLE"
            End If
            
            .FailureReasons = reasons
        End With
    Next i
    
    ValidateLoans = results
End Function

'---------------------------------------------------------------------
' WriteValidationResults: Outputs results array to sheet
'---------------------------------------------------------------------
Private Sub WriteValidationResults(ByRef results() As ValidationResult)
    Dim ws As Worksheet
    Dim outputArr() As Variant
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Validation Results")
    
    ' Clear previous results
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 3 Then
        ws.Range("A4:M" & lastRow).ClearContents
        ws.Range("A4:M" & lastRow).ClearFormats
    End If
    
    ' Build output array
    ReDim outputArr(1 To UBound(results), 1 To 13)
    
    For i = 1 To UBound(results)
        With results(i)
            outputArr(i, 1) = .LoanID
            outputArr(i, 2) = .Borrower
            outputArr(i, 3) = .Country
            outputArr(i, 4) = .LoanEUR
            outputArr(i, 5) = .PassLoanSize
            outputArr(i, 6) = .PassRevenue
            outputArr(i, 7) = .PassEmployees
            outputArr(i, 8) = .PassMaturity
            outputArr(i, 9) = .PassInterestRate
            outputArr(i, 10) = .PassGuarantee
            outputArr(i, 11) = .PassOrigination
            outputArr(i, 12) = .OverallResult
            outputArr(i, 13) = .FailureReasons
        End With
    Next i
    
    ' Batch write (fast)
    ws.Range("A4").Resize(UBound(results), 13).Value = outputArr
    
    ' Format EUR column
    ws.Range("D4:D" & 3 + UBound(results)).NumberFormat = "#,##0"
End Sub

'---------------------------------------------------------------------
' SafeDbl: Safe conversion to Double with default
'---------------------------------------------------------------------
Private Function SafeDbl(val As Variant, defaultVal As Double) As Double
    If IsNumeric(val) And Not IsEmpty(val) Then
        SafeDbl = CDbl(val)
    Else
        SafeDbl = defaultVal
    End If
End Function
