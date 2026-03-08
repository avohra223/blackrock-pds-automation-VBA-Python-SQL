Attribute VB_Name = "modIntegrityChecks"
'=============================================================================
' Module: modIntegrityChecks
' Purpose: Pre-validation data integrity checks (missing fields, duplicates,
'          format errors, negative values, text in numeric fields)
' Author: Akhil Vohra
'=============================================================================
Option Explicit

Public Type IntegrityIssue
    CheckNum As Long
    CheckType As String
    LoanID As String
    Field As String
    Issue As String
    Severity As String
    Action As String
End Type

'---------------------------------------------------------------------
' RunIntegrityChecks: Executes all data quality checks
'---------------------------------------------------------------------
Public Function RunIntegrityChecks(ByRef loans() As LoanRecord) As Long
    Dim issues() As IntegrityIssue
    Dim issueCount As Long
    Dim i As Long, j As Long
    Dim criticalCount As Long, warningCount As Long
    
    issueCount = 0
    criticalCount = 0
    warningCount = 0
    ReDim issues(1 To UBound(loans) * 5)  ' Pre-allocate generously
    
    For i = 1 To UBound(loans)
        ' CHECK: Missing Loan ID
        If Trim(loans(i).LoanID) = "" Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Missing Field", "Row " & loans(i).RowNum, _
                     "Loan ID", "Loan ID is blank", "Critical", "Add unique Loan ID"
            criticalCount = criticalCount + 1
        End If
        
        ' CHECK: Missing Borrower Name
        If Trim(loans(i).Borrower) = "" Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Missing Field", loans(i).LoanID, _
                     "Borrower Name", "Borrower name is blank", "Critical", "Add borrower name"
            criticalCount = criticalCount + 1
        End If
        
        ' CHECK: Missing Tax ID
        If Trim(loans(i).TaxID) = "" Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Missing Field", loans(i).LoanID, _
                     "Tax ID", "Borrower Tax ID is blank", "Warning", "Add tax identification"
            warningCount = warningCount + 1
        End If
        
        ' CHECK: Negative loan amount
        If IsNumeric(loans(i).LoanAmtLocal) Then
            If CDbl(loans(i).LoanAmtLocal) < 0 Then
                issueCount = issueCount + 1
                AddIssue issues, issueCount, "Invalid Value", loans(i).LoanID, _
                         "Loan Amount", "Negative amount: " & loans(i).LoanAmtLocal, _
                         "Critical", "Correct to positive value"
                criticalCount = criticalCount + 1
            End If
        ElseIf Not IsEmpty(loans(i).LoanAmtLocal) Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Format Error", loans(i).LoanID, _
                     "Loan Amount", "Non-numeric value: " & loans(i).LoanAmtLocal, _
                     "Critical", "Enter numeric amount"
            criticalCount = criticalCount + 1
        End If
        
        ' CHECK: Revenue - missing or non-numeric
        If IsEmpty(loans(i).Revenue) Or loans(i).Revenue = "" Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Missing Field", loans(i).LoanID, _
                     "Annual Revenue", "Revenue is blank", "Warning", "Add revenue figure"
            warningCount = warningCount + 1
        ElseIf Not IsNumeric(loans(i).Revenue) Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Format Error", loans(i).LoanID, _
                     "Annual Revenue", "Non-numeric value: " & loans(i).Revenue, _
                     "Critical", "Enter numeric revenue"
            criticalCount = criticalCount + 1
        End If
        
        ' CHECK: Employees - non-numeric
        If Not IsEmpty(loans(i).Employees) Then
            If Not IsNumeric(loans(i).Employees) Then
                issueCount = issueCount + 1
                AddIssue issues, issueCount, "Format Error", loans(i).LoanID, _
                         "Employees", "Non-numeric value: " & loans(i).Employees, _
                         "Critical", "Enter numeric headcount"
                criticalCount = criticalCount + 1
            End If
        End If
        
        ' CHECK: Maturity date - not a date
        If Not IsEmpty(loans(i).MaturityDate) Then
            If Not IsDate(loans(i).MaturityDate) Then
                issueCount = issueCount + 1
                AddIssue issues, issueCount, "Format Error", loans(i).LoanID, _
                         "Maturity Date", "Invalid date: " & loans(i).MaturityDate, _
                         "Critical", "Enter valid date"
                criticalCount = criticalCount + 1
            End If
        End If
        
        ' CHECK: Origination date - not a date
        If Not IsEmpty(loans(i).OriginationDate) Then
            If Not IsDate(loans(i).OriginationDate) Then
                issueCount = issueCount + 1
                AddIssue issues, issueCount, "Format Error", loans(i).LoanID, _
                         "Origination Date", "Invalid date: " & loans(i).OriginationDate, _
                         "Critical", "Enter valid date"
                criticalCount = criticalCount + 1
            End If
        End If
        
        ' CHECK: Missing currency
        If Trim(loans(i).Currency) = "" Then
            issueCount = issueCount + 1
            AddIssue issues, issueCount, "Missing Field", loans(i).LoanID, _
                     "Currency", "Currency code is blank", "Critical", "Add ISO currency code"
            criticalCount = criticalCount + 1
        End If
        
        ' CHECK: Duplicate Loan IDs
        If Trim(loans(i).LoanID) <> "" Then
            For j = 1 To i - 1
                If loans(j).LoanID = loans(i).LoanID Then
                    issueCount = issueCount + 1
                    AddIssue issues, issueCount, "Duplicate", loans(i).LoanID, _
                             "Loan ID", "Duplicate found at rows " & loans(j).RowNum & " and " & loans(i).RowNum, _
                             "Critical", "Assign unique Loan IDs"
                    criticalCount = criticalCount + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' Write issues to sheet
    WriteIntegrityResults issues, issueCount, UBound(loans), criticalCount, warningCount
    
    RunIntegrityChecks = issueCount
End Function

'---------------------------------------------------------------------
' AddIssue: Helper to populate issue array
'---------------------------------------------------------------------
Private Sub AddIssue(ByRef issues() As IntegrityIssue, ByVal idx As Long, _
                     ByVal checkType As String, ByVal loanID As String, _
                     ByVal field As String, ByVal issue As String, _
                     ByVal severity As String, ByVal action As String)
    If idx > UBound(issues) Then
        ReDim Preserve issues(1 To idx + 50)
    End If
    With issues(idx)
        .CheckNum = idx
        .CheckType = checkType
        .LoanID = loanID
        .Field = field
        .Issue = issue
        .Severity = severity
        .Action = action
    End With
End Sub

'---------------------------------------------------------------------
' WriteIntegrityResults: Outputs integrity check results to sheet
'---------------------------------------------------------------------
Private Sub WriteIntegrityResults(ByRef issues() As IntegrityIssue, _
                                   ByVal issueCount As Long, _
                                   ByVal totalRecords As Long, _
                                   ByVal criticalCount As Long, _
                                   ByVal warningCount As Long)
    Dim ws As Worksheet
    Dim outputArr() As Variant
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Data Integrity")
    
    ' Clear previous results
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 4 Then
        ws.Range("A5:G" & lastRow).ClearContents
        ws.Range("A5:G" & lastRow).ClearFormats
    End If
    
    ' Write issues
    If issueCount > 0 Then
        ReDim outputArr(1 To issueCount, 1 To 7)
        For i = 1 To issueCount
            outputArr(i, 1) = issues(i).CheckNum
            outputArr(i, 2) = issues(i).CheckType
            outputArr(i, 3) = issues(i).LoanID
            outputArr(i, 4) = issues(i).Field
            outputArr(i, 5) = issues(i).Issue
            outputArr(i, 6) = issues(i).Severity
            outputArr(i, 7) = issues(i).Action
        Next i
        ws.Range("A5").Resize(issueCount, 7).Value = outputArr
    Else
        ws.Range("A5").Value = "No integrity issues found"
        ws.Range("A5").Font.Italic = True
        ws.Range("A5").Font.Color = RGB(0, 128, 0)
    End If
    
    ' Update summary (row 8 onwards)
    Dim summaryRow As Long
    summaryRow = issueCount + 7
    ws.Cells(summaryRow, 1).Value = "Integrity Summary"
    ws.Cells(summaryRow, 1).Font.Bold = True
    ws.Cells(summaryRow, 1).Font.Size = 12
    
    ws.Cells(summaryRow + 1, 1).Value = "Total Records Checked"
    ws.Cells(summaryRow + 1, 2).Value = totalRecords
    ws.Cells(summaryRow + 2, 1).Value = "Records with Issues"
    ws.Cells(summaryRow + 2, 2).Value = issueCount
    ws.Cells(summaryRow + 3, 1).Value = "Critical Issues"
    ws.Cells(summaryRow + 3, 2).Value = criticalCount
    ws.Cells(summaryRow + 3, 2).Font.Color = RGB(231, 76, 60)
    ws.Cells(summaryRow + 4, 1).Value = "Warning Issues"
    ws.Cells(summaryRow + 4, 2).Value = warningCount
    ws.Cells(summaryRow + 4, 2).Font.Color = RGB(243, 156, 18)
    ws.Cells(summaryRow + 5, 1).Value = "Clean Records"
    ws.Cells(summaryRow + 5, 2).Value = totalRecords - issueCount
    ws.Cells(summaryRow + 5, 2).Font.Color = RGB(39, 174, 96)
End Sub
