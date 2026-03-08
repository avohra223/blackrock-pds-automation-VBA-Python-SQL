Attribute VB_Name = "frmEligibilityChecker"
'=============================================================================
' UserForm: frmEligibilityChecker
' Purpose: Main user interface for running eligibility checks, selecting
'          criteria sets, adding loan records, and exporting results
' Author: Akhil Vohra
'
' NOTES FOR IMPLEMENTATION:
' This file contains the code-behind for a UserForm that should be created
' in the VBA Editor with the following controls:
'
' Form Properties:
'   Name: frmEligibilityChecker
'   Caption: "SME Loan Eligibility Checker"
'   Width: 420, Height: 380
'   BackColor: &H00FFFFFF (White)
'
' Controls Layout:
'   lblTitle        - Label (top banner, "SME Loan Eligibility Checker")
'   lblCriteriaSet  - Label ("Active Criteria Set:")
'   cmbCriteriaSet  - ComboBox (criteria set dropdown)
'   fraActions       - Frame ("Actions")
'     btnRunCheck    - CommandButton ("Run Eligibility Check")
'     btnAddLoan     - CommandButton ("Add Loan Record")
'     btnRefreshFX   - CommandButton ("Refresh FX Rates")
'   fraExport        - Frame ("Export")
'     btnExportXLSX  - CommandButton ("Export Report (.xlsx)")
'     btnExportCSV   - CommandButton ("Export Results (.csv)")
'   lblProgress      - Label ("Progress:")
'   barProgress      - Label (simulated progress bar, width changes dynamically)
'   barBackground    - Label (progress bar background track)
'   lblStatus        - Label (status text, e.g. "Processing loan 15 of 30...")
'   btnClose         - CommandButton ("Close")
'   lblLastRun       - Label ("Last run: <timestamp>")
'=============================================================================
Option Explicit

'---------------------------------------------------------------------
' Form Initialize: Set up defaults and populate dropdowns
'---------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Set title styling
    With lblTitle
        .Caption = "SME Loan Eligibility Checker"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(27, 42, 74)
    End With
    
    ' Populate criteria set dropdown
    With cmbCriteriaSet
        .AddItem "EU SME Guarantee Facility 2025"
        .AddItem "EIB InvestEU Programme"
        .AddItem "Custom Criteria Set"
        .ListIndex = 0
    End With
    
    ' Set progress bar initial state
    barProgress.Width = 0
    barProgress.BackColor = RGB(39, 174, 96)
    barBackground.BackColor = RGB(217, 217, 217)
    lblStatus.Caption = "Ready"
    lblProgress.Caption = ""
    
    ' Show last run time from audit trail
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("Audit Trail")
    Dim lastRow As Long
    lastRow = wsAudit.Cells(wsAudit.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 4 Then
        lblLastRun.Caption = "Last run: " & Format(wsAudit.Cells(lastRow, 2).Value, "DD/MM/YYYY HH:MM")
    Else
        lblLastRun.Caption = "No previous runs"
    End If
    
    ' Style buttons
    StyleButton btnRunCheck, RGB(39, 174, 96), True   ' Green, primary
    StyleButton btnAddLoan, RGB(46, 80, 144), False
    StyleButton btnRefreshFX, RGB(243, 156, 18), False
    StyleButton btnExportXLSX, RGB(46, 80, 144), False
    StyleButton btnExportCSV, RGB(46, 80, 144), False
    StyleButton btnClose, RGB(128, 128, 128), False
End Sub

'---------------------------------------------------------------------
' StyleButton: Helper to format command buttons
'---------------------------------------------------------------------
Private Sub StyleButton(ByRef btn As MSForms.CommandButton, _
                         ByVal clr As Long, ByVal isPrimary As Boolean)
    With btn
        .BackColor = clr
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = isPrimary
        .Font.Size = IIf(isPrimary, 11, 10)
    End With
End Sub

'---------------------------------------------------------------------
' Run Eligibility Check button
'---------------------------------------------------------------------
Private Sub btnRunCheck_Click()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Run eligibility check on all loan records?" & vbCrLf & _
                    "Criteria set: " & cmbCriteriaSet.Value, _
                    vbYesNo + vbQuestion, "Confirm")
    
    If answer = vbNo Then Exit Sub
    
    ' Update criteria set name in sheet
    ThisWorkbook.Sheets("Eligibility Criteria").Cells(2, 2).Value = cmbCriteriaSet.Value
    
    ' Validate FX rates first
    Dim fxStatus As String
    fxStatus = ValidateFXRates()
    If fxStatus <> "OK" Then
        answer = MsgBox(fxStatus & vbCrLf & vbCrLf & _
                        "Continue anyway? (affected loans will show EUR 0)", _
                        vbYesNo + vbExclamation, "FX Rate Warning")
        If answer = vbNo Then Exit Sub
    End If
    
    ' Update progress bar
    UpdateProgress 0, "Initializing..."
    DoEvents
    
    ' Run the full validation
    UpdateProgress 10, "Loading loan data..."
    DoEvents
    
    ' Call the main validation engine
    RunFullValidation
    
    UpdateProgress 100, "Complete!"
    lblLastRun.Caption = "Last run: " & Format(Now(), "DD/MM/YYYY HH:MM")
End Sub

'---------------------------------------------------------------------
' Add Loan Record button - opens input form
'---------------------------------------------------------------------
Private Sub btnAddLoan_Click()
    ' Open a simple input sequence for adding a single loan
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Loan Portfolio")
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Generate next Loan ID
    Dim newID As String
    newID = "LN-" & Format(nextRow - 1, "0000")
    
    ' Simple input boxes for key fields
    Dim borrower As String
    borrower = InputBox("Borrower Name:", "Add Loan Record")
    If borrower = "" Then Exit Sub
    
    Dim country As String
    country = InputBox("Country:", "Add Loan Record")
    If country = "" Then Exit Sub
    
    Dim sector As String
    sector = InputBox("Sector:", "Add Loan Record")
    
    Dim amount As String
    amount = InputBox("Loan Amount (local currency):", "Add Loan Record")
    If Not IsNumeric(amount) Then
        MsgBox "Invalid amount.", vbExclamation
        Exit Sub
    End If
    
    Dim ccy As String
    ccy = InputBox("Currency (e.g., EUR, PLN, HUF):", "Add Loan Record", "EUR")
    
    ' Write to sheet
    ws.Cells(nextRow, 1).Value = newID
    ws.Cells(nextRow, 2).Value = borrower
    ws.Cells(nextRow, 3).Value = country
    ws.Cells(nextRow, 4).Value = sector
    ws.Cells(nextRow, 9).Value = CDbl(amount)
    ws.Cells(nextRow, 10).Value = UCase(ccy)
    ws.Cells(nextRow, 19).Value = "Active"
    
    ' Set EUR formula
    ws.Cells(nextRow, 11).Formula = _
        "=IF(J" & nextRow & "=""EUR"",I" & nextRow & ",I" & nextRow & "/VLOOKUP(J" & nextRow & ",'FX Rates'!A:B,2,FALSE))"
    
    MsgBox "Loan " & newID & " added for " & borrower & ".", vbInformation
End Sub

'---------------------------------------------------------------------
' Refresh FX Rates button
'---------------------------------------------------------------------
Private Sub btnRefreshFX_Click()
    RefreshEURAmounts
End Sub

'---------------------------------------------------------------------
' Export buttons
'---------------------------------------------------------------------
Private Sub btnExportXLSX_Click()
    ExportEligibilityReport
End Sub

Private Sub btnExportCSV_Click()
    ExportToCSV
End Sub

'---------------------------------------------------------------------
' Close button
'---------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

'---------------------------------------------------------------------
' UpdateProgress: Animates the progress bar and status text
'---------------------------------------------------------------------
Public Sub UpdateProgress(ByVal pct As Long, ByVal statusText As String)
    ' barProgress width proportional to percentage
    ' Assuming barBackground.Width = 340 (the full track width)
    Dim maxWidth As Single
    maxWidth = barBackground.Width
    
    barProgress.Width = (pct / 100) * maxWidth
    lblProgress.Caption = pct & "%"
    lblStatus.Caption = statusText
    
    ' Color changes as progress advances
    If pct < 30 Then
        barProgress.BackColor = RGB(243, 156, 18)   ' Amber
    ElseIf pct < 70 Then
        barProgress.BackColor = RGB(46, 80, 144)     ' Blue
    Else
        barProgress.BackColor = RGB(39, 174, 96)     ' Green
    End If
    
    DoEvents
End Sub

'---------------------------------------------------------------------
' ShowForm: Public entry point to display the UserForm
'---------------------------------------------------------------------
Public Sub ShowEligibilityForm()
    frmEligibilityChecker.Show vbModeless
End Sub
