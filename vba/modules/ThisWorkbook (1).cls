Attribute VB_Name = "ThisWorkbook"
'=============================================================================
' ThisWorkbook Events
' Purpose: Workbook-level event handlers for auto-setup and cleanup
' Author: Akhil Vohra
'=============================================================================
Option Explicit

Private Sub Workbook_Open()
    ' Create the custom toolbar when workbook opens
    CreateToolbar
    
    ' Navigate to Instructions sheet
    ThisWorkbook.Sheets("Instructions").Activate
    
    ' Show welcome message
    If MsgBox("Welcome to the SME Loan Eligibility Checker." & vbCrLf & vbCrLf & _
              "Would you like to open the Control Panel?", _
              vbYesNo + vbInformation, "Eligibility Checker") = vbYes Then
        ShowEligibilityForm
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clean up toolbar
    RemoveToolbar
End Sub
