Attribute VB_Name = "modAuditExport"
'=============================================================================
' Module: modAuditExport
' Purpose: Audit trail logging and report export functionality
' Author: Akhil Vohra
'=============================================================================
Option Explicit

'---------------------------------------------------------------------
' LogAuditTrail: Appends a new entry to the Audit Trail sheet
'---------------------------------------------------------------------
Public Sub LogAuditTrail(ByVal totalRecords As Long, _
                          ByRef results() As ValidationResult, _
                          ByVal integrityIssues As Long, _
                          ByVal elapsed As Double)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim eligible As Long, ineligible As Long
    Dim i As Long
    Dim criteriaSet As String
    
    Set ws = ThisWorkbook.Sheets("Audit Trail")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 4 Then nextRow = 4
    
    ' Count eligible/ineligible
    eligible = 0
    ineligible = 0
    For i = 1 To UBound(results)
        If results(i).OverallResult = "ELIGIBLE" Then
            eligible = eligible + 1
        Else
            ineligible = ineligible + 1
        End If
    Next i
    
    ' Get active criteria set name
    criteriaSet = CStr(ThisWorkbook.Sheets("Eligibility Criteria").Cells(2, 2).Value)
    If criteriaSet = "" Then criteriaSet = "Default"
    
    ' Determine run number
    Dim runNum As Long
    If nextRow = 4 Then
        runNum = 1
    Else
        runNum = ws.Cells(nextRow - 1, 1).Value + 1
    End If
    
    ' Write audit entry
    ws.Cells(nextRow, 1).Value = runNum
    ws.Cells(nextRow, 2).Value = Now()
    ws.Cells(nextRow, 2).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    ws.Cells(nextRow, 3).Value = Environ("USERNAME")
    ws.Cells(nextRow, 4).Value = criteriaSet
    ws.Cells(nextRow, 5).Value = totalRecords
    ws.Cells(nextRow, 6).Value = eligible
    ws.Cells(nextRow, 6).Font.Color = RGB(39, 174, 96)
    ws.Cells(nextRow, 7).Value = ineligible
    ws.Cells(nextRow, 7).Font.Color = RGB(231, 76, 60)
    ws.Cells(nextRow, 8).Value = integrityIssues
    ws.Cells(nextRow, 9).Value = "Completed in " & Format(elapsed, "0.00") & "s"
    
    ' Apply borders
    Dim c As Long
    For c = 1 To 9
        With ws.Cells(nextRow, c).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(217, 217, 217)
        End With
    Next c
End Sub

'---------------------------------------------------------------------
' ExportEligibilityReport: Exports results to a new workbook
'---------------------------------------------------------------------
Public Sub ExportEligibilityReport()
    Dim wbExport As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim fileName As String
    Dim savePath As Variant
    
    Set wsSource = ThisWorkbook.Sheets("Validation Results")
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 4 Then
        MsgBox "No validation results to export. Run the eligibility check first.", vbExclamation
        Exit Sub
    End If
    
    ' Create new workbook
    Set wbExport = Workbooks.Add(xlWBATWorksheet)
    Set wsTarget = wbExport.Sheets(1)
    wsTarget.Name = "Eligibility Report"
    
    ' Copy headers and data
    wsSource.Range("A3:M" & lastRow).Copy wsTarget.Range("A1")
    
    ' Add report header
    wsTarget.Rows(1).Insert
    wsTarget.Cells(1, 1).Value = "SME Loan Eligibility Report"
    wsTarget.Cells(1, 1).Font.Bold = True
    wsTarget.Cells(1, 1).Font.Size = 14
    
    wsTarget.Rows(2).Insert
    wsTarget.Cells(2, 1).Value = "Generated: " & Format(Now(), "DD/MM/YYYY HH:MM")
    wsTarget.Cells(2, 2).Value = "Criteria: " & _
        CStr(ThisWorkbook.Sheets("Eligibility Criteria").Cells(2, 2).Value)
    
    wsTarget.Rows(3).Insert  ' Blank row
    
    ' Add summary sheet
    Dim wsSummary As Worksheet
    Set wsSummary = wbExport.Sheets.Add(After:=wsTarget)
    wsSummary.Name = "Summary"
    
    ' Copy dashboard data
    ThisWorkbook.Sheets("Dashboard").Range("A1:H30").Copy wsSummary.Range("A1")
    
    ' Add concentration sheet
    Dim wsConc As Worksheet
    Set wsConc = wbExport.Sheets.Add(After:=wsSummary)
    wsConc.Name = "Concentration"
    
    ThisWorkbook.Sheets("Concentration Analysis").Range("A1:F45").Copy wsConc.Range("A1")
    
    ' Auto-fit columns
    Dim ws As Worksheet
    For Each ws In wbExport.Worksheets
        ws.Columns.AutoFit
    Next ws
    
    ' Prompt for save location
    fileName = "Eligibility_Report_" & Format(Now(), "YYYYMMDD_HHMM") & ".xlsx"
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=fileName, _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    If savePath <> False Then
        Application.DisplayAlerts = False
        wbExport.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        MsgBox "Report exported successfully to:" & vbCrLf & savePath, vbInformation
    End If
End Sub

'---------------------------------------------------------------------
' ExportToCSV: Quick CSV export of validation results
'---------------------------------------------------------------------
Public Sub ExportToCSV()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim csvContent As String
    Dim i As Long, j As Long
    Dim filePath As Variant
    
    Set ws = ThisWorkbook.Sheets("Validation Results")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 4 Then
        MsgBox "No results to export.", vbExclamation
        Exit Sub
    End If
    
    ' Build CSV string
    csvContent = ""
    For i = 3 To lastRow
        For j = 1 To lastCol
            If j > 1 Then csvContent = csvContent & ","
            Dim cellVal As String
            cellVal = CStr(ws.Cells(i, j).Value & "")
            ' Escape commas and quotes
            If InStr(cellVal, ",") > 0 Or InStr(cellVal, """") > 0 Then
                cellVal = """" & Replace(cellVal, """", """""") & """"
            End If
            csvContent = csvContent & cellVal
        Next j
        csvContent = csvContent & vbCrLf
    Next i
    
    ' Save
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Eligibility_Results_" & Format(Now(), "YYYYMMDD") & ".csv", _
        FileFilter:="CSV Files (*.csv), *.csv")
    
    If filePath <> False Then
        Dim ff As Integer
        ff = FreeFile
        Open filePath For Output As #ff
        Print #ff, csvContent
        Close #ff
        MsgBox "CSV exported to:" & vbCrLf & filePath, vbInformation
    End If
End Sub
