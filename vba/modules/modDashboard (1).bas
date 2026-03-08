Attribute VB_Name = "modDashboard"
'=============================================================================
' Module: modDashboard
' Purpose: Updates Dashboard sheet with KPIs and breakdown tables;
'          Applies dynamic conditional formatting across result sheets
' Author: Akhil Vohra
'=============================================================================
Option Explicit

'---------------------------------------------------------------------
' UpdateDashboard: Populates dashboard KPIs and breakdown tables
'---------------------------------------------------------------------
Public Sub UpdateDashboard(ByRef results() As ValidationResult, _
                            ByVal integrityIssues As Long)
    Dim ws As Worksheet
    Dim i As Long
    Dim eligible As Long, ineligible As Long
    Dim totalEUR As Double, eligibleEUR As Double
    Dim failureDict As Object
    Dim countryElig As Object, countryInelig As Object
    Dim concBreaches As Long
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Set failureDict = CreateObject("Scripting.Dictionary")
    Set countryElig = CreateObject("Scripting.Dictionary")
    Set countryInelig = CreateObject("Scripting.Dictionary")
    
    eligible = 0
    ineligible = 0
    totalEUR = 0
    eligibleEUR = 0
    
    For i = 1 To UBound(results)
        totalEUR = totalEUR + results(i).LoanEUR
        
        If results(i).OverallResult = "ELIGIBLE" Then
            eligible = eligible + 1
            eligibleEUR = eligibleEUR + results(i).LoanEUR
        Else
            ineligible = ineligible + 1
        End If
        
        ' Track failure reasons
        If results(i).FailureReasons <> "" Then
            Dim reasons() As String
            Dim r As Long
            reasons = Split(results(i).FailureReasons, ";")
            For r = 0 To UBound(reasons)
                Dim reason As String
                reason = Trim(reasons(r))
                If reason <> "" Then
                    If failureDict.Exists(reason) Then
                        failureDict(reason) = failureDict(reason) + 1
                    Else
                        failureDict.Add reason, 1
                    End If
                End If
            Next r
        End If
        
        ' Track country breakdown
        Dim ctry As String
        ctry = results(i).Country
        If results(i).OverallResult = "ELIGIBLE" Then
            If countryElig.Exists(ctry) Then
                countryElig(ctry) = countryElig(ctry) + 1
            Else
                countryElig.Add ctry, 1
                If Not countryInelig.Exists(ctry) Then countryInelig.Add ctry, 0
            End If
        Else
            If countryInelig.Exists(ctry) Then
                countryInelig(ctry) = countryInelig(ctry) + 1
            Else
                countryInelig.Add ctry, 1
                If Not countryElig.Exists(ctry) Then countryElig.Add ctry, 0
            End If
        End If
    Next i
    
    ' Count concentration breaches
    concBreaches = CountConcentrationBreaches()
    
    ' Clear old data
    ws.Range("A3:H30").ClearContents
    
    ' Write KPIs
    ' Row 3: Total | Eligible | Ineligible
    ws.Cells(3, 1).Value = "Total Loans Submitted"
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 2).Value = UBound(results)
    ws.Cells(3, 2).Font.Size = 14
    ws.Cells(3, 2).Font.Bold = True
    
    ws.Cells(3, 3).Value = "Eligible Loans"
    ws.Cells(3, 3).Font.Bold = True
    ws.Cells(3, 4).Value = eligible
    ws.Cells(3, 4).Font.Size = 14
    ws.Cells(3, 4).Font.Color = RGB(39, 174, 96)
    ws.Cells(3, 4).Font.Bold = True
    
    ws.Cells(3, 5).Value = "Ineligible Loans"
    ws.Cells(3, 5).Font.Bold = True
    ws.Cells(3, 6).Value = ineligible
    ws.Cells(3, 6).Font.Size = 14
    ws.Cells(3, 6).Font.Color = RGB(231, 76, 60)
    ws.Cells(3, 6).Font.Bold = True
    
    ' Row 5: Rate | Total EUR | Eligible EUR
    ws.Cells(5, 1).Value = "Eligibility Rate"
    ws.Cells(5, 1).Font.Bold = True
    If UBound(results) > 0 Then
        ws.Cells(5, 2).Value = eligible / UBound(results)
    Else
        ws.Cells(5, 2).Value = 0
    End If
    ws.Cells(5, 2).NumberFormat = "0.0%"
    ws.Cells(5, 2).Font.Size = 14
    ws.Cells(5, 2).Font.Bold = True
    
    ws.Cells(5, 3).Value = "Total EUR Exposure"
    ws.Cells(5, 3).Font.Bold = True
    ws.Cells(5, 4).Value = totalEUR
    ws.Cells(5, 4).NumberFormat = "#,##0"
    ws.Cells(5, 4).Font.Bold = True
    
    ws.Cells(5, 5).Value = "Eligible EUR Exposure"
    ws.Cells(5, 5).Font.Bold = True
    ws.Cells(5, 6).Value = eligibleEUR
    ws.Cells(5, 6).NumberFormat = "#,##0"
    ws.Cells(5, 6).Font.Bold = True
    
    ' Row 7: Integrity | Concentration | Last Run
    ws.Cells(7, 1).Value = "Data Integrity Issues"
    ws.Cells(7, 1).Font.Bold = True
    ws.Cells(7, 2).Value = integrityIssues
    If integrityIssues > 0 Then
        ws.Cells(7, 2).Font.Color = RGB(243, 156, 18)
    Else
        ws.Cells(7, 2).Font.Color = RGB(39, 174, 96)
    End If
    
    ws.Cells(7, 3).Value = "Concentration Breaches"
    ws.Cells(7, 3).Font.Bold = True
    ws.Cells(7, 4).Value = concBreaches
    If concBreaches > 0 Then
        ws.Cells(7, 4).Font.Color = RGB(231, 76, 60)
    End If
    
    ws.Cells(7, 5).Value = "Last Run"
    ws.Cells(7, 5).Font.Bold = True
    ws.Cells(7, 6).Value = Now()
    ws.Cells(7, 6).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    
    ' Failure reason breakdown (row 12+)
    ws.Cells(10, 1).Value = "Failure Reason Breakdown"
    ws.Cells(10, 1).Font.Bold = True
    ws.Cells(10, 1).Font.Size = 12
    
    ws.Cells(11, 1).Value = "Failure Reason"
    ws.Cells(11, 2).Value = "Count"
    ws.Cells(11, 3).Value = "% of Ineligible"
    ws.Cells(11, 1).Font.Bold = True
    ws.Cells(11, 2).Font.Bold = True
    ws.Cells(11, 3).Font.Bold = True
    
    Dim fKeys() As Variant
    Dim rowIdx As Long
    rowIdx = 12
    
    If failureDict.Count > 0 Then
        fKeys = failureDict.keys
        Dim fk As Long
        For fk = 0 To UBound(fKeys)
            ws.Cells(rowIdx, 1).Value = fKeys(fk)
            ws.Cells(rowIdx, 2).Value = failureDict(fKeys(fk))
            If ineligible > 0 Then
                ws.Cells(rowIdx, 3).Value = failureDict(fKeys(fk)) / ineligible
            End If
            ws.Cells(rowIdx, 3).NumberFormat = "0.0%"
            rowIdx = rowIdx + 1
        Next fk
    End If
    
    ' Country breakdown (row 10, column 5+)
    ws.Cells(10, 5).Value = "Country Breakdown"
    ws.Cells(10, 5).Font.Bold = True
    ws.Cells(10, 5).Font.Size = 12
    
    ws.Cells(11, 5).Value = "Country"
    ws.Cells(11, 6).Value = "Eligible"
    ws.Cells(11, 7).Value = "Ineligible"
    ws.Cells(11, 8).Value = "Rate"
    ws.Cells(11, 5).Font.Bold = True
    ws.Cells(11, 6).Font.Bold = True
    ws.Cells(11, 7).Font.Bold = True
    ws.Cells(11, 8).Font.Bold = True
    
    Dim cKeys() As Variant
    rowIdx = 12
    
    If countryElig.Count > 0 Then
        cKeys = countryElig.keys
        Dim ck As Long
        For ck = 0 To UBound(cKeys)
            Dim cElig As Long, cInelig As Long
            cElig = countryElig(cKeys(ck))
            If countryInelig.Exists(cKeys(ck)) Then
                cInelig = countryInelig(cKeys(ck))
            Else
                cInelig = 0
            End If
            
            ws.Cells(rowIdx, 5).Value = cKeys(ck)
            ws.Cells(rowIdx, 6).Value = cElig
            ws.Cells(rowIdx, 7).Value = cInelig
            If (cElig + cInelig) > 0 Then
                ws.Cells(rowIdx, 8).Value = cElig / (cElig + cInelig)
            End If
            ws.Cells(rowIdx, 8).NumberFormat = "0.0%"
            rowIdx = rowIdx + 1
        Next ck
    End If
End Sub

'---------------------------------------------------------------------
' CountConcentrationBreaches: Counts BREACH entries in Concentration sheet
'---------------------------------------------------------------------
Private Function CountConcentrationBreaches() As Long
    Dim ws As Worksheet
    Dim cell As Range
    Dim count As Long
    
    Set ws = ThisWorkbook.Sheets("Concentration Analysis")
    count = 0
    
    For Each cell In ws.Range("F5:F45")
        If CStr(cell.Value) = "BREACH" Then
            count = count + 1
        End If
    Next cell
    
    CountConcentrationBreaches = count
End Function

'---------------------------------------------------------------------
' ApplyConditionalFormatting: Applies dynamic formatting across sheets
'---------------------------------------------------------------------
Public Sub ApplyConditionalFormatting()
    ApplyResultsFormatting
    ApplyIntegrityFormatting
End Sub

'---------------------------------------------------------------------
' ApplyResultsFormatting: Colour-codes Validation Results sheet
'---------------------------------------------------------------------
Private Sub ApplyResultsFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, c As Long
    
    Set ws = ThisWorkbook.Sheets("Validation Results")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 4 Then Exit Sub
    
    For i = 4 To lastRow
        ' Per-criterion columns (E through K)
        For c = 5 To 11
            Select Case UCase(CStr(ws.Cells(i, c).Value))
                Case "PASS"
                    ws.Cells(i, c).Interior.Color = RGB(232, 245, 233)
                    ws.Cells(i, c).Font.Color = RGB(39, 174, 96)
                Case "FAIL"
                    ws.Cells(i, c).Interior.Color = RGB(255, 235, 238)
                    ws.Cells(i, c).Font.Color = RGB(231, 76, 60)
                Case "N/A"
                    ws.Cells(i, c).Interior.Color = RGB(255, 248, 225)
                    ws.Cells(i, c).Font.Color = RGB(243, 156, 18)
            End Select
        Next c
        
        ' Overall result column (L)
        Select Case UCase(CStr(ws.Cells(i, 12).Value))
            Case "ELIGIBLE"
                ws.Cells(i, 12).Interior.Color = RGB(39, 174, 96)
                ws.Cells(i, 12).Font.Color = RGB(255, 255, 255)
                ws.Cells(i, 12).Font.Bold = True
            Case "INELIGIBLE"
                ws.Cells(i, 12).Interior.Color = RGB(231, 76, 60)
                ws.Cells(i, 12).Font.Color = RGB(255, 255, 255)
                ws.Cells(i, 12).Font.Bold = True
        End Select
    Next i
End Sub

'---------------------------------------------------------------------
' ApplyIntegrityFormatting: Colour-codes Data Integrity sheet
'---------------------------------------------------------------------
Private Sub ApplyIntegrityFormatting()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Data Integrity")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 5 Then Exit Sub
    
    For i = 5 To lastRow
        Select Case UCase(CStr(ws.Cells(i, 6).Value))
            Case "CRITICAL"
                ws.Cells(i, 6).Interior.Color = RGB(255, 235, 238)
                ws.Cells(i, 6).Font.Color = RGB(231, 76, 60)
                ws.Cells(i, 6).Font.Bold = True
            Case "WARNING"
                ws.Cells(i, 6).Interior.Color = RGB(255, 248, 225)
                ws.Cells(i, 6).Font.Color = RGB(243, 156, 18)
                ws.Cells(i, 6).Font.Bold = True
        End Select
    Next i
End Sub
