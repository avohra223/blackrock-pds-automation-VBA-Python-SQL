Attribute VB_Name = "modConcentration"
'=============================================================================
' Module: modConcentration
' Purpose: Portfolio-level concentration analysis (borrower, sector, country)
'          Checks against configurable limits from Eligibility Criteria sheet
' Author: Akhil Vohra
'=============================================================================
Option Explicit

Private Type ConcentrationEntry
    Name As String
    LoanCount As Long
    EURExposure As Double
    PctOfPortfolio As Double
    Limit As Double
    Status As String
End Type

'---------------------------------------------------------------------
' RunConcentrationAnalysis: Main concentration check routine
'---------------------------------------------------------------------
Public Sub RunConcentrationAnalysis(ByRef loans() As LoanRecord, _
                                     ByRef criteria As EligibilityCriteria)
    Dim ws As Worksheet
    Dim wsCrit As Worksheet
    Dim totalEUR As Double
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Concentration Analysis")
    Set wsCrit = ThisWorkbook.Sheets("Eligibility Criteria")
    
    ' Clear previous data
    ws.Range("A5:F16").ClearContents
    ws.Range("A19:F30").ClearContents
    ws.Range("A33:F45").ClearContents
    
    ' Calculate total EUR exposure
    totalEUR = 0
    For i = 1 To UBound(loans)
        If IsNumeric(loans(i).LoanAmtEUR) Then
            totalEUR = totalEUR + CDbl(loans(i).LoanAmtEUR)
        End If
    Next i
    
    If totalEUR = 0 Then Exit Sub
    
    ' Read concentration limits from criteria sheet
    Dim singleBorrowerLimit As Double
    Dim top10Limit As Double
    Dim sectorLimit As Double
    Dim countryLimit As Double
    
    singleBorrowerLimit = SafeDblConc(wsCrit.Cells(34, 2).Value, 0.05)
    top10Limit = SafeDblConc(wsCrit.Cells(35, 2).Value, 0.3)
    sectorLimit = SafeDblConc(wsCrit.Cells(36, 2).Value, 0.25)
    countryLimit = SafeDblConc(wsCrit.Cells(37, 2).Value, 0.35)
    
    ' If stored as percentage (>1), convert to decimal
    If singleBorrowerLimit > 1 Then singleBorrowerLimit = singleBorrowerLimit / 100
    If top10Limit > 1 Then top10Limit = top10Limit / 100
    If sectorLimit > 1 Then sectorLimit = sectorLimit / 100
    If countryLimit > 1 Then countryLimit = countryLimit / 100
    
    ' 1. BORROWER CONCENTRATION
    AnalyseBorrowerConcentration loans, ws, totalEUR, singleBorrowerLimit, top10Limit
    
    ' 2. SECTOR CONCENTRATION
    AnalyseDimensionConcentration loans, ws, totalEUR, sectorLimit, "Sector", 19
    
    ' 3. COUNTRY CONCENTRATION
    AnalyseDimensionConcentration loans, ws, totalEUR, countryLimit, "Country", 33
End Sub

'---------------------------------------------------------------------
' AnalyseBorrowerConcentration: Top borrower exposure analysis
'---------------------------------------------------------------------
Private Sub AnalyseBorrowerConcentration(ByRef loans() As LoanRecord, _
                                          ByRef ws As Worksheet, _
                                          ByVal totalEUR As Double, _
                                          ByVal singleLimit As Double, _
                                          ByVal top10Limit As Double)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim key As String
    
    ' Aggregate by borrower
    For i = 1 To UBound(loans)
        key = loans(i).Borrower
        If IsNumeric(loans(i).LoanAmtEUR) Then
            If dict.Exists(key) Then
                dict(key) = dict(key) + CDbl(loans(i).LoanAmtEUR)
            Else
                dict.Add key, CDbl(loans(i).LoanAmtEUR)
            End If
        End If
    Next i
    
    ' Sort by exposure (descending) - simple bubble sort for top 10
    Dim keys() As Variant, vals() As Variant
    keys = dict.keys
    vals = dict.Items
    
    Dim temp As Variant
    Dim tempStr As String
    Dim j As Long
    
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If vals(j) > vals(i) Then
                temp = vals(i): vals(i) = vals(j): vals(j) = temp
                tempStr = keys(i): keys(i) = keys(j): keys(j) = tempStr
            End If
        Next j
    Next i
    
    ' Write top 10 (or all if fewer)
    Dim writeCount As Long
    writeCount = Application.WorksheetFunction.Min(10, dict.Count)
    Dim top10Sum As Double
    top10Sum = 0
    
    For i = 0 To writeCount - 1
        Dim r As Long
        r = 5 + i
        Dim pct As Double
        pct = vals(i) / totalEUR
        top10Sum = top10Sum + vals(i)
        
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = keys(i)
        ws.Cells(r, 3).Value = vals(i)
        ws.Cells(r, 3).NumberFormat = "#,##0"
        ws.Cells(r, 4).Value = pct
        ws.Cells(r, 4).NumberFormat = "0.0%"
        ws.Cells(r, 5).Value = singleLimit
        ws.Cells(r, 5).NumberFormat = "0.0%"
        
        If pct > singleLimit Then
            ws.Cells(r, 6).Value = "BREACH"
            ws.Cells(r, 6).Font.Color = RGB(231, 76, 60)
            ws.Cells(r, 6).Font.Bold = True
        Else
            ws.Cells(r, 6).Value = "OK"
            ws.Cells(r, 6).Font.Color = RGB(39, 174, 96)
        End If
    Next i
    
    ' Top 10 combined check
    r = 5 + writeCount + 1
    ws.Cells(r, 1).Value = ""
    ws.Cells(r, 2).Value = "Top 10 Combined"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = top10Sum
    ws.Cells(r, 3).NumberFormat = "#,##0"
    ws.Cells(r, 4).Value = top10Sum / totalEUR
    ws.Cells(r, 4).NumberFormat = "0.0%"
    ws.Cells(r, 5).Value = top10Limit
    ws.Cells(r, 5).NumberFormat = "0.0%"
    
    If top10Sum / totalEUR > top10Limit Then
        ws.Cells(r, 6).Value = "BREACH"
        ws.Cells(r, 6).Font.Color = RGB(231, 76, 60)
    Else
        ws.Cells(r, 6).Value = "OK"
        ws.Cells(r, 6).Font.Color = RGB(39, 174, 96)
    End If
End Sub

'---------------------------------------------------------------------
' AnalyseDimensionConcentration: Generic concentration for sector/country
'---------------------------------------------------------------------
Private Sub AnalyseDimensionConcentration(ByRef loans() As LoanRecord, _
                                           ByRef ws As Worksheet, _
                                           ByVal totalEUR As Double, _
                                           ByVal limit As Double, _
                                           ByVal dimension As String, _
                                           ByVal startRow As Long)
    Dim dict As Object
    Dim countDict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set countDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim key As String
    
    For i = 1 To UBound(loans)
        If dimension = "Sector" Then
            key = loans(i).Sector
        Else
            key = loans(i).Country
        End If
        
        If IsNumeric(loans(i).LoanAmtEUR) Then
            If dict.Exists(key) Then
                dict(key) = dict(key) + CDbl(loans(i).LoanAmtEUR)
                countDict(key) = countDict(key) + 1
            Else
                dict.Add key, CDbl(loans(i).LoanAmtEUR)
                countDict.Add key, 1
            End If
        End If
    Next i
    
    ' Sort descending
    Dim keys() As Variant, vals() As Variant
    keys = dict.keys
    vals = dict.Items
    Dim temp As Variant, tempStr As String, j As Long
    
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If vals(j) > vals(i) Then
                temp = vals(i): vals(i) = vals(j): vals(j) = temp
                tempStr = keys(i): keys(i) = keys(j): keys(j) = tempStr
            End If
        Next j
    Next i
    
    ' Write results
    For i = 0 To UBound(keys)
        Dim r As Long
        r = startRow + i
        Dim pct As Double
        pct = vals(i) / totalEUR
        
        ws.Cells(r, 1).Value = keys(i)
        ws.Cells(r, 2).Value = countDict(keys(i))
        ws.Cells(r, 3).Value = vals(i)
        ws.Cells(r, 3).NumberFormat = "#,##0"
        ws.Cells(r, 4).Value = pct
        ws.Cells(r, 4).NumberFormat = "0.0%"
        ws.Cells(r, 5).Value = limit
        ws.Cells(r, 5).NumberFormat = "0.0%"
        
        If pct > limit Then
            ws.Cells(r, 6).Value = "BREACH"
            ws.Cells(r, 6).Font.Color = RGB(231, 76, 60)
            ws.Cells(r, 6).Font.Bold = True
        Else
            ws.Cells(r, 6).Value = "OK"
            ws.Cells(r, 6).Font.Color = RGB(39, 174, 96)
        End If
    Next i
End Sub

Private Function SafeDblConc(val As Variant, defaultVal As Double) As Double
    If IsNumeric(val) And Not IsEmpty(val) Then
        SafeDblConc = CDbl(val)
    Else
        SafeDblConc = defaultVal
    End If
End Function
