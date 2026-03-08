Attribute VB_Name = "modFXConversion"
'=============================================================================
' Module: modFXConversion
' Purpose: Multi-currency handling - converts local currency amounts to EUR
'          using rates from the FX Rates sheet
' Author: Akhil Vohra
'=============================================================================
Option Explicit

Public Function ConvertToEUR(ByVal amount As Double, ByVal currCode As String) As Double
    Dim rate As Double
    
    If UCase(Trim(currCode)) = "EUR" Then
        ConvertToEUR = amount
        Exit Function
    End If
    
    rate = GetFXRate(currCode)
    
    If rate = 0 Then
        ConvertToEUR = 0
    Else
        ConvertToEUR = amount / rate
    End If
End Function

Public Function GetFXRate(ByVal currCode As String) As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FX Rates")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    currCode = UCase(Trim(currCode))
    
    For i = 4 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = currCode Then
            If IsNumeric(ws.Cells(i, 2).Value) Then
                GetFXRate = CDbl(ws.Cells(i, 2).Value)
            Else
                GetFXRate = 0
            End If
            Exit Function
        End If
    Next i
    
    GetFXRate = 0
End Function

Public Sub RefreshEURAmounts()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim localAmt As Variant
    Dim ccy As String
    Dim eurAmt As Double
    Dim updated As Long
    
    Set ws = ThisWorkbook.Sheets("Loan Portfolio")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    updated = 0
    For i = 2 To lastRow
        localAmt = ws.Cells(i, 9).Value
        ccy = CStr(ws.Cells(i, 10).Value)
        
        If IsNumeric(localAmt) And Trim(ccy) <> "" Then
            eurAmt = ConvertToEUR(CDbl(localAmt), ccy)
            ws.Cells(i, 11).Value = eurAmt
            ws.Cells(i, 11).NumberFormat = "#,##0"
            updated = updated + 1
        End If
    Next i
    
    MsgBox updated & " EUR amounts refreshed using current FX rates.", vbInformation
End Sub

Public Function ValidateFXRates() As String
    Dim ws As Worksheet
    Dim wsLoans As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim requiredCurrencies As Object
    Dim availableCurrencies As Object
    Dim missing As String
    Dim ccy As String
    
    Set requiredCurrencies = CreateObject("Scripting.Dictionary")
    Set availableCurrencies = CreateObject("Scripting.Dictionary")
    
    Set wsLoans = ThisWorkbook.Sheets("Loan Portfolio")
    lastRow = wsLoans.Cells(wsLoans.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        ccy = UCase(Trim(CStr(wsLoans.Cells(i, 10).Value)))
        If ccy <> "" And ccy <> "EUR" Then
            If Not requiredCurrencies.Exists(ccy) Then
                requiredCurrencies.Add ccy, 1
            End If
        End If
    Next i
    
    Set ws = ThisWorkbook.Sheets("FX Rates")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 4 To lastRow
        ccy = UCase(Trim(CStr(ws.Cells(i, 1).Value)))
        If ccy <> "" Then
            If Not availableCurrencies.Exists(ccy) Then
                availableCurrencies.Add ccy, CDbl(ws.Cells(i, 2).Value)
            End If
        End If
    Next i
    
    missing = ""
    Dim keys() As Variant
    If requiredCurrencies.Count > 0 Then
        keys = requiredCurrencies.keys
        For i = 0 To UBound(keys)
            If Not availableCurrencies.Exists(keys(i)) Then
                If missing <> "" Then missing = missing & ", "
                missing = missing & keys(i)
            End If
        Next i
    End If
    
    If missing <> "" Then
        ValidateFXRates = "Missing FX rates for: " & missing
    Else
        ValidateFXRates = "OK"
    End If
End Function
