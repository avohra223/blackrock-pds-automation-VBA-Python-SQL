Attribute VB_Name = "modRibbon"
'=============================================================================
' Module: modRibbon
' Purpose: Custom ribbon tab callbacks and toolbar button handlers
'          Note: Full ribbon customization requires customUI XML in the .xlsm
'          This module provides the callback procedures and a fallback
'          toolbar-based approach using CommandBar controls
' Author: Akhil Vohra
'
' RIBBON XML (to be added to customUI/customUI14.xml in the .xlsm package):
'
' <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
'   <ribbon>
'     <tabs>
'       <tab id="tabEligibility" label="Eligibility Checker" insertAfterMso="TabHome">
'         <group id="grpValidation" label="Validation">
'           <button id="btnRunCheck" label="Run Check"
'                   imageMso="AcceptInvitation" size="large"
'                   onAction="OnRunCheck" screentip="Run full eligibility validation"/>
'           <button id="btnOpenForm" label="Control Panel"
'                   imageMso="ControlWizards" size="large"
'                   onAction="OnOpenForm" screentip="Open the eligibility checker control panel"/>
'         </group>
'         <group id="grpData" label="Data">
'           <button id="btnRefreshFX" label="Refresh FX"
'                   imageMso="RecurrenceEdit" size="normal"
'                   onAction="OnRefreshFX" screentip="Refresh EUR amounts using current FX rates"/>
'           <button id="btnAddLoan" label="Add Loan"
'                   imageMso="RecordsAddFromOutlook" size="normal"
'                   onAction="OnAddLoan" screentip="Add a new loan record"/>
'         </group>
'         <group id="grpExport" label="Export">
'           <button id="btnExportXLSX" label="Export Report"
'                   imageMso="ExportExcel" size="normal"
'                   onAction="OnExportXLSX" screentip="Export results to a new workbook"/>
'           <button id="btnExportCSV" label="Export CSV"
'                   imageMso="ExportTextFile" size="normal"
'                   onAction="OnExportCSV" screentip="Export results as CSV"/>
'         </group>
'         <group id="grpHelp" label="Help">
'           <button id="btnInstructions" label="Instructions"
'                   imageMso="Help" size="normal"
'                   onAction="OnShowInstructions" screentip="View the instructions sheet"/>
'         </group>
'       </tab>
'     </tabs>
'   </ribbon>
' </customUI>
'=============================================================================
Option Explicit

'---------------------------------------------------------------------
' Ribbon Callback Procedures
' These are called by the custom ribbon XML
'---------------------------------------------------------------------
Public Sub OnRunCheck(control As IRibbonControl)
    RunFullValidation
End Sub

Public Sub OnOpenForm(control As IRibbonControl)
    ShowEligibilityForm
End Sub

Public Sub OnRefreshFX(control As IRibbonControl)
    RefreshEURAmounts
End Sub

Public Sub OnAddLoan(control As IRibbonControl)
    ' Quick add via InputBox sequence
    ShowEligibilityForm
End Sub

Public Sub OnExportXLSX(control As IRibbonControl)
    ExportEligibilityReport
End Sub

Public Sub OnExportCSV(control As IRibbonControl)
    ExportToCSV
End Sub

Public Sub OnShowInstructions(control As IRibbonControl)
    ThisWorkbook.Sheets("Instructions").Activate
End Sub

'---------------------------------------------------------------------
' CreateToolbar: Fallback toolbar creation for environments without
' custom ribbon XML support. Creates a CommandBar toolbar.
'---------------------------------------------------------------------
Public Sub CreateToolbar()
    Dim cb As Object
    Dim btn As Object
    
    On Error Resume Next
    ' Remove existing toolbar if present
    Application.CommandBars("EligibilityChecker").Delete
    On Error GoTo 0
    
    ' Create new toolbar
    Set cb = Application.CommandBars.Add(Name:="EligibilityChecker", _
                                          Position:=msoBarTop, Temporary:=True)
    cb.Visible = True
    
    ' Run Check button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Run Check"
        .Style = msoButtonCaption
        .FaceId = 2151
        .OnAction = "RunFullValidation"
        .TooltipText = "Run full eligibility validation"
    End With
    
    ' Control Panel button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Control Panel"
        .Style = msoButtonCaption
        .FaceId = 548
        .OnAction = "ShowEligibilityForm"
        .TooltipText = "Open eligibility checker control panel"
    End With
    
    ' Separator
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.BeginGroup = True
    
    ' Refresh FX button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Refresh FX"
        .Style = msoButtonCaption
        .OnAction = "RefreshEURAmounts"
        .TooltipText = "Refresh EUR conversions"
    End With
    
    ' Export Report button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Export Report"
        .Style = msoButtonCaption
        .OnAction = "ExportEligibilityReport"
        .TooltipText = "Export to new workbook"
    End With
    
    ' Export CSV button
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Export CSV"
        .Style = msoButtonCaption
        .OnAction = "ExportToCSV"
        .TooltipText = "Export as CSV"
    End With
End Sub

'---------------------------------------------------------------------
' RemoveToolbar: Clean up on workbook close
'---------------------------------------------------------------------
Public Sub RemoveToolbar()
    On Error Resume Next
    Application.CommandBars("EligibilityChecker").Delete
    On Error GoTo 0
End Sub
