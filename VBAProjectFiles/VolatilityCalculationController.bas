Attribute VB_Name = "VolatilityCalculationController"
Option Explicit

    Private diWsDataLastCol                 As Long
    Private diWsDataLastRow                 As Long
    
    Private errorCol                        As Integer
    Private annualizationFactor             As Integer
    
    Private calculationMethodologySelected  As String
    
    ' This module provides the shell where the volatility calculation service is initiated.
    ' It initiates the appropriate data mapping that has to be done and builds the appropriate
    ' variables which have to be used from the actual volatility calculation methods hosted in
    ' volatility calculation methodologies service (module name: VolCalcMethodologies Service).
    ' =========================================================================================

Public Sub VolatilityCalculationController()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Cursor = xlNorthwestArrow
    End With
    
    Call ColumnDataMappingService.DataImportSheetDataMapping
    Call ColumnDataMappingService.CalculationResultsSheetDataMapping
    
    With diWs
        diWsDataLastCol = .Range("iy1").End(xlToLeft).Column
        diWsDataLastRow = .Range(Split(diWs.Cells(1, diDateCol).Address, "$")(1) & .Rows.Count).End(xlUp).Row
    End With
    
    ' Following there is an error hanlder in place which checks the validity of the dates
    ' provided in the Data Import sheet. Rule is that the dates have to be ordered with
    ' descending order. If otherwise, the check will prone the user for the error before the
    ' calculation service loop will start.
    '========================================================================================
    
    Dim i As Long: i = 0
    For i = 2 To diWsDataLastRow
        If (diWs.Cells(i, diDateCol).Value <= diWs.Cells(i + 1, diDateCol).Value) Then
            GoTo DatesErrorHandler
        End If
    Next i
    
    ' Volatility Calculation Controller initiation
    ' ============================================
    
    annualizationFactor = crWs.Cells(8, crAnnualizationFactorCol).Value
    
    ' Checking whether the inputs that will be used are valid numbers
    ' ===============================================================
    errorCol = vbEmpty
    For i = 2 To diWsDataLastRow
        Select Case False
            Case IsNumeric(diWs.Cells(i, diCloseCol).Value), IsNumeric(diWs.Cells(i, diOpenCol).Value), IsNumeric(diWs.Cells(i, diHighCol).Value), IsNumeric(diWs.Cells(i, diLowCol).Value)
                GoTo DataInputErrorHandler
        End Select
    Next i
    
    ' Close to Close Model
    ' ====================
    crWs.Cells(4, crCalcResCloseToCloseCol).Value = vbNullString
    crWs.Cells(4, crCalcResCloseToCloseCol).Value = VolatilityCalculationService.getCloseToCloseVolatility(diWsDataLastRow, annualizationFactor)
    
    ' Garman - Klass Model
    ' ====================
    crWs.Cells(4, crCalcResGarmanKlassCol).Value = vbNullString
    crWs.Cells(4, crCalcResGarmanKlassCol).Value = VolatilityCalculationService.getGarmanKlassVolatility(diWsDataLastRow, annualizationFactor)
    
    ' Rogers - Satchell Model
    ' =======================
    crWs.Cells(4, crCalcResRogersSatcellCol).Value = vbNullString
    crWs.Cells(4, crCalcResRogersSatcellCol).Value = VolatilityCalculationService.getRogersSatchellVolatility(diWsDataLastRow, annualizationFactor)
    
SubExit:

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Cursor = xlDefault
    End With

Exit Sub

DatesErrorHandler:

    MsgBox "There seems to be a problem in the sequence with the dates are provided in the 'Data Import' sheet." _
           & vbNewLine _
           & vbNewLine _
           & "Dates have to be laid out with descending order." _
           & vbNewLine _
           & vbNewLine _
           & "However the following ones are not." _
           & vbNewLine _
           & vbNewLine _
           & "First row in the problematic sequence: " & i _
           & vbNewLine _
           & vbNewLine _
           & "Next row in the problematic sequence: " & i - 1 _
           & vbNewLine _
           & vbNewLine _
           & "The volatility calculation will now stop." _
           & vbNewLine _
           & vbNewLine _
           & "Please remedy the prolem and restart the process."
    GoTo SubExit
    
DataInputErrorHandler:

    MsgBox "There seems to be a problem in the data input format of the data which have to be processed for the calculation method chosen." _
           & vbNewLine _
           & vbNewLine _
           & "The data have to be numeric. However, some of them are not." _
           & vbNewLine _
           & vbNewLine _
           & "Please remedy the prolem and restart the process."
    GoTo SubExit

           
End Sub
