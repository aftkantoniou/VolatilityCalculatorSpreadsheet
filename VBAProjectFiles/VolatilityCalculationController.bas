Attribute VB_Name = "VolatilityCalculationController"
Option Explicit

    Private diWsDataLastCol                 As Long
    Private diWsDataLastRow                 As Long
    
    Private errorCol                        As Integer
    
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
    
    ' Checking whether the inputs that will be used are valid numbers
    ' ===============================================================
    errorCol = vbEmpty
    For i = 2 To diWsDataLastRow
        If IsNumeric(diWs.Cells(i, diCloseCol).Value) = False Then
            errorCol = diCloseCol
            GoTo DataInputErrorHandler
        End If
    Next i
    
    
    
    crWs.Cells(4, crCalcResCloseToCloseCol).Value = vbNullString
    crWs.Cells(4, crCalcResCloseToCloseCol).Value = VolatilityCalculationService.getCloseToCloseVolatility(diWsDataLastRow)
    
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
           & "Column where the problem occured: " & errorCol _
           & vbNewLine _
           & vbNewLine _
           & "Row there the problem occured: " & i _
           & vbNewLine _
           & vbNewLine _
           & "Please remedy the prolem and restart the process."
    GoTo SubExit

           
End Sub
