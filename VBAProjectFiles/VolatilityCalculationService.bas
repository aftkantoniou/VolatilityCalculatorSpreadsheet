Attribute VB_Name = "VolatilityCalculationService"
Option Explicit

Public Function getCloseToCloseVolatility(ByVal dataLastRow As Long)

    ' This function computes the historical volatility by the close to close method.
    ' ==============================================================================
    
    ' Symbolic links to variable names:
    ' logReturnsSdFmSum = sum of squared deviations from mean logReturn
    ' =================================================================
    
    Dim i                       As Long
    Dim arrayCounter            As Long
    
    Dim annualizationFactor     As Integer
    
    Dim logReturn               As Double: logReturn = vbEmpty
    Dim logReturnsSum           As Double: logReturnsSum = vbEmpty
    Dim logReturnsNumber        As Double: logReturnsNumber = vbEmpty
    Dim logReturnsMean          As Double: logReturnsMean = vbEmpty
    Dim logReturnsSdFmSum       As Double: logReturnsSdFmSum = vbEmpty
    
    Dim logReturns()            As Double
    
    i = 2
    Do While (i + 1 <= dataLastRow)
        
        ' The reason for the i / i+1 sequence is that the data are
        ' order in descending order. Thus, the latest data point is (i)
        ' and the previous data point is (i+1)
        ' ============================================================
        
        ' The following check is done in order to solidify that no negative or zero number
        ' will be inputs to the log function.
        ' ================================================================================
        
        If (diWs.Cells(i, diCloseCol).Value > 0 And diWs.Cells(i + 1, diCloseCol).Value > 0) Then
            logReturn = Log(diWs.Cells(i, diCloseCol).Value) - Log(diWs.Cells(i + 1, diCloseCol).Value)
            ReDim Preserve logReturns(arrayCounter)
            logReturns(arrayCounter) = logReturn
        End If
        
        logReturnsSum = logReturnsSum + logReturn
        logReturnsNumber = logReturnsNumber + 1
        
        arrayCounter = arrayCounter + 1
        i = i + 1
    Loop
    
    ' Calculation of mean of log returns
    ' ==================================
    logReturnsMean = logReturnsSum / logReturnsNumber
    
    ' this logreturn will be added the logReturns Array
        
    ' Calculation of sum sum of squared deviations from mean logReturn
    ' Here the difference from mean log return is not "powered to 2".
    ' Instead is multiplied with itself because this operation is faster.
    ' ===================================================================
        
    If UtilitiesArrays.isArrayAllocated(logReturns) = True Then
        For i = LBound(logReturns) To UBound(logReturns)
            logReturnsSdFmSum = logReturnsSdFmSum + ((logReturns(i) - logReturnsMean) * (logReturns(i) - logReturnsMean))
        Next i
    End If
    
    ' Calculation of period's standard deviation
    ' ==========================================
    annualizationFactor = crWs.Cells(8, crAnnualizationFactorCol).Value
    getCloseToCloseVolatility = Sqr(logReturnsSdFmSum / (logReturnsNumber - 1)) * Sqr(annualizationFactor)

End Function
