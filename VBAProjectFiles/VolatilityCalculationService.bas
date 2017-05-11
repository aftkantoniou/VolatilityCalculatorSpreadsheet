Attribute VB_Name = "VolatilityCalculationService"
Option Explicit

' This module provides the collection of functions which are used to calculate historical volatility
' with different valuation methods.
'
' WARNING
' -------
'
' A generic characteristic of the code used is that a "non negative entry" check is applied before
' every loop which is used to calculate historical volatility with all these different methods.
' The aforementioned check is used to guarantee that no negative or zero numbers will be used as
' inputs to the log functions used in the models applied.
'
' In the case that a negative or zero number is found the record is skipped and no error will be thrown.
' If you validate the resuls of these calculation and you see discrepancies with your validation reference
' without an error on another obvious reason please check the data input for negative entries or ultimately
' debug this code in order to find out if and where an error or miscalculation occurs.
'
'
' Referance implementation note:
' ------------------------------
'
' The calculations made below were done using the implementations of the respective models presented in the
' "Measuring Historical Volatility" paper distributed by Santander on February 3, 2012.
' =========================================================================================================

Public Function getCloseToCloseVolatility(ByVal dataLastRow As Long, ByVal annualizationFactor As Integer) As Double

    ' This function computes the historical volatility by the close to close model.
    ' =============================================================================
    
    ' Symbolic links to variable names:
    ' logReturnsSdFmSum = sum of squared deviations from mean logReturn
    ' =================================================================
    
    Dim i                       As Long
    Dim arrayCounter            As Long
    
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
    getCloseToCloseVolatility = Sqr(logReturnsSdFmSum / (logReturnsNumber - 1)) * Sqr(annualizationFactor)

End Function

Public Function getGarmanKlassVolatility(ByVal dataLastRow As Long, ByVal annualizationFactor As Integer) As Double

    ' This function computes the historical volatility using the Garman-Klass model.
    ' ==============================================================================

    Dim i       As Long
    Dim sumGk   As Double: sumGk = vbEmpty
    
    For i = 2 To dataLastRow
        With diWs
            If (.Cells(i, diCloseCol).Value > 0) And (.Cells(i, diOpenCol).Value > 0) And (.Cells(i, diHighCol).Value > 0) And (.Cells(i, diLowCol).Value > 0) Then
                sumGk = sumGk + ((1 / 2 * ((Log(.Cells(i, diHighCol).Value) - Log(.Cells(i, diLowCol).Value)) * (Log(.Cells(i, diHighCol).Value) - Log(.Cells(i, diLowCol).Value)))) _
                              - ((2 * Log(2) - 1) _
                              * (((Log(.Cells(i, diCloseCol).Value) - Log(.Cells(i, diOpenCol).Value)) * (Log(.Cells(i, diCloseCol).Value) - Log(.Cells(i, diOpenCol).Value))))))
            End If
        End With
    Next i
        
    getGarmanKlassVolatility = Sqr(sumGk / (dataLastRow - 1)) * Sqr(annualizationFactor)

End Function

Public Function getRogersSatchellVolatility(dataLastRow, annualizationFactor) As Double

    ' This function computes the historical volatility using the Rogers-Satchell model.
    ' =================================================================================

    Dim i       As Long
    Dim sumRs   As Double: sumRs = vbEmpty
    
    For i = 2 To dataLastRow
        With diWs
            If (.Cells(i, diCloseCol).Value > 0) And (.Cells(i, diOpenCol).Value > 0) And (.Cells(i, diHighCol).Value > 0) And (.Cells(i, diLowCol).Value > 0) Then
                sumRs = sumRs + ((Log(.Cells(i, diHighCol).Value) - Log(.Cells(i, diCloseCol).Value)) * (Log(.Cells(i, diHighCol).Value) - Log(.Cells(i, diOpenCol).Value))) _
                              + ((Log(.Cells(i, diLowCol).Value) - Log(.Cells(i, diCloseCol).Value)) * (Log(.Cells(i, diLowCol).Value) - Log(.Cells(i, diOpenCol).Value)))
            End If
        End With
    Next i
    
    getRogersSatchellVolatility = Sqr(sumRs / (dataLastRow - 1)) * Sqr(annualizationFactor)
    
End Function

Public Function getGarmanKlassYangZhangVolatility(dataLastRow, annualizationFactor) As Double
    
    Dim i       As Long
    Dim sumGkYz As Double: sumGkYz = vbEmpty
    
    For i = 2 To (dataLastRow - 1)
        With diWs
            If (.Cells(i, diCloseCol).Value > 0) And (.Cells(i, diOpenCol).Value > 0) And (.Cells(i, diHighCol).Value > 0) And (.Cells(i, diLowCol).Value > 0) And (.Cells(i + 1, diCloseCol).Value > 0) Then
                sumGkYz = sumGkYz + ((Log(.Cells(i, diOpenCol).Value) - Log(.Cells(i + 1, diCloseCol).Value)) * (Log(.Cells(i, diOpenCol).Value) - Log(.Cells(i + 1, diCloseCol).Value)) _
                                  + ((1 / 2 * ((Log(.Cells(i, diHighCol).Value) - Log(.Cells(i, diLowCol).Value)) * (Log(.Cells(i, diHighCol).Value) - Log(.Cells(i, diLowCol).Value)))) _
                                  - ((2 * Log(2) - 1) _
                                  * (((Log(.Cells(i, diCloseCol).Value) - Log(.Cells(i, diOpenCol).Value)) * (Log(.Cells(i, diCloseCol).Value) - Log(.Cells(i, diOpenCol).Value)))))))
            End If
        End With
    Next i
    
    getGarmanKlassYangZhangVolatility = Sqr(sumGkYz / (dataLastRow - 1)) * Sqr(annualizationFactor)

End Function

Public Function getYangZhangVolatility(dataLastRow, annualizationFactor) As Double

    Dim i                   As Long
    Dim n                   As Long: n = vbEmpty
    
    Dim k                   As Double: k = vbEmpty
    Dim sumYz               As Double: sumYz = vbEmpty
    Dim overnigthVol        As Double: overnigthVol = vbEmpty
    Dim openToCloseVol      As Double: openToCloseVol = vbEmpty
    Dim overnightSum        As Double: overnightSum = vbEmpty
    Dim openToCloseSum      As Double: openToCloseSum = vbEmpty
    Dim overnightVolMean    As Double: overnightVolMean = vbEmpty
    Dim openToCloseVolMean  As Double: openToCloseVolMean = vbEmpty
    
    n = dataLastRow - 1
    
    ' Constant used to minimize the variance of the estimator
    ' =======================================================
    k = 0.34 / (1.34 + ((n + 1) / (n - 1)))
    
    ' Loop used to calculate the mean values needed for the generic calculation
    ' =========================================================================
    For i = 2 To n
        With diWs
            If (.Cells(i, diCloseCol).Value > 0) And (.Cells(i, diOpenCol).Value > 0) And (.Cells(i + 1, diCloseCol).Value > 0) Then
                overnightSum = overnightSum + (Log(.Cells(i, diOpenCol).Value) - Log(.Cells(i + 1, diCloseCol).Value))
                openToCloseSum = openToCloseSum + (Log(.Cells(i, diCloseCol).Value) - Log(.Cells(i, diOpenCol).Value))
            End If
        End With
    
End Function
