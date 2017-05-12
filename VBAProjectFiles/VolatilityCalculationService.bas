Attribute VB_Name = "VolatilityCalculationService"
Option Explicit

    Private sumRs As Double

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
' If you validate the resuls of these calculations and you see discrepancies with your validation reference
' without an error popping up on another obvious reason please check the data input for negative entries or ultimately
' debug this code in order to find out if and where an error or a miscalculation occurs.
'
'
' Reference implementation note:
' ------------------------------
'
' The calculations made below were done using the formulas of the respective models presented in the
' "Measuring Historical Volatility" paper distributed by Santander on February 3, 2012.
'
'
' Parameters used in the following functions:
' -------------------------------------------
'
' @param dataLastRow Last row number of the given input data located in the "Data Import" sheet
' @param annualizationFactor Factor to be used in all the calculations in order to annualize the volatility number computed
'
'
' Generic implementation notes concerning the function below:
' -----------------------------------------------------------
'
' "Power of 2" calculation is not done with the traditional way by squaring a quantity (x ^ 2).
' Instead repeated multiplication is used (x * x) for the reasons below:
'
' 1. As this calculation will be used within recursive sumation routines, by following the approach described above, we minimize
'    the roundoff error which will occur during the sumation process.
' 2. The above implementation is faster calculation wise.
'
'
' Use of "two-pass" algorithms where is needed.
' ---------------------------------------------
' For the models which require computation of mean values the decision was to use two-pass algorithms, and not one-pass (online algorithms),
' due to the fact that the majority of them present numericaly unstable behaviour except some specific implementations. Also, these
' algorithms are chosen to be used in cases where the data sets are extremely large or the calculation is required to be done as the
' data is generated. Here we do not deal with the latter cases.
' However, if there will be a need for very large data sets calculation or better accuracy in the produced results, some of the
' implementations will be revisited. Before that though, is advisable to revisit the whole framework, as this code plays the role of a
' prototyping tool and it probably will be better to change to a different languange ecosystem when the above needs started to come to
' the surface.
'
' Two-pass algorithms provide safety in two fronts:
'
' 1. Majority of one-pass algorithms are prone to provide poor results in the presence of rounding errors,
'    as they compute variances as differences of two positive numbers. Therefore, can suffer severe cancellation that leaves the
'    computed result dominated by roundoff. Also sometimes, the computed result can be even negative. In contrast, two pass algorithms
'    yield a very accurate and nonnegative result, unless N is very large.
'
' 2. There are one-pass algorithms, such as Welford algorithm, that can produce accurate results. However, for the
'    scope of this tool were not needed due to the size of the data we are dealing with and also because their error bound is not as small as
'    the one for the two-pass algorithm.
' ============================================================================================================================================

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
        
        ' The reason for the i / i+1 sequence is that the data is
        ' ordered in descending order. Thus, the latest data point is (i)
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
    ' ================================================================
    If UtilitiesArrays.isArrayAllocated(logReturns) = True Then
        For i = LBound(logReturns) To UBound(logReturns)
            logReturnsSdFmSum = logReturnsSdFmSum + ((logReturns(i) - logReturnsMean) * (logReturns(i) - logReturnsMean))
        Next i
    End If
    
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
    
    sumRs = vbEmpty
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

    ' This function computes the historical volatility using the Garman-Klass Yang-Zhang Extension model.
    ' ===================================================================================================
    
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

    ' This function computes the historical volatility using the Yang-Zhang model.
    ' ============================================================================

    Dim i                   As Long
    Dim n                   As Long: n = vbEmpty
    
    Dim k                   As Double: k = vbEmpty
    Dim sumYz               As Double: sumYz = vbEmpty
    Dim rsVariance          As Double: rsVariance = vbEmpty
    Dim overnigthVariance   As Double: overnigthVariance = vbEmpty
    Dim openToCloseVariance As Double: openToCloseVariance = vbEmpty
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
    Next i
    
    ' Calculation of the means needed for the generic calculation
    ' ===========================================================
    overnightVolMean = overnightSum / n
    openToCloseVolMean = openToCloseVolMean / n
    
    ' Calculation of the respective sums which map to the different volatilities used for the generic calculation
    ' ===========================================================================================================
    overnightSum = vbEmpty
    openToCloseSum = vbEmpty
    For i = 2 To n
        With diWs
            If (.Cells(i, diCloseCol).Value > 0) And (.Cells(i, diOpenCol).Value > 0) And (.Cells(i + 1, diCloseCol).Value > 0) Then
                overnightSum = overnightSum + (((Log(.Cells(i, diOpenCol).Value) - Log(.Cells(i + 1, diCloseCol).Value)) - overnightVolMean) * ((Log(.Cells(i, diOpenCol).Value) - Log(.Cells(i + 1, diCloseCol).Value)) - overnightVolMean))
                openToCloseSum = openToCloseSum + (((Log(.Cells(i, diCloseCol).Value) - Log(.Cells(i, diOpenCol).Value)) - openToCloseVolMean) * ((Log(.Cells(i, diOpenCol).Value) - Log(.Cells(i + 1, diCloseCol).Value)) - overnightVolMean))
            End If
        End With
    Next i
    
    overnigthVariance = overnightSum / (n - 1)
    openToCloseVariance = openToCloseSum / (n - 1)
    rsVariance = sumRs / n
    
    getYangZhangVolatility = Sqr(overnigthVariance + (k * openToCloseVariance) + ((1 - k) * rsVariance)) * Sqr(annualizationFactor)
    
End Function
