Attribute VB_Name = "VolatilityCalculationService"
Option Explicit

Public Function getCloseToCloseVolatility(ByVal dataLastRow As Long)

    ' This function computes the historical volatility by the close to close method.
    ' ==============================================================================
    
    Dim logReturns() As Double

    Dim i As Long: i = 0
    For i = 2 To dataLastRow
        ' the log return for two returns will be computed.
        
        ' this logreturn will be added the logReturns Array
        
        ' then sum of distances from the mean logReturn
        
        ' then vol (square root of the mean of the above distances)
        
        ' check for wrong inputs in log returns (0, or less than 0)

End Function
