Attribute VB_Name = "ColumnDataMappingService"
Option Explicit

    ' This module maps the columns which are used in all the
    ' sheets of volatility calculator workbook.
    '
    '
    ' Symbolic links to variable names:
    ' vc --> volatilityCalculator
    ' di --> dataImport
    ' cr --> calculationResults
    ' ================================
    
    Public vcWb                                 As Workbook
    Public diWs                                 As Worksheet
    Public crWs                                 As Worksheet
    
    Public diDateCol                            As Integer
    Public diOpenCol                            As Integer
    Public diCloseCol                           As Integer
    Public diHighCol                            As Integer
    Public diLowCol                             As Integer
    
    Public crCalcResCloseToCloseCol             As Integer
    Public crCalcResGarmanKlassCol              As Integer
    Public crCalcResRogersSatcellCol            As Integer
    Public crCalcResGarmanKlassYangZhangCol     As Integer
    Public crCalcResYangZhangCol                As Integer
    Public crAnnualizationFactorCol             As Integer
    
Public Sub DataImportSheetDataMapping()

    Set vcWb = ThisWorkbook
    Set diWs = vcWb.Sheets("Data Import")
    
    With diWs
        diDateCol = .Rows("1:1").Find(what:="date", LookAt:=xlWhole).Column
        diOpenCol = .Rows("1:1").Find(what:="open", LookAt:=xlWhole).Column
        diCloseCol = .Rows("1:1").Find(what:="close", LookAt:=xlWhole).Column
        diHighCol = .Rows("1:1").Find(what:="high", LookAt:=xlWhole).Column
        diLowCol = .Rows("1:1").Find(what:="low", LookAt:=xlWhole).Column
    End With

End Sub

Public Sub CalculationResultsSheetDataMapping()

    Set vcWb = ThisWorkbook
    Set crWs = vcWb.Sheets("Calculation Results")
    
    With crWs
        crAnnualizationFactorCol = .Rows("7:7").Find(what:="Specify Annualization Factor Below", LookAt:=xlWhole).Column
        crCalcResCloseToCloseCol = .Rows("3:3").Find(what:="Close to Close", LookAt:=xlWhole).Column
        crCalcResGarmanKlassCol = .Rows("3:3").Find(what:="Garman - Klass", LookAt:=xlWhole).Column
        crCalcResRogersSatcellCol = .Rows("3:3").Find(what:="Rogers - Satchell", LookAt:=xlWhole).Column
        crCalcResGarmanKlassYangZhangCol = .Rows("3:3").Find(what:="Garman - Klass Yang - Zhang", LookAt:=xlWhole).Column
        crCalcResYangZhangCol = .Rows("3:3").Find(what:="Yang - Zhang", LookAt:=xlWhole).Column
    End With

End Sub
