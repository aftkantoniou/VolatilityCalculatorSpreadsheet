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
    
    Public vcWb                         As Workbook
    Public diWs                         As Worksheet
    Public crWs                         As Worksheet
    
    Public diDateCol                    As Integer
    Public diOpenCol                    As Integer
    Public diCloseCol                   As Integer
    Public diHighCol                    As Integer
    Public diLowCol                     As Integer
    
    Public crCalcMethCol                As Integer
    Public crCalcResCol                 As Integer
    Public crAnnualizationFactorCol     As Integer
    
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
        crCalcMethCol = .Rows("3:3").Find(what:="Select Volatility Calculation Method Below", LookAt:=xlWhole).Column
        crAnnualizationFactorCol = .Rows("7:7").Find(what:="Specify Annualization Factor Below", LookAt:=xlWhole).Column
        crCalcResCol = .Rows("3:3").Find(what:="Calculation Result", LookAt:=xlWhole).Column
    End With

End Sub
