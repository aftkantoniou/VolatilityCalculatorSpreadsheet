VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

    Private vcWb    As Workbook
    Private crWs    As Worksheet
    
Private Sub Workbook_Open()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ' Initializing global format config for all the sheets in the workbook
    ' ====================================================================
    
    Debug.Print "Passed from Workbook Open: Start"
    
    Set vcWb = ThisWorkbook
    Set crWs = vcWb.Sheets("Calculation Results")
    
    With crWs
         Debug.Print "Passed from crWs: Start"
         .Unprotect Password:="aft123"
         .Cells.Locked = True
         .Range("B4:F4").Locked = False
         .Range("B8:F8").Locked = False
         .Range("J4:L4").Locked = False
         .Range("J4:L4").Value = ""
         .Protect Password:="aft123", UserInterFaceOnly:=True
         Debug.Print "Passed from crWs: End"
    End With
    
    With Application
         .EnableEvents = True
         .ScreenUpdating = True
    End With
        
    Debug.Print "Passed from Workbook Open: End"
    
End Sub
