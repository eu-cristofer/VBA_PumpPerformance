Attribute VB_Name = "a7_Report"
Option Explicit


Sub Report(control As IRibbonControl)

    Dim strSheetsToReport(0) As String
    
    strSheetsToReport(0) = ActiveWorkbook.ActiveSheet.Name
    
    Call WordReport(strSheetsToReport)

End Sub

