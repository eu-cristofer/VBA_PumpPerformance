Attribute VB_Name = "a2_ClonePump"
Option Explicit

Sub ClonePump(control As IRibbonControl)

    ' Start of optimization actions
    Call OptimizationStart
    
    Dim oNewSheet As Object
    
    ' Clone
    ActiveSheet.Copy _
        After:=ActiveSheet
        
    Call ResetMain(Worksheets(ActiveSheet.Index))
    
    ' End of optimization actions
    Call OptimizationEnd
    
End Sub
