Attribute VB_Name = "a5_ResetCalcFields"
Option Explicit

Sub ResetCalcFields(control As IRibbonControl)
'
' Sub to clean all the calculated data into the InputTestData
'
    ' Start of optimization actions
    Call OptimizationStart
    
    Call ResetMain(ActiveSheet)
    
    ' End of optimization actions
    Call OptimizationEnd

End Sub
