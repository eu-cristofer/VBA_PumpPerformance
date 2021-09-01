Attribute VB_Name = "a3_RunPerformance"
Option Explicit

Sub RunPerformance(control As IRibbonControl)

    ' Start of optimization actions
    Call OptimizationStart
    
    ' Create the NamedRanges for read the input data
    Call NameRangesCreator

    Call PumpPerformanceMain

    ' End of optimization actions
    Call OptimizationEnd

End Sub
