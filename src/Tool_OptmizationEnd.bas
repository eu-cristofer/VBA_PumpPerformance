Attribute VB_Name = "Tool_OptmizationEnd"
Option Explicit
'
' This sub ends optimization measures to improve
' computational efficiency in the code
'
Sub OptimizationEnd()
    
    Application.ScreenUpdating = True
    
    Application.Calculation = xlCalculationAutomatic
    
    ActiveSheet.Calculate
    
    Debug.Print "OptimizationEnd Module:" & Chr(10) & "      Read!"
    
End Sub
