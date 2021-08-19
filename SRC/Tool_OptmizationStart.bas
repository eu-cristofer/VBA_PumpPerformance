Attribute VB_Name = "Tool_OptmizationStart"
Option Explicit
'
' This sub start optimization measures to improve
' computational efficiency in the code
'
Sub OptimizationStart()

    Application.ScreenUpdating = False
    
    Application.Calculation = xlCalculationManual
    
    Debug.Print "OptimizationStart Module:" & Chr(10) & "     Read!"
    
End Sub
