Attribute VB_Name = "a1_NewPump"
Option Explicit

Sub NewPump(control As IRibbonControl)

    ' Start of optimization actions
    Call OptimizationStart
    
    ' Create the Sheet names string
    Const strTD As String = "InputTestData"
    Const strGS As String = "GettingStarted"
    
    ' Create the Add-in workbook object
    Dim oAddWB As Object
    Set oAddWB = Application.Workbooks("vba-pump-performance.xlam")
    
    'Copy the Getting Started and Input data sheet to the active workbook
    With ActiveWorkbook
    
        ' Check if there is getting started tab
        Dim wsWS As Worksheet
        On Error Resume Next
        Set wsWS = .Worksheets(strGS)
        On Error GoTo 0
        
        ' Add getting started
        If wsWS Is Nothing Then
            oAddWB.Worksheets(strGS).Copy _
                Before:=.Worksheets(1)
        End If
        
        ' Add new input data sheet
        oAddWB.Worksheets(strTD).Copy _
            After:=.Worksheets(1)
    
    End With
    
    ' End of optimization actions
    Call OptimizationEnd

End Sub
