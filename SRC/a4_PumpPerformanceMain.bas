Attribute VB_Name = "a4_PumpPerformanceMain"
Option Explicit
'
' Main script sample
' This script runs the process of performance computation
'
Sub PumpPerformanceMain()
    
    Debug.Print _
    "=================" & Chr(10) _
    ; "VBA PUMP PERFORMANCE"
  
    ' Start of optimization actions
    Call OptimizationStart
    
    ' Create the NamedRanges for read the input data
    Call NameRangesCreator
    
    ' Create the Sheet names string
    Dim strDS As String
    Dim strTD As String
    
    strDS = "InputDataSheet"
    strTD = "InputTestData"
    
    ' Create Sheet objects call precisely the ranges
    Dim oDS As Worksheet
    Dim oTD As Worksheet
    
    Set oDS = ThisWorkbook.Worksheets(strDS)
    Set oTD = ThisWorkbook.Worksheets(strTD)
    
    ' Create a Pump instancy
    Dim NewPump As Pump
    
    Set NewPump = New Pump
    
    ' Add pump Design Data
    NewPump.D0 = 25.4 / 1000 * Replace(oDS.Range("'" & strDS & "'!PumpD0"), "''", "")      'm
    NewPump.D3 = 25.4 / 1000 * Replace(oDS.Range("'" & strDS & "'!PumpD3"), "''", "")      'm
    NewPump.TAG = oDS.Range("'" & strDS & "'!PumpTAG")
    
    ' Add pump Design Data
    NewPump.Z0 = oTD.Range("'" & strTD & "'!ApparatusZ0")      'm
    NewPump.Z3 = oTD.Range("'" & strTD & "'!ApparatusZ3")      'm
    NewPump.ZM0 = oTD.Range("'" & strTD & "'!ApparatusZM0")      'm
    NewPump.ZM3 = oTD.Range("'" & strTD & "'!ApparatusZM3")      'm
    
    ' Read Teste Input Data
    Call InputData(NewPump, oDS, oTD, strDS, strTD)
    
    ' Calc test points
        ' Update test points calculation
        NewPump.Update
        
        ' imprimindo o número de pontos coletados
        Debug.Print "Test points: " & NewPump.TestPoints.Count; ""
        
        ' Executando a correção de velocidade densidade
        NewPump.SpeedCorrection
    
        ' Teste do procedimento: executar duas vezes consecutivas a operação
        NewPump.SpeedCorrection
        
        NewPump.ViscosityCorrection
     
    ' Print Ouput Data
    Call OutputData(NewPump, oDS, oTD, strDS, strTD)
    
    NewPump.PolyCurveFit _
        Collection:="TestPoints", _
        Propertie:="Head"
    
    ' End of optimization actions
    Call OptimizationEnd
    
    Debug.Print "=================="
End Sub
