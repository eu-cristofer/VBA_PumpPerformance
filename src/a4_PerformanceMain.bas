Attribute VB_Name = "a4_PerformanceMain"
Option Explicit
'
' Main script sample
' This script runs the process of performance computation
'
Sub PumpPerformanceMain()

    ' Create the Sheet names string
    Dim strTD As String
    strTD = Application.ActiveSheet.Name
    
    ' Create Sheet objects call precisely the ranges
    Dim oTD As Worksheet
    Set oTD = ActiveWorkbook.Worksheets(strTD)
    
    ' Create a Pump instancy
    Dim NewPump As New Pump
    
    '
    ' Add pump Design Data
    '
    NewPump.D0 = _
        25.4 / 1000 * Replace(oTD.Range("'" & strTD & "'!PumpD0"), "''", "")      'm
    
    NewPump.D3 = _
        25.4 / 1000 * Replace(oTD.Range("'" & strTD & "'!PumpD3"), "''", "")      'm
    
    NewPump.TAG = _
        oTD.Range("'" & strTD & "'!PumpTAG")
    
    '
    ' Add test apparatus data
    '
    NewPump.Z0 = _
        oTD.Range("'" & strTD & "'!ApparatusZ0")      'm
    
    NewPump.Z3 = _
        oTD.Range("'" & strTD & "'!ApparatusZ3")      'm
    
    NewPump.ZM0 = _
        oTD.Range("'" & strTD & "'!ApparatusZM0")      'm
    
    NewPump.ZM3 = _
        oTD.Range("'" & strTD & "'!ApparatusZM3")      'm
    
    '
    ' Read Teste Input Data
    '
    Call InputData(NewPump, oTD, strTD)
    
    ' Calc test points
        ' Update test points calculation
        NewPump.Update
        
        ' printout the number of collected points
        Debug.Print "Test points: " & NewPump.TestPoints.Count; ""
        
        ' Executando a correção de velocidade densidade
        NewPump.SpeedCorrection
    
        ' Teste do procedimento: executar duas vezes consecutivas a operação
        NewPump.SpeedCorrection
        
        NewPump.ViscosityCorrection
     
    ' Print Ouput Data
    Call OutputData(NewPump, oTD, strTD)

End Sub
