Attribute VB_Name = "Sub_DataOutput"
Option Explicit
'
' This sub reads the input data
'
Sub OutputData( _
    PumpObject As Pump, _
    oTD As Worksheet, _
    strTD As String _
    )
    '
    ' Test raw data output
    '
    oTD.Range("'" & strTD & "'!PumpD0m") = PumpObject.D0 & " m"
    oTD.Range("'" & strTD & "'!PumpD3m") = PumpObject.D3 & " m"
    If oTD.Range("'" & strTD & "'!RatedPointDinVisc") <> 0 Then
        oTD.Range("'" & strTD & "'!AproxBEP") = PumpObject.BEP
    Else
        oTD.Range("'" & strTD & "'!AproxBEP") = "-"
    End If
    
    ' Teste points head
    PumpObject.PrintMultiPointVar _
        VarIndex:=1, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointHead")
    
    ' Test points efficiency
    PumpObject.PrintMultiPointVar _
        VarIndex:=2, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointEfficiency")
    '
    ' Test correct data
    '
    ' Teste rated head
    oTD.Range("'" & strTD & "'!RatedPointHeadPoly") = PumpObject.PolyAprox(1)
    oTD.Range("'" & strTD & "'!RatedPointHeadSpline") = PumpObject.SplineAprox(1)
    
     ' Teste rated driver power
    oTD.Range("'" & strTD & "'!RatedPointDriverPowerPoly") = PumpObject.PolyAprox(2)
    oTD.Range("'" & strTD & "'!RatedPointDriverPowerSpline") = PumpObject.SplineAprox(2)
    
    ' Teste rated driver power
    oTD.Range("'" & strTD & "'!RatedPointEfficiencyPoly") = PumpObject.PolyAprox(3)
    oTD.Range("'" & strTD & "'!RatedPointEfficiencySpline") = PumpObject.SplineAprox(3)
    
    ' Test points flow
    PumpObject.PrintMultiPointVar _
        VarIndex:=3, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorQ")
        
    ' Test points head
    PumpObject.PrintMultiPointVar _
        VarIndex:=4, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorHead")
        
    ' Test points driver power
    PumpObject.PrintMultiPointVar _
        VarIndex:=5, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorDriverPower")
        
     ' Test points speed
    PumpObject.PrintMultiPointVar _
        VarIndex:=6, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorNSpeed")
        
     ' Test points efficiency
    PumpObject.PrintMultiPointVar _
        VarIndex:=7, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorEfficiency")
        
     ' Test points NPSH3
    PumpObject.PrintMultiPointVar _
        VarIndex:=8, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorNPSH3")
        
    ' Test points CQ
    PumpObject.PrintMultiPointVar _
        VarIndex:=9, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorCQ")
        
     ' Test points CH
    PumpObject.PrintMultiPointVar _
        VarIndex:=10, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorCH")
        
     ' Test points CEff
    PumpObject.PrintMultiPointVar _
        VarIndex:=11, _
        RangeToPrint:=oTD.Range("'" & strTD & "'!TestPointCorCEff")
    
    ' Print the performance chart
    PumpObject.PlotPerformance _
        LeftCorner:=oTD.Range("'" & strTD & "'!ChartLefCorner"), _
        RightMid:=oTD.Range("'" & strTD & "'!ChartRightMid"), _
        RightCorner:=oTD.Range("'" & strTD & "'!ChartRightCorner")
        
    ' Print the efficiency chart
    PumpObject.PlotEfficiency _
        LeftCorner:=oTD.Range("'" & strTD & "'!ChartLefCorner"), _
        RightMid:=oTD.Range("'" & strTD & "'!ChartRightMid"), _
        RightCorner:=oTD.Range("'" & strTD & "'!ChartRightCorner")
    
End Sub
