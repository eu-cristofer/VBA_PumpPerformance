Attribute VB_Name = "Sub_DataOutput"
Option Explicit
'
' This sub reads the input data
'
Sub OutputData( _
    PumpObject As Pump, _
    oDS As Worksheet, _
    oTD As Worksheet, _
    strDS As String, _
    strTD As String _
    )
    '
    ' Test raw data output
    '
    oTD.Range("'" & strTD & "'!PumpD0") = PumpObject.D0
    oTD.Range("'" & strTD & "'!PumpD3") = PumpObject.D3
    
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
