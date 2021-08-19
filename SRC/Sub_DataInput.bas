Attribute VB_Name = "Sub_DataInput"
Option Explicit
'
' This sub reads the input data
'
Sub InputData( _
    PumpObject As Pump, _
    oDS As Worksheet, _
    oTD As Worksheet, _
    strDS As String, _
    strTD As String _
    )
    '
    ' Add Rated Point data
    '
    ' Parameters may be supplied at the creation of the Rated Point
    ' Required parameters: Head; and Flow.
    PumpObject.AddRatedPoint _
        Head:=oDS.Range("'" & strDS & "'!RatedPointHead"), _
        Flow:=oDS.Range("'" & strDS & "'!RatedPointQ"), _
        InletPressure:=oDS.Range("'" & strDS & "'!RatedPointP0"), _
        OutletPressure:=oDS.Range("'" & strDS & "'!RatedPointP3"), _
        NSpeed:=oDS.Range("'" & strDS & "'!RatedPointN"), _
        DriverPower:=oDS.Range("'" & strDS & "'!RatedPointDriverPower"), _
        Efficiency:=oDS.Range("'" & strDS & "'!RatedPointEfficiency")
    ' Optional parameters can be added later
    PumpObject.RatedPoint.DinVisc = oDS.Range("'" & strDS & "'!RatedPointDinVisc")
    PumpObject.RatedPoint.Density = oDS.Range("'" & strDS & "'!RatedPointDensity")
    '
    ' Add supplier BEP, if supplied
    '
    PumpObject.SupplierBEP = 0
    '
    ' Add test points
    '
    ' Individually input:
    ' PumpObject.AddTestPoint _
        ' Flow:=23, _
        ' InletPressure:=45, _
        ' OutletPressure:=90
        ' Name:="MCSF"
    '
    ' Multi point data read:
    PumpObject.AddMultiTestPoint _
        FlowRange:=oTD.Range("'" & strTD & "'!TestPointQ"), _
        InletPressureRange:=oTD.Range("'" & strTD & "'!TestPointP0"), _
        OutletPressureRange:=oTD.Range("'" & strTD & "'!TestPointP3")
    
    PumpObject.AddMultiPointDriverPower _
        DriverPowerRange:=oTD.Range("'" & strTD & "'!TestPointDriverPower")
    
    PumpObject.AddMultiPointNSpeed _
        NRange:=oTD.Range("'" & strTD & "'!TestPointNSpeed")
    
    PumpObject.AddMultiPointTemp _
        Temp:=oTD.Range("'" & strTD & "'!TestPointTemp")
        
    PumpObject.AddMultiPointNPSH3 _
        NPSH3_Range:=oTD.Range("'" & strTD & "'!TestPointNPSH3")
    
End Sub
