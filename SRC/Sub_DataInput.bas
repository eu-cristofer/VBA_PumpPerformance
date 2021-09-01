Attribute VB_Name = "Sub_DataInput"
Option Explicit
'
' This sub reads the input data
'
Sub InputData( _
    PumpObject As Pump, _
    oTD As Worksheet, _
    strTD As String _
    )
    '
    ' Add Rated Point data
    '
    ' Parameters may be supplied at the creation of the Rated Point
    ' Required parameters: Head; and Flow.
    PumpObject.AddRatedPoint _
        Head:=oTD.Range("'" & strTD & "'!RatedPointHead"), _
        Flow:=oTD.Range("'" & strTD & "'!RatedPointQ"), _
        NSpeed:=oTD.Range("'" & strTD & "'!RatedPointN"), _
        DriverPower:=oTD.Range("'" & strTD & "'!RatedPointDriverPower"), _
        Efficiency:=oTD.Range("'" & strTD & "'!RatedPointEfficiency")
    ' Optional parameters can be added later
    PumpObject.RatedPoint.DinVisc = oTD.Range("'" & strTD & "'!RatedPointDinVisc")
    PumpObject.RatedPoint.Density = oTD.Range("'" & strTD & "'!RatedPointDensity")
    '
    ' Add supplier BEP, if supplied
    '
    PumpObject.SupplierBEP = oTD.Range("'" & strTD & "'!SupplierBEP")
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
