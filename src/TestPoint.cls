VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Member variables:
'
Private Type typTestPoint
    strName As String
    dDriverPower As Double
    dHead As Double
    dNPSH3 As Double
    dN As Double
    dP0 As Double
    dP3 As Double
    dQ As Double
    dRho As Double
    dTemp As Double
    dU0 As Double
    dU3 As Double
    dEff As Double
End Type
Private ThisTestPoint As typTestPoint
Private Const dg As Double = 9.81       'm/s^2
'
' Events, triggered when class created:
'
Private Sub Class_Initialize()
    
    ' Optional fullfilment variable initialization
    ThisTestPoint.dRho = 995     ' kg/m^3

End Sub
'
' Properties
'
Public Property Get Name() As String
' Name - Read and write
    Name = ThisTestPoint.strName
End Property
Public Property Let Name(ByVal PointName As String)
    ThisTestPoint.strName = PointName
End Property

Public Property Get Q() As Double
' Q - Read and write
    Q = ThisTestPoint.dQ      ' m^3/h
End Property
Public Property Let Q(ByVal Flow As Double)
    ThisTestPoint.dQ = Flow      ' m^3/h
End Property

Public Property Get P0() As Double
' P0 - Read and write
    P0 = ThisTestPoint.dP0     ' kgf/cm^2
End Property
Public Property Let P0(ByVal InletPressure As Double)
    ThisTestPoint.dP0 = InletPressure    ' kgf/cm^2
End Property

Public Property Get P3() As Double
' P3 - Read and write
    P3 = ThisTestPoint.dP3    ' kgf/cm^2
End Property
Public Property Let P3(ByVal OutletPressure As Double)
    ThisTestPoint.dP3 = OutletPressure    ' kgf/cm^2
End Property

Public Property Get DriverPower() As Double
 ' DriverPower- Read and write
    DriverPower = ThisTestPoint.dDriverPower     ' kW
End Property
Public Property Let DriverPower(ByVal Power As Double)
    ThisTestPoint.dDriverPower = Power     ' kW
End Property

Public Property Get Head() As Double
' Head - Read and write
    Head = ThisTestPoint.dHead     ' m
End Property
Public Property Let Head(ByVal Value As Double)
    ThisTestPoint.dHead = Value    ' m
End Property

Public Property Get NPSH3() As Double
' NPSH3 - Read and write
    NPSH3 = ThisTestPoint.dNPSH3     ' m
End Property
Public Property Let NPSH3(ByVal Value As Double)
    ThisTestPoint.dNPSH3 = Value    ' m
End Property

Public Property Get Density() As Double
' Density - Read and write
    Density = ThisTestPoint.dRho / 995
End Property
Public Property Let Density(ByVal Value As Double)
    ThisTestPoint.dRho = Value * 995
End Property

Public Property Get Rho() As Double
' Rho - Read and write
    Rho = ThisTestPoint.dRho
End Property
Public Property Let Rho(ByVal Value As Double)
    ThisTestPoint.dRho = Value
End Property

Public Property Get U0() As Double
' U0 - Read and write
    U0 = ThisTestPoint.dU0
End Property
Public Property Let U0(ByVal Value As Double)
    ThisTestPoint.dU0 = Value
End Property

Public Property Get U3() As Double
' U3 - Read and write
    U3 = ThisTestPoint.dU3
End Property
Public Property Let U3(ByVal Value As Double)
    ThisTestPoint.dU3 = Value
End Property

Public Property Get Efficiency() As Double
' Efficiency - Read only
    Me.RunCalc
    Efficiency = ThisTestPoint.dEff
End Property
Public Property Let Efficiency(ByVal Value As Double)
    ThisTestPoint.dEff = Value
End Property

Public Property Get NSpeed() As Double
' N - Read and write
    NSpeed = ThisTestPoint.dN
End Property
Public Property Let NSpeed(ByVal RotateSpeed As Double)
    ThisTestPoint.dN = RotateSpeed
End Property

Public Property Get Temp() As Double
' Temp - Read and write
    Temp = ThisTestPoint.dTemp
End Property
Public Property Let Temp(ByVal Temperature As Double)
    ThisTestPoint.dTemp = Temperature
End Property
'
' Methods
'
Public Sub Update( _
    ByRef A0 As Double, _
    ByRef A3 As Double, _
    ByRef Z0 As Double, _
    ByRef Z3 As Double, _
    ByRef ZM0 As Double, _
    ByRef ZM3 As Double _
    )
    '
    ' Function to update main data
    '
    With Me
        .U0 = .Q / 3600 / A0
        .U3 = .Q / 3600 / A3
        .Head = Z3 - Z0 _
            + (.P3 - .P0) / (.Rho * dg) * 98000 _
            + ZM3 - ZM0 _
            + (.U3 ^ 2 - .U0 ^ 2) / (2 * dg)
        
        .RunCalc
    End With

End Sub

Public Sub RunCalc()
    '
    ' Run calc
    '
    With Me
        If .DriverPower <> 0 Then
            .Efficiency = 100 * .Rho * dg * .Q * .Head / (.DriverPower * 1000 * 3600)
        Else
            .Efficiency = 0
        End If
    End With
End Sub
