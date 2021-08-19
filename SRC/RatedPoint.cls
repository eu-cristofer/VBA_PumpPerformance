VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RatedPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Member variables:
'
Private Type typRatedPoint
    strName As String
    dDriverPower As Double
    dHead As Double
    dN As Double
    dP0 As Double
    dP3 As Double
    dQ As Double
    dRho As Double
    dTemp As Double
    dEff As Double
    dDinVisc As Double
End Type
Private ThisRatedPoint As typRatedPoint
'
' Events, triggered when class created:
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Private Sub Class_Initialize()
    ThisRatedPoint.strName = "Design Rated Point"
    
    ' Optional variable initialization
    ThisRatedPoint.dRho = 995     ' kg/m^3
End Sub
'
' Properties
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
Public Property Get Name() As String
' Name - Read only
    Name = ThisRatedPoint.strName
End Property

Public Property Get Q() As Double
' Q - Read and write
    Q = ThisRatedPoint.dQ      ' m^3/h
End Property
Public Property Let Q(ByVal Flow As Double)
    ThisRatedPoint.dQ = Flow      ' m^3/h
End Property

Public Property Get P0() As Double
' P0 - Read and write
    P0 = ThisRatedPoint.dP0     ' kgf/cm^2
End Property
Public Property Let P0(ByVal InletPressure As Double)
    ThisRatedPoint.dP0 = InletPressure    ' kgf/cm^2
End Property

Public Property Get P3() As Double
' P3 - Read and write
    P3 = ThisRatedPoint.dP3    ' kgf/cm^2
End Property
Public Property Let P3(ByVal OutletPressure As Double)
    ThisRatedPoint.dP3 = OutletPressure    ' kgf/cm^2
End Property

Public Property Get DriverPower() As Double
 ' DriverPower- Read and write
    DriverPower = ThisRatedPoint.dDriverPower     ' kW
End Property
Public Property Let DriverPower(ByVal Power As Double)
    ThisRatedPoint.dDriverPower = Power     ' kW
End Property

Public Property Get Head() As Double
' Head - Read and write
    Head = ThisRatedPoint.dHead     ' m
End Property
Public Property Let Head(ByVal Value As Double)
    ThisRatedPoint.dHead = Value      ' m
End Property

Public Property Get Density() As Double
' Density - Read and write
    Density = ThisRatedPoint.dRho / 995
End Property
Public Property Let Density(ByVal Value As Double)
    ThisRatedPoint.dRho = Value * 995
End Property

Public Property Get Rho() As Double
' Rho - Read and write
    Rho = ThisRatedPoint.dRho
End Property
Public Property Let Rho(ByVal Value As Double)
    ThisRatedPoint.dRho = Value
End Property

Public Property Get Efficiency() As Double
' Efficiency - Read and write
    Efficiency = ThisRatedPoint.dEff
End Property
Public Property Let Efficiency(ByVal Value As Double)
    ThisRatedPoint.dEff = Value
End Property

Public Property Get NSpeed() As Double
' N - Read and write
    NSpeed = ThisRatedPoint.dN
End Property
Public Property Let NSpeed(ByVal RotateSpeed As Double)
    ThisRatedPoint.dN = RotateSpeed
End Property

Public Property Get Temp() As Double
' Temp - Read and write
    Temp = ThisRatedPoint.dTemp
End Property
Public Property Let Temp(ByVal Temperature As Double)
    ThisRatedPoint.dTemp = Temperature
End Property

Public Property Get DinVisc() As Double
' DinVisc - Read and write
    DinVisc = ThisRatedPoint.dDinVisc
End Property
Public Property Let DinVisc(ByVal Value As Double)
    ThisRatedPoint.dDinVisc = Value
End Property
