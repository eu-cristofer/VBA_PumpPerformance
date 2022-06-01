VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPointCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Member variables:
'
Private Type typTestPointCor
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
    dEff As Double
    dCQ As Double
    dCEff As Double
    dCH As Double
End Type
Private ThisCorrectedPoint As typTestPointCor
'
' Events, triggered when class created:
'
Private Sub Class_Initialize()
    ' Optional variable initialization
    Me.Rho = 995     ' kg/m^3
    Me.CQ = 1
    Me.CEff = 1
    Me.CH = 1
End Sub
'
' Properties
'
Public Property Get Name() As String
' Name - Read and write
    Name = ThisCorrectedPoint.strName
End Property

Public Property Let Name(ByVal Value As String)
    ThisCorrectedPoint.strName = Value
End Property


Public Property Get Q() As Double
' Q - Read and write
    Q = ThisCorrectedPoint.dQ      ' m^3/h
End Property

Public Property Let Q(ByVal Flow As Double)
    ThisCorrectedPoint.dQ = Flow      ' m^3/h
End Property


Public Property Get P0() As Double
' P0 - Read and write
    P0 = ThisCorrectedPoint.dP0     ' kgf/cm^2
End Property

Public Property Let P0(ByVal InletPressure As Double)
    ThisCorrectedPoint.dP0 = InletPressure    ' kgf/cm^2
End Property


Public Property Get P3() As Double
' P3 - Read and write
    P3 = ThisCorrectedPoint.dP3    ' kgf/cm^2
End Property

Public Property Let P3(ByVal OutletPressure As Double)
    ThisCorrectedPoint.dP3 = OutletPressure    ' kgf/cm^2
End Property


Public Property Get DriverPower() As Double
 ' DriverPower- Read and write
    DriverPower = ThisCorrectedPoint.dDriverPower     ' kW
End Property

Public Property Let DriverPower(ByVal Power As Double)
    ThisCorrectedPoint.dDriverPower = Power     ' kW
End Property


Public Property Get Head() As Double
' Head - Read and write
    Head = ThisCorrectedPoint.dHead     ' m
End Property

Public Property Let Head(ByVal Value As Double)
    ThisCorrectedPoint.dHead = Value      ' m
End Property


Public Property Get NPSH3() As Double
' NPSH3 - Read and write
    NPSH3 = ThisCorrectedPoint.dNPSH3     ' m
End Property

Public Property Let NPSH3(ByVal Value As Double)
    ThisCorrectedPoint.dNPSH3 = Value    ' m
End Property


Public Property Get Density() As Double
' Density - Read and write
    Density = ThisCorrectedPoint.dRho / 995
End Property

Public Property Let Density(ByVal Value As Double)
    ThisCorrectedPoint.dRho = Value * 995
End Property


Public Property Get Rho() As Double
' Rho - Read and write
    Rho = ThisCorrectedPoint.dRho
End Property

Public Property Let Rho(ByVal Value As Double)
    ThisCorrectedPoint.dRho = Value
End Property


Public Property Get Efficiency() As Double
' Efficiency - Read and write
    Efficiency = ThisCorrectedPoint.dEff
End Property

Public Property Let Efficiency(ByVal Value As Double)
    ThisCorrectedPoint.dEff = Value
End Property


Public Property Get NSpeed() As Double
' N - Read and write
    NSpeed = ThisCorrectedPoint.dN
End Property

Public Property Let NSpeed(ByVal RotateSpeed As Double)
    ThisCorrectedPoint.dN = RotateSpeed
End Property


Public Property Get Temp() As Double
' Temp - Read and write
    Temp = ThisCorrectedPoint.dTemp
End Property

Public Property Let Temp(ByVal Temperature As Double)
    ThisCorrectedPoint.dTemp = Temperature
End Property


Public Property Get CQ() As Double
' CQ - Read and write
    CQ = ThisCorrectedPoint.dCQ
End Property

Public Property Let CQ(ByVal Value As Double)
    ThisCorrectedPoint.dCQ = Value
End Property


Public Property Get CEff() As Double
' CEff - Read and write
    CEff = ThisCorrectedPoint.dCEff
End Property

Public Property Let CEff(ByVal Value As Double)
    ThisCorrectedPoint.dCEff = Value
End Property


Public Property Get CH() As Double
' CH - Read and write
    CH = ThisCorrectedPoint.dCH
End Property

Public Property Let CH(ByVal Value As Double)
    ThisCorrectedPoint.dCH = Value
End Property
