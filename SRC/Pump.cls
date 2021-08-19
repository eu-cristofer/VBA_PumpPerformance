VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Member variables:
'
Private Type typPump
    strTAG As String
    dD0 As Double
    dD3 As Double
    dA0 As Double
    dA3 As Double
    dZ0 As Double
    dZ3 As Double
    dZM0 As Double
    dZM3 As Double
    dSupBEP As Double
    rpRatedPoint As RatedPoint
    colTestPoints As Collection
    colTestPointsCor As Collection
End Type
Private ThisPump As typPump
Private Const dg As Double = 9.81 'm/s^2
'
' Events, triggered when class created:
'
Private Sub Class_Initialize()
    Debug.Print _
        "* * * * * * * * * * * * * * * * * * " & Chr(10) _
        ; "VBA Pump Performance - Pump.cls" & Chr(10) _
        & "Vesion 0.0.0" & Chr(10) _
        ; "* * * * * * * * * * * * * * * * * * "
    Set ThisPump.colTestPoints = New Collection
    Set ThisPump.colTestPointsCor = New Collection
End Sub
'
' Properties:
'
Public Property Let TAG(ByVal TAG_name As String)
' TAG - Read and write
    ThisPump.strTAG = TAG_name
End Property
Public Property Get TAG() As String
    TAG = ThisPump.strTAG
End Property

Public Property Let D0(ByVal InternalDiameter As Double)
 ' D0 - Read and write
    ThisPump.dD0 = InternalDiameter     ' m
    ThisPump.dA0 = WorksheetFunction.pi * ThisPump.dD0 ^ 2 / 4       'm^2
End Property
Public Property Get D0() As Double
    D0 = ThisPump.dD0       ' m
End Property
Public Property Get A0() As Double
    A0 = ThisPump.dA0       'm^2
End Property

Public Property Let D3(ByVal InternalDiameter As Double)
' D3  - Read and write
    ThisPump.dD3 = InternalDiameter      'm
    ThisPump.dA3 = WorksheetFunction.pi * ThisPump.dD3 ^ 2 / 4       'm^2
End Property
Public Property Get D3() As Double
    D3 = ThisPump.dD3       ' m
End Property
Public Property Get A3() As Double
    A3 = ThisPump.dA3       'm^2
End Property

Public Property Let Z0(ByVal InletTapElevation As Double)
' Z0  - Read and write
    ThisPump.dZ0 = InletTapElevation        ' m
End Property
Public Property Get Z0() As Double
    Z0 = ThisPump.dZ0       ' m
End Property

Public Property Let Z3(ByVal OutletTapElevation As Double)
' Z3  - Read and write
    ThisPump.dZ3 = OutletTapElevation        ' m
End Property
Public Property Get Z3() As Double
    Z3 = ThisPump.dZ3       ' m
End Property

Public Property Let ZM0(ByVal InletAboveTapElevation As Double)
' ZM0  - Read and write
    ThisPump.dZM0 = InletAboveTapElevation        ' m
End Property
Public Property Get ZM0() As Double
    ZM0 = ThisPump.dZM0       ' m
End Property

Public Property Let ZM3(ByVal OutletAboveTapElevation As Double)
' ZM3 - Read and write
    ThisPump.dZM3 = OutletAboveTapElevation        ' m
End Property
Public Property Get ZM3() As Double
    ZM3 = ThisPump.dZM3       ' m
End Property

Public Property Let SupplierBEP(ByVal Value As Double)
' ZM3 - Read and write
    ThisPump.dSupBEP = Value        ' m^3/h
End Property
Public Property Get SupplierBEP() As Double
    ZM3 = ThisPump.dSupBEP       ' m^3/h
End Property

Property Get TestPoints() As Collection
' Test points collection
    Set TestPoints = ThisPump.colTestPoints
End Property

Property Get TestPointsCor() As Collection
' Corrected test points collection
    Set TestPointsCor = ThisPump.colTestPointsCor
End Property

Property Get RatedPoint() As RatedPoint
' Index of test points collection
    Set RatedPoint = ThisPump.rpRatedPoint
End Property
'
' Methods:
'
Public Sub AddRatedPoint( _
    ByVal Head As Double, _
    ByVal Flow As Double, _
    Optional ByVal InletPressure As Double, _
    Optional ByVal OutletPressure As Double, _
    Optional ByVal Rho As Double, _
    Optional ByVal Density As Double, _
    Optional ByVal DinVisc As Double, _
    Optional ByVal NSpeed As Double, _
    Optional ByVal DriverPower As Double, _
    Optional ByVal Efficiency As Double _
    )
'
' Create the Rated point
'
    ' Instantiating the rated point
    Set ThisPump.rpRatedPoint = New RatedPoint
    
    With Me.RatedPoint
        ' Loading required variables
        .Head = Head
        .Q = Flow
        .P0 = InletPressure
        .P3 = OutletPressure
        .NSpeed = NSpeed
        If Rho <> 0 Then .Rho = Rho
        If Density <> 0 Then .Density = Density
        If DinVisc <> 0 Then .DinVisc = DinVisc
        .DriverPower = DriverPower
        .Efficiency = Efficiency
    End With
End Sub

Public Sub AddTestPoint( _
    ByVal Flow As Double, _
    ByVal InletPressure As Double, _
    ByVal OutletPressure As Double, _
    Optional ByVal Name As String _
    )
'
' Add a new test point
'
    Dim tpTP As TestPoint
    
    ' Instantiating new test point
    Set tpTP = New TestPoint
    
    With tpTP
        ' Loading required variables
        .Q = Flow
        .P0 = InletPressure
        .P3 = OutletPressure
        
        ' Loading optional variable
        If Name = "" Then
            .Name = "TestPoint" & Me.TestPoints.Count + 1
        Else
            .Name = Name
        End If
    
    End With
    
    Me.TestPoints.Add _
        Item:=tpTP, _
        Key:=tpTP.Name

End Sub

Public Sub AddMultiTestPoint( _
    ByRef FlowRange As Range, _
    ByRef InletPressureRange As Range, _
    ByRef OutletPressureRange As Range, _
    Optional ByRef NameRange As Range _
    )
'
' Add a block of new points to test points collection
'
    ' Validation of the ranges
    Dim blC1 As Boolean
    Dim blC2 As Boolean
    Dim blC3 As Boolean
    
    blC1 = FlowRange.Count <> InletPressureRange.Count
    blC2 = FlowRange.Count <> OutletPressureRange.Count
  
    If NameRange Is Nothing Then
        blC3 = False
    Else
        blC3 = FlowRange.Count <> NameRange.Count
    End If
    
    If blC1 Or blC2 Or blC3 Then
        ' Error message
        MsgBox "The block of input data (Test point name, Q, P0, P3) shall have the same dimensions."
        Call ErrorExit
    End If
    
    ' Iterative block
    Dim n As Integer
    Dim i As Integer
    
    n = FlowRange.Count
    
    For i = 1 To n
        If NameRange Is Nothing Then
            Me.AddTestPoint _
                Flow:=FlowRange(i).Value, _
                InletPressure:=InletPressureRange(i).Value, _
                OutletPressure:=OutletPressureRange(i).Value
        Else
            Me.AddTestPoint _
                Flow:=FlowRange(i).Value, _
                InletPressure:=InletPressureRange(i).Value, _
                OutletPressure:=OutletPressureRange(i).Value, _
                Name:=NameRange(i).Value
        End If
    Next i

End Sub

Public Sub AddMultiPointDriverPower(ByVal DriverPowerRange As Range)
'
' Add the Driver Power to all the test points
'
    ' Validation of the ranges
    If Me.TestPoints.Count <> DriverPowerRange.Count Then
        ' Error message
        MsgBox "The block of DriverPower input data shall have the same dimensions of (Test point name, Q, P0, P3)."
        Call ErrorExit
    End If
    
    ' Iterative block
    Dim n As Integer
    Dim i As Integer
    
    n = Me.TestPoints.Count
    
    For i = 1 To n
        Me.TestPoints(i).DriverPower = DriverPowerRange(i).Value
    Next i

End Sub

Public Sub AddMultiPointNSpeed(ByVal NRange As Range)
'
' Add the Speed to all the test points
'
    ' Validation block
    If ThisPump.colTestPoints.Count <> NRange.Count Then
        ' Error message
        MsgBox "The block of N input data shall have the same dimensions of (Test point name, Q, P0, P3)."
        Call ErrorExit
    End If
    
    ' Iterative block
    Dim n As Integer
    Dim i As Integer
    
    n = Me.TestPoints.Count
    
    For i = 1 To n
        Me.TestPoints(i).NSpeed = NRange(i).Value
    Next i

End Sub

Public Sub AddMultiPointNPSH3(ByVal NPSH3_Range As Range)
'
' Add the Speed to all the test points
'
    ' Validation block
    If ThisPump.colTestPoints.Count <> NPSH3_Range.Count Then
        ' Error message
        MsgBox "The block of NPSH3 input data shall have the same dimensions of (Test point name, Q, P0, P3)."
        Call ErrorExit
    End If
    
    ' Iterative block
    Dim n As Integer
    Dim i As Integer
    
    n = Me.TestPoints.Count
    
    For i = 1 To n
        Me.TestPoints(i).NPSH3 = NPSH3_Range(i).Value
    Next i

End Sub

Public Sub AddMultiPointTemp(ByVal Temp As Double)
'
' Methodo to add a block of Temperature data
'
    Dim n As Integer
    Dim i As Integer
    
    n = Me.TestPoints.Count
    
    For i = 1 To n
        Me.TestPoints(i).Temp = Temp
    Next i

End Sub


Public Sub PrintMultiPointVar( _
    VarIndex As Integer, _
    RangeToPrint As Range _
    )
'
' Print the var
'
    ' Validation of the range
    If Me.TestPoints.Count <> RangeToPrint.Count Then
        ' Error message
        MsgBox "The block to output data shall have the same dimensions of (Test point name, Q, P0, P3)."
        Call ErrorExit
    End If
    
    ' Iterative block
    Dim n As Integer
    Dim i As Integer
    
    n = Me.TestPoints.Count
    
    For i = 1 To n
        Select Case VarIndex
        
        Case 1      ' Teste point Head - m
            RangeToPrint(i).Value = Me.TestPoints(i).Head
        
        Case 2      ' Teste point Efficiency - %
            RangeToPrint(i).Value = Me.TestPoints(i).Efficiency
            
        Case 3      ' Corrected Flow - m^3/h
            RangeToPrint(i).Value = Me.TestPointsCor(i).Q
            
        Case 4      ' Corrected Head - m
            RangeToPrint(i).Value = Me.TestPointsCor(i).Head
            
        Case 5      ' Corrected Driver Power - kW
            RangeToPrint(i).Value = Me.TestPointsCor(i).DriverPower
            
        Case 6      ' Corrected Speed - rpm
            RangeToPrint(i).Value = Me.TestPointsCor(i).NSpeed
            
        Case 7      ' Corrected Efficiency - %
            RangeToPrint(i).Value = Me.TestPointsCor(i).Efficiency
            
         Case 8      ' Corrected NPSH3 - m
            RangeToPrint(i).Value = Me.TestPointsCor(i).NPSH3
            
        Case 9      ' Flow viscous correction factor
            RangeToPrint(i).Value = Me.TestPointsCor(i).CQ
            
        Case 10      ' Head viscous correction factor
            RangeToPrint(i).Value = Me.TestPointsCor(i).CH
            
        Case 11     ' Efficiency viscous correction factor
            RangeToPrint(i).Value = Me.TestPointsCor(i).CEff
        
        End Select
            
    Next i
    
End Sub

Public Sub Update()
'
' Update Test Points calculations
'
    Dim tpTP As TestPoint
    Dim colErase As New Collection
    Dim strName As Variant
    
    With Me
        For Each tpTP In .TestPoints
            
            tpTP.Update _
                A0:=.A0, _
                A3:=.A3, _
                Z0:=.Z0, _
                Z3:=.Z3, _
                ZM0:=.ZM0, _
                ZM3:=.ZM3
                
            If tpTP.P0 = 0 _
                And tpTP.P3 = 0 _
                And tpTP.Q = 0 Then
            
                colErase.Add (tpTP.Name)
            End If
        Next tpTP
        
        For Each strName In colErase
            Me.TestPoints.Index(strName).Remove
        Next strName
    End With
End Sub

Public Sub SpeedCorrection()
'
' Correct the the test points to the rated speed and density
'
    ' Update all test points
    Me.Update
    
    ' Check if there is any corrected points into the collection
    Dim blFirstCorrection As Boolean
    If Me.TestPointsCor.Count = 0 Then
        blFirstCorrection = True
    Else
        blFirstCorrection = False
    End If
    
    Dim i As Integer
    i = 1
    
    Dim tpCorPoint As TestPointCor
    
    Dim tpTP As TestPoint
    For Each tpTP In Me.TestPoints
    
        ' Act this way if is necessarry to create the points
        If blFirstCorrection Then
            ' Instantiating new test point
            Set tpCorPoint = New TestPointCor
            tpCorPoint.Name = "Corrected" & tpTP.Name
            Me.TestPointsCor.Add _
                Item:=tpCorPoint, _
                Key:=tpCorPoint.Name
        End If
        
        With Me.TestPointsCor(i)
            ' Loading required variables
            .NSpeed = Me.RatedPoint.NSpeed
            .Rho = Me.RatedPoint.Rho
            .Efficiency = tpTP.Efficiency
            .Q = tpTP.Q * .NSpeed / tpTP.NSpeed
            .Head = tpTP.Head * (.NSpeed / tpTP.NSpeed) ^ 2
            .NPSH3 = tpTP.NPSH3 * (.NSpeed / tpTP.NSpeed) ^ 2
            .DriverPower = .Rho / tpTP.Rho * tpTP.DriverPower * (.NSpeed / tpTP.NSpeed) ^ 3
                                    
        End With
        
        i = i + 1
     
     Next tpTP
End Sub

Public Sub ViscosityCorrection()
'
' Correct the test points to the rated viscosity
'
    Dim blFirstCorrection As Boolean      ' Check if viscosity is supplied
    If Me.RatedPoint.DinVisc <> 0 Then
        
        Dim dQBEP As Double      ' BEP flow determination
        If Me.SupplierBEP = 0 Then
            dQBEP = Me.BEP
        Else
            dQBEP = Me.SupplierBEP
        End If
                
        Dim PolyH As Variant     ' BEP Head curvefit process
        PolyH = Me.PolyCurveFit( _
            Collection:="TestPoints", _
            Propertie:="Efficiency" _
            )
            
        Dim dHBEP As Double     ' BEP Head determination
        dHBEP = Me.PolyEval( _
            Var:=dQBEP, _
            Coefficients:=PolyH _
            )
            
        Dim dTestAverageSpeed As Double      ' Determination of test BEP speed
        Dim tpTP As TestPoint      ' Variable to iterate testpoints data
        
        dTestAverageSpeed = 0
        For Each tpTP In Me.TestPoints
            dTestAverageSpeed = dTestAverageSpeed + tpTP.NSpeed
        Next tpTP
        dTestAverageSpeed = dTestAverageSpeed / Me.TestPoints.Count
        
        Dim dns As Double     ' Specific speed
        dns = dTestAverageSpeed * dQBEP ^ 0.5 / (dHBEP ^ 0.75 * 3600)
        
        
        Dim dKinVisc As Double        ' Kinematic viscosity computation
        dKinVisc = 1000 * Me.RatedPoint.DinVisc / Me.RatedPoint.Rho     'cST
          
        
        Dim dB As Double      ' B parameter computation
        dB = 16.5 * (dKinVisc ^ 0.5 * _
            dHBEP ^ 0.0625) _
            / (dQBEP ^ 0.375 * dTestAverageSpeed ^ 0.25)
        '
        ' Establish the procedure base in B parameter
        '
        
        Dim i As Integer
        i = 1
        
        Dim tpCorPoint As TestPointCor      ' Variable to update corrected points
        
        If dB >= 40 Then
            
            ' Print the error
            MsgBox " B parameter > = 40. Viscous corrections not applied to the testpoints." _
                & Chr(10) & " H.I. 9.6 Loss analysis required. "
            
            ' No viscosity correction, only speed and density correction are applied
            Me.SpeedCorrection
        
        ElseIf dB <= 1 Then
            
            ' No viscosity correction, only speed and density correction are applied
            Me.SpeedCorrection
            
            ' H.I. Step 5. calculate the values for viscous pump shaft power
            i = 1
            For Each tpTP In Me.TestPoints
                With Me.TestPointsCor(i)
                    If .Q <> 0 Then
                        .DriverPower = .Q * .Head * .Density / (367 * .Efficiency / 100)
                    End If
                End With
                i = i + 1
             Next tpTP
        
        Else
            
            Dim dCQ As Double
            Dim dCEff As Double
            Dim dCH As Double
            
            ' Computation of corection factors
            dCQ = 2.71 ^ (-0.165 * (Log(dB) / Log(10#)) ^ 3.15)
            dCEff = dB ^ (-0.0547 * dB ^ 0.69)
            
            ' Speed and density
            Me.SpeedCorrection
            i = 1
            For Each tpTP In Me.TestPoints
               
               'Head correction coefficient for each flowrate
               dCH = 1 - ((1 - dCQ) * (tpTP.Q / dQBEP) ^ 0.75)
               
                With Me.TestPointsCor(i)
                    If .Q <> 0 Then
                        .CQ = dCQ
                        .CEff = dCEff
                        .CH = dCH
                        .Head = .Head * dCH
                        .Q = .Q * dCQ
                        .Efficiency = .Efficiency * dCEff
                        .DriverPower = .Q * .Head * .Density / (367 * .Efficiency / 100)
                    End If
                End With
                
                i = i + 1
            
            Next tpTP
        End If
    End If
End Sub

Public Function PolyCurveFit(ByVal Collection As String, ByVal Propertie As String) As Variant
'
' Polynomial curve fit
'
    Dim x() As Double
    Dim y() As Double
    Dim colSel As Collection
    
    ' Redimensionando o m da matriz para adequar ao tamanho do bloco
    ReDim x(0 To Me.TestPoints.Count - 1, 0 To 3)
    ReDim y(0 To Me.TestPoints.Count - 1, 0)
    
    Dim i As Integer
    Dim j As Integer
    
    Select Case Collection
        Case "TestPoints"
            Set colSel = Me.TestPoints
        Case "TestPointsCor"
            Set colSel = Me.TestPointsCor
    End Select
     
    ' Inicializando a matriz de coeficientes X
    For i = 1 To TestPoints.Count
        y(i - 1, 0) = CallByName(colSel(i), Propertie, VbGet)
        For j = 0 To 3
            x(i - 1, j) = CallByName(colSel(i), "Q", VbGet) ^ (j + 1)
        Next j
    Next i

    ' Aplicando m’nimos quadrados
    PolyCurveFit = WorksheetFunction.LinEst(y, x)

End Function

Public Function PolyEval(ByVal Var As Double, ByVal Coefficients As Variant) As Double
'
' Evaluates a polynomial function
'
    Dim i As Integer
    PolyEval = 0
    For i = 1 To UBound(Coefficients)
        PolyEval = PolyEval + Coefficients(i) * Var ^ (UBound(Coefficients) - i)
    Next i
End Function

Public Function Der(ByVal Order As Integer, ByVal Var As Double, ByVal Coefficients As Variant) As Double
'
' Take the derivative of a polynomial function
' Order: 1 and 2
'
    Dim i As Integer
    Der = 0
    For i = 1 To (UBound(Coefficients) - Order)
        Der = Der + (UBound(Coefficients) - Order - i + 1) * Coefficients(i) * Var ^ (UBound(Coefficients) - Order - i)
    Next i
End Function
   
Public Function BEP() As Double
    '
    ' BEP determination
    '
    Dim Funcao As Double
    Dim DerFuncao As Double
    Dim X0 As Double
    Dim Xi As Double
    Dim i As Integer
    Dim Poly As Variant
    
    Poly = Me.PolyCurveFit( _
        Collection:="TestPoints", _
        Propertie:="Efficiency" _
        )
    
    X0 = Me.RatedPoint.Q
    Xi = 0
    i = 1
    
    Do Until Abs(X0 - Xi) < 0.001
        ' Derivada do Polinomino (funcao a ser zerada)
        Funcao = Me.Der(Order:=1, Var:=X0, Coefficients:=Poly)
            
        ' Derivada segunda do Polinomio (funcao a ser zerada)
        DerFuncao = Me.Der(Order:=2, Var:=X0, Coefficients:=Poly)
        
        Xi = X0 - Funcao / DerFuncao
        X0 = Xi
        i = i + 1
    Loop
    
    Debug.Print "Newton loops: " & i; ""

    ' BEP flow
    ' If Xi < 0 Then Xi = WorksheetFunction.Max(dQ)
    Debug.Print "BEP Flow: " & Xi
    BEP = Xi
End Function

Sub PlotPerformance( _
    ByRef LeftCorner As Range, _
    ByRef RightMid As Range, _
    ByRef RightCorner As Range _
    )
'
' Power and performance plot
'
    '
    ' Loading the values to plot
    '
    Dim iPoints As Integer
    iPoints = Me.TestPoints.Count
    Dim i As Integer
    
    ' Test points head load
    Dim dTestQ() As Double     'Test volume flow
    Dim dTestHead() As Double     'Test head
    ReDim dTestHead(0 To iPoints - 1)
    ReDim dTestQ(0 To iPoints - 1)
    i = 0
    Dim tpTP As TestPoint      ' Test points
    For Each tpTP In Me.TestPoints
        dTestHead(i) = tpTP.Head
        dTestQ(i) = tpTP.Q
        i = i + 1
    Next tpTP
    
    ' Corrected test points head load
    Dim dTestQCor() As Double     'Corrected volume flow
    Dim dTestHeadCor() As Double     'Corrected head
    ReDim dTestHeadCor(0 To iPoints - 1)
    ReDim dTestQCor(0 To iPoints - 1)
    i = 0
    Dim tpCorPoint As TestPointCor      ' Corrected test points
    For Each tpCorPoint In Me.TestPointsCor
        dTestHeadCor(i) = tpCorPoint.Head
        dTestQCor(i) = tpCorPoint.Q
        i = i + 1
    Next tpCorPoint
    
    '  Polynomial aprox test points head load
    Dim vPolyTestHead As Variant
    vPolyTestHead = Me.PolyCurveFit("TestPoints", "Head")
    Dim dTestHeadPoly() As Double     'Test head polynomial aprox
    ReDim dTestHeadPoly(0 To iPoints - 1)
    For i = 0 To iPoints - 1
        dTestHeadPoly(i) = Me.PolyEval(dTestQ(i), vPolyTestHead)
    Next i
    
     '  Polynomial aprox corrected test points head load
    Dim vPolyTestHeadCor As Variant
     vPolyTestHeadCor = Me.PolyCurveFit("TestPointsCor", "Head")
    Dim dTestHeadCorPoly() As Double     'Corrected head polynomial aprox
    ReDim dTestHeadCorPoly(0 To iPoints - 1)
    For i = 0 To iPoints - 1
        dTestHeadCorPoly(i) = Me.PolyEval(dTestQCor(i), vPolyTestHeadCor)
    Next i
    
    ' Test points driver power load
    Dim dTestDriverPower() As Double     'Test driver power
    ReDim dTestDriverPower(0 To iPoints - 1)
    i = 0
    For Each tpTP In Me.TestPoints
        dTestDriverPower(i) = tpTP.DriverPower
        i = i + 1
    Next tpTP
    
    ' Corrected test points driver power load
    Dim dTestDriverPowerCor() As Double     'Corrected driver power
    ReDim dTestDriverPowerCor(0 To iPoints - 1)
    i = 0
    For Each tpCorPoint In Me.TestPointsCor
        dTestDriverPowerCor(i) = tpCorPoint.DriverPower
        i = i + 1
    Next tpCorPoint
    
    '  Polynomial aprox test points driver power load
    Dim vPolyTestDriverPower As Variant
    vPolyTestDriverPower = Me.PolyCurveFit("TestPoints", "DriverPower")
    Dim dTestDriverPowerPoly() As Double     'Test driver power polynomial aprox
    ReDim dTestDriverPowerPoly(0 To iPoints - 1)
    For i = 0 To iPoints - 1
        dTestDriverPowerPoly(i) = Me.PolyEval(dTestQ(i), vPolyTestDriverPower)
    Next i
    
     '  Polynomial aprox corrected test points driver power load
    Dim vPolyTestDriverPowerCor As Variant
     vPolyTestDriverPowerCor = Me.PolyCurveFit("TestPointsCor", "DriverPower")
    Dim dTestDriverPowerCorPoly() As Double     'Corrected head polynomial aprox
    ReDim dTestDriverPowerCorPoly(0 To iPoints - 1)
    For i = 0 To iPoints - 1
        dTestDriverPowerCorPoly(i) = Me.PolyEval(dTestQCor(i), vPolyTestDriverPowerCor)
    Next i
    
     '
     ' Chart creation
     '
     Dim wsWS As Worksheet
     Set wsWS = LeftCorner.Worksheet
     
     ' Check if there is any other graph into the sheet
    If wsWS.ChartObjects.Count > 0 Then
        wsWS.ChartObjects.Delete
    End If
    
    ' Chart position
    Dim iGrLeft As Integer, iGrTop As Integer, iGrWidth As Integer, iGrHeight As Integer
    
    ' Lef and top position to insert the graph
    With LeftCorner
        iGrLeft = .Left + 1
        iGrTop = .Top + 1
    End With
    
    ' Width and Height of the chart
    With RightMid
        iGrWidth = .Left + .Width - iGrLeft
        iGrHeight = .Top - iGrTop - 1
    End With
    
    'Chart object creation
    Dim coPer As ChartObject
    Set coPer = wsWS.ChartObjects.Add(iGrLeft, iGrTop, iGrWidth, iGrHeight)
    coPer.Name = "Performance " & Me.TAG
        
    With coPer.Chart

        ' Title settings
        .HasTitle = True
        .ChartTitle.Text = "Performance " & Me.TAG
            
        .ChartArea.Format.Line.Visible = msoFalse
        .HasLegend = False
        
        ' Chart type
        .ChartType = xlXYScatter
            
        ' Two axis creation
        .HasAxis(xlCategory, xlPrimary) = True
        .HasAxis(xlCategory, xlSecondary) = True
        .HasAxis(xlValue, xlPrimary) = True
        .HasAxis(xlValue, xlSecondary) = True
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDash
            .MajorGridlines.Border.Color = RGB(180, 180, 180)
        End With
        
        With .Axes(xlCategory, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDash
            .MajorGridlines.Border.Color = RGB(180, 180, 180)
            .TickLabelPosition = xlTickLabelPositionHigh
            .Border.LineStyle = xlNone
            .MajorTickMark = xlNone
            .MinorTickMark = xlNone
        End With
        
        Debug.Print "width: " & .PlotArea.InsideWidth
        Debug.Print iGrWidth
        .PlotArea.InsideWidth = 0.8 * iGrWidth
        .PlotArea.InsideLeft = 0.1 * iGrWidth

        
        ' Plotting test points head
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = dTestQ
            .Values = dTestHead
            .AxisGroup = xlPrimary
            .MarkerBackgroundColor = RGB(0, 102, 204)
            .MarkerForegroundColor = RGB(0, 102, 204)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
            
        ' Plotting corrected test points head
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .XValues = dTestQCor
            .Values = dTestHeadCor
            .AxisGroup = xlPrimary
            .MarkerBackgroundColor = RGB(0, 133, 66)
            .MarkerForegroundColor = RGB(0, 133, 66)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
        
        ' Plotting test points head polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(3)
            .XValues = dTestQ
            .Values = dTestHeadPoly
            .AxisGroup = xlPrimary
            .Format.Line.DashStyle = msoLineLongDash
            .Format.Line.Weight = 0.6
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
            
        ' Plotting corrected test points head polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(4)
            .XValues = dTestQCor
            .Values = dTestHeadCorPoly
            .AxisGroup = xlPrimary
            .Format.Line.Style = msoLineSingle
            .Format.Line.Weight = 0.5
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
           
        ' Rated point head
        .SeriesCollection.NewSeries
        With .SeriesCollection(5)
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestHead)
            .AxisGroup = xlPrimary
            .MarkerForegroundColor = RGB(51, 51, 153)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
            
        ' Rated point correct head
        .SeriesCollection.NewSeries
        With .SeriesCollection(6)
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestHeadCor)
            .AxisGroup = xlPrimary
            .MarkerForegroundColor = RGB(255, 0, 0)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
        
        ' Plotting test points driver power
        .SeriesCollection.NewSeries
        With .SeriesCollection(7)
            .XValues = dTestQ
            .Values = dTestDriverPower
            .AxisGroup = xlSecondary
            .MarkerBackgroundColor = RGB(0, 102, 204)
            .MarkerForegroundColor = RGB(0, 102, 204)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
            
        ' Plotting corrected test points driver power
        .SeriesCollection.NewSeries
        With .SeriesCollection(8)
            .XValues = dTestQCor
            .Values = dTestDriverPowerCor
            .AxisGroup = xlSecondary
            .MarkerBackgroundColor = RGB(0, 133, 66)
            .MarkerForegroundColor = RGB(0, 133, 66)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
        
        ' Plotting test points driver power polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(9)
            .XValues = dTestQ
            .Values = dTestDriverPowerPoly
            .AxisGroup = xlSecondary
            .Format.Line.DashStyle = msoLineLongDash
            .Format.Line.Weight = 0.6
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
            
        ' Plotting corrected test points driver power polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(10)
            .XValues = dTestQCor
            .Values = dTestDriverPowerCorPoly
            .AxisGroup = xlSecondary
            .Format.Line.Style = msoLineSingle
            .Format.Line.Weight = 0.5
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
           
        ' Rated point driver power
        .SeriesCollection.NewSeries
        With .SeriesCollection(11)
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestDriverPower)
            .AxisGroup = xlSecondary
            .MarkerForegroundColor = RGB(51, 51, 153)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
            
        ' Rated point correct driver power
        .SeriesCollection.NewSeries
        With .SeriesCollection(12)
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestDriverPowerCor)
            .AxisGroup = xlSecondary
            .MarkerForegroundColor = RGB(255, 0, 0)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
        
        ' Ajustando as escalas
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Head (m)"
            .MinimumScale = Round(WorksheetFunction.Min(dTestHead) * 0.2 / 10, 0) * 10
            .MaximumScale = Round(WorksheetFunction.Max(dTestHeadCor) * 1.1 / 10, 0) * 10
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "Driver Power (kW)"
            .MinimumScale = Round(WorksheetFunction.Min(dTestDriverPowerCor) * 0.7 / 10, 0) * 10
            .MaximumScale = Round(WorksheetFunction.Max(dTestDriverPower) * 2.5 / 10, 0) * 10
        End With
           
    End With
End Sub

Sub PlotEfficiency( _
    ByRef LeftCorner As Range, _
    ByRef RightMid As Range, _
    ByRef RightCorner As Range _
    )
'
' Efficiency and NPSH plot
'
    '
    ' Loading the values to plot
    '
    Dim iPoints As Integer
    Dim bQ0 As Boolean
    iPoints = Me.TestPoints.Count
    Dim i As Integer
    
    ' Test points efficiency load
    Dim dTestQ() As Double     'Test volume flow
    Dim dTestEfficiency() As Double     'Test efficiency
    ReDim dTestEfficiency(0 To iPoints - 1)
    ReDim dTestQ(0 To iPoints - 1)
    i = 0
    Dim tpTP As TestPoint      ' Test points
    For Each tpTP In Me.TestPoints
        
        If tpTP.Q = 0 Then bQ0 = True      ' Boolean to identify flowarate = 0 mö3/h
    
        dTestEfficiency(i) = tpTP.Efficiency
        dTestQ(i) = tpTP.Q
        i = i + 1
    Next tpTP
    
    ' Corrected test points Efficiency load
    Dim dTestQCor() As Double     'Corrected volume flow
    Dim dTestEfficiencyCor() As Double     'Corrected Efficiency
    ReDim dTestEfficiencyCor(0 To iPoints - 1)
    ReDim dTestQCor(0 To iPoints - 1)
    i = 0
    Dim tpCorPoint As TestPointCor      ' Corrected test points
    For Each tpCorPoint In Me.TestPointsCor
        dTestEfficiencyCor(i) = tpCorPoint.Efficiency
        dTestQCor(i) = tpCorPoint.Q
        i = i + 1
    Next tpCorPoint
    
    '  Polynomial aprox test points Efficiency load
    Dim vPolyTestEfficiency As Variant
    vPolyTestEfficiency = Me.PolyCurveFit("TestPoints", "Efficiency")
    Dim dTestEfficiencyPoly() As Double     'Test Efficiency polynomial aprox
    ReDim dTestEfficiencyPoly(0 To iPoints - 1)
    For i = 0 To iPoints - 1
        dTestEfficiencyPoly(i) = Me.PolyEval(dTestQ(i), vPolyTestEfficiency)
    Next i
    
     '  Polynomial aprox corrected test points Efficiency load
    Dim vPolyTestEfficiencyCor As Variant
     vPolyTestEfficiencyCor = Me.PolyCurveFit("TestPointsCor", "Efficiency")
    Dim dTestEfficiencyCorPoly() As Double     'Corrected Efficiency polynomial aprox
    ReDim dTestEfficiencyCorPoly(0 To iPoints - 1)
    For i = 0 To iPoints - 1
        dTestEfficiencyCorPoly(i) = Me.PolyEval(dTestQCor(i), vPolyTestEfficiencyCor)
    Next i
    
    ' Test points NPSH3 resize variables
    Dim iPoints2 As Integer
    
    If bQ0 Then
        iPoints2 = iPoints - 1
    Else
        iPoints2 = iPoints
    End If
    
    ' Test points NPSH3 load
    Dim dTestNPSH3() As Double     'Test driver power
    Dim dTestNPSH3Q() As Double
    ReDim dTestNPSH3(0 To iPoints2 - 1)
    ReDim dTestNPSH3Q(0 To iPoints2 - 1)
    i = 0
    For Each tpTP In Me.TestPoints
        If tpTP.Q <> 0 Then
            dTestNPSH3Q(i) = tpTP.Q
            dTestNPSH3(i) = tpTP.NPSH3
            i = i + 1
        End If
    Next tpTP
    
    ' Corrected test points driver power load
    Dim dTestNPSH3Cor() As Double     'Corrected driver power
    Dim dTestNPSH3QCor() As Double
    ReDim dTestNPSH3Cor(0 To iPoints2 - 1)
    ReDim dTestNPSH3QCor(0 To iPoints2 - 1)
    i = 0
    For Each tpCorPoint In Me.TestPointsCor
        If tpCorPoint.Q <> 0 Then
            dTestNPSH3QCor(i) = tpCorPoint.Q
            dTestNPSH3Cor(i) = tpCorPoint.NPSH3
            i = i + 1
        End If
    Next tpCorPoint
    
    '  Polynomial aprox test points driver power load
    Dim vPolyTestNPSH3 As Variant
    vPolyTestNPSH3 = Me.PolyCurveFit("TestPoints", "NPSH3")
    Dim dTestNPSH3Poly() As Double     'Test driver power polynomial aprox
    ReDim dTestNPSH3Poly(0 To iPoints2 - 1)
    For i = 0 To iPoints2 - 1
        dTestNPSH3Poly(i) = Me.PolyEval(dTestNPSH3Q(i), vPolyTestNPSH3)
    Next i
    
     '  Polynomial aprox corrected test points driver power load
    Dim vPolyTestNPSH3Cor As Variant
     vPolyTestNPSH3Cor = Me.PolyCurveFit("TestPointsCor", "NPSH3")
    Dim dTestNPSH3CorPoly() As Double     'Corrected Efficiency polynomial aprox
    ReDim dTestNPSH3CorPoly(0 To iPoints2 - 1)
    For i = 0 To iPoints2 - 1
        dTestNPSH3CorPoly(i) = Me.PolyEval(dTestNPSH3QCor(i), vPolyTestNPSH3Cor)
    Next i
    
    '
    ' Chart creation
    '
    Dim wsWS As Worksheet
    Set wsWS = LeftCorner.Worksheet
    
    ' Chart position
    Dim iGrLeft As Integer, iGrTop As Integer, iGrWidth As Integer, iGrHeight As Integer
    
    ' Left position to insert the graph
    iGrLeft = LeftCorner.Left + 1
    
    ' Top position to insert the graph
    iGrTop = RightMid.Top
    
    ' Width and Height of the chart
    With RightCorner
        iGrWidth = .Left + .Width - iGrLeft
        iGrHeight = .Top + .Height - iGrTop - 1
    End With
    
    'Chart object creation
    Dim coPer As ChartObject
    Set coPer = wsWS.ChartObjects.Add(iGrLeft, iGrTop, iGrWidth, iGrHeight)
    coPer.Name = "Efficiency " & Me.TAG
        
    With coPer.Chart

        ' Title settings
        .HasTitle = False
            
        .ChartArea.Format.Line.Visible = msoFalse
        
        ' Legends data
        Const strTestData As String = "Test data"
        Const strCorData As String = "Corrected data"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' Chart type
        .ChartType = xlXYScatter
            
        ' Two axis creation
        .HasAxis(xlCategory, xlPrimary) = True
        .HasAxis(xlCategory, xlSecondary) = True
        .HasAxis(xlValue, xlPrimary) = True
        .HasAxis(xlValue, xlSecondary) = True
        
        With .Axes(xlValue, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDash
            .MajorGridlines.Border.Color = RGB(180, 180, 180)
        End With
        
        With .Axes(xlCategory, xlPrimary)
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDash
            .MajorGridlines.Border.Color = RGB(180, 180, 180)
            .HasTitle = True
            .AxisTitle.Caption = "Volume Flow (m^3/h)"
        End With
        
        Debug.Print "width: " & .PlotArea.InsideWidth
        Debug.Print iGrWidth
        .PlotArea.InsideWidth = 0.8 * iGrWidth
        .PlotArea.InsideLeft = 0.1 * iGrWidth
        .PlotArea.InsideTop = 0
        
        
        
        
        ' Plotting test points Efficiency
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = strTestData
            .XValues = dTestQ
            .Values = dTestEfficiency
            .AxisGroup = xlPrimary
            .MarkerBackgroundColor = RGB(0, 102, 204)
            .MarkerForegroundColor = RGB(0, 102, 204)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
            
        ' Plotting corrected test points Efficiency
        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .Name = strCorData
            .XValues = dTestQCor
            .Values = dTestEfficiencyCor
            .AxisGroup = xlPrimary
            .MarkerBackgroundColor = RGB(0, 133, 66)
            .MarkerForegroundColor = RGB(0, 133, 66)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
        
        ' Plotting test points Efficiency polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(3)
            .Name = strTestData
            .XValues = dTestQ
            .Values = dTestEfficiencyPoly
            .AxisGroup = xlPrimary
            .Format.Line.DashStyle = msoLineLongDash
            .Format.Line.Weight = 0.6
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
            
        ' Plotting corrected test points Efficiency polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(4)
            .Name = strCorData
            .XValues = dTestQCor
            .Values = dTestEfficiencyCorPoly
            .AxisGroup = xlPrimary
            .Format.Line.Style = msoLineSingle
            .Format.Line.Weight = 0.5
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
           
        ' Rated point Efficiency
        .SeriesCollection.NewSeries
        With .SeriesCollection(5)
            .Name = "Rated Point"
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestEfficiency)
            .AxisGroup = xlPrimary
            .MarkerForegroundColor = RGB(51, 51, 153)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
            
        ' Rated point correct Efficiency
        .SeriesCollection.NewSeries
        With .SeriesCollection(6)
            .Name = "Guarantee point"
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestEfficiencyCor)
            .AxisGroup = xlPrimary
            .MarkerForegroundColor = RGB(255, 0, 0)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
        
        ' Plotting test points NPSH3
        .SeriesCollection.NewSeries
        With .SeriesCollection(7)
            .Name = strTestData
            .XValues = dTestNPSH3Q
            .Values = dTestNPSH3
            .AxisGroup = xlSecondary
            .MarkerBackgroundColor = RGB(0, 102, 204)
            .MarkerForegroundColor = RGB(0, 102, 204)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
            
        ' Plotting corrected test points NPSH3
        .SeriesCollection.NewSeries
        With .SeriesCollection(8)
            .Name = strCorData
            .XValues = dTestNPSH3QCor
            .Values = dTestNPSH3Cor
            .AxisGroup = xlSecondary
            .MarkerBackgroundColor = RGB(0, 133, 66)
            .MarkerForegroundColor = RGB(0, 133, 66)
            .MarkerSize = 7
            .MarkerStyle = xlMarkerStyleCircle
        End With
        
        ' Plotting test points NPSH3 polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(9)
            .Name = strTestData
            .XValues = dTestNPSH3Q
            .Values = dTestNPSH3Poly
            .AxisGroup = xlSecondary
            .Format.Line.DashStyle = msoLineLongDash
            .Format.Line.Weight = 0.6
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
            
        ' Plotting corrected test points NPSH3 polynomial aprox
        .SeriesCollection.NewSeries
        With .SeriesCollection(10)
            .Name = strCorData
            .XValues = dTestNPSH3QCor
            .Values = dTestNPSH3CorPoly
            .AxisGroup = xlSecondary
            .Format.Line.Style = msoLineSingle
            .Format.Line.Weight = 0.5
            .MarkerStyle = xlMarkerStyleNone
            .Smooth = True
        End With
           
        ' Rated point NPSH3
        .SeriesCollection.NewSeries
        With .SeriesCollection(11)
            .Name = strTestData
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestNPSH3)
            .AxisGroup = xlSecondary
            .MarkerForegroundColor = RGB(51, 51, 153)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
            
        ' Rated point correct NPSH3
        .SeriesCollection.NewSeries
        With .SeriesCollection(12)
            .Name = "Remove"
            .XValues = Me.RatedPoint.Q
            .Values = Me.PolyEval(Me.RatedPoint.Q, vPolyTestNPSH3Cor)
            .AxisGroup = xlSecondary
            .MarkerForegroundColor = RGB(255, 0, 0)
            .MarkerSize = 15
            .MarkerStyle = xlMarkerStyleX
        End With
        
        ' Ajustando as escalas
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Efficiency (%)"
            .MinimumScale = Round(WorksheetFunction.Min(0) * 0.2 / 10, 0) * 10
            .MaximumScale = Round(WorksheetFunction.Max(dTestEfficiency) * 1.1 / 10, 0) * 10
        End With
        
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "NPSH3 (m)"
            .MinimumScale = Round(WorksheetFunction.Min(dTestNPSH3Cor) * 0.7 / 10, 0) * 10
            .MaximumScale = Round(WorksheetFunction.Max(dTestNPSH3) * 2.5 / 10, 0) * 10
        End With
        
        ' Removing some legends entries
        For i = .Legend.LegendEntries.Count To .Legend.LegendEntries.Count - 5 Step -1
                .Legend.LegendEntries(i).Delete
        Next i

           
    End With
End Sub
