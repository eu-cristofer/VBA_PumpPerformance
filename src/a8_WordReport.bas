Attribute VB_Name = "a8_WordReport"
Option Explicit

Dim oWordApp As Object
Dim oWordDoc As Object

Sub WordReport(ByRef strSheets As Variant)
    
    ' Start of optimization actions
    Call OptimizationStart
    
    ' Check OS and define path to save auxiliary files
    Dim strPath As String
    Dim bOSX As Boolean
    
    If Application.OperatingSystem Like "*Mac*" Then
        strPath = Environ("HOME") & Application.PathSeparator
        bOSX = True
    ElseIf Application.OperatingSystem Like "*Wind*" Then
        strPath = Application.ThisWorkbook.Path & Application.PathSeparator
        bOSX = False
    Else
        MsgBox "Not possible to run in this system"
        Call ErrorExit
    End If
    
    ' Create word instance
    Call CreateDoc(bOSX)
    
    Call PageSetup
    
    Dim strSheet As Variant
    Dim strSheetName As String
    Dim strTag As String
    Dim strTitle As String
    Dim strBody As String
    
    
    For Each strSheet In strSheets
        ActiveWorkbook.Sheets(strSheet).Activate
        
        ' Update the calc
        Call PumpPerformanceMain
        
        strSheetName = "'" & strSheet & "'!"
        strTag = Range(strSheetName & "PumpTAG").Value
    
        strTitle = GetTitle(strTag)
        
        strBody = GetBodyWord(strSheetName)
        
        With oWordApp
            .Selection.TypeText strTitle
            
            .Selection.TypeParagraph
            
            .Selection.TypeText strBody
        End With
    
        ' Performance chart insertion procedure
        Dim strChartName As String
        strChartName = strPath & "ChartReport.jpg"
    
        Call ExportChart(strChartName, strTag)
        
        ' Inserindo a figura
        With oWordDoc
            .InlineShapes.AddPicture _
                FileName:=strChartName, _
                LinkToFile:=False, _
                savewithdocument:=True, _
                Range:=oWordApp.Selection.Range
        End With
        
        oWordApp.Selection.EndKey Unit:=5, Extend:=0
        
        oWordApp.Selection.TypeParagraph
            
    Next strSheet

    ' End of optimization actions
    Call OptimizationEnd
    
End Sub

Private Function GetTitle(strTagName As String) As String

    GetTitle = "Resultados do Teste de Performance - " & strTagName
    
End Function

Private Function GetBodyWord(strTD As String) As String
    
    Dim strBody As String

    ' Head text
    strBody = "Head no ponto garantido" & Chr(10) _
        & "     Valor contratual: " & Range(strTD & "RatedPointHead").Value & " m." & Chr(10) _
        & "     Aproximação polinomial: " & Range(strTD & "RatedPointHeadPoly").Value & " m." & Chr(10) _
        & "     Aproximação spline: " & Range(strTD & "RatedPointHeadSpline").Value & " m." & Chr(10)
        

    ' Power text
    strBody = strBody & Chr(10) & "Potência no ponto garantido" & Chr(10) _
        & "     Valor contratual: " & Range(strTD & "RatedPointDriverPower").Value & " kW." & Chr(10) _
        & "     Aproximação polinomial: " & Range(strTD & "RatedPointDriverPowerPoly").Value & " kW." & Chr(10) _
        & "     Aproximação spline: " & Range(strTD & "RatedPointDriverPowerSpline").Value & " kW." & Chr(10)

   ' Efficiency text
    strBody = strBody & Chr(10) & "Eficiência no ponto garantido" & Chr(10) _
        & "     Valor contratual: " & Range(strTD & "RatedPointEfficiency").Value & " %." & Chr(10) _
        & "     Aproximação polinomial: " & Range(strTD & "RatedPointEfficiencyPoly").Value & " %." & Chr(10) _
        & "     Aproximação spline: " & Range(strTD & "RatedPointEfficiencySpline").Value & " %." & Chr(10)

    GetBodyWord = strBody
        
End Function

Private Sub CreateDoc(ByRef IsOSX As Boolean)

    
    On Error Resume Next
        Set oWordApp = GetObject(, "Word.Application")        'gives error 429 if Word is not open
        If Err = 429 Then
            Set oWordApp = CreateObject("Word.Application")        'creates a Word Application
            Err.Clear
        End If
    On Error GoTo 0
    
    If Not IsOSX Then
        oWordApp.Visible = True
    End If

    Set oWordDoc = oWordApp.Documents.Add

End Sub


Private Sub PageSetup()
        
        ' Margins setup
        With oWordApp.ActiveDocument.PageSetup
            .TopMargin = oWordApp.MillimetersToPoints(30)
            .BottomMargin = oWordApp.MillimetersToPoints(20)
            .LeftMargin = oWordApp.MillimetersToPoints(20)
            .RightMargin = oWordApp.MillimetersToPoints(20)
        End With

End Sub

Private Sub ExportChart(ByRef ChartName As String, ByRef sTagName As String)

    Dim myShape As Shape
    Dim iShapeCounter As Integer
    Dim chtObj As ChartObject
    Dim ZoomIni As Integer, ZoomPrint As Integer
    
    ' Grouping the itens
    With ActiveSheet.Shapes
        .Range(Array("Performance " & sTagName, "Efficiency " & sTagName)).Group
        iShapeCounter = .Count
    End With
    
    Set myShape = ActiveWorkbook.ActiveSheet.Shapes(iShapeCounter)
    
    ' Adjust the zoom before make a copy
    ZoomIni = ActiveWindow.Zoom
    ZoomPrint = 150
    ActiveWindow.Zoom = ZoomPrint
    
    ' Copy the shape
    myShape.Select
    Selection.CopyPicture xlPrinter
    
    ' Create a new chartobject with the same dimensions as the source shape
    Set chtObj = ActiveSheet.ChartObjects.Add(myShape.Left, myShape.Top, myShape.Width, myShape.Height)
    
    chtObj.Select

    chtObj.Chart.Paste
    
    chtObj.Chart.ChartArea.Format.Line.Visible = msoFalse
    
    ' Export the chart to the specified directory, using the specified name, and save the chart
    chtObj.Chart.Export _
        FileName:=ChartName, _
        Filtername:="JPG"
    
    ' Ungroup thecharts
    myShape.Ungroup
    
    '  empty the object
    Set chtObj = Nothing

    ' Reset zoom settings
    ActiveWindow.Zoom = ZoomIni

End Sub
