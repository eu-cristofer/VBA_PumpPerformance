Attribute VB_Name = "a6_ResetMain"
Option Explicit

Sub ResetMain(ByRef wsWS As Worksheet)
'
' Sub to clean all the calculated data into the InputTestData
'
    ' Erase chart objects
    Dim coCO As Object
    For Each coCO In wsWS.ChartObjects
        coCO.Delete
    Next coCO
    
    ' Erase all the named ranges in gray
    Dim strSheet As String
    Dim N As Integer
    Dim ranCel As Range
    Dim wbWB As Workbook
    
    strSheet = "ResetFields"
    
    ' Add-in workbook
    Set wbWB = Application.Workbooks("vba-pump-performance.xlam")
    Set ranCel = wbWB.Worksheets(strSheet).Range("A1")
    N = LastRow(ranCel)
    
    Dim i As Integer
    Dim strAux As String
    Dim strName As String
    
    For i = 1 To N
        strAux = wbWB.Sheets(strSheet).Cells(i, 1).Value
        strName = wsWS.Name
        wsWS.Range("'" & strName & "'" & strAux).ClearContents
    Next i

End Sub
