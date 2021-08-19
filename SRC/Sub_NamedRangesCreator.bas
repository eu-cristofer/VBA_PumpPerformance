Attribute VB_Name = "Sub_NamedRangesCreator"
Option Explicit
'
' This script creates the NamedRanges for read the input data
'
Sub NameRangesCreator()
    
    Const strSheet As String = "NamedRangesList"
    Dim n As Integer
    Dim i As Integer
    Dim ranCell As Range
    Dim oName As Object
    Dim oWB As Workbook
    Dim strSheetRef As String
    Dim strCellRef As String
    Dim intRColor As Integer
    Dim intGColor As Integer
    Dim intBColor As Integer
    
    Set oWB = Application.ThisWorkbook
    
    ' Erase all other names in the workbook
    For Each oName In oWB.Sheets(strSheet).Names
        oName.Delete
    Next
    
    Set ranCell = oWB.Sheets(strSheet).Range("A1")
    
    n = LastRow(ranCell)
     
    For i = 1 To n
        oWB.Names.Add _
            Name:=oWB.Sheets(strSheet).Cells(i, 2).Value, _
            RefersTo:=oWB.Sheets(strSheet).Cells(i, 1).Value
        
        ' Painting the cell interior
        strSheetRef = oWB.Sheets(strSheet).Cells(i, 4).Value
        strCellRef = oWB.Sheets(strSheet).Cells(i, 5).Value
        intRColor = oWB.Sheets(strSheet).Cells(i, 6).Value
        intGColor = oWB.Sheets(strSheet).Cells(i, 7).Value
        intBColor = oWB.Sheets(strSheet).Cells(i, 8).Value
        
        Application.ThisWorkbook _
            .Sheets(strSheetRef).Range(strCellRef).Interior.Color = _
                RGB(intRColor, intGColor, intBColor)
    Next i
    
    Debug.Print "NameRangesCreator Module" & Chr(10) & "     Read"
    
End Sub
