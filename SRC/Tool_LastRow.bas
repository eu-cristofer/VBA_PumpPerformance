Attribute VB_Name = "Tool_LastRow"
Option Explicit
'
' Function to find the last row of column
'
Public Function LastRow(ranCell As Range) As Integer
    
    If Not IsEmpty(ranCell.Offset(1, 0)) Then
        LastRow = ranCell.End(xlDown).Row
    Else
        LastRow = 1
    End If

    Debug.Print "LastRow Module:" & Chr(10) & "     Read"
    Debug.Print "Named Ranges: " & LastRow

End Function
