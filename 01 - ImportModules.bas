Attribute VB_Name = "Tool_ModulesImport"
Option Explicit
'
' VBA Pump Performance
' Dev. tool to import all VBAComponents from a folder
'
Public Sub ImportModules()
    Dim strWB As String
    Dim strAddress As String
    Const strSheet As String = "ModList"
    Dim n As Integer
    Dim ranCell As Range
    Dim i As Integer
    Dim oVBP As Object ' VBProject
    Dim oWS As Object

    Dim strFiles() As String

    Set oWS = Application.ThisWorkbook.Sheets(strSheet)

    strAddress = "/Users/cristofercosta/Desktop/Teste/"

    Set ranCell = oWS.Range("A1")

    n = LastRow(ranCell)

    ReDim strFiles(0 To n - 1)

    For i = 0 To n - 1
        strFiles(i) = strAddress & oWS.Cells(i + 1, 1)
    Next i

    If Application.OperatingSystem Like "*Mac*" Then
        If RequestFileAccess(strFiles) Then
        Else
            MsgBox "Access denied"
            Exit Sub
        End If
    End If

    Set oVBP = Application.ThisWorkbook.VBProject.VBComponents

    For i = 0 To n - 1
        oVBP.Import strFiles(i)
        Debug.Print strFiles(i)
    Next i

    Debug.Print "Imported modules: " & n

End Sub

Private Function LastRow(ranCell As Range) As Integer
'
' Function to find the last row of column
'
    If Not IsEmpty(ranCell.Offset(1, 0)) Then
        LastRow = ranCell.End(xlDown).Row
    Else
        LastRow = 1
    End If

    Debug.Print "VBA_PumpPerformance Modules: " & LastRow

End Function

Private Function RequestFileAccess(strFileName As Variant) As Boolean
'
' Request permissions Mac OS only
'
    ' Request access from user
    ' Argument: an array with file paths for the permissions that are needed
    RequestFileAccess = GrantAccessToMultipleFiles(strFileName)

End Function
