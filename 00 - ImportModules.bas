Attribute VB_Name = "Tool_ImportModules"
Option Explicit

Sub ImportModules()
' Code snippet to import modules listed the text file
' named 00_module_list.txt from folder \src
'
' VBA Pump Performance
' https://github.com/eu-cristofer/vba-pump-performance

    ' Open the folder selection dialog
    Dim selectedFolder As FileDialog
    Set selectedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    selectedFolder.Title = "Select src folder"
    
    ' Check if the user clicked OK
    If selectedFolder.Show = -1 Then
        Dim folderPath As String
        folderPath = selectedFolder.SelectedItems(1)
        
        ' Specify the path to the text file containing module names
        Dim txtFilePath As String
        txtFilePath = folderPath & "\00_module_list.txt"
        
        ' Check if the text file exists
        If Dir(txtFilePath) <> "" Then
            ' Read the content of the text file
            Dim txtFileContent As String
            Open txtFilePath For Input As #1
            txtFileContent = Input$(LOF(1), 1)
            Close #1
            
            ' Split the content into an array of module names
            Dim moduleNames() As String
            moduleNames = Split(txtFileContent, vbCrLf)
            
            ' Import each module
            Dim moduleName As String
            Dim i As Long
            For i = LBound(moduleNames) To UBound(moduleNames)
                moduleName = Trim(moduleNames(i))
                If moduleName <> "" Then
                    ImportModuleFromFile folderPath & "\" & moduleName
                End If
            Next i
        Else
            MsgBox "Text file 'module_list.txt' not found in the selected folder.", vbExclamation
        End If
    Else
        MsgBox "Try again. Pick the src folder and click OK"
    End If
End Sub

Private Sub ImportModuleFromFile(ByRef filePath As String)
    
    Dim vbComp As Object
    ' Check if the module file exists
    If Dir(filePath) <> "" Then
        ' Import the module
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
    Else
        MsgBox "Module file not found: " & filePath, vbExclamation
    End If
End Sub