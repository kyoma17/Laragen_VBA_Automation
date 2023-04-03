Attribute VB_Name = "zExportGitHub"
Option Explicit

Sub ExportModulesAsBAS()
    Dim VBComp As Object
    Dim ExportPath As String
    Dim FileName As String
    
    ' Change the export path to your desired location
    ExportPath = "C:\Users\SERVER\Desktop\VBA_EXPORT\"
    
    ' Loop through each module in the workbook
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        ' Skip any components that are not modules
        If VBComp.Type = vbext_ct_StdModule Then
            ' Get the name of the module
            FileName = VBComp.name & ".bas"
            ' Export the module as a BAS file to the specified path
            VBComp.Export ExportPath & FileName
        End If
    Next VBComp
    
    MsgBox "All modules have been exported as BAS files to " & ExportPath
End Sub
