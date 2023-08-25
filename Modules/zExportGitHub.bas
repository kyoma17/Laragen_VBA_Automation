Attribute VB_Name = "zExportGitHub"
Option Explicit

'Exports each module as a .bas file for easy upload to github.
Sub ExportModulesAsBAS()
    Dim VBComp As Object
    Dim ExportPath As String
    Dim FileName As String
    
    ' Get the path of the active workbook
    ExportPath = ActiveWorkbook.Path & "\Modules\"
    
    ' Create the Modules folder if it doesn't exist
    If Dir(ExportPath, vbDirectory) = "" Then
        MkDir ExportPath
    End If
    
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

'Import the BAS files Do not use, broken.
Sub ImportBASFilesAsModules()
    Dim ImportPath As String
    Dim FileName As String
    Dim ModName As String
    Dim ModCode As String
    Dim ExistingMod As Object
    
    ' Get the path of the active workbook
    ImportPath = ActiveWorkbook.Path & "\Modules\"
    
    ' Loop through each BAS file in the specified path
    FileName = Dir(ImportPath & "*.bas")
    Do While FileName <> ""
        ' Get the name of the module from the file name
        ModName = Left(FileName, InStrRev(FileName, ".") - 1)
        ' Read the code from the BAS file and remove the Attribute VB_Name statement
        Open ImportPath & FileName For Input As #1
        ModCode = Input$(LOF(1), #1)
        Close #1
        ModCode = Replace(ModCode, "Attribute VB_Name =", "")
        ' Check if a module with the same name already exists
        On Error Resume Next
        Set ExistingMod = ActiveWorkbook.VBProject.VBComponents(ModName)
        On Error GoTo 0
        ' If the module exists, overwrite it with the imported code
        If Not ExistingMod Is Nothing Then
            ExistingMod.CodeModule.DeleteLines 1, ExistingMod.CodeModule.CountOfLines
            ExistingMod.CodeModule.AddFromString ModCode
        ' If the module doesn't exist, add a new module with the imported code
        Else
            Set ExistingMod = ActiveWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
            ExistingMod.name = ModName
            ExistingMod.CodeModule.AddFromString ModCode
        End If
        ' Get the next BAS file in the folder
        FileName = Dir()
    Loop
    
    MsgBox "All BAS files have been imported as modules."
End Sub
