Attribute VB_Name = "AutoPath"
Option Explicit

Function GetExcelFilePath()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get the path of the current Excel file
    Dim ExcelFilePath As String
    ExcelFilePath = ThisWorkbook.Path
    
    Debug.Print (ExcelFilePath)
    
    Dim RootFolder As String
    
    Dim lastBackslash As Integer
    lastBackslash = InStrRev(ExcelFilePath, "\")
    
    ' Extract the directory part
    If lastBackslash > 0 Then
        RootFolder = Left(ExcelFilePath, lastBackslash - 1)
    End If
        
    Debug.Print (RootFolder)
    
    
    GetExcelFilePath = ExcelFilePath

End Function


Function GetComputerName()
    Dim ComputerName As String
    ComputerName = Environ("COMPUTERNAME") ' Get the computer name
    
    Debug.Print (ComputerName)
    GetComputerName = ComputerName
    
End Function
    
    
Sub AutoPath()
    Dim StartFolderName As Integer
    StartFolderName = "30"
    
    Dim StartCellPath As Integer
    StartCellPath = "44"

    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("READ_ME")
    
    Dim Folder As range
    Dim i As Integer
    
    For i = 0 To 10
        Set Folder = ws.Cells(StartFolderName + i, 13)
        
        'Debug.Print (Folder)
        Debug.Print (ws.Cells(StartCellPath + i, 13))
        
        
        If GetComputerName = "LARAGEN-DEV" Then
            Dim ExcelFilePath As String
            ExcelFilePath = GetExcelFilePath()
            
            ws.Cells(StartCellPath + i, 13).Value = ExcelFilePath & Folder.Value
            
        End If
        
        
    Next i

End Sub
