Attribute VB_Name = "TaconicConversion"
'Purpose To Convert the Tactonic samplesheets into the Laragen Result sheets.
'Written by Kenny Ma. Contact: kyoma17@gmail.com Cell number:(626)-246-2233
'Version 1.1 July 21, 2021

Public Sub GenerateFiles()
    turnOff
    
    Dim summaryMessage As String
    summaryMessage = "Taconic Conversion Summary:" & vbNewLine
    
    Dim folderName  As String
    Dim OutputfolderName As String
    Dim FSOLibrary  As FileSystemObject
    Dim FSOFolder   As Object
    Dim FSOFile     As Object
    
    Dim wss         As Sheets
    Dim READ_Me     As Worksheet
    
    'Assigns the Read_Me Worksheet
    Set wss = ThisWorkbook.Worksheets
    Set READ_Me = wss("READ_ME")
    Dim EntryDate   As String
    
    'Temporary Worksheets
    
    'Seperate the file name into Project Number and Plate Number
    Dim fileSplitName As Variant
    
    Dim strainProjectNumber As Variant
    Dim plateNumber As Variant
    Dim plainName   As Variant
    
    'Target Column Variables
    Dim TargetSampleName As Variant
    Dim TargetWellPlate As Variant
    Dim TargetSerialNumber As Variant
    
    'SourcePlate Column Variables
    Dim SourceProjectNumber As Variant
    Dim SourceAscesioningNumber As Variant
    Dim SourceSampleName As Variant
    Dim SourceWellPlate As Variant
    Dim SourceSerialNumber As Variant
    Dim last_row    As Integer
    Dim indexNum    As Integer
    
    '!!! Set the path of working folder!!!
    'InputfolderName = "C:\\Users\\\SERVER\\Desktop\\INPUT\"
    'OutputfolderName = "C:\\Users\\\SERVER\\Desktop\\OUTPUT\"
    'TemplateFolder = "C:\\Users\\\SERVER\\Desktop\\Taconic_Strain_Template\"
    
    'This section will Refer to the Excel Read-Me Fields"
    TemplateFolder = READ_Me.Cells(19, 2)
    InputfolderName = READ_Me.Cells(18, 2)
    OutputfolderName = READ_Me.Cells(14, 2)
    
    EntryDate = InputBox("Please Enter the Date:")
    
    'Set all the references to the INPUT FSO Library
    Set FSOLibrary = New FileSystemObject
    Set FSOFolder = FSOLibrary.GetFolder(InputfolderName)
    Set FSOFile = FSOFolder.Files
    
    'Set all the references to the TEMPLATE FSO Library
    Set TemplateFSOLibrary = New FileSystemObject
    Set TemplateFSOFolder = TemplateFSOLibrary.GetFolder(TemplateFolder)
    Set TemplateFSOFile = TemplateFSOFolder.Files
    
    Dim Index       As Integer
    Dim found       As Boolean
    Dim match       As Boolean
    Dim Current     As String
    
    'Use For Each loop to loop through each file in the folder
    For Each EachFSOFile In FSOFile
        If InStr(EachFSOFile, ".csv") > 0 And InStr(EachFSOFile, "~$") < 1 Then
            fileSplitName = Split(EachFSOFile.name, "_")        'Splits the Strain and Plate from File name
            plainName = Replace(EachFSOFile.name, ".csv", "")        'Removes .csv from filiename
            
            If InStr(fileSplitName(0), "CHDI") Then
                strainProjectNumber = fileSplitName(1)
                plateNumber = fileSplitName(2)
            Else
                strainProjectNumber = fileSplitName(0)
                plateNumber = fileSplitName(1)
            End If
            
            Workbooks.OpenText FileName:=Replace(EachFSOFile, "\", "\\"), Local:=True
            Set inputbook = ActiveWorkbook
            
            'Verification if project number inside title matches the project number inside the file.
            projectNumVerificatoin = SearchColumnSource("project", inputbook)
            projectnameinsidefile = inputbook.Sheets(1).Cells(2, projectNumVerificatoin).Value()
            If projectnameinsidefile = "" Then
                projectnameinsidefile = inputbook.Sheets(1).Cells(3, projectNumVerificatoin).Value()
            End If
            
            
            If Not Str(projectnameinsidefile) = Str(strainProjectNumber) Then
                summaryMessage = summaryMessage & plainName & "'s title does not match Project Number " & projectnameinsidefile & _
                                 " inside file." & vbNewLine
                
            End If
            
            found = False
            
            For Each EachTemplateFSOFile In TemplateFSOFile
                If InStr(EachTemplateFSOFile, ".xlsx") > 0 And InStr(EachTemplateFSOFile, "~$") < 1 Then
                    StrainName = Replace(EachTemplateFSOFile.name, ".xlsx", "")        'Removes .xlxs from filiename
                    
                    If StrainName = strainProjectNumber Then
                        'Grabs Column Data from CSV and enters it into the Results Document.
                        Workbooks.OpenText FileName:=Replace(EachTemplateFSOFile, "\", "\\"), Local:=True
                        Set OutputTemplate = ActiveWorkbook
                        
                        'Grabbing ColumnIndex from TargetBook
                        TargetSampleName = SearchColumnTarget("animalid", OutputTemplate)
                        TargetWellPlate = SearchColumnTarget("plate#", OutputTemplate)
                        TargetSerialNumber = SearchColumnTarget("serial#", OutputTemplate)
                        
                        'Grabbing ColumnIndex from Source book
                        SourceSampleName = SearchColumnSource("samplename", inputbook)
                        SourceWellPlate = SearchColumnSource("wellplate", inputbook)
                        SourceSerialNumber = SearchColumnSource("serial#", inputbook)
                        
                        'Grabbing Number of samples in Input File
                        lastRow = inputbook.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 2
                        
                        'Adds The Index to the Output based on the number of Samples from Input
                        addIndex lastRow, OutputTemplate
                        
                        
                        copyData lastRow, OutputTemplate, inputbook, SourceSampleName, TargetSampleName
                        copyData lastRow, OutputTemplate, inputbook, SourceWellPlate, TargetWellPlate
                        copyData lastRow, OutputTemplate, inputbook, SourceSerialNumber, TargetSerialNumber
                        
                        'Adds Date to the Sheet Assuming Box is in J4
                        OutputTemplate.Worksheets(1).range("J4") = EntryDate
                        
                        inputbook.Close
                        OutputTemplate.SaveAs _
                                              FileName:=OutputfolderName & plainName & ".xlsx"
                        
                        OutputTemplate.Close
                        found = True
                        Exit For
                    End If
                End If
                
            Next
            If found = False Then
                summaryMessage = summaryMessage & plainName & "'s Template File Not Found for Strain #" _
                               & strainProjectNumber & ". Please Add Template File To Folder" & vbNewLine
                
            End If
            'Looks up the Strain number in the StrainTemplates Page
        End If
        
    Next
    
    'Release the memory
    Set FSOLibrary = Nothing
    Set FSOFolder = Nothing
    Set FSOFile = Nothing

    reset
    
    CreateTextFileBasic "TacConversion", summaryMessage, READ_Me
    MsgBox (summaryMessage)
End Sub

'''''''''''''''''''''''''''''HELPER FUNCTIONS''''''''''''''''''''''''''''''''

Sub autoFormat()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With


End Sub

'This function Copies the Data from the Source to the Output file.
Public Function copyData(lastRow, OutputWB, InputWB, source, target)
    Dim indexNum, i As Integer
    Dim insert  As String
    Dim splitInsert
    
    indexNum = 18
    For i = 2 To lastRow + 2
        insert = InputWB.Worksheets(1).Cells(i, source)
        
        
        'If detects that cell has been convertedd to a date, this portion will revert it
        If InStr(insert, "/") > 0 Then
            splitInsert = Split(insert, "/")
            insert = splitInsert(2) & "-" & splitInsert(0) & "-" & splitInsert(1)
            OutputWB.Sheets(1).Cells(indexNum, target).NumberFormat = "@"
        End If
        
        If Not insert = "" Then
            
            OutputWB.Sheets(1).Cells(indexNum, target) = insert
            
            indexNum = indexNum + 1
            
        End If
        
    Next
End Function

'This function will look for the names in the Target file and return the index of that column
Public Function SearchColumnTarget(search, wb) As Integer
    Dim columnIndex As Integer
    columnIndex = 1
    'The Template rows has to start at 17 for Taconic
    For Each field In wb.Sheets(1).range("A17", "K17")
        If Replace(LCase(field), " ", "") = search Then
            SearchColumnTarget = columnIndex
        End If
        columnIndex = columnIndex + 1
    Next
End Function

'This function will add Rows the index numbers to the Target sheet
Function addIndex(lastRow, wb)
    Dim indexNum    As Integer
    indexNum = 1
    For i = 18 To lastRow + 18
        wb.Sheets(1).Cells(i, 1) = indexNum
        indexNum = indexNum + 1
    Next
End Function



