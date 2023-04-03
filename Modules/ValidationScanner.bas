Attribute VB_Name = "ValidationScanner"
'Checks the number of samples In result sheets and number of samples in analysis SS to make sure there is a correct number of samples or if there area extras.

Option Explicit
Dim mainWB          As Sheets
Dim TestWB          As Workbook
Dim READ_Me, TAC_TEST_LIST, JAX_TEST_LIST, SETTINGS, COPY_PASTE As Worksheet

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object
Dim AnalysisSS, ExtractionSS, TacResultSS, JaxResultSS, MargeResultSS, TestDB, summaryMessage  As String

Sub runAllValidation()
    CreateTextFileBasic "ExtraSamples", "Taconic:" & TACverifyTotalSamples & vbNewLine & _
                        "Jax:" & vbNewLine & JAXverifyTotalSamples & _
                        "Marge:" & vbNewLine & MargeVerifyTotalSamples, _
                        READ_Me
    MsgBox ("Complete Validation, Please confirm Text file")
End Sub

Function MargeVerifyTotalSamples() As String
    turnOff
    
    summaryMessage = "Marge Extra Sample:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set TacResultSS = READ_Me.Cells(14, 2)
    Set JaxResultSS = READ_Me.Cells(15, 2)
    Set MargeResultSS = READ_Me.Cells(16, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    'Sets up the Global Variables Using the Settings Page
    
    Dim last_row    As Integer
    Dim indexNum    As Integer
    Dim resultTextLocation As Integer
    'Analysis SS
    
    Set asFSOLibrary = New FileSystemObject
    Set asFSOFolder = asFSOLibrary.GetFolder(AnalysisSS)
    Set asFSOFile = asFSOFolder.Files
    
    'Extraction SS
    Set esFSOLibrary = New FileSystemObject
    Set esFSOFolder = esFSOLibrary.GetFolder(ExtractionSS)
    Set esFSOFile = esFSOFolder.Files
    
    'Result SS
    Set rsFSOLibrary = New FileSystemObject
    Set rsFSOFolder = rsFSOLibrary.GetFolder(MargeResultSS)
    Set rsFSOFile = rsFSOFolder.Files
    
    'Use For Each loop to loop through each file in the folder
    
    Dim resultBook, analysisBook, exBook As Workbook
    Dim reqNumber   As String
    Dim resultfile, exFile As Variant
    Dim samplenameColumn, lastRowResults, totalResultSample, totalExSample As Integer
    Dim Exsample, secondExsample, cell As Variant
    
    Dim complete    As Boolean
    
    Dim ResultSSStartRow As Integer
    
    For Each resultfile In rsFSOFile
        complete = False
        If InStr(resultfile, ".xlsx") > 0 And InStr(resultfile, "~$") < 1 And InStr(resultfile, "desktop.ini") < 1 Then
            Workbooks.Open FileName:=resultfile
            Set resultBook = ActiveWorkbook
            
            reqNumber = Replace(resultBook.Sheets(1).range("A3"), "Req #: ", "")
            
            samplenameColumn = SearchColumnTargetU("mouseid", resultBook, "Marge")
            lastRowResults = resultBook.Sheets(1).Cells(Rows.count, samplenameColumn).End(xlUp).Row
            
            ResultSSStartRow = FindFirstHashRow(resultBook) + 1
            
            
            totalResultSample = lastRowResults - ResultSSStartRow
            
            For Each exFile In esFSOFile
                'Set up Extraction Files
                If complete Then
                    'Skip
                ElseIf InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 Then
                    
                    Workbooks.Open FileName:=exFile
                    Set exBook = ActiveWorkbook
                    'Req Number Matching
                    For Each cell In Union(exBook.Sheets(1).range("A11", "M12"), exBook.Sheets(1).range("N1", "P13"))        ' Need to add the other stray ones on the right side as well
                        If cell = "" Or complete Then
                            
                        ElseIf InStr(cell, reqNumber) > 0 Then
                            totalExSample = 0
                            complete = True
                            For Each Exsample In exBook.Sheets(1).range("B3:M10")
                                If cell.Font.color = Exsample.Font.color And Not Exsample = "" Then
                                    totalExSample = totalExSample + 1
                                End If
                            Next
                            
                            If Not totalExSample = totalResultSample Then
                                summaryMessage = summaryMessage & "Marge: There are extra samples For " & resultfile.name & "(" & totalResultSample & ")" & _
                                                 " Please see ExtSS: " & exFile.name & "(" & totalExSample & ")" & vbNewLine
                            End If
                            
                        End If
                        
                    Next
                    exBook.Close savechanges:=False
                End If
            Next
            resultBook.Close savechanges:=False
            
        End If
    Next
    
    reset
    MargeVerifyTotalSamples = summaryMessage
    
End Function

Public Function SearchColumnTargetU(search, wb, org) As Integer
    ' Dynamic Implemntation for Start Row
    Dim sampleStartRow As Integer
    sampleStartRow = FindFirstHashRow(wb)
    
    Dim orgRange    As Variant
    
    If org = "Taconic" Then
        orgRange = wb.Sheets(1).range("A17", "K17")
    ElseIf org = "JAX" Then
        orgRange = wb.Sheets(1).range("A" & sampleStartRow, "Z" & sampleStartRow) 'CHANGE When new Row
    ElseIf org = "Marge" Then
        orgRange = wb.Sheets(1).range("A" & sampleStartRow, "Z" & sampleStartRow)
    Else
        orgRange = wb.Sheets(1).range("A17", "K17")
    End If
    Dim columnIndex As Integer
    columnIndex = 1
    Dim field       As Variant
    For Each field In orgRange        'The Template rows has to start at 17
        If Replace(LCase(field), " ", "") = search Then
            SearchColumnTargetU = columnIndex
        End If
        columnIndex = columnIndex + 1
    Next
End Function



Function TACverifyTotalSamples() As String
    turnOff
    
    summaryMessage = "Taconic Extra Sample:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set TacResultSS = READ_Me.Cells(14, 2)
    Set JaxResultSS = READ_Me.Cells(15, 2)
    Set MargeResultSS = READ_Me.Cells(16, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    'Sets up the Global Variables Using the Settings Page
    
    Dim last_row    As Integer
    Dim indexNum    As Integer
    Dim resultTextLocation As Integer
    'Analysis SS
    
    Set asFSOLibrary = New FileSystemObject
    Set asFSOFolder = asFSOLibrary.GetFolder(AnalysisSS)
    Set asFSOFile = asFSOFolder.Files
    
    'Extraction SS
    Set esFSOLibrary = New FileSystemObject
    Set esFSOFolder = esFSOLibrary.GetFolder(ExtractionSS)
    Set esFSOFile = esFSOFolder.Files
    
    'Result SS
    Set rsFSOLibrary = New FileSystemObject
    Set rsFSOFolder = rsFSOLibrary.GetFolder(TacResultSS)
    Set rsFSOFile = rsFSOFolder.Files
    
    'Use For Each loop to loop through each file in the folder
    
    Dim resultBook, analysisBook, exBook As Workbook
    Dim reqNumber   As String
    Dim resultfile, exFile As Variant
    Dim samplenameColumn, lastRowResults, totalResultSample, totalExSample As Integer
    Dim Exsample, secondExsample As Variant
    
    Dim complete    As Boolean
    For Each resultfile In rsFSOFile
        complete = False
        If InStr(resultfile, ".xlsx") > 0 And InStr(resultfile, "~$") < 1 And InStr(resultfile, "desktop.ini") < 1 Then
            Workbooks.Open FileName:=resultfile
            Set resultBook = ActiveWorkbook
            
            samplenameColumn = SearchColumnTargetU("animalid", resultBook, "Taconic")
            lastRowResults = resultBook.Sheets(1).Cells(Rows.count, samplenameColumn).End(xlUp).Row
            totalResultSample = lastRowResults - 17
            For Each exFile In esFSOFile
                'Set up Extraction Files
                If complete Then
                    'Skip
                ElseIf InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 Then
                    
                    Workbooks.Open FileName:=exFile
                    Set exBook = ActiveWorkbook
                    If Not complete Then
                        
                        For Each Exsample In exBook.Sheets(1).range("B3:M10")
                            
                            If resultBook.Sheets(1).Cells(lastRowResults, samplenameColumn) = Exsample And Not secondExsample = "" Then
                                totalExSample = 0
                                complete = True
                                For Each secondExsample In exBook.Sheets(1).range("B3:M10")
                                    If secondExsample.Font.color = Exsample.Font.color Then
                                        totalExSample = totalExSample + 1
                                    End If
                                Next
                                If Not totalExSample = totalResultSample Then
                                    summaryMessage = summaryMessage & "Tac: There are extra samples For " & resultfile.name & "(" & totalResultSample & ")" & _
                                                     " Please see ExtSS: " & exFile.name & "(" & totalExSample & ")" & vbNewLine
                                End If
                            End If
                        Next
                    End If
                    
                    exBook.Close savechanges:=False
                End If
            Next
            resultBook.Close savechanges:=False
            
        End If
    Next
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    TACverifyTotalSamples = summaryMessage
    
End Function

Function JAXverifyTotalSamples() As String
    turnOff
    
    summaryMessage = "Jax Extra Sample:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set TacResultSS = READ_Me.Cells(14, 2)
    Set JaxResultSS = READ_Me.Cells(15, 2)
    Set MargeResultSS = READ_Me.Cells(16, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    'Sets up the Global Variables Using the Settings Page
    
    Dim last_row    As Integer
    Dim indexNum    As Integer
    Dim resultTextLocation As Integer
    'Analysis SS
    
    Set asFSOLibrary = New FileSystemObject
    Set asFSOFolder = asFSOLibrary.GetFolder(AnalysisSS)
    Set asFSOFile = asFSOFolder.Files
    
    'Extraction SS
    Set esFSOLibrary = New FileSystemObject
    Set esFSOFolder = esFSOLibrary.GetFolder(ExtractionSS)
    Set esFSOFile = esFSOFolder.Files
    
    'Result SS
    Set rsFSOLibrary = New FileSystemObject
    Set rsFSOFolder = rsFSOLibrary.GetFolder(JaxResultSS)
    Set rsFSOFile = rsFSOFolder.Files
    
    'Use For Each loop to loop through each file in the folder
    
    Dim resultBook, analysisBook, exBook As Workbook
    Dim reqNumber   As String
    Dim resultfile, exFile As Variant
    Dim samplenameColumn, lastRowResults, totalResultSample, totalExSample As Integer
    Dim Exsample, secondExsample, cell As Variant
    
    Dim ResultSSStartRow As Integer
    
    
    
    Dim complete    As Boolean
    For Each resultfile In rsFSOFile
        complete = False
        If InStr(resultfile, ".xlsx") > 0 And InStr(resultfile, "~$") < 1 And InStr(resultfile, "desktop.ini") < 1 Then
            Workbooks.Open FileName:=resultfile
            Set resultBook = ActiveWorkbook
            samplenameColumn = SearchColumnTargetU("mouseid", resultBook, "JAX")
            lastRowResults = resultBook.Sheets(1).Cells(Rows.count, samplenameColumn).End(xlUp).Row
            
            ResultSSStartRow = FindFirstHashRow(resultBook) + 1
            
            totalResultSample = lastRowResults - ResultSSStartRow 'Need to edit this if changed
            
            For Each exFile In esFSOFile
                'Set up Extraction Files
                If complete Then
                    'Skip
                ElseIf InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 Then
                    
                    Workbooks.Open FileName:=exFile
                    Set exBook = ActiveWorkbook
                    If Not complete Then
                        
                        For Each Exsample In exBook.Sheets(1).range("B3:M10")
                            If complete Then
                                
                            ElseIf resultBook.Sheets(1).Cells(lastRowResults, samplenameColumn) = Exsample Then
                                totalExSample = 0
                                complete = True
                                For Each secondExsample In exBook.Sheets(1).range("B3:M10")
                                    If secondExsample.Font.color = Exsample.Font.color And Not secondExsample = "" Then
                                        totalExSample = totalExSample + 1
                                        
                                    End If
                                Next
                                If Not totalExSample = totalResultSample Then
                                    summaryMessage = summaryMessage & "JAX: There are extra samples For " & resultfile.name & "(" & totalResultSample & ")" & _
                                                     " Please see ExtSS: " & exFile.name & "(" & totalExSample & ")" & vbNewLine
                                End If
                            End If
                        Next
                        
                    End If
                    
                    exBook.Close savechanges:=False
                End If
            Next
            resultBook.Close savechanges:=False
            
        End If
    Next
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    JAXverifyTotalSamples = summaryMessage
End Function



