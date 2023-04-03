Attribute VB_Name = "PIReqTransfer"
'Purpose To Transfer PI and Plate Information onto the Analysis SampleSheets
'Written by Kenny Ma. Contact: kyoma17@gmail.com Cell number:(626)-246-2233
'Version 1.0 July 21, 2021
Option Explicit

'Global Variables
Dim mainWB          As Sheets
Dim READ_Me         As Worksheet
Dim AnalysisSS, ExtractionSS, ResultSS As range

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object

Dim analysisFile, exFile, sample, Exsample As Variant
Dim analysisBook, exBook As Workbook
Dim samplenameColumn, stypeColumn, PIcolumn, ReqColumn, lastRow As Integer
Dim AnalysisRange1, AnalysisRange2  As String

Dim fontColor, PI, reqNumber As String

Dim exBookArray()   As Variant


'Main Script
Public Sub PiREQ()
    turnOff
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    
    'Sets up the Global Variables Using the Settings Page
    'Analysis SS
    
    Set asFSOLibrary = New FileSystemObject
    Set asFSOFolder = asFSOLibrary.GetFolder(AnalysisSS)
    Set asFSOFile = asFSOFolder.Files
    
    'Extraction SS
    Set esFSOLibrary = New FileSystemObject
    Set esFSOFolder = esFSOLibrary.GetFolder(ExtractionSS)
    Set esFSOFile = esFSOFolder.Files
    
    Dim i As Integer

    'Use For Each loop to loop through each file in the folder
    For Each analysisFile In asFSOFile
        
        If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then
            
            Workbooks.OpenText FileName:=analysisFile, Local:=True
            Set analysisBook = ActiveWorkbook
            
            addTransferColumn analysisBook
            checkPIheader analysisBook
            
            samplenameColumn = SearchColumnSource("samplename", analysisBook)
            stypeColumn = SearchColumnSource("stype", analysisBook)
            PIcolumn = SearchColumnSource("pi", analysisBook)
            ReqColumn = SearchColumnSource("req", analysisBook)
            checkPIheader analysisBook
            PIcolumn = SearchColumnSource("pi", analysisBook)
            ReqColumn = SearchColumnSource("req", analysisBook)
            
            lastRow = analysisBook.Sheets(1).Cells(Rows.count, samplenameColumn).End(xlUp).Row
            
            AnalysisRange1 = Split(Cells(1, samplenameColumn).Address, "$")(1) & 2
            AnalysisRange2 = Split(Cells(1, samplenameColumn).Address, "$")(1) & lastRow
            
            
            
            Dim matcher As Variant
            
            Dim found As Boolean
            Dim explate As String
            
            For Each sample In analysisBook.Sheets(1).range(AnalysisRange1, AnalysisRange2)
                found = False
                
            
                If Not analysisBook.Sheets(1).Cells(sample.Row, PIcolumn) = "" Then
                    'Skip
                ElseIf sample = "r62" Or sample = "HOMO" Or sample = "HET" Or sample = "" Or sample = "NTC" Or sample = "WT" Then
                    'Skip
                Else
                    explate = Split(sample, "$")(0)
                    
                    
                    
                    For Each exFile In esFSOFile
                        If found Then
                            'Skip
                        ElseIf explate = Replace(Split(exFile.name, "-")(1), ".xlsx", "") And InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 Then
                                
                                Workbooks.OpenText FileName:=exFile, Local:=True
                                Set exBook = ActiveWorkbook
                                For Each Exsample In exBook.Worksheets(1).range("B16:M23")
                                    If found Then
                                        'Skip
                                    ElseIf sample = Exsample Then
                                       
                                        For Each matcher In Union(exBook.Sheets(1).range("A11:R11"), exBook.Sheets(1).range("N1:N14"))
                                            
                                            If found Then
                                                'Skip
                                            ElseIf matcher = "" Then
                                                'skip
                                            ElseIf matcher.Font.color = Exsample.Font.color Then
                                                found = True
                                                For i = 1 To 6
                                                
                                                    If exBook.Worksheets(1).Cells(matcher.Row, matcher.column + i) = "" Then
                                                        'skip
                                                    ElseIf matcher.Font.color = exBook.Worksheets(1).Cells(matcher.Row, matcher.column + i).Font.color Then
                                                        
                                                        PI = exBook.Worksheets(1).Cells(matcher.Row, matcher.column + i)
                                                        reqNumber = exBook.Worksheets(1).Cells(matcher.Row + 1, matcher.column + i)

                                                    
                                                    End If
                                                Next i

                                                analysisBook.Sheets(1).Cells(sample.Row, PIcolumn) = PI
                                                analysisBook.Sheets(1).Cells(sample.Row, ReqColumn) = reqNumber
                                                
                                                
                                            End If
                                        Next
                                    End If
                                Next
                                exBook.Close savechanges:=False
    
                        End If
                        
                    Next
                End If
            Next
            analysisBook.Close savechanges:=True
        End If
        
    Next
    
    'Release the memory
    Set asFSOLibrary = Nothing
    Set asFSOFolder = Nothing
    Set asFSOFile = Nothing
    Set rsFSOLibrary = Nothing
    Set rsFSOFolder = Nothing
    Set rsFSOFile = Nothing
    Set esFSOLibrary = Nothing
    Set esFSOFolder = Nothing
    Set esFSOFile = Nothing
    
    reset
    MsgBox ("PIREQ Complete")
End Sub

'Adds PI and Req Header if does not exists
Sub checkPIheader(aWB)
    Dim stypeBpi As Integer
    stypeBpi = SearchColumnSource("stype", analysisBook)
    
    If SearchColumnSource("pi", analysisBook) = 0 Then
        aWB.Sheets(1).range(Split(Cells(1, stypeBpi + 1).Address, "$")(1) & ":" & Split(Cells(1, stypeBpi + 1).Address, "$")(1)).EntireColumn.insert
        aWB.Sheets(1).Cells(1, stypeBpi + 1) = "PI"
        
    End If
    
    If SearchColumnSource("req", analysisBook) = 0 Then
        aWB.Sheets(1).range(Split(Cells(1, stypeBpi + 2).Address, "$")(1) & ":" & Split(Cells(1, stypeBpi + 2).Address, "$")(1)).EntireColumn.insert
        aWB.Sheets(1).Cells(1, stypeBpi + 2) = "REQ"
        
    End If
End Sub

