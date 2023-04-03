Attribute VB_Name = "TransferReport"
Option Explicit

Dim summaryMessage  As String
Dim mainWB          As Sheets
Dim TestWB          As Workbook
Dim READ_Me, TAC_TEST_LIST, SETTINGS As Worksheet
Dim AnalysisSS, ExtractionSS, ResultSS, TestDB As range
Dim extraTestNames, extraStypes, standardtestnames As range

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object

Dim analysisFile, exFile, sample, Exsample As Variant
Dim analysisBook, exBook As Workbook

Dim SSN, tran, PI, REQ, lastRow As Integer
Dim tran1, tran2    As String

Dim MissingExSS, MissingReSS, AnalysisRange1, AnalysisRange2 As String

Sub TransferReport()
    turnOff

    Dim summaryMessage As String
    Dim answer      As String

    summaryMessage = "Summary:" & vbNewLine
    MissingExSS = "Please locate ExtractionSS And Re-Run PiReqTransfer:" & vbNewLine
    MissingReSS = "Please Locate ResultSS Paperwork:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    
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
    
    Dim workSample  As Variant
    
    'Use For Each loop to loop through each file in the folder
    For Each analysisFile In asFSOFile
        
        If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then
            
            Workbooks.OpenText FileName:=analysisFile, Local:=True
            Set analysisBook = ActiveWorkbook
            
            SSN = SearchColumnSource("samplename", analysisBook)
            tran = SearchColumnSource("tran", analysisBook)
            PI = SearchColumnSource("pi", analysisBook)
            REQ = SearchColumnSource("req", analysisBook)
            
            If tran = 0 Then
                summaryMessage = summaryMessage & "Transfer column Not found For " & Left(analysisFile.name, 20)
            ElseIf PI = 0 Or REQ = 0 Then
                summaryMessage = summaryMessage & "PiReq columns Not found For " & Left(analysisFile.name, 20)
            Else
                lastRow = analysisBook.Sheets(1).Cells(Rows.count, SSN).End(xlUp).Row
                
                For Each workSample In analysisBook.Sheets(1).range(Cells(2, SSN), Cells(lastRow, SSN))
                    
                    If InStr(workSample, "$") < 1 Then
                        
                    ElseIf analysisBook.Sheets(1).Cells(workSample.Row, tran) = "" And analysisBook.Sheets(1).Cells(workSample.Row, PI) = "" Then
                        
                        If InStr(MissingExSS, Split(workSample, "$")(0)) < 1 Then
                            MissingExSS = MissingExSS & Split(workSample, "$")(0) & vbNewLine
                        End If
                        
                    ElseIf analysisBook.Sheets(1).Cells(workSample.Row, tran) = "" And Not analysisBook.Sheets(1).Cells(workSample.Row, PI) = "" Then
                        If InStr(MissingReSS, analysisBook.Sheets(1).Cells(workSample.Row, REQ)) < 1 Then
                            
                            MissingReSS = MissingReSS & analysisBook.Sheets(1).Cells(workSample.Row, PI) & " " _
                                        & analysisBook.Sheets(1).Cells(workSample.Row, REQ) & " For " & Left(analysisFile.name, 20) & vbNewLine
                        End If
                        
                    End If
                Next
            End If
            analysisBook.Close savechanges:=False
        End If
        
    Next
    Set rsFSOLibrary = Nothing
    Set rsFSOFolder = Nothing
    Set rsFSOFile = Nothing
    
    reset
    
    CreateTextFileBasic "TransferReport", MissingExSS & vbNewLine & MissingReSS & vbNewLine & summaryMessage, READ_Me
    
    MsgBox (MissingExSS & vbNewLine & MissingReSS)
End Sub

'Adds transfer column header if does not exist
Public Sub addTransferColumn(aWB)
    addLastcolumn aWB, "TRAN"
End Sub



