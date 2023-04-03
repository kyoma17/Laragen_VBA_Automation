Attribute VB_Name = "RepeatTransfer"
Option Explicit
Dim mainWB          As Sheets
Dim TestWB          As Workbook
Dim READ_Me, TAC_TEST_LIST, JAX_TEST_LIST, SETTINGS, COPY_PASTE As Worksheet
Dim summaryMessage, notTransferred, missingRT As String

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object
Dim AnalysisSS, ExtractionSS, TacResultSS, JaxResultSS, MargeResultSS, RepeatSS, TestDB As String

Dim repeatGT, repeatSN, repeatPCR1, repeatPCR2, repeatStype, repeatTran As Integer
Dim analysisGT, analysisSN, analysisPCR1, analysisPCR2, analysisStype As Integer
Dim resultBook, analysisBook, exBook, repeatBook As Workbook

Dim AnalysisLastRow As Integer
Dim AnalysisRange1, AnalysisRange2, repeatRange1, repeatRange2 As String



Dim repeatfile, analysisFile As Variant


Sub RepeatTransfer()
    turnOff
    
    summaryMessage = "Repeat Transfer Summary:" & vbNewLine
    notTransferred = "Not Transfferred:" & vbNewLine
    missingRT = "Missing Repeat Samples:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set RepeatSS = READ_Me.Cells(21, 2)

    
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
    Set rsFSOFolder = rsFSOLibrary.GetFolder(RepeatSS)
    Set rsFSOFile = rsFSOFolder.Files
    
    'Use For Each loop to loop through each file in the folder
    

    Dim reqNumber   As String

    Dim samplenameColumn, lastRowResults, repeatLastRow, totalResultSample, totalExSample As Integer
    Dim Exsample, secondExsample, cell As Variant
    
    Dim holdingString As String

    
    
    Dim repeatSample, analysisSample As Variant
    
    Dim complete    As Boolean
    
    For Each repeatfile In rsFSOFile
        If InStr(repeatfile, ".xlsx") > 0 And InStr(repeatfile, "~$") < 1 Then
            Workbooks.Open FileName:=repeatfile
            Set repeatBook = ActiveWorkbook
            
                addTransferColumn repeatBook
                repeatTran = SearchColumnSource("tran", repeatBook)
            
                repeatGT = SearchColumnSource("gt", repeatBook)
                repeatSN = SearchColumnSource("samplename", repeatBook)
                repeatStype = SearchColumnSource("stype", repeatBook)
                repeatPCR1 = SearchColumnSource("pcr1", repeatBook)
                repeatPCR2 = SearchColumnSource("pcr2", repeatBook)
                
                repeatLastRow = repeatBook.Sheets(1).Cells(Rows.count, repeatSN).End(xlUp).Row
                repeatRange1 = Split(Cells(1, repeatSN).Address, "$")(1) & 2
                repeatRange2 = Split(Cells(1, repeatSN).Address, "$")(1) & repeatLastRow
                
                
                            
            For Each analysisFile In asFSOFile
                If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then
                    
                    Workbooks.Open FileName:=analysisFile
                    Set analysisBook = ActiveWorkbook
                    
                    
                    createRepeated repeatBook, analysisBook

                    analysisGT = SearchColumnSource("gt", analysisBook)
                    analysisSN = SearchColumnSource("samplename", analysisBook)
                    analysisStype = SearchColumnSource("stype", analysisBook)
                    analysisPCR1 = SearchColumnSource("pcr1", analysisBook)
                    analysisPCR2 = SearchColumnSource("pcr2", analysisBook)
                    
                    AnalysisLastRow = analysisBook.Sheets(1).Cells(Rows.count, analysisSN).End(xlUp).Row
                    AnalysisRange1 = Split(Cells(1, analysisSN).Address, "$")(1) & 2
                    AnalysisRange2 = Split(Cells(1, analysisSN).Address, "$")(1) & AnalysisLastRow
                
                    MatchDataRepeat repeatBook, analysisBook
                    
                    If Not SearchColumnSource("rt", analysisBook) = 0 Then
                        sortBook analysisBook, "rt"
                    End If
                    
                    holdingString = verificationRepeatTransfer2
                    If InStr(missingRT, holdingString) < 1 Then
                        missingRT = missingRT & verificationRepeatTransfer2
                    End If
                    
                    analysisBook.Close savechanges:=True
                End If
            Next
            notTransferred = notTransferred & verifyRepeatTransfer & vbNewLine
            
            repeatBook.Close savechanges:=True
        End If
    Next
    
    reset
    
    CreateTextFileBasic "RepeatTransfer", summaryMessage & vbNewLine & notTransferred & vbNewLine & missingRT, READ_Me
    MsgBox ("Complete, Please review Summary Output")
    
End Sub

Function SampleNameSearchRepeats(name, aWB, stype) As Integer
    Dim sampType, Rowindex, columnIndex, sampleColumn As Integer
    
    sampType = SearchColumnSource("stype", aWB)
    sampleColumn = SearchColumnSource("samplename", aWB)
    Dim sbool As Boolean
    
    
    Rowindex = 2
    
    'Change this when doing JaX
    Dim field       As Variant
    
    'Loops through analysis Sheet and searches for Sample Name
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)
        If field = name And LCase(aWB.Sheets(1).Cells(field.Row, sampType)) = LCase(stype) Then
            
            SampleNameSearchRepeats = Rowindex
            Exit For
        End If
        Rowindex = Rowindex + 1
    Next
End Function

Sub MatchDataRepeat(repeatBook, analysisBook)
    Dim lastRow, Index, rowdata  As Integer
    
    Dim sample      As Variant
    
    'Loop through samples in ResultSS and match with AnalysisSS rows.
    'For Each Sample In resultBook.Worksheets(1).Range("B18", "B114") 'Fix this to make dynamic
    
    For Each sample In repeatBook.Sheets(1).range(repeatRange1, repeatRange2)
        If sample = "" Or sample = "HET" Or sample = "HOMO" Or sample = "WT" Or sample = "R62" Or sample = "CAR" Or sample = "NTC" Then
            'skip
        Else
        
        'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
        rowdata = SampleNameSearchRepeats(sample, analysisBook, repeatBook.Sheets(1).Cells(sample.Row, repeatStype))

        
        If Not rowdata = 0 Then
            
            analysisBook.Sheets(1).Cells(rowdata, analysisGT) = repeatBook.Sheets(1).Cells(sample.Row, repeatGT)
            
            
            analysisBook.Sheets(1).Cells(rowdata, analysisSN).Interior.color = vbGreen
            
            If Not analysisPCR1 = 0 And Not repeatPCR1 = 0 Then
                If Not repeatBook.Sheets(1).Cells(sample.Row, repeatPCR1) = "" Then
                    analysisBook.Sheets(1).Cells(rowdata, analysisPCR1) = repeatBook.Sheets(1).Cells(sample.Row, repeatPCR1)
                End If
            End If
            repeatBook.Sheets(1).Cells(sample.Row, repeatTran) = "CP_to_" & analysisBook.name
            
            
        End If
    
        
        
        Index = Index + 1
        
    Next
    
    
    
End Sub



Sub createRepeated(rpWB, aWB)
    Dim repeatCol, repeatLastColumn As Integer
    Dim RanalysisGT, RanalysisSN, RanalysisPCR1, RanalysisPCR2 As Integer
    repeatCol = SearchColumnSource("repeat->", aWB)
    repeatLastColumn = aWB.Sheets(1).UsedRange.Columns.count
    
    RanalysisGT = SearchColumnSource("gt", aWB)
    RanalysisSN = SearchColumnSource("samplename", aWB)
    RanalysisPCR1 = SearchColumnSource("pcr1", aWB)
    RanalysisPCR2 = SearchColumnSource("pcr2", aWB)

    If repeatCol = 0 Then
        
        aWB.Sheets(1).range(Split(Cells(1, repeatLastColumn + 1).Address, "$")(1) & ":" & Split(Cells(1, repeatLastColumn + 1).Address, "$")(1)).EntireColumn.insert
        aWB.Sheets(1).Cells(1, repeatLastColumn + 1) = "REPEAT->"
        
        
        aWB.Sheets(1).range(Split(Cells(1, RanalysisSN).Address, "$")(1) & ":" & Split(Cells(1, RanalysisSN).Address, "$")(1)).Copy _
            range(Split(Cells(1, repeatLastColumn + 2).Address, "$")(1) & ":" & Split(Cells(1, repeatLastColumn + 2).Address, "$")(1))
        
        aWB.Sheets(1).range(Split(Cells(1, RanalysisGT).Address, "$")(1) & ":" & Split(Cells(1, RanalysisGT).Address, "$")(1)).Copy _
            range(Split(Cells(1, repeatLastColumn + 3).Address, "$")(1) & ":" & Split(Cells(1, repeatLastColumn + 3).Address, "$")(1))

            
        If Not RanalysisPCR1 = 0 Then
        
        aWB.Sheets(1).range(Split(Cells(1, RanalysisPCR1).Address, "$")(1) & ":" & Split(Cells(1, RanalysisPCR1).Address, "$")(1)).Copy _
            range(Split(Cells(1, repeatLastColumn + 4).Address, "$")(1) & ":" & Split(Cells(1, repeatLastColumn + 4).Address, "$")(1))
        End If
        
        If Not RanalysisPCR2 = 0 Then
        
        aWB.Sheets(1).range(Split(Cells(1, RanalysisPCR2).Address, "$")(1) & ":" & Split(Cells(1, RanalysisPCR2).Address, "$")(1)).Copy _
            range(Split(Cells(1, repeatLastColumn + 5).Address, "$")(1) & ":" & Split(Cells(1, repeatLastColumn + 5).Address, "$")(1))

        End If
        
        
    End If
    
End Sub


Function verifyRepeatTransfer() As String
        Dim eachSamp As Variant
        
        For Each eachSamp In repeatBook.Sheets(1).range(repeatRange1, repeatRange2)
        
            If repeatBook.Sheets(1).Cells(eachSamp.Row, repeatTran) = "" Then
                verifyRepeatTransfer = verifyRepeatTransfer & eachSamp & "  From " & Left(repeatfile.name, 20) & vbNewLine
            End If
        Next

End Function

Function verificationRepeatTransfer2() As String
        Dim eachSamp As Variant
        
        For Each eachSamp In analysisBook.Sheets(1).range(AnalysisRange1, AnalysisRange2)
            If Not SearchColumnSource("rt", analysisBook) = 0 Then
            
                If Not analysisBook.Sheets(1).Cells(eachSamp.Row, SearchColumnSource("rt", analysisBook)) = "" And _
                Not analysisBook.Sheets(1).Cells(eachSamp.Row, SearchColumnSource("samplename", analysisBook)).Interior.color = vbGreen Then
                
                    verificationRepeatTransfer2 = verificationRepeatTransfer2 & eachSamp & "  From " & Left(analysisFile.name, 20) & vbNewLine
                End If
            End If
            
        Next

End Function

