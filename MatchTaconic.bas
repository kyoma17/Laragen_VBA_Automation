Attribute VB_Name = "MatchTaconic"
'Purpose To Transfer Results from Analysis Sample Sheets to Taconic Result Sample Sheets
'Written by Kenny Ma. Contact: kyoma17@gmail.com Cell number:(626)-246-2233
'Version 2.2 July 27, 2021
Option Explicit

'Global Variables
Dim summaryMessage, TacCAGCheck As String
Dim mainWB          As Sheets
Dim TestWB          As Workbook
Dim READ_Me, TAC_TEST_LIST, SETTINGS As Worksheet
Dim AnalysisSS, ExtractionSS, ResultSS, TestDB As range
Dim extraTestNames, extraStypes, standardtestnames, gelTestName As range

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object

Dim resultHeaderIndex, resultStartIndex  As Integer

'Variables for ResultBook
Dim TargetName, TargetGenotype, TargetPCR1, TargetgmCAG, TargetseqCAG As Integer
'Variables for Analysis Book
Dim SSname, genotype, PCR1, gmcag, seqCAG, stype, tran As Integer

Dim AnalysisLastRow, ResultlastRow, AnalysisLastColumn As Integer
Dim resultRange1, resultRange2, AnalysisRange1, AnalysisRange2 As String

Dim resultfile, analysisFile As Variant
Dim resultBook, analysisBook As Workbook

'Main Script
Public Sub TaconicMatchResults()
    
    turnOff
    
    summaryMessage = "Run Summary:" & vbNewLine
    TacCAGCheck = "Counter Issues:"
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set ResultSS = READ_Me.Cells(14, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    Workbooks.OpenText FileName:=TestDB, Local:=True
    Set TestWB = ActiveWorkbook
    
    Set TAC_TEST_LIST = TestWB.Sheets("TACONIC_TEST_LIST")
    
    'this sets up the List of Extra Assay Tests
    Set extraTestNames = TAC_TEST_LIST.range("A2:A50")
    Set extraStypes = TAC_TEST_LIST.range("B2:B50")
    Set standardtestnames = TAC_TEST_LIST.range("F2:F50")
    Set gelTestName = TAC_TEST_LIST.range("I2:I50")
    
    'Sets up the Global Variables Using the Settings Page
    
    'Seperate the file name into Project Number and Plate Number
    Dim fileSplitName, strainProjectNumber, plateNumber, plainName As Variant
    
    'Target Column Variables
    Dim TargetSampleName, TargetWellPlate, TargetSerialNumber As Variant
    
    'SourcePlate Column Variables
    Dim SourceProjectNumber, SourceAscesioningNumber, SourceSampleName, SourceWellPlate, SourceSerialNumber As Variant
    
    Dim last_row    As Integer
    Dim indexNum    As Integer
    
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
    Set rsFSOFolder = rsFSOLibrary.GetFolder(ResultSS)
    Set rsFSOFile = rsFSOFolder.Files
    
    'Use For Each loop to loop through each file in the folder
    checkAnalysisSSHeaders asFSOFile

    
    For Each resultfile In rsFSOFile
        If InStr(resultfile, ".xlsx") > 0 And InStr(resultfile, "~$") < 1 Then
        
        'SET UP RESULT FILE
        fileSplitName = Split(resultfile.name, "_")        'Splits the Strain and Plate from File name
        plainName = Replace(resultfile.name, ".csv", "")        'Removes .csv from filiename
        
        If InStr(fileSplitName(0), "CHDI") Then
            strainProjectNumber = fileSplitName(1)
            plateNumber = fileSplitName(2)
        Else
            strainProjectNumber = fileSplitName(0)
            plateNumber = fileSplitName(1)
        End If
        
        Workbooks.OpenText FileName:=Replace(resultfile, "\", "\\"), Local:=True
        Set resultBook = ActiveWorkbook
        
        TACcheckTests resultBook
  
       

        
        'Set up the Target Row for copyPaste
        TargetName = SearchColumnTarget("animalid", resultBook)
        TargetGenotype = SearchTestColumn(resultBook)
        TargetPCR1 = SearchColumnTarget("pcr1", resultBook)
        TargetgmCAG = SearchColumnTarget("gmcag1", resultBook)
        TargetseqCAG = SearchColumnTarget("seqcag1", resultBook)
        
        ResultlastRow = resultBook.Sheets(1).Cells(Rows.count, TargetName).End(xlUp).Row
        resultRange1 = Split(Cells(1, TargetName).Address, "$")(1) & 18
        resultRange2 = Split(Cells(1, TargetName).Address, "$")(1) & 18 + ResultlastRow
        
        For Each analysisFile In asFSOFile        'B
            'SET UP Analysis FILE
            If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then        'B1
            Workbooks.OpenText FileName:=analysisFile, Local:=True
            Set analysisBook = ActiveWorkbook
            
            'Set up TRANSFRER column
            addTransferColumn analysisBook
            tran = SearchColumnSource("tran", analysisBook)
            
            'Set up Analysis Book SS Columns.
            SSname = SearchColumnSource("ss", analysisBook)
            If SSname = 0 Then
                SSname = SearchColumnSource("samplename", analysisBook)
            End If
            
            genotype = SearchColumnSource("gt", analysisBook)
            PCR1 = SearchColumnSource("pcr1", analysisBook)
            gmcag = SearchColumnSource("gm", analysisBook)
            seqCAG = SearchColumnSource("seq", analysisBook)
            stype = SearchColumnSource("stype", analysisBook)
            
            If Not SSname = 0 Then
                AnalysisLastRow = analysisBook.Sheets(1).Cells(Rows.count, SSname).End(xlUp).Row - 1
                AnalysisRange1 = Split(Cells(1, SSname).Address, "$")(1) & 2
                AnalysisRange2 = Split(Cells(1, SSname).Address, "$")(1) & 2 + AnalysisLastRow
                
                'Runs Matching Function.
                
                MatchData resultBook, analysisBook
                
                MatchExtra resultBook, analysisBook
                
            End If
            
            analysisBook.Close savechanges:=True
            
        End If
    Next
    
    Dim tempVerify As String
    tempVerify = countCAGsamplesTac(resultBook)
    If Not tempVerify = "" Then TacCAGCheck = TacCAGCheck & vbNewLine & tempVerify
    
    resultBook.Save
    resultBook.Close
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

TACverifyTotalSamples

TestWB.Close savechanges:=False
summaryMessage = summaryMessage & vbNewLine & TACverifyTotalSamples & vbNewLine & TacCAGCheck
CreateTextFileBasic "TacMatch", summaryMessage, READ_Me
MsgBox (summaryMessage)
End Sub

Function searchSection(name, range) As Variant
    Dim search As Variant
    For Each search In range
        If InStr(cleanInput(search), name) > 0 Then
            searchSection = Array(search.Row, search.column)
        End If
    Next
End Function


'Validation of Taconic CAG and Genotyping
Function countCAGsamplesTac(resultBook) As String
    Dim pcrHeader, pcrSample As Variant
    Dim pcrRange, genotype As String
    Dim pcrCounter, totalCounter, homoCounter, hetCounter, carCounter, wtCounter As Integer

    pcrCounter = 0
    For Each pcrHeader In resultBook.Sheets(1).range("A17:Z17")
        If cleanInput(pcrHeader) = "pcr1" Then
     
            pcrRange = Split(Cells(1, pcrHeader.column).Address, "$")(1)
            
            For Each pcrSample In resultBook.Sheets(1).range(pcrRange & 18 & ":" & pcrRange & ResultlastRow)
                'PCR1 counting
                If Not pcrSample = "" Then pcrCounter = pcrCounter + 1
                
                'Genotype Counting
                genotype = cleanInput(resultBook.Sheets(1).Cells(pcrSample.Row, pcrSample.column - 1))
                If genotype = "wt" Then
                    wtCounter = wtCounter + 1
                ElseIf genotype = "het" Then
                    hetCounter = hetCounter + 1
                ElseIf genotype = "homo" Then
                    homoCounter = homoCounter + 1
                ElseIf genotype = "car" Then
                    carCounter = carCounter + 1
                End If
            Next
            
        End If
    Next
    
    Dim testLocation As Variant
    testLocation = searchSection("test", resultBook.Sheets(1).range("C1:G17"))

    If homoCounter + hetCounter + carCounter = pcrCounter Then
        'Skip all good

    Else
        resultBook.Sheets(1).Cells(testLocation(0), testLocation(1)).Interior.color = vbRed
        countCAGsamplesTac = "CAG and GT count issue on " & resultBook.name
    End If
    
    
End Function


'Sorts Analysis workbook by old inex
Sub sortAnalysisBook(aWB)
    Dim oldIndexcol As Integer
    Dim oldIndexAdress As String
    
    oldIndexcol = SearchColumnSource("oldindex", aWB)
    oldIndexAdress = Split(Cells(1, oldIndexcol).Address, "$")(1) & 1
    
    aWB.Sheets(1).range("A1:BB600").Sort Key1:=range(oldIndexAdress), Order1:=xlAscending, header:=xlYes
End Sub

'''''''''''''''''''''''''''''HELPER FUNCTIONS''''''''''''''''''''''''''''''''
Function TACcheckTests(rwb) As Boolean
    TACcheckTests = False
    Dim cell, test As Variant
    Dim notFound As String
    
    For Each cell In rwb.Sheets(1).range("A17", "AZ17").Cells

            For Each test In Union(standardtestnames, standardtestnames, gelTestName)
                If cell = "Index" Or _
                cell = "Animal ID" Or _
                cell = "PCR 1" Or _
                cell = "GM CAG 1" Or _
                cell = "SEQ CAG 1" Or _
                cell = "Plate #" Or _
                cell = "Serial #" Or _
                cell = "Comment" Then
                    'skip
                ElseIf cell = "" Or test = "" Then
                    'SkipMsgBox (cell & test)
                ElseIf cell = test Then
                    TACcheckTests = True
                Else
                    notFound = cell
                End If
            Next

    Next
    If TACcheckTests = False Then
        summaryMessage = summaryMessage & notFound & " from " & resultfile.name & " not found in Test DB" & vbNewLine
    End If
    
End Function


'Will match the extra asccessory tests.
Sub MatchExtra(resultBook, analysisBook)
    'Pulls data
    Dim InsertColumnIndex As Integer
    
    'Matching Loop
    Dim i, SampleIndex, rowdata As Integer
    Dim cell, sample As Variant
    InsertColumnIndex = 1
    
    For Each cell In resultBook.Sheets(1).range("A17", "AZ17").Cells
        
        For i = 1 To 50
            If extraTestNames(i) = "" Then
                'skip
            ElseIf extraTestNames(i) = cell Then
                SampleIndex = 18
                'Loop through samples in ResultSS and match with AnalysisSS rows.
                
                For Each sample In resultBook.Worksheets(1).range(resultRange1, resultRange2)
                    
                    If Not sample = "" Then
                        'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
                        rowdata = SampleNameSearchExtra(sample, analysisBook, extraStypes(i))
                    End If

                    If sample = "" Then
                        'Skip
                    ElseIf Not rowdata = 0 And Not sample = "" Then
                        
                        resultBook.Worksheets(1).Cells(SampleIndex, InsertColumnIndex) = _
                                                                    analysisBook.Worksheets(1).Cells(rowdata, genotype)
                        
                        'Transfer Mark
                        analysisBook.Worksheets(1).Cells(rowdata, tran) = "CP_Tac:"
                        
                    End If
                    
                    SampleIndex = SampleIndex + 1
                Next
            End If
            
        Next
        InsertColumnIndex = InsertColumnIndex + 1
    Next
    
End Sub

Function SampleNameSearchExtra(name, aWB, stype) As Integer
    Dim sampType, Rowindex, columnIndex, sampleColumn As Integer
    
    sampType = SearchColumnSource("stype", aWB)
    sampleColumn = SearchColumnSource("samplename", aWB)
    
    Rowindex = 2
    
    'Change this when doing JaX
    Dim field       As Variant
    
    'Loops through analysis Sheet and searches for Sample Name
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)
        
        If field = name And checkExtraStype(aWB.Sheets(1).Cells(Rowindex, sampType)) Then
            SampleNameSearchExtra = Rowindex
            Exit For
        End If
        Rowindex = Rowindex + 1
    Next
End Function

'Will check if Stype is Not an Extra Test
Function checkExtraStype(stype) As Boolean
    checkExtraStype = False
    Dim cell        As Variant
    For Each cell In extraStypes.Cells
        If LCase(stype) = cell Then
            checkExtraStype = True
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function

'Match Data will Copy results from Analysis file to Result File.
Sub MatchData(resultBook, analysisBook)
    Dim lastRow, Index, rowdata  As Integer
    
    Dim sample      As Variant
    
    Index = 18
    'Loop through samples in ResultSS and match with AnalysisSS rows.
    'For Each Sample In resultBook.Worksheets(1).Range("B18", "B114") 'Fix this to make dynamic
    
    For Each sample In resultBook.Worksheets(1).range(resultRange1, resultRange2)
        If Not sample = "" Then
        'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
            rowdata = SampleNameSearch(sample, analysisBook)        'Grabs the Row from SS sheet.
        End If
        
        
        If sample = "" Then
            'Skip
        ElseIf Not rowdata = 0 And Not TargetPCR1 = 0 And Not PCR1 = 0 Then
            resultBook.Worksheets(1).Cells(Index, TargetGenotype) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, genotype)
            resultBook.Worksheets(1).Cells(Index, TargetPCR1) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, PCR1)
            resultBook.Worksheets(1).Cells(Index, TargetgmCAG) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, gmcag)
            resultBook.Worksheets(1).Cells(Index, TargetseqCAG) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, seqCAG)
            
            'Transfer Mark
            analysisBook.Worksheets(1).Cells(rowdata, tran) = "CP_Tac"
        End If
        
        Index = Index + 1
        
    Next
    
End Sub

'Function to search for Standard qPCR Laragen Tests.
Function SearchTestColumn(wb) As Integer
    Dim columnIndex As Integer
    columnIndex = 1
    
    Dim found       As Boolean
    found = False
    
    Dim field, test As Variant
    
    For Each field In wb.Sheets(1).range("A17", "K17")        'The Template rows has to start at 17
        For Each test In standardtestnames.Cells
            If test = field Then
                SearchTestColumn = columnIndex
                found = True
                Exit For
            End If
        Next
        If found Then
            Exit For
        End If
        columnIndex = columnIndex + 1
    Next
End Function

'Returns the Row number in the Analysis File matching the Name. Will Ignore Extra Tests Stype.
Function SampleNameSearch(name, aWB) As Integer
    
    Dim lastRow, Rowindex, sampleColumn, sampType, columnIndex As Integer
    Dim field       As Variant
    
    lastRow = aWB.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 2        'Use this to stop the loop early.
    
    sampleColumn = SearchColumnSource("Sample Name", aWB)
    sampType = SearchColumnSource("stype", aWB)
    
    columnIndex = 2
    
    'Change this when doing JaX
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)        ' Make This dynamic, not SS for now
        
        If field = name And checkStype(aWB.Sheets(1).Cells(columnIndex, sampType)) Then
            SampleNameSearch = columnIndex
            Exit For
        End If
        columnIndex = columnIndex + 1
    Next
    
End Function

'Will check if Stype is Not an Extra Test (Final)
Function checkStype(stype) As Boolean
    checkStype = True
    Dim cell        As Variant
    For Each cell In extraStypes.Cells
        If LCase(stype) = cell Then
            checkStype = False
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function

'This function will look for the names in the Target file and return the index of that column
Function SearchColumnTarget(search, wb) As Integer
    Dim columnIndex As Integer
    Dim field       As Variant
    columnIndex = 1
    For Each field In wb.Sheets(1).range("A17", "K17")        'The Template rows has to start at 17
        If Replace(LCase(field), " ", "") = search Then
            SearchColumnTarget = columnIndex
        End If
        columnIndex = columnIndex + 1
    Next
End Function




