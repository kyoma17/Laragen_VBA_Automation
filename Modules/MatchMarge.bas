Attribute VB_Name = "MatchMarge"
'Purpose To Convert the Tactonic samplesheets into the Laragen Result sheets.
'Written by Kenny Ma. Contact: kyoma17@gmail.com Cell number:(626)-246-2233
'Version 1.3 July 21, 2021
Option Explicit

'Global Variables
Dim summaryMessage, MargeCAGcheck  As String
Dim mainWB          As Sheets
Dim TestWB          As Workbook
Dim READ_Me, TAC_TEST_LIST, MARGE_TEST_LIST, SETTINGS, COPY_PASTE As Worksheet
Dim AnalysisSS, ExtractionSS, ResultSS, TestDB As range
Dim extraTestNames, extraStypes, standardtestnames, MARGETests, MARGEStypes, MARGEExtraTests, MARGEExtraStypes, margeGelTests _
    , laragenTestHeader As range
Dim replaceMARGEHeader, sampleCounterMARGE As range

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object

Dim resultHeaderIndex, resultStartIndex  As Integer

'Variables for ResultBook
Dim TargetName, TargetIndex, TargetGenotype, TargetPCR1, TargetgmCAG, TargetseqCAG, TargetPCR2, TargetgmCAG2 As Integer
'Variables for Analysis Book
Dim SSname, genotype, PCR1, gmcag, seqCAG, stype, pcr2, gmcag2, tran As Integer

Dim AnalysisLastRow, ResultlastRow As Integer
Dim resultRange1, resultRange2, AnalysisRange1, AnalysisRange2, fileExSplitName As String

Dim reqNumber       As String


Dim resultfile, analysisFile As Variant
Dim resultBook, analysisBook, exBook As Workbook

'Result Book Variables
Dim ResultSSRangeMergeCheck, ResultSSRangeTestName, ResultSSRangeHeader As range
Dim ResultSSSampleRow, ResultSSHeaderRow, ResultSSTestRow As Integer

'Main Script
Public Sub MargeMatchResults()
    turnOff
    
    summaryMessage = "Summary:" & vbNewLine
    MargeCAGcheck = "Counter Issues:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set ResultSS = READ_Me.Cells(16, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    Workbooks.OpenText FileName:=TestDB, Local:=True
    Set TestWB = ActiveWorkbook
    
    Set MARGE_TEST_LIST = TestWB.Sheets("MARGE_TEST_LIST")
    Set TAC_TEST_LIST = TestWB.Sheets("TACONIC_TEST_LIST")
    Set COPY_PASTE = TestWB.Sheets("COPY_PASTE")
    
    'this sets up the List of Extra Assay Tests
    Set extraTestNames = TAC_TEST_LIST.range("A2:A50")
    Set extraStypes = TAC_TEST_LIST.range("B2:B50")
    Set standardtestnames = TAC_TEST_LIST.range("F2:F50")
    Set MARGETests = MARGE_TEST_LIST.range("A2:A50")
    Set MARGEStypes = MARGE_TEST_LIST.range("B2:B50")
    Set MARGEExtraTests = MARGE_TEST_LIST.range("D2:D50")
    Set MARGEExtraStypes = MARGE_TEST_LIST.range("E2:E50")
    Set margeGelTests = MARGE_TEST_LIST.range("G2:G50")
    Set laragenTestHeader = COPY_PASTE.range("B1:E1")
    Set sampleCounterMARGE = COPY_PASTE.range("B11:C17")
    
    'Sets up the Global Variables Using the Settings Page
    
    'Seperate the file name into Project Number and Plate Number
    Dim fileSplitName, laragenJoinedName, strainProjectNumber, plateNumber, plainName As Variant
    
    'Target Column Variables
    Dim TargetSampleName, TargetWellPlate, TargetSerialNumber As Variant
    
    'SourcePlate Column Variables
    Dim SourceProjectNumber, SourceAscesioningNumber, SourceSampleName, SourceWellPlate, SourceSerialNumber As Variant
    
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
    Set rsFSOFolder = rsFSOLibrary.GetFolder(ResultSS)
    Set rsFSOFile = rsFSOFolder.Files
    
    'Use For Each loop to loop through each file in the folder

    checkAnalysisSSHeaders asFSOFile
    
    For Each resultfile In rsFSOFile
        If InStr(resultfile, ".xlsx") > 0 And InStr(resultfile, "~$") < 1 And InStr(resultfile, "desktop.ini") < 1 Then
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
            
            Workbooks.OpenText FileName:=resultfile, Local:=True
            Set resultBook = ActiveWorkbook
            
            
            'Result SS Row Index range and where the Sample Table starts
            ResultSSSampleRow = FindFirstHashRow(resultBook) + 2
            ResultSSHeaderRow = ResultSSSampleRow - 2
            ResultSSTestRow = ResultSSSampleRow - 3
        
           
            Set ResultSSRangeMergeCheck = range("A" & ResultSSSampleRow - 4 & ":AZ" & ResultSSSampleRow - 4)
            
            Debug.Print (ResultSSRangeMergeCheck)
            
            Set ResultSSRangeTestName = range("A" & ResultSSTestRow & ":AZ" & ResultSSTestRow)
            Set ResultSSRangeHeader = range("A" & ResultSSHeaderRow & ":AZ" & ResultSSHeaderRow)
            
            MARGEcheckTests resultBook

            'Set up the Target Row for copyPaste
            TargetName = SearchColumnTarget("#", resultBook)
            ResultlastRow = resultBook.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row
            
            'Sets up the header and the formula
            
            resultTextLocation = SearchResultColumnTestMARGE(resultBook)
            
            'Set up for Single or Double CAG Peaks
            setUpResultBook resultBook, resultTextLocation
            
            reqNumber = Split(resultBook.Sheets(1).Cells(3, 1), " ")(2)
            
            'Set up the Target Row for copyPaste
            TargetName = SearchColumnTarget("mouseid", resultBook)
            TargetIndex = SearchColumnTargetFirst("#", resultBook)
            TargetGenotype = SearchColumnTarget("gt", resultBook)
            TargetPCR1 = SearchColumnTarget("pcr1", resultBook)
            TargetgmCAG = SearchColumnTarget("gmcag1", resultBook)
            
            TargetPCR2 = SearchColumnTarget("pcr2", resultBook)
            TargetgmCAG2 = SearchColumnTarget("gmcag2", resultBook)
            
            ResultlastRow = resultBook.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row
            resultRange1 = Split(Cells(1, TargetIndex).Address, "$")(1) & ResultSSSampleRow
            resultRange2 = Split(Cells(1, TargetIndex).Address, "$")(1) & ResultSSSampleRow + ResultlastRow
            
            Debug.Print (resultRange1)
            Debug.Print (resultRange2)
                    
    
            Dim exFile, cell As Variant
            
            For Each exFile In esFSOFile
                'Set up Extraction Files
                If InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 Then
                    Workbooks.OpenText FileName:=exFile, Local:=True
                    Set exBook = ActiveWorkbook
                    
                    fileExSplitName = Replace(Split(exFile.name, "-")(1), ".xlsx", "")
                    
                    'Req Number Matching
                    For Each cell In Union(exBook.Sheets(1).range("A11", "M12"), exBook.Sheets(1).range("N1", "P13"))        ' Need to add the other stray ones on the right side as well
                        
                        If cell = "" Then
                            'Skip
                
                        ElseIf InStr(cell, reqNumber) > 0 Then
                            For Each analysisFile In asFSOFile
                                If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then
                                    'SET UP Analysis FILE
                                    Workbooks.OpenText FileName:=Replace(analysisFile, "\", "\\"), Local:=True
                                    Set analysisBook = ActiveWorkbook
                                    
                                    'Set up TRANSFRER column
                                    addTransferColumn analysisBook
                                    tran = SearchColumnSource("tran", analysisBook)
                                    
                                    'Set up Analysis Book SS Columns.
                                    SSname = SearchColumnSource("samplename", analysisBook)
                                    genotype = SearchColumnSource("gt", analysisBook)
                                    PCR1 = SearchColumnSource("pcr1", analysisBook)
                                    gmcag = SearchColumnSource("gm", analysisBook)
                                    stype = SearchColumnSource("stype", analysisBook)
                                    pcr2 = SearchColumnSource("pcr2", analysisBook)
                                    gmcag2 = SearchColumnSource("gm2", analysisBook)
                                    
                                    
                                    
                                    
                                    
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
                            Exit For
                            
                        End If
                    Next
                    exBook.Close savechanges:=False
                End If
                
            Next
            
            'Validation of CAG and Genotyping
            Dim tempVerify As String
            tempVerify = countCAGsamples(resultBook, resultTextLocation - 1)
            If Not tempVerify = "" Then MargeCAGcheck = MargeCAGcheck & vbNewLine & tempVerify
            
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
    
    TestWB.Close
    
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    summaryMessage = summaryMessage & vbNewLine & MargeVerifyTotalSamples & MargeCAGcheck
    
    CreateTextFileBasic "MargeMatch", summaryMessage, READ_Me
    MsgBox (summaryMessage)
End Sub

'Validation of CAG and Genotyping
Function countCAGsamples(resultBook, resultTextLocation) As String
    Dim pcrHeader, pcrSample As Variant
    Dim pcrRange As String
    Dim pcrCounter, total, homo, het, car, wt, cag As Integer
    
    total = 3
    car = 4
    homo = 5
    het = 6
    wt = 7
    cag = 9
    
    pcrCounter = 0
    For Each pcrHeader In resultBook.Sheets(1).range(ResultSSRangeHeader.Address)
        If pcrHeader = "PCR1" Then
            pcrRange = Split(Cells(1, pcrHeader.column).Address, "$")(1)
            
            For Each pcrSample In resultBook.Sheets(1).range(pcrRange & ResultSSSampleRow & ":" & pcrRange & ResultlastRow)
                If Not pcrSample = "" Then pcrCounter = pcrCounter + 1
            Next
            
        End If
    Next
    
    If resultBook.Sheets(1).Cells(cag, resultTextLocation + 4) = pcrCounter Then
        'Skip all good
    Else
        resultBook.Sheets(1).Cells(cag, resultTextLocation + 4).Interior.color = vbRed
        countCAGsamples = "CAG and GT count issue on " & resultBook.name
    End If
        
End Function


Function MARGEcheckTests(rwb) As Boolean
    MARGEcheckTests = False
    Dim cell, test As Variant
    Dim notFound As String
    
    For Each cell In rwb.Sheets(1).range(ResultSSRangeTestName.Address).Cells

            For Each test In Union(MARGETests, MARGEExtraTests, margeGelTests)
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
                    MARGEcheckTests = True
                Else
                    notFound = cell
                End If
            Next

    Next
    If MARGEcheckTests = False Then
        summaryMessage = summaryMessage & notFound & " from " & resultfile.name & " not found in Test DB" & vbNewLine
    End If
End Function

'Will match the extra asccessory tests.
Sub MatchExtra(resultBook, analysisBook)
    'Pulls data
    Dim InsertColumnIndex, replaceColumnIndex As Integer
    
    'Matching Loop
    Dim i, SampleIndex, rowdata As Integer
    Dim cell, sample As Variant
    InsertColumnIndex = 1
    
    For Each cell In resultBook.Sheets(1).range(ResultSSRangeTestName.Address).Cells
        
        For i = 1 To 50
            If MARGEExtraTests(i) = "" Then
                
                Exit For
            ElseIf MARGEExtraTests(i) = cell Then
                
                If resultBook.Sheets(1).Cells(ResultSSHeaderRow, InsertColumnIndex + 3) = "GT" Then
                    replaceColumnIndex = SearchColumnTarget("#", resultBook) + 1
                    
                End If
                
                replaceColumnIndex = InsertColumnIndex
                
                resultBook.Sheets(1).Cells(ResultSSHeaderRow, replaceColumnIndex + 3) = MARGEExtraStypes(i) & " GT"
                
                SampleIndex = ResultSSSampleRow
                'Loop through samples in ResultSS and match with AnalysisSS rows.
                
                For Each sample In resultBook.Worksheets(1).range(resultRange1, resultRange2)
                    
                    If Not sample = "" Then
                        
                        'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
                        
                        rowdata = SampleNameSearchExtra(sample, analysisBook, MARGEExtraStypes(i))        ''Issues in ehre
                        
                        If Not rowdata = 0 Then
                            
                            resultBook.Worksheets(1).Cells(SampleIndex, replaceColumnIndex + 3) = _
                                                                        analysisBook.Worksheets(1).Cells(rowdata, genotype)
                            
                            'Transfer Mark
                            analysisBook.Worksheets(1).Cells(rowdata, tran) = "CP_Marge"
                        End If
                    End If
                    
                    SampleIndex = SampleIndex + 1
                Next
            End If
            
        Next
        InsertColumnIndex = InsertColumnIndex + 1
    Next
    
End Sub

Public Function SampleNameSearchExtra(name, aWB, stype) As Integer
    'MsgBox (name & stype)
    Dim lastRow, Rowindex, sampleColumn, sampType, columnIndex As Integer
    Dim comboName   As String
    Dim field       As Variant
    
    lastRow = aWB.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 2        'Use this to stop the loop early.
    
    sampleColumn = SearchColumnSource("samplename", aWB)
    sampType = SearchColumnSource("stype", aWB)
    
    Rowindex = 2
    
    comboName = fileExSplitName & "$" & reqNumber & "-" & name
    
    'Change this when doing MARGe
    
    'Loops through analysis Sheet and searches for Sample Name
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)
        If field = "" Then
            'Skip
        ElseIf field = comboName And checkExtraStype(aWB.Sheets(1).Cells(Rowindex, sampType)) Then
            SampleNameSearchExtra = Rowindex
            Debug.Print ("matched" & Rowindex)
            
            Exit For
        End If
        Rowindex = Rowindex + 1
    Next
End Function
'Returns the Row number in the Analysis File matching the Name. Will Ignore Extra Tests Stype.
Function SampleNameSearch(name, rwb, aWB) As Integer
    
    Dim lastRow, Rowindex, sampleColumn, sampType, columnIndex As Integer
    Dim comboName   As String
    Dim field       As Variant
    
    lastRow = aWB.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 2        'Use this to stop the loop early.
    
    sampleColumn = SearchColumnSource("samplename", aWB)
    sampType = SearchColumnSource("stype", aWB)
    
    Rowindex = 2
    
    comboName = fileExSplitName & "$" & reqNumber & "-" & name
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)        ' Make This dynamic, not SS for now
        
        If field = "" Then
            'Skip
        ElseIf field = comboName And checkStype(aWB.Sheets(1).Cells(Rowindex, sampType)) Then
            SampleNameSearch = Rowindex
            Exit For
        End If
        Rowindex = Rowindex + 1
    Next
    
End Function

'Match Data will Copy results from Analysis file to Result File.
Sub MatchData(resultBook, analysisBook)
    Dim lastRow, Index, rowdata  As Integer
    
    Dim sample      As Variant
    
    lastRow = SearchColumnTarget("#", resultBook)
    
    Index = ResultSSSampleRow
    'Loop through samples in ResultSS and match with AnalysisSS rows.
    'For Each Sample In resultBook.Worksheets(1).Range("B18", "B114") 'Fix this to make dynamic

    
    
    
    For Each sample In resultBook.Worksheets(1).range(resultRange1, resultRange2)
        Debug.Print (sample)
        If sample = "" Then
            Exit For
        End If
        
        'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
        rowdata = SampleNameSearch(sample, resultBook, analysisBook)        'Grabs the Row from SS sheet.
        
        If TargetGenotype = 0 Then
            'Skip
            Debug.Print ("skip")
            
        ElseIf Not rowdata = 0 Then
            resultBook.Worksheets(1).Cells(Index, TargetGenotype) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, genotype)
            resultBook.Worksheets(1).Cells(Index, TargetPCR1) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, PCR1)
            resultBook.Worksheets(1).Cells(Index, TargetgmCAG) = _
                                                  analysisBook.Worksheets(1).Cells(rowdata, gmcag)
            
            If Not pcr2 = 0 And Not gmcag2 = 0 Then
                
                resultBook.Worksheets(1).Cells(Index, lastRow + 1) = _
                                                      analysisBook.Worksheets(1).Cells(rowdata, pcr2)
                resultBook.Worksheets(1).Cells(Index, lastRow + 2) = _
                                                      analysisBook.Worksheets(1).Cells(rowdata, gmcag2)
                
            End If
            'Transfer Mark
            analysisBook.Worksheets(1).Cells(rowdata, tran) = "CP_MARGE"
        End If
        
        Index = Index + 1
        
    Next
    
End Sub

'This function will look for the names in the Target file and return the index of that column (LAST)
Public Function SearchColumnTarget(search, wb) As Integer
    Dim columnIndex As Integer
    Dim field       As Variant
    columnIndex = 1
    For Each field In wb.Sheets(1).range(ResultSSRangeHeader.Address)        'The Template rows has to start at 18
        If Replace(LCase(field), " ", "") = search Then
            SearchColumnTarget = columnIndex
        End If
        columnIndex = columnIndex + 1
    Next
End Function

'This function will look for the names in the Target file and return the index of that column (FIRST)
Public Function SearchColumnTargetFirst(search, wb) As Integer
    Dim columnIndex As Integer
    Dim field As Variant
    
    If search = "" Then
        SearchColumnTargetFirst = "Search string is empty."
        Exit Function
    End If
    
    columnIndex = 1
    
    For Each field In wb.Sheets(1).range(ResultSSRangeHeader.Address)
        If Replace(LCase(field), " ", "") = LCase(Replace(search, " ", "")) Then
            SearchColumnTargetFirst = columnIndex
            Exit Function
        End If
        columnIndex = columnIndex + 1
    Next field
    
End Function

'Changes the headers on the resultbooks, Also adds the counter formula
Public Sub setUpResultBook(wb, resultTextLocation)
    
    If Not resultTextLocation = 0 Then
        
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation) = "PCR1"
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation + 1) = "GM CAG 1"
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation + 2) = "GT"
        
        sampleCounterMARGE.Copy wb.Worksheets(1).Cells(3, resultTextLocation + 2)
        
    End If
    
End Sub

Public Sub setUpResultBooksDouble(wb, resultTextLocation)
    
    If Not resultTextLocation = 0 Then
        
    End If
    
End Sub

Public Function SearchResultColumnTestMARGE(wb) As Integer
    Dim columnIndex As Integer
    Dim field, test As Variant
    Dim found       As Boolean
    found = False
    columnIndex = 0
    
    For Each field In wb.Sheets(1).range(ResultSSRangeTestName.Address)        'The Template rows has to start at 19 for MARGE
        If field = "" Then
            columnIndex = columnIndex + 1
        Else
            columnIndex = columnIndex + 1
            For Each test In MARGETests
                If test = field Then
                    SearchResultColumnTestMARGE = columnIndex
                    found = True
                    Exit For
                End If
            Next
            
        End If
        
        If found Then
            Exit For
        End If
        
    Next
    
End Function

'''''''''''''''''''''''''''''HELPER FUNCTIONS''''''''''''''''''''''''''''''''

'Will check if Stype is Not an Extra Test
Function checkExtraStype(stype) As Boolean
    checkExtraStype = False
    Dim cell        As Variant
    For Each cell In MARGEExtraStypes.Cells
        If LCase(stype) = cell Then
            checkExtraStype = True
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function

'Function to search for Standard qPCR Laragen Tests.
Function SearchTestColumn(wb) As Integer
    Dim columnIndex As Integer
    columnIndex = 1
    
    Dim found       As Boolean
    found = False
    
    Dim field, test As Variant
    
    For Each field In wb.Sheets(1).range("A17", "K17")        'The Template rows has to start at 17
        For Each test In standardtestnames.Cells
            If test = Replace(LCase(field), " ", "") Then
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

'Will check if Stype is Not an Extra Test (Final)
Function checkStype(stype) As Boolean
    checkStype = True
    Dim cell        As Variant
    For Each cell In MARGEExtraStypes.Cells
        If LCase(stype) = cell Then
            checkStype = False
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function



