Attribute VB_Name = "MatchJAX"
'Purpose To Match and Copy Samples from AnalysisSS to JAX Result Paperwork.
'Written by Kenny Ma. Contact: kyoma17@gmail.com Cell number:(626)-246-2233
'Version 2.3 August 2, 2021
Option Explicit

'Global Variables
Dim summaryMessage, JaxCAGCheck  As String
Dim mainWB          As Sheets
Dim TestWB          As Workbook
Dim READ_Me, TAC_TEST_LIST, JAX_TEST_LIST, SETTINGS, COPY_PASTE As Worksheet

Dim AnalysisSS, ExtractionSS, ResultSS, TestDB As range
Dim extraTestNames, extraStypes, standardtestnames, jaxTests, jaxStypes, jaxExtraTests, jaxExtraStypes, jaxGelTests _
    , laragenTestHeader As range
Dim replaceJaxHeader, sampleCounterJax As range

Dim asFSOLibrary, esFSOLibrary, rsFSOLibrary As FileSystemObject
Dim asFSOFolder, asFSOFile, esFSOFolder, esFSOFile, rsFSOFolder, rsFSOFile  As Object

Dim resultHeaderIndex, resultStartIndex  As Integer

'Variables for ResultBook
Dim TargetName, TargetGenotype, TargetPCR1, TargetgmCAG, TargetseqCAG As Integer
Dim TargetPCR2, TargetgmCAG2, TargetseqCAG2 As Integer

'Variables for Analysis Book
Dim SSname, genotype, PCR1, gmcag, seqCAG, stype, pcr2, seqcag2, gmcag2, tran As Integer

Dim AnalysisLastRow, ResultlastRow, AnalysisLastColumn As Integer
Dim resultRange1, resultRange2, AnalysisRange1, AnalysisRange2, fileExSplitName As String

Dim resultfile, analysisFile As Variant

'Result Book Variables
Dim ResultSSRangeMergeCheck, ResultSSRangeTestName, ResultSSRangeHeader As range
Dim ResultSSSampleRow, ResultSSHeaderRow, ResultSSTestRow As Integer
    

'Main Script
Public Sub JaxMatchResults()
    turnOff
    
    summaryMessage = "Run Summary:" & vbNewLine
    JaxCAGCheck = "Counter Issues:" & vbNewLine
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set ExtractionSS = READ_Me.Cells(13, 2)
    Set ResultSS = READ_Me.Cells(15, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    Workbooks.OpenText FileName:=TestDB, Local:=True
    Set TestWB = ActiveWorkbook
    
    Set JAX_TEST_LIST = TestWB.Sheets("JAX_TEST_LIST")
    Set TAC_TEST_LIST = TestWB.Sheets("TACONIC_TEST_LIST")
    Set COPY_PASTE = TestWB.Sheets("COPY_PASTE")
    
    'this sets up the List of Extra Assay Tests
    Set extraTestNames = TAC_TEST_LIST.range("A2:A50")
    Set extraStypes = TAC_TEST_LIST.range("B2:B50")
    Set standardtestnames = TAC_TEST_LIST.range("F2:F50")
    Set jaxTests = JAX_TEST_LIST.range("A2:A50")
    Set jaxStypes = JAX_TEST_LIST.range("B2:B50")
    Set jaxExtraTests = JAX_TEST_LIST.range("D2:D50")
    Set jaxExtraStypes = JAX_TEST_LIST.range("E2:E50")
    Set jaxGelTests = JAX_TEST_LIST.range("G2:G50")
    Set laragenTestHeader = COPY_PASTE.range("B1:E1")
    Set sampleCounterJax = COPY_PASTE.range("B11:C17")
    
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
    
    
    
    Dim resultBook, analysisBook, exBook As Workbook
    Dim reqNumber   As String
    
    For Each resultfile In rsFSOFile

        If InStr(resultfile, "~") < 1 And InStr(resultfile, ".xlsx") > 0 And InStr(resultfile, "~$") < 1 And InStr(resultfile, "desktop.ini") < 1 Then
            
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
            
            'Debug.Print ("result sample row" & ResultSSSampleRow)
            ResultSSHeaderRow = ResultSSSampleRow - 2
            ResultSSTestRow = ResultSSSampleRow - 3
        
           
            Set ResultSSRangeMergeCheck = range("A" & ResultSSSampleRow - 4 & ":AZ" & ResultSSSampleRow - 4)
            
            Debug.Print ResultSSRangeMergeCheck.Address
            
            Set ResultSSRangeTestName = range("A" & ResultSSTestRow & ":AZ" & ResultSSTestRow)
            Set ResultSSRangeHeader = range("A" & ResultSSHeaderRow & ":AZ" & ResultSSHeaderRow)

            
            JAXcheckTests resultBook
            If JAXCheckFormatting(resultBook) Then
                'summaryMessage = summaryMessage & "Formatting Issue with " & resultfile.name & vbCrLf
            End If
             
            
            
            'Set up the Target Row for copyPaste
            TargetName = SearchColumnTarget("#", resultBook)
            ResultlastRow = resultBook.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 1
            
            'Sets up the header and the formula sections
            
            resultTextLocation = SearchResultColumnTestJax(resultBook)
            
            setUpResultBooks resultBook, resultTextLocation
            Debug.Print (resultBook.Sheets(1).Cells(3, 1))
            
            Dim cellValue As String
            cellValue = resultBook.Sheets(1).Cells(3, 1).Value ' Replace with your cell reference
    
            reqNumber = ExtractReqNumber(cellValue)
            
            Debug.Print (reqNumber)
                    
            
            
            'Set up the Target Row for copyPaste
            TargetName = SearchColumnTarget("mouseid", resultBook)
            
            'Debug.Print (TargetName)
            'Debug.Print (reqNumber)
            
            
            TargetGenotype = SearchColumnTarget("gt", resultBook)
            TargetPCR1 = SearchColumnTarget("pcr1", resultBook)
            TargetgmCAG = SearchColumnTarget("gmcag1", resultBook)
            TargetseqCAG = SearchColumnTarget("seqcag1", resultBook)
            
            TargetPCR2 = SearchColumnTarget("pcr2", resultBook)
            TargetgmCAG2 = SearchColumnTarget("gmcag2", resultBook)
            TargetseqCAG2 = SearchColumnTarget("seqcag2", resultBook)
            
            resultRange1 = Split(Cells(1, TargetName).Address, "$")(1) & ResultSSSampleRow
            resultRange2 = Split(Cells(1, TargetName).Address, "$")(1) & ResultlastRow
            
            If checkResultDuplicates3(resultBook) Then
                summaryMessage = summaryMessage + "Duplicates in " & resultfile.name & ". Skipped" & vbNewLine
            Else
                
                Dim exFile, cell As Variant
                
                For Each exFile In esFSOFile
                    'Set up Extraction Files
                    If InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 Then
                        
                        Workbooks.OpenText FileName:=exFile, Local:=True
                        Set exBook = ActiveWorkbook
                        
                        fileExSplitName = Replace(Split(exFile.name, "-")(1), ".xlsx", "")
                        
                        'Req Number Matching
                        For Each cell In Union(exBook.Sheets(1).range("A11", "M12"), exBook.Sheets(1).range("N1", "P13"))
                            If cell = "" Then
                                
                            ElseIf InStr(cell, reqNumber) > 0 Then
                                For Each analysisFile In asFSOFile
                                    'SET UP Analysis FILE
                                    If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then
                                        
                                        Workbooks.OpenText FileName:=analysisFile, Local:=True
                                        Set analysisBook = ActiveWorkbook
                                        
                                        'Set up TRANSFRER column
                                        addTransferColumn analysisBook
                                        tran = SearchColumnSource("tran", analysisBook)
                                        
                                        'Set up Analysis Book SS Columns.
                                        SSname = SearchColumnSource("samplename", analysisBook)
                                        genotype = SearchColumnSource("gt", analysisBook)
                                        PCR1 = SearchColumnSource("pcr1", analysisBook)
                                        gmcag = SearchColumnSource("gm", analysisBook)
                                        seqCAG = SearchColumnSource("seq", analysisBook)
                                        stype = SearchColumnSource("stype", analysisBook)
                                        
                                        pcr2 = SearchColumnSource("pcr2", analysisBook)
                                        gmcag2 = SearchColumnSource("gm2", analysisBook)
                                        seqcag2 = SearchColumnSource("seq2", analysisBook)
                                        
                                        
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
            End If
            'counts PCR1 results
            'Validation of CAG and Genotyping
            
            Dim tempVerify As String
            tempVerify = countCAGsamples(resultBook, resultTextLocation)
            If Not tempVerify = "" Then JaxCAGCheck = JaxCAGCheck & vbNewLine & tempVerify
            
            resultBook.Save
            resultBook.Close
            
        End If
    Next
    
    TestWB.Close
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
    
    summaryMessage = summaryMessage & vbNewLine & JAXverifyTotalSamples & vbNewLine & JaxCAGCheck
    
    CreateTextFileBasic "JaxMatch", summaryMessage, READ_Me
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


Function JAXCheckFormatting(rwb) As Boolean
    JAXCheckFormatting = False
    Dim items As Variant
    For Each items In rwb.Sheets(1).range(ResultSSRangeMergeCheck.Address)
        'Debug.Print (items)
        If items.MergeCells Then
            JAXCheckFormatting = True
        End If
    Next
End Function

Function JAXcheckTests(rwb) As Boolean
    JAXcheckTests = False
    Dim cell, test As Variant
    Dim notFound As String
    
    For Each cell In rwb.Sheets(1).range(ResultSSRangeTestName.Address).Cells

            For Each test In Union(jaxTests, jaxExtraTests, jaxGelTests)
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
                    JAXcheckTests = True

                Else
                    notFound = cell
                End If
            Next

    Next
    If JAXcheckTests = False Then
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
    replaceColumnIndex = 1
    
    For Each cell In resultBook.Sheets(1).range(ResultSSRangeTestName.Address).Cells
        
        For i = 1 To 50
            If jaxExtraTests(i) = "" Or cell = "" Then
                'skip
            
            ElseIf jaxExtraTests(i) = cell Then
                
                If resultBook.Sheets(1).Cells(ResultSSHeaderRow, InsertColumnIndex + 3) = "GT" Then
                    replaceColumnIndex = SearchColumnTarget("#", resultBook) + 1
                Else
                    replaceColumnIndex = InsertColumnIndex
                End If
                
                resultBook.Sheets(1).Cells(ResultSSHeaderRow, replaceColumnIndex + 3) = jaxExtraStypes(i) & " GT"
                
                SampleIndex = ResultSSSampleRow
                'Loop through samples in ResultSS and match with AnalysisSS rows.
                
                For Each sample In resultBook.Worksheets(1).range(resultRange1, resultRange2)
                    
                    
                    If Not sample = "" Then
                        
                        'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
                        
                        rowdata = SampleNameSearchExtra(sample, analysisBook, jaxExtraStypes(i))        ''Issues in ehre
                        
                        If Not rowdata = 0 Then
                            
                            resultBook.Worksheets(1).Cells(SampleIndex, replaceColumnIndex + 3) = _
                                                                        analysisBook.Worksheets(1).Cells(rowdata, genotype)
                            
                            'Transfer Mark
                            analysisBook.Worksheets(1).Cells(rowdata, tran) = "CP_JAX"
                            
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
    Dim sampType, Rowindex, columnIndex, sampleColumn As Integer
    Dim comboName   As String
    
    sampType = SearchColumnSource("stype", aWB)
    sampleColumn = SearchColumnSource("samplename", aWB)
    
    Rowindex = 2
    
    'Change this when doing JaX
    Dim field       As Variant
    comboName = fileExSplitName & "$" & name
    
    'Loops through analysis Sheet and searches for Sample Name
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)

        If comboName = field Then

        End If

        
        If field = comboName And checkExtraStype(aWB.Sheets(1).Cells(Rowindex, sampType)) Then
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
    For Each cell In jaxExtraStypes.Cells
        If Replace(LCase(stype), " ", "") = Replace(LCase(cell), " ", "") Then
            checkExtraStype = True
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function

'if there is stuff on the right side of the result sheet such as PCR2, create extra columns and move data.
Sub movePCR2(rwb)
    Dim lastinrow   As Integer
    lasltinrow = searchTargetColumn("#", rwb)
    
End Sub

Sub deletePCR2(rwb, resultLocation)
    Dim pcr2Column  As Integer
    pcr2Column = SearchColumnTarget("pcr2", rwb)
    
    Dim pcr2range1, pcr2range2 As String
    Dim Rng         As range
    Dim myArray     As Variant
    
    pcr2range1 = Split(Cells(1, pcr2Column).Address, "$")(1) & ResultSSSampleRow
    pcr2range2 = Split(Cells(1, pcr2Column).Address, "$")(1) & ResultlastRow
    
    Dim deleteColumn As String
    
    Set Rng = rwb.Sheets(1).range(pcr2range1, pcr2range2)
    myArray = WorksheetFunction.Unique(Rng)
    
    If Not pcr2Column = 0 Then
        If UBound(myArray) <= 1 Then
            
            deleteColumn = Split(Cells(1, pcr2Column).Address, "$")(1)
            
            rwb.Sheets(1).Columns(deleteColumn).Delete
            rwb.Sheets(1).Columns(deleteColumn).Delete
            rwb.Sheets(1).Columns(deleteColumn).Delete
            
        End If
        
    End If
End Sub

'Willcheck for for duplicate sample names

Function checkResultDuplicates3(rwb) As Boolean
    Dim range_1     As range
    Dim array_1     As Variant
    Set range_1 = rwb.Sheets(1).range(resultRange1, resultRange2)
    checkResultDuplicates3 = False
    
    Dim sID, sID2   As Variant
    
    Dim startindex, count As Integer
    startindex = range_1(0).Row
    For Each sID In range_1
        count = 0
        If checkResultDuplicates3 Then
            Exit For
        Else
            For Each sID2 In range_1
                If checkResultDuplicates3 = True Then
                    Exit For
                ElseIf sID = sID2 Then
                    count = count + 1
                    If count > 1 Then
                        checkResultDuplicates3 = True
                        
                    End If
                End If
            Next
        End If
    Next
End Function

Function checkResultDuplicates2(rwb) As Boolean
    Dim Rng         As range
    checkResultDuplicates2 = False
    Set Rng = rwb.Sheets(1).range(resultRange1, resultRange2)
    'Test Range for Duplicates
    
    MsgBox (Replace("NOT(AND((COUNTIF(@,@)=1)))", "@", Rng.Address))
    
    If Evaluate(Replace("NOT(AND((COUNTIF(@,@)=1)))", "@", Rng.Address)) = True Then
        checkResultDuplicates2 = True
    End If
End Function

Public Function checkResultDuplicates(rwb) As Boolean
    Dim Rng         As range
    Dim myArray     As Variant
    checkResultDuplicates = False
    
    'Store Range to search through
    Set Rng = rwb.Sheets(1).range(resultRange1, resultRange2)
    
    'Use Unique Function and store result to an Array variable
    myArray = WorksheetFunction.Unique(Rng)
    
    'Test Range for Duplicates
    If UBound(myArray) < Rng.count Then
        summaryMessage = summaryMessage & resultfile.name & " Contains Duplicate Sample Names. Please Check." & vbNewLine
        
    End If
End Function

'After matching, Will check if all of the genotype in the result book are filled.
Public Function checkCompleteGenotype(rwb, gtColumn) As Boolean
    
    Dim gtLastRow, numSample As Integer
    Dim gtrange1, gtrange2 As String
    Dim Rng         As range
    Dim myArray     As Variant
    
    gtLastRow = rwb.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 1
    numSample = ResultlastRow
    gtrange1 = Split(Cells(1, gtColumn).Address, "$")(1) & ResultSSSampleRow
    gtrange2 = Split(Cells(1, gtColumn).Address, "$")(1) & gtLastRow
    
    Set Rng = rwb.Sheets(1).range(gtrange1, gtrange2)
    myArray = WorksheetFunction.Unique(Rng)
    
    'Test Range for Duplicates
    If Not UBound(myArray) = numSample Then
        '"Duplicates Found!
        MsgBox (gtrange1 & gtrange2)
        MsgBox (UBound(myArray) & "  " & numSample)
        MsgBox ("samplesheet Not filled")
        
    End If
End Function

'Returns the Row number in the Analysis File matching the Name. Will Ignore Extra Tests Stype.
Function SampleNameSearch(name, aWB) As Integer
    
    Dim lastRow, Rowindex, sampleColumn, sampType, columnIndex As Integer
    Dim comboName   As String
    Dim field       As Variant
    
    lastRow = aWB.Sheets(1).Cells(Rows.count, 1).End(xlUp).Row - 2        'Use this to stop the loop early.
    
    sampleColumn = SearchColumnSource("samplename", aWB)
    sampType = SearchColumnSource("stype", aWB)
    
    columnIndex = 2
    
    'Change this when doing JaX
    For Each field In aWB.Sheets(1).range(AnalysisRange1, AnalysisRange2)        ' Make This dynamic, not SS for now
        comboName = fileExSplitName & "$" & name
        

        
        
        If field = "" Then
            'Skip
        ElseIf field = comboName And checkJaxExtraStype(aWB.Sheets(1).Cells(columnIndex, sampType)) Then

            SampleNameSearch = columnIndex
            
            Exit For
        End If
        columnIndex = columnIndex + 1
    Next
    
End Function

'Will check if Stype is Not an Extra Test
Function checkJaxExtraStype(stype) As Boolean
    checkJaxExtraStype = True
    Dim cell        As Variant
    For Each cell In jaxExtraStypes.Cells
        
        If Replace(LCase(stype), " ", "") = Replace(LCase(cell), " ", "") Then
            checkJaxExtraStype = False
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function

'Match Data will Copy results from Analysis file to Result File.
Sub MatchData(resultBook, analysisBook)
    Dim lastRow, Index, rowdata  As Integer
    
    lastRow = SearchColumnTarget("#", resultBook)
    
    Dim sample      As Variant
    
    Index = ResultSSSampleRow
    'Loop through samples in ResultSS and match with AnalysisSS rows.
    'For Each Sample In resultBook.Worksheets(1).Range("B18", "B114") 'Fix this to make dynamic
    
    For Each sample In resultBook.Worksheets(1).range(resultRange1, resultRange2)
        If sample = "" Then
            
        Else
            
            'Where the Magic Happens. Matches samples to AnalysisSS. Will Ignore Extra Test Stypes
            rowdata = SampleNameSearch(sample, analysisBook)        'Grabs the Row from SS sheet.
            'Debug.Print ("Row" & rowdata)
            'Debug.Print ("TG" & TargetGenotype)
            'Debug.Print ("PCR" & PCR1)
            
            
            'PCR1 Section
            If Not rowdata = 0 And Not TargetGenotype = 0 And Not PCR1 = 0 Then

                resultBook.Worksheets(1).Cells(Index, TargetGenotype) = analysisBook.Worksheets(1).Cells(rowdata, genotype)
                resultBook.Worksheets(1).Cells(Index, TargetPCR1) = analysisBook.Worksheets(1).Cells(rowdata, PCR1)
                resultBook.Worksheets(1).Cells(Index, TargetgmCAG) = analysisBook.Worksheets(1).Cells(rowdata, gmcag)
                resultBook.Worksheets(1).Cells(Index, TargetseqCAG) = analysisBook.Worksheets(1).Cells(rowdata, seqCAG)
                                                      
                'PCR2 Section
                If Not pcr2 = 0 And Not gmcag2 = 0 And Not seqcag2 = 0 Then
                    
                    resultBook.Worksheets(1).Cells(Index, lastRow + 1) = analysisBook.Worksheets(1).Cells(rowdata, pcr2)
                    
                    resultBook.Worksheets(1).Cells(Index, lastRow + 2) = analysisBook.Worksheets(1).Cells(rowdata, gmcag2)
                    
                    resultBook.Worksheets(1).Cells(Index, lastRow + 3) = analysisBook.Worksheets(1).Cells(rowdata, seqcag2)
                    
                    'Transfer Mark
                    
                End If
                analysisBook.Worksheets(1).Cells(rowdata, tran) = "CP_JAX"
            End If
            
        End If
        
        Index = Index + 1
        
    Next
    
End Sub

'This function will look for the names in the Target file and return the index of that column
Public Function SearchColumnTarget(search, wb) As Integer
    Dim columnIndex As Integer
    Dim field       As Variant
    columnIndex = 1
    For Each field In wb.Sheets(1).range(ResultSSRangeHeader.Address)        'The Template rows has to start at 18
        If Replace(LCase(field), " ", "") = Replace(LCase(search), " ", "") Then
            SearchColumnTarget = columnIndex
        End If
        columnIndex = columnIndex + 1
    Next
End Function

'Sorts Analysis workbook by old inex

Sub sortAnalysisBook(aWB)
    Dim oldIndexcol As Integer
    Dim oldIndexAdress As String
    
    oldIndexcol = SearchColumnSource("oldindex", aWB)
    If Not oldIndexcol = 0 Then
        
        oldIndexAdress = Split(Cells(1, oldIndexcol).Address, "$")(1) & 1
        
        aWB.Sheets(1).range("A1:BB600").Sort Key1:=range(oldIndexAdress), Order1:=xlAscending, header:=xlYes
        
    End If
    
End Sub

'Changes the headers on the resultbooks, Also adds the counter formula
Public Sub setUpResultBooks(wb, resultTextLocation)
    Dim gtColumn    As String
    Dim gtnumber    As Integer
    
    If Not resultTextLocation = 0 Then
        
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation) = "PCR1"
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation + 1) = "GM CAG 1"
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation + 2) = "SEQ CAG 1"
        wb.Worksheets(1).Cells(ResultSSHeaderRow, resultTextLocation + 3) = "GT"
        
        Dim CounterFormulaBox As range
        
        Set CounterFormulaBox = wb.Worksheets(1).Cells(3, resultTextLocation + 3)
        
        If CounterFormulaBox.MergeCells Then
            UnmergeMergedRangeBasedOnCell wb.Worksheets(1), Cells(3, resultTextLocation + 3)

        End If
        
        
        Debug.Print (CounterFormulaBox)
        

        
        sampleCounterJax.Copy wb.Worksheets(1).Cells(3, resultTextLocation + 3)
        
    End If
    
End Sub

Public Function SearchResultColumnTestJax(wb) As Integer
    Dim columnIndex As Integer
    Dim field, test As Variant
    Dim found       As Boolean
    found = False
    columnIndex = 0
    
    For Each field In wb.Sheets(1).range(ResultSSRangeTestName.Address)        'The Template rows has to start at 19 for Jax
        If field = "" Then
            columnIndex = columnIndex + 1
        Else
            columnIndex = columnIndex + 1
            For Each test In jaxTests
                If test = field Then
                    SearchResultColumnTestJax = columnIndex
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

'Function to search for Standard qPCR Laragen Tests.
Function SearchTestColumn(wb) As Integer
    Dim columnIndex As Integer
    columnIndex = 1
    
    Dim found       As Boolean
    found = False
    
    Dim field, test As Variant
    
    For Each field In wb.Sheets(1).range(ResultSSRangeTestName.Address)        'The Template rows has to start at 17
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

'Will check if Stype is Not an Extra Test (Final)
Function checkStype(stype) As Boolean
    checkStype = True
    Dim cell        As Variant
    For Each cell In extraStypes.Cells
        If stype = cell Then
            checkStype = False
        ElseIf cell = "" Then
            Exit For
        End If
    Next
End Function

Function FindFirstHashRow(wb) As Integer
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim i As Integer
    
    ' Set the worksheet where the data is located
    Set ws = wb.Sheets(1)
    
    ' Find the last row of the data in column A
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' Loop through each row of column A
    For i = 1 To lastRow
        ' Check if the cell contains only "#" and return the row number if found
        If ws.Cells(i, "A").Value = "#" And Len(ws.Cells(i, "A").Value) = 1 Then
            FindFirstHashRow = i
            Exit Function
        End If
    Next i
    
    ' Return 0 if no match found
    FindFirstHashRow = 0
End Function



Function ExtractReqNumber(inputText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the regex pattern to match variations of "Req#:" or "Req #:"
    regex.Pattern = "Req\s*#?\s*:\s*(\d+)"
    
    ' Check if the pattern matches the input text
    If regex.test(inputText) Then
        ' If there's a match, extract the captured group (the number)
        Dim match As Object
        Set match = regex.Execute(inputText)(0)
        ExtractReqNumber = match.SubMatches(0)
    Else
        ' If no match is found, return an empty string or handle the error as needed
        ExtractReqNumber = ""
    End If
End Function

Sub UnmergeMergedRangeBasedOnCell(ByVal ws As Worksheet, ByVal targetCell As range)
    Dim cell As range
    Dim mergedRange As range
    
    ' Loop through all cells in the worksheet
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            ' Check if the target cell is within the merged range of the current cell
            If Not Intersect(targetCell, cell.MergeArea) Is Nothing Then
                ' If it is, store the merged range in the mergedRange variable
                Set mergedRange = cell.MergeArea
                Exit For ' Exit the loop since we found the merged range
            End If
        End If
    Next cell
    
    ' Check if a merged range was found
    If Not mergedRange Is Nothing Then
        mergedRange.UnMerge
    Else
        MsgBox "No merged range found for the target cell."
    End If
End Sub

