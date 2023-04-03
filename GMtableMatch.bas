Attribute VB_Name = "GMtableMatch"
'Transfer Data from the GM table to the Analysis SS

Option Explicit


'Global Variables
Dim mainWB          As Sheets
Dim READ_Me, PAGE2, TAC_TEST_LIST, JAX_TEST_LIST, MARGE_TEST_LIST, TestProtocol As Worksheet
Dim TestWB          As Workbook
Dim gmFSOLibrary, asFSOLibrary, osFSOLibrary As FileSystemObject
Dim gmFSOFolder, gmFSOFile, asFSOFolder, asFSOFile, osFSOFolder, osFSOFile  As Object

Dim AnalysisSS, GmTableSS, TestDB As range

Dim gmBook, analysisBook As Workbook
Dim gmSSN, gmPCR1, gmPH, gmPA, gmSEQ, gmGM, gmPCR2, gmPH2, gmPA2, gmSEQ2, gmGM2, tran As Integer
Dim aSSN, aPCR1, aPH, aPA, aSEQ, aGM, aPCR2, aPH2, aPA2, aSEQ2, aGM2 As Integer
Dim gmLastRow, gmRange1, gmHexRange, anLastRow, aRange1, aRange2, extraStypes As String


Dim hexSSN, hexPCR1, hexPH, hexPA, hexPCR2

Dim allExtraStypes As Object


Sub checkGMHeader()
    Dim gmHeaders() As Variant
    gmHeaders = Array("Sample Name", "Size 1", "Height ", "Area", "Size 2")
    headerChecker gmHeaders, gmBook
End Sub

Sub checkSSHeader()
    Dim ssHeaders() As Variant
    ssHeaders = Array("Sample Name", "Stype")
    headerChecker ssHeaders, analysisBook
End Sub


Sub GMTransfer()
    turnOff
    
    
    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    Set PAGE2 = mainWB("PAGE2")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set AnalysisSS = READ_Me.Cells(12, 2)
    Set GmTableSS = READ_Me.Cells(22, 2)
    Set TestDB = READ_Me.Cells(17, 2)
    
    Workbooks.Open FileName:=TestDB, Local:=True
    Set TestWB = ActiveWorkbook
    
    Set JAX_TEST_LIST = TestWB.Sheets("JAX_TEST_LIST")
    Set TAC_TEST_LIST = TestWB.Sheets("TACONIC_TEST_LIST")
    Set MARGE_TEST_LIST = TestWB.Sheets("MARGE_TEST_LIST")
    
    extraStypes = ""
    Dim exStype As Variant
    
    'combine all the extra stypes
    For Each exStype In JAX_TEST_LIST.range("E2:E50")
        If Not exStype = "" And InStr(extraStypes, exStype) = 0 Then
            extraStypes = extraStypes & "," & cleanInput(exStype)
        End If
    
    Next
    For Each exStype In MARGE_TEST_LIST.range("E2:E50")
        If Not exStype = "" And InStr(extraStypes, exStype) = 0 Then
            extraStypes = extraStypes & "," & cleanInput(exStype)
        End If
    
    Next
    For Each exStype In TAC_TEST_LIST.range("B2:B50")
        If Not exStype = "" And InStr(extraStypes, exStype) = 0 Then
            extraStypes = extraStypes & "," & cleanInput(exStype)
        End If
    
    Next
    
    
    
    'gmTableSS
    Set gmFSOLibrary = New FileSystemObject
    Set gmFSOFolder = gmFSOLibrary.GetFolder(GmTableSS)
    Set gmFSOFile = gmFSOFolder.Files
    
    'AnalysisSS
    Set asFSOLibrary = New FileSystemObject
    Set asFSOFolder = asFSOLibrary.GetFolder(AnalysisSS)
    Set asFSOFile = asFSOFolder.Files
    
    
    Dim gmTable, analysisFile As Variant
    Dim rowMatch As Integer
    Dim gmSample As Variant
    

    
    
    For Each gmTable In gmFSOFile
        
        If InStr(gmTable, ".xlsx") > 0 And InStr(gmTable, "~$") < 1 And InStr(gmTable, "desktop.ini") < 1 Then
        
        
            'Setting up GM Book for Transfer
            Workbooks.OpenText FileName:=gmTable, Local:=True
            Set gmBook = ActiveWorkbook
            
            checkGMHeader
            
            'Rox GM singleload samples refer to checkGMHeader
            gmSSN = SearchColumnSource("samplename", gmBook)
            gmPCR1 = SearchColumnSource("size1", gmBook)
            gmPH = SearchColumnSource("height", gmBook)
            gmPA = SearchColumnSource("area", gmBook)
            gmGM = SearchColumnSource("gm", gmBook)
            gmSEQ = SearchColumnSource("seq", gmBook)
            
            gmPCR2 = SearchColumnSource("size2", gmBook)
            gmPH2 = SearchColumnSource("height2", gmBook)
            gmPA2 = SearchColumnSource("area2", gmBook)
            gmGM2 = SearchColumnSource("gm2", gmBook)
            gmSEQ2 = SearchColumnSource("seq2", gmBook)
            gmLastRow = gmBook.Sheets(1).Cells(Rows.count, gmSSN).End(xlUp).Row
            gmRange1 = Split(Cells(1, gmSSN).Address, "$")(1)
            
            
            'Hex GM Doubleload Samples
            hexSSN = SearchColumnSource("hexsamplename", gmBook)
            hexPCR1 = SearchColumnSource("hexsize1", gmBook)
            hexPA = SearchColumnSource("hexarea", gmBook)
            hexPH = SearchColumnSource("hexheight1", gmBook)
            hexPCR2 = SearchColumnSource("hexsize2", gmBook)

            'Adds Transfer Column
            addTransferColumn gmBook
            tran = SearchColumnSource("tran", gmBook)

            hexdoubleload
            gmLastRow = gmBook.Sheets(1).Cells(Rows.count, gmSSN).End(xlUp).Row
            
            
            '---------------------------------------------------------------------------------------------------------------------
            For Each analysisFile In asFSOFile
        
                If InStr(analysisFile, ".xlsx") > 0 And InStr(analysisFile, "~$") < 1 Then
                    Workbooks.OpenText FileName:=analysisFile, Local:=True
                    Set analysisBook = ActiveWorkbook
                    sortBook analysisBook, "oldindex"
                    
                    checkSSHeader
                    
                    addGeneMapperColumn analysisBook
                    
                    'AnalysisBookSetup REFER to addGeneMapperColumn Setup
                    aSSN = SearchColumnSource("samplename", analysisBook)
                    aPCR1 = SearchColumnSource("pcr1", analysisBook)
                    aPH = SearchColumnSource("peakheight", analysisBook)
                    aPA = SearchColumnSource("peakarea", analysisBook)
                    aGM = SearchColumnSource("gm", analysisBook)
                    aSEQ = SearchColumnSource("seq", analysisBook)
                    
                    aPCR2 = SearchColumnSource("pcr2", analysisBook)
                    aGM2 = SearchColumnSource("gm2", analysisBook)
                    aSEQ2 = SearchColumnSource("seq2", analysisBook)
                    
                    anLastRow = analysisBook.Sheets(1).Cells(Rows.count, aSSN).End(xlUp).Row
                    aRange1 = Split(Cells(1, aSSN).Address, "$")(1)
                    
                    For Each gmSample In gmBook.Sheets(1).range(gmRange1 & 2 & ":" & gmRange1 & gmLastRow)
                        'Single Load Samples Only Rox

                        If VarType(gmSample) = 8 Then
                            'Skip non strings
                        
                            If Not gmSample = "" Then
                            
                                rowMatch = gmMatching(gmSample, analysisBook)
                                
                                If Not rowMatch = 0 Then
                                    If Not gmBook.Worksheets(1).Cells(gmSample.Row, gmPCR1) = "" Then
                                    
                                    
                                        analysisBook.Worksheets(1).Cells(rowMatch, aPCR1) = _
                                                                                gmBook.Worksheets(1).Cells(gmSample.Row, gmPCR1)
                                        analysisBook.Worksheets(1).Cells(rowMatch, aPH) = _
                                                                                gmBook.Worksheets(1).Cells(gmSample.Row, gmPH)
                                        analysisBook.Worksheets(1).Cells(rowMatch, aPA) = _
                                                                                gmBook.Worksheets(1).Cells(gmSample.Row, gmPA)
        
                                        'GM CAG and SEQ Formulas can be added here in the future
                                        cagFormula (analysisBook.Worksheets(1).Cells(rowMatch, aPCR1))
                                        
                                                                                
                                        If Not gmPCR2 = 0 Then
                                            analysisBook.Worksheets(1).Cells(rowMatch, aPCR2) = _
                                                                                gmBook.Worksheets(1).Cells(gmSample.Row, gmPCR2)
                                            cagFormula (analysisBook.Worksheets(1).Cells(rowMatch, aPCR2))
        
                                        End If
                                        
                                        'Adds copy marker for this row
                                        gmBook.Worksheets(1).Cells(gmSample.Row, tran) = "CP_" & Left(analysisFile.name, 20)
                                    
                                    End If
                                End If
                            End If
                        End If
                        
                    Next
                    
                    formatColumn aPCR1, analysisBook.Sheets(1)
                    formatColumn aGM, analysisBook.Sheets(1)
                    formatColumn aSEQ, analysisBook.Sheets(1)
                    formatColumn aPCR2, analysisBook.Sheets(1)
                    formatColumn aGM2, analysisBook.Sheets(1)
                    formatColumn aSEQ2, analysisBook.Sheets(1)
                    
                    analysisBook.Close savechanges:=True
                End If
            Next
            gmBook.Close savechanges:=True
        End If
    Next
    
    
    TestWB.Close savechanges:=False
    
    Set gmFSOLibrary = Nothing
    Set gmFSOFolder = Nothing
    Set gmFSOFile = Nothing
    Set asFSOLibrary = Nothing
    Set asFSOFolder = Nothing
    Set asFSOFile = Nothing
    
    reset
End Sub

'Adds the GM Data Headers to the Analysis Workbooks
Public Sub addGeneMapperColumn(aWB)
    addLastcolumn aWB, "PCR1"
    addLastcolumn aWB, "PeakHeight"
    addLastcolumn aWB, "PeakArea"
    addLastcolumn aWB, "GM"
    addLastcolumn aWB, "SEQ"
    
    addLastcolumn aWB, "PCR2"
    addLastcolumn aWB, "GM2"
    addLastcolumn aWB, "SEQ2"
End Sub

Sub hexdoubleload()
    Dim hexSample As Variant

    If Not hexSSN = 0 Then
        gmHexRange = Split(Cells(1, hexSSN).Address, "$")(1)
    
        For Each hexSample In range(gmHexRange & 2 & ":" & gmHexRange & gmLastRow)
            If hexSample = "None" Or hexSample = "" Then
                'Skip
            ElseIf existsColumn(hexSample) Then
                'Skip
            Else
                gmBook.Sheets(1).Cells(gmLastRow + 1, gmSSN) = gmBook.Sheets(1).Cells(hexSample.Row, hexSSN)
                gmBook.Sheets(1).Cells(gmLastRow + 1, gmSSN).Font.color = vbGreen
                gmBook.Sheets(1).Cells(gmLastRow + 1, gmPCR1) = gmBook.Sheets(1).Cells(hexSample.Row, hexPCR1)
                gmBook.Sheets(1).Cells(gmLastRow + 1, gmPH) = gmBook.Sheets(1).Cells(hexSample.Row, hexPH)
                gmBook.Sheets(1).Cells(gmLastRow + 1, gmPA) = gmBook.Sheets(1).Cells(hexSample.Row, hexPA)
                gmBook.Sheets(1).Cells(gmLastRow + 1, gmPCR2) = gmBook.Sheets(1).Cells(hexSample.Row, hexPCR2)
                gmLastRow = gmLastRow + 1
            End If
        Next
    End If

End Sub


Function existsColumn(name) As Boolean
    Dim checkCell As Variant
    Dim checkRange As String
    Dim lastSSN As Integer
    lastSSN = gmBook.Sheets(1).Cells(Rows.count, gmSSN).End(xlUp).Row
    checkRange = Split(Cells(1, gmSSN).Address, "$")(1)
    existsColumn = False
    
    For Each checkCell In gmBook.Sheets(1).range(checkRange & 2 & ":" & checkRange & lastSSN)
        If Not WorksheetFunction.IsFormula(checkCell) Then
            If checkCell = name Then existsColumn = True
        End If
    Next
End Function


'Calculate an shove the cag gm and seq data into the analysis sheet. This is the lazy way.
Sub cagFormula(data)

    Dim cagStype As String
    cagStype = analysisBook.Sheets(1).Cells(data.Row, SearchColumnSource("stype", analysisBook))
    
    If data = "" Then
        'Skip
    ElseIf IsNumeric(data) Then
        'BAC CAG (PCR-206)/3
        cagFormulaHelper (InStr(cleanInput(cagStype), "bac") > 0), data, cagBAC(data)
        'N107 cag  (PCR-211)/3
        cagFormulaHelper (InStr(cleanInput(cagStype), "n107") > 0), data, cagN107(data)
        'Common (PCR-122)/3
        cagFormulaHelper True, data, cagCommon(data)
    End If
    
End Sub

Sub cagFormulaHelper(argument, data, gmData)

    If argument Then
        If analysisBook.Sheets(1).Cells(1, data.column) = "PCR1" Then
            analysisBook.Sheets(1).Cells(data.Row, SearchColumnSource("gm", analysisBook)) = gmData
            analysisBook.Sheets(1).Cells(data.Row, SearchColumnSource("seq", analysisBook)) = seqCAG(gmData)
        ElseIf analysisBook.Sheets(1).Cells(1, data.column) = "PCR2" Then
            analysisBook.Sheets(1).Cells(data.Row, SearchColumnSource("gm2", analysisBook)) = gmData
            analysisBook.Sheets(1).Cells(data.Row, SearchColumnSource("seq2", analysisBook)) = seqCAG(gmData)
        End If
    End If
End Sub

Function cagColorCode(data, color)
    analysisBook.Sheets(1).Cells(data.Row, SearchColumnSource("peakarea", analysisBook)).Interior.color = color
End Function


Function cagCommon(size) As Double
    cagCommon = (size - 122) / 3
     cagColorCode size, vbBlue
End Function
    
Function cagN107(size) As Double
    cagN107 = (size - 206) / 3
    cagColorCode size, vbCyan
End Function
Function cagBAC(size) As Double
    cagBAC = (size - 211) / 3
    cagColorCode size, vbMagenta
End Function
Function seqCAG(gmCAGinput) As Double
    seqCAG = gmCAGinput * 1.0424 + 1.2088
End Function


'prevent copying to extra stype.
Function gmMatching(gmSample, aWB) As Integer
    Dim aSample As Variant
    gmMatching = 0
    
    For Each aSample In aWB.Sheets(1).range(aRange1 & 2 & ":" & aRange1 & anLastRow)
        If aSample = "" Then
            'skip
        
        ElseIf aSample = gmSample And InStr(extraStypes, cleanInput(aWB.Sheets(1).Cells(aSample.Row, SearchColumnSource("stype", aWB)))) < 1 Then
            gmMatching = aSample.Row
        End If
    Next
                    
End Function


Function extraStypeGMcopy(aSample, aWB) As Boolean
    extraStypeGMcopy = True
    If aWB.Sheets(1).Cells(aSample.Row, SearchColumnSource("stype", aWB)) Then
        
    End If
    
End Function




        
        

