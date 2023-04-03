Attribute VB_Name = "ProtocolGeneration"
Option Explicit

'Global Variables
Dim mainWB          As Sheets
Dim READ_Me, PAGE2, TestProtocol As Worksheet
Dim TestWB          As Workbook
Dim tsFSOLibrary, esFSOLibrary, osFSOLibrary As FileSystemObject
Dim tsFSOFolder, tsFSOFile, esFSOFolder, esFSOFile, osFSOFolder, osFSOFile  As Object

Dim ExtractionSS, TestDB, ProtocolFiles, ProtocolOutput, testNickNames As range
Dim pseudoTestArray, pseudoColorArray As String
Dim currentDate As String

Dim protocolNotFound As String



Sub protocolGen()
    turnOff
    
    
    currentDate = InputBox("Please Enter the Date in MMDDYY Format:")
    If currentDate = "" Then
        currentDate = "NoDate"
    End If
    

    'Assigns the Read_Me ExtraTests and Settings Worksheet
    Set mainWB = ThisWorkbook.Worksheets
    Set READ_Me = mainWB("READ_ME")
    Set PAGE2 = mainWB("PAGE2")
    
    'This Sets up the Folder Directories from the Read_Me page
    Set ExtractionSS = PAGE2.Cells(12, 2)
    Set TestDB = PAGE2.Cells(13, 2)
    Set ProtocolFiles = PAGE2.Cells(14, 2)
    Set ProtocolOutput = PAGE2.Cells(15, 2)
    
    Workbooks.OpenText FileName:=TestDB, Local:=True
    Set TestWB = ActiveWorkbook
    

    Set TestProtocol = TestWB.Sheets("TEST_PROTOCOL")
    Set testNickNames = TestProtocol.range("B2:B500")
    
    Set tsFSOLibrary = New FileSystemObject
    Set tsFSOFolder = tsFSOLibrary.GetFolder(ProtocolFiles)
    Set tsFSOFile = tsFSOFolder.Files
    
    'Extraction SS
    Set esFSOLibrary = New FileSystemObject
    Set esFSOFolder = esFSOLibrary.GetFolder(ExtractionSS)
    Set esFSOFile = esFSOFolder.Files
    
    Set osFSOLibrary = New FileSystemObject
    Set osFSOFolder = osFSOLibrary.GetFolder(ProtocolOutput)
    Set osFSOFile = osFSOFolder.Files
    
    
    Dim proBook, exBook As Workbook
    Dim exFile, proFile, matcher, nickname, i, proto, unTest As Variant

    Dim splitTestArray As Variant
    Dim outputFile As String
    
    
    For Each exFile In esFSOFile
        pseudoTestArray = ""
        If InStr(exFile, ".xlsx") > 0 And InStr(exFile, "~$") < 1 And InStr(exFile, "desktop.ini") < 1 Then
            Workbooks.OpenText FileName:=exFile, Local:=True
            Set exBook = ActiveWorkbook
            
            'Grab all the unique tests from the extraction plate
            For Each matcher In Union(exBook.Sheets(1).range("A11:R11"), exBook.Sheets(1).range("N1:N14"))
                Debug.Print (matcher)
                If Not matcher = 0 Then
                    For Each nickname In testNickNames
                        If cleanInput(matcher) = cleanInput(nickname) Then
                            If InStr(pseudoTestArray, cleanInput(matcher)) < 1 Then
                                pseudoTestArray = pseudoTestArray & cleanInput(matcher) & ","
                                
                            End If
                        End If
    
                    Next
                End If
            Next
            'Cycle through each unqiue test and perform the plate copy
            For Each unTest In Split(pseudoTestArray, ",")
                
                For Each nickname In testNickNames
                    
                    If unTest = "" Or nickname = "" Then
                        'Skip
                    ElseIf cleanInput(unTest) = cleanInput(nickname) Then
                        For i = 1 To 7
                            proto = TestProtocol.Cells(nickname.Row, nickname.column + i)
                            If Not proto = "" Then
                                'open workbook
                                
                                If checkExistProtocol(proto) Then
                                    Workbooks.OpenText FileName:=ProtocolOutput & currentDate & "-" & proto & ".xlsx", Local:=True
                                    Set proBook = ActiveWorkbook
                                Else
                                    Workbooks.OpenText FileName:=ProtocolFiles & proto & ".xlsx", Local:=True
                                    Set proBook = ActiveWorkbook
                                End If
                                
                                If Not checkExistExtPlate(exBook, proBook) Then
                                    copyEngine exBook, proBook, nickname, proto
                                End If

                                sampleCounter proBook
                                Application.DisplayAlerts = False
                                proBook.SaveAs FileName:=ProtocolOutput & currentDate & "-" & proto & ".xlsx"
                                proBook.Close
                                Application.DisplayAlerts = True

           
                            End If
                        Next
                    End If

                Next
            Next
            exBook.Close savechanges:=False
            
        End If
    Next
    
    Set tsFSOLibrary = Nothing
    Set tsFSOFolder = Nothing
    Set tsFSOFile = Nothing
    Set osFSOLibrary = Nothing
    Set osFSOFolder = Nothing
    Set osFSOFile = Nothing
    Set esFSOLibrary = Nothing
    Set esFSOFolder = Nothing
    Set esFSOFile = Nothing
    
    
    
    TestWB.Close savechanges:=False
    
    
    reset
End Sub

Sub copyEngine(eWB, pWB, nickname, protocol)
    pWB.Sheets(1).Cells(1, 1) = Now()
    pWB.Sheets(1).range("O:R").EntireColumn.Hidden = True
    
    Dim startRow As Integer
    Dim colorKeep, testKeep As String
    Dim delColor, testColor As Variant
    Dim lastEmpty As Boolean
    lastEmpty = True
    
    startRow = 17
    
    'Moves target to last empty location
    While lastEmpty
        If pWB.Sheets(1).Cells(startRow, 1) = "" Then
            lastEmpty = False
        Else
            startRow = startRow + 13
        End If
    Wend
    
    'Group 1 Extraction, not edited
    eWB.Sheets(1).range("A14:Q25").Copy pWB.Sheets(1).range("A" & startRow)
    eWB.Sheets(1).range("N1:Q10").Copy pWB.Sheets(1).range("N" & startRow)
                                
    'Group 2 Test Map, remove samples from different tests.
    eWB.Sheets(1).range("A14:Q25").Copy pWB.Sheets(1).range("S" & startRow - 13)
    eWB.Sheets(1).range("N1:Q10").Copy pWB.Sheets(1).range("AF" & startRow - 13)
    
    'Picks up colors that belong to this protocol
    'Looks at the PI and Strain Names
    For Each testColor In Union(pWB.Sheets(1).range("A" & startRow + 10 & ":" & "N" & startRow + 10), pWB.Sheets(1).range("N" & startRow & ":" & "N" & startRow + 10))
        
        If testColor = "" Then
            'Skip
            
        'If the NicknameTest on extraction sheet also shares a Test with the Protocol OR The nickname matches the testDB nickname
        ElseIf cleanInput(testColor) = cleanInput(nickname) Or testMatch(testColor, protocol) Then
            colorKeep = colorKeep & testColor.Font.color & ","
            testKeep = testKeep & testColor & ","
            

        End If
    Next
    'Deletes all non-colors
    For Each delColor In pWB.Sheets(1).range("T" & startRow - 11 & ":" & "AI" & startRow - 1)
        delColor.Interior.ColorIndex = xlNone
        If InStr(colorKeep, delColor.Font.color) = 0 Then
            delColor.Value = ""
        End If
    Next
    
    'Adds Test Header to Protocol Map
    pWB.Sheets(1).Cells(startRow - 13, 19) = pWB.Sheets(1).Cells(startRow - 13, 19) & " (" & pWB.Sheets(1).Cells(2, 1) & ")"
End Sub

Function testMatch(test, protocol) As Boolean
    Dim test1, file As Variant
     testMatch = False
    
    For Each test1 In testNickNames
        
        For file = 1 To 7

            If TestProtocol.Cells(test1.Row, test1.column + file) = "" Then
                ' skip
            ElseIf protocol = TestProtocol.Cells(test1.Row, test1.column + file) And cleanInput(test) = cleanInput(test1) Then
                testMatch = True
                
            End If
        
        Next
    Next
End Function

Function checkExistExtPlate(eWB, pWB) As Boolean

    Dim startPWB As Integer
    startPWB = 14
    checkExistExtPlate = False
    
    Dim checker As Variant
    
    For Each checker In pWB.Sheets(1).range("A14:A600")
        If checker = eWB.Sheets(1).Cells(1, 1) Then
            checkExistExtPlate = True
        End If
    Next
        
    
End Function



'Checks if the output folder has the same protocl file.
Function checkExistProtocol(protocol) As Boolean
    Dim osFile As Variant
    checkExistProtocol = False
    For Each osFile In osFSOFile
        If InStr(osFile, protocol) And InStr(osFile, currentDate) Then
            checkExistProtocol = True
        End If
    Next
End Function


Sub sampleCounter(pWB)
    Dim startCounter, sampleCount, counterEnd As Integer
    Dim TotalFinder, sampleFinder As Variant
    startCounter = 4
    sampleCount = 0
    
    Dim theEnd, theCounterEnd As Boolean
    theEnd = False
    theCounterEnd = False
    
    While Not theEnd
    
        If pWB.Sheets(1).range("S" & startCounter) = "" Then
            theEnd = True
        Else
            For Each sampleFinder In pWB.Sheets(1).range("T" & startCounter + 2 & ":" & "AE" & startCounter + 9)

                If Not sampleFinder = "" Then
                    sampleCount = sampleCount + 1
                End If
            Next
            
            startCounter = startCounter + 13
        End If
        
    Wend
    
    For Each TotalFinder In pWB.Sheets(1).range("A1:H17")
        If InStr(LCase(TotalFinder.text), "total") > 0 Then
            counterEnd = False
            counterEnd = 0
            While Not theCounterEnd
                If Application.WorksheetFunction.IsNumber(pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 1)) _
                And Application.WorksheetFunction.IsNumber(pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 2)) _
                And Application.WorksheetFunction.IsNumber(pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 3)) Then
                    
                    
                    pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 2) = sampleCount
                    ' Volumn Calculations
                    If sampleCount <= 27 Then
                        'If less than 27 samples add 3 more samples worth
                        pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 3) = _
                            (sampleCount + 3) * pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 1)
                        
                    Else
                        'If greater than 27 samples multiplay total volumes by 120 percent
                        pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 3) = _
                            (sampleCount * pWB.Sheets(1).Cells(TotalFinder.Row - counterEnd, TotalFinder.column + 1)) * 1.2
                    
                    End If
                    
                Else
                    theCounterEnd = True
                End If
            
                counterEnd = counterEnd + 1
            Wend
        End If
    Next


End Sub
