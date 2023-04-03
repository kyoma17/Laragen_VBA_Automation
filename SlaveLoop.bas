Attribute VB_Name = "SlaveLoop"
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


Sub deletestuff()
    turnOff
    
    
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
    Set testNickNames = TestProtocol.range("B2:B100")
    
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
    Dim tsFile, exFile, proFile, matcher, nickname, i, proto, unTest As Variant

    Dim splitTestArray As Variant
    Dim outputFile As String
    
    
    For Each tsFile In tsFSOFile


        

        
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


        
