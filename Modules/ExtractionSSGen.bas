Attribute VB_Name = "ExtractionSSGen"
Option Explicit

'Global Variables
Dim mainWB          As Sheets
Dim READ_Me, PAGE2, TestProtocol As Worksheet
Dim TestWB, AllInOne As Workbook
Dim tsFSOLibrary, esFSOLibrary, FSOLibrary As FileSystemObject
Dim tsFSOFolder, tsFSOFile, esFSOFolder, esFSOFile, FSOFolder, FSOFile  As Object

Dim ExtractionSS, TestDB, ProtocolFiles, ProtocolOutput, testNickNames As range
Dim pseudoTestArray, pseudoColorArray As String
Dim currentDate As String


Sub CreateExtractionSS()
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
    Set AllInOne = Workbooks.Add
    allInOneHeader
    
    

    Set TestProtocol = TestWB.Sheets("TEST_PROTOCOL")
    Set testNickNames = TestProtocol.range("B2:B100")
    
    Set tsFSOLibrary = New FileSystemObject
    Set tsFSOFolder = tsFSOLibrary.GetFolder(ProtocolFiles)
    Set tsFSOFile = tsFSOFolder.Files
    
    'Extraction SS
    Set esFSOLibrary = New FileSystemObject
    Set esFSOFolder = esFSOLibrary.GetFolder(ExtractionSS)
    Set esFSOFile = esFSOFolder.Files
    
    Set FSOLibrary = New FileSystemObject
    Set FSOFolder = FSOLibrary.GetFolder("C:\Users\SERVER\Desktop\RAW")

    Set FSOFile = FSOFolder.Files
    
    
    Dim rawInput As Variant
    Dim rawBook As Workbook
    
    For Each rawInput In FSOFile
        Workbooks.OpenText FileName:=rawInput, Local:=True
        Set rawBook = ActiveWorkbook
        
    
    
    Next
    
    Set tsFSOLibrary = Nothing
    Set tsFSOFolder = Nothing
    Set tsFSOFile = Nothing
    Set FSOLibrary = Nothing
    Set FSOFolder = Nothing
    Set FSOFile = Nothing
    Set esFSOLibrary = Nothing
    Set esFSOFolder = Nothing
    Set esFSOFile = Nothing
    TestWB.Close savechanges:=False
    
    
    reset
End Sub
Sub allInOneHeader()
    Dim headerArray As String
    Dim counter As Integer
    Dim head As Variant
    counter = 1
    headerArray = (",0,1,2,3,4,5,6,7,8,9,10,11,12,test,req,strain,strain_number,numberOfSample")
    For Each head In Split(headerArray, ",")
        AllInOne.Sheets(1).Cells(1, counter) = head
        counter = counter + 1
    Next
    

End Sub


Function wellPosToLocation(wellpos) As String

End Function


Function listIndexToLocation(wellpos) As String

End Function
        


