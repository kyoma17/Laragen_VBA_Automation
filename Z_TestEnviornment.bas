Attribute VB_Name = "Z_TestEnviornment"
Sub templateFunction()

End Sub








Public Sub test1()
    Dim Program, Version As String
    Program = "TaconicConversion"
    Version = "2.0"
    
    'checkVersion program, version
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    Dim summaryMessage As String
    summaryMessage = "Taconic Conversion Summary:" & vbNewLine
    
    Dim folderName  As String
    Dim OutputfolderName As String
    Dim FSOLibrary  As FileSystemObject
    Dim FSOFolder   As Object
    Dim FSOFile     As Object
    
    Dim wss         As Sheets
    Dim READ_Me     As Worksheet
    
    'Assigns the Read_Me Worksheet
    Set wss = ThisWorkbook.Worksheets
    Set READ_Me = wss("READ_ME")
    Dim EntryDate   As String
    
    'Temporary Worksheets
    
    'Seperate the file name into Project Number and Plate Number
    Dim fileSplitName As Variant
    
    Dim strainProjectNumber As Variant
    Dim plateNumber As Variant
    Dim plainName   As Variant
    
    'Target Column Variables
    Dim TargetSampleName As Variant
    Dim TargetWellPlate As Variant
    Dim TargetSerialNumber As Variant
    
    'SourcePlate Column Variables
    Dim SourceProjectNumber As Variant
    Dim SourceAscesioningNumber As Variant
    Dim SourceSampleName As Variant
    Dim SourceWellPlate As Variant
    Dim SourceSerialNumber As Variant
    Dim last_row    As Integer
    Dim indexNum    As Integer
    
    '!!! Set the path of working folder!!!
    'InputfolderName = "C:\\Users\\\SERVER\\Desktop\\INPUT\"
    'OutputfolderName = "C:\\Users\\\SERVER\\Desktop\\OUTPUT\"
    'TemplateFolder = "C:\\Users\\\SERVER\\Desktop\\Taconic_Strain_Template\"
    
    'This section will Refer to the Excel Read-Me Fields"
    TemplateFolder = READ_Me.Cells(19, 2)
    InputfolderName = READ_Me.Cells(18, 2)
    OutputfolderName = READ_Me.Cells(14, 2)
    
    EntryDate = InputBox("Please Enter the Date:")
    
    'Set all the references to the INPUT FSO Library
    Set FSOLibrary = New FileSystemObject
    Set FSOFolder = FSOLibrary.GetFolder(InputfolderName)
    Set FSOFile = FSOFolder.Files
    
    'Set all the references to the TEMPLATE FSO Library
    Set TemplateFSOLibrary = New FileSystemObject
    Set TemplateFSOFolder = TemplateFSOLibrary.GetFolder(TemplateFolder)
    Set TemplateFSOFile = TemplateFSOFolder.Files
    
    Dim Index       As Integer
    Dim found       As Boolean
    Dim Match       As Boolean
    Dim Current     As String
    
    'Use For Each loop to loop through each file in the folder
    For Each EachFSOFile In FSOFile
        If InStr(EachFSOFile, ".csv") > 0 And InStr(EachFSOFile, "~$") < 1 Then
            fileSplitName = Split(EachFSOFile.name, "_")        'Splits the Strain and Plate from File name
            plainName = Replace(EachFSOFile.name, ".csv", "")        'Removes .csv from filiename
            
            If InStr(fileSplitName(0), "CHDI") Then
                strainProjectNumber = fileSplitName(1)
                plateNumber = fileSplitName(2)
            Else
                strainProjectNumber = fileSplitName(0)
                plateNumber = fileSplitName(1)
            End If
            
            OpenCsvAsText EachFSOFile
            Set inputbook = ActiveWorkbook
            

            
            End
        End If
    Next
End Sub

Sub OpenCsvAsText(ByVal strFilepath As String)

    Dim intFileNo As Integer
    Dim iCol As Long
    Dim nCol As Long
    Dim strLine As String
    Dim varColumnFormat As Variant
    Dim varTemp As Variant

    '// Read first line of file to figure out how many columns there are
    intFileNo = FreeFile()
    Open strFilepath For Input As #intFileNo
    Line Input #intFileNo, strLine
    Close #intFileNo
    varTemp = Split(strLine, ",")
    nCol = UBound(varTemp) + 1

    '// Prepare description of column format
    ReDim varColumnFormat(0 To nCol - 1)
    For iCol = 1 To nCol
        varColumnFormat(iCol - 1) = Array(iCol, xlTextFormat)
        ' What's this? See VBA help for OpenText method (FieldInfo argument).
    Next iCol

    '// Open the file using the specified column formats
    Workbooks.OpenText _
            FileName:=strFilepath, _
            DataType:=xlDelimited, _
            ConsecutiveDelimiter:=False, Comma:=True, _
            FieldInfo:=varColumnFormat

End Sub

