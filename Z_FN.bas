Attribute VB_Name = "Z_FN"
Option Explicit

Sub headerChecker(listOfHeaders, wbs)
    Dim unit        As Variant
    For Each unit In listOfHeaders
        If SearchColumnSource(unit, wbs) = 0 Then
            MsgBox ("Missing <" & unit & "> Header")
            reset
            End
        End If
    Next
    
End Sub

Sub formatColumn(column, wbs)
    Dim LastxRow    As Integer
    Dim formatRangeStart, formatRangeEnd As String
    Dim formatCell  As Variant
    
    LastxRow = wbs.Cells(Rows.count, column).End(xlUp).Row
    formatRangeStart = Split(Cells(1, column).Address, "$")(1) & 1
    formatRangeEnd = Split(Cells(1, column).Address, "$")(1) & 385
    
    Formatter (wbs.range(formatRangeStart, formatRangeEnd))
    
    'For Each formatCell In wbs.range(formatRangeStart, formatRangeEnd)
    '    Formatter (formatCell)
    'Next
    
End Sub

Sub turnOff()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Sub reset()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


Sub Formatter(range)
    With range
        'Adds border edges
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        
        .FormatConditions.Delete
        
        'Remove all leading decimals
        .NumberFormat = "0"
        
        'Center contents
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    
End Sub

Function cleanInput(inputText) As String
    cleanInput = LCase(inputText)
    cleanInput = Replace(cleanInput, " ", "")        ' Kill Ghost "space"
    cleanInput = Replace(cleanInput, " ", "")        ' Regular Space
    cleanInput = Replace(cleanInput, "_", "")
    cleanInput = Replace(cleanInput, "-", "")
End Function

Function checkAnalysisSSHeaders(fso) As Boolean
    Dim anBook      As Workbook
    Dim temp        As Variant
    Dim arrayMessage As String
    checkAnalysisSSHeaders = False
    For Each temp In fso
        If InStr(temp, "~") < 1 Then
            Workbooks.OpenText FileName:=temp, Local:=True
            Set anBook = ActiveWorkbook
            If Not SearchColumnSource("pcr1", anBook) = 0 Then
                
                If SearchColumnSource("SS", anBook) = 0 Then
                    arrayMessage = arrayMessage & "ss header missing from " & Left(anBook.name, 20) & vbNewLine
                    checkAnalysisSSHeaders = True
                End If
                
                If SearchColumnSource("samplename", anBook) = 0 Then
                    arrayMessage = arrayMessage & "Sample Name header missing from " & Left(anBook.name, 20) & vbNewLine
                    checkAnalysisSSHeaders = True
                End If
                
                If SearchColumnSource("gm", anBook) = 0 Then
                    arrayMessage = arrayMessage & "GM header missing from " & Left(anBook.name, 20) & vbNewLine
                    checkAnalysisSSHeaders = True
                End If
                
                If SearchColumnSource("seq", anBook) = 0 Then
                    arrayMessage = arrayMessage & "SEQ header missing from " & Left(anBook.name, 20) & vbNewLine
                    checkAnalysisSSHeaders = True
                    
                End If
            End If
            
            If Not SearchColumnSource("pcr2", anBook) = 0 Then
                If SearchColumnSource("gm2", anBook) = 0 Then
                    arrayMessage = arrayMessage & "GM2 header missing from " & Left(anBook.name, 20) & vbNewLine
                    checkAnalysisSSHeaders = True
                End If
                
                If SearchColumnSource("seq2", anBook) = 0 Then
                    arrayMessage = arrayMessage & "SEQ2 header missing from " & Left(anBook.name, 20) & vbNewLine
                    checkAnalysisSSHeaders = True
                End If
            End If
            
            anBook.Close savechanges:=False
        End If
    Next
    
    If checkAnalysisSSHeaders Then
        MsgBox (arrayMessage & vbNewLine & "Please Correct this, Ending Program")
        End
    End If
    
End Function

'This function will look for the names in the Source File and return the index of that column. Looks for the Last Instance
Public Function SearchColumnSource(search, wb) As Integer
    Dim columnIndex As Integer
    Dim field       As Variant
    
    columnIndex = 1
    For Each field In wb.Sheets(1).range("A1", "AZ1")
        If cleanInput(field) = cleanInput(search) Then
            SearchColumnSource = columnIndex
        End If
        columnIndex = columnIndex + 1
    Next
End Function

'Checks if column exisit, if not will add to the last column.
Sub addLastcolumn(book, title)
    Dim AnLCol      As Integer
    If SearchColumnSource(title, book) = 0 Then
        AnLCol = book.Sheets(1).UsedRange.Columns.count
        book.Sheets(1).range(Split(Cells(1, AnLCol + 1).Address, "$")(1) & ":" & Split(Cells(1, AnLCol + 1).Address, "$")(1)).EntireColumn.insert
        book.Sheets(1).Cells(1, AnLCol + 1) = title
    End If
End Sub

Sub checkVersion()
    
    Dim Version, Program As String
    Set Version = 1.4
    Set Program = "LaragenAutomation"
    
    Dim versionFolder As String
    Dim vFSOlibrary, vFSOFolder, vFSOFile As Object
    versionFolder = "READ_Me.Cells(15, 2)"
    Dim vfile       As Variant
    Set vFSOlibrary = New FileSystemObject
    Set vFSOFolder = vFSOlibrary.GetFolder(versionFolder)
    Set vFSOFile = vFSOFolder.Files
    For Each vfile In vFSOFile
        If Split(vfile.name, "_")(0) = Program Then
            
            If Replace(Split(vfile.name, "_")(2), ".txt", "") = ver Then
                'Good, continue with program
            Else:
                MsgBox ("Warning!!! Your " & Program & " version " & ver & " Is out of date. Please use version " & Replace(Split(vfile.name, "_")(2), ".txt", "") & ". The program will Not run.")
                End
            End If
        Else
        End If
    Next
End Sub

Sub CreateTextFileBasic(title, text, mainWB)
    'Must add reference to Tools > References > Microsoft Scripting Runtime
    ' (1) Open an instance of the FileSystemObject.
    ' (2) Open an instance of the FileSystemObject TextStream.
    ' (3) Write two example lines of text to the TextStream object.
    ' (4) Close the TextStream object.
    
    Dim fso         As Scripting.FileSystemObject
    Dim tsTxtFile   As Scripting.TextStream
    
    Dim timeDate    As String
    timeDate = Format(Now, "mmm_dd_hhmm")
    
    ' (1) Open an instance of the FileSystemObject.
    Set fso = New Scripting.FileSystemObject
    
    ' (2) Open an instance of the FileSystemObject TextStream class.
    Set tsTxtFile = fso.CreateTextFile(mainWB.range("B20") & title & "_" & timeDate & ".txt", True)
    
    ' (3) Write two example lines of text to the TextStream object.
    tsTxtFile.WriteLine text
    
    ' (4) Close the textstream object
    tsTxtFile.Close
    
End Sub

Sub sortBook(aWB, title)
    Dim rtIndexcol  As Integer
    Dim rtIndexAdress As String
    
    rtIndexcol = SearchColumnSource(title, aWB)
    rtIndexAdress = Split(Cells(1, rtIndexcol).Address, "$")(1) & 1
    
    aWB.Sheets(1).range("A1:BZ800").Sort Key1:=range(rtIndexAdress), Order1:=xlAscending, header:=xlYes
End Sub
