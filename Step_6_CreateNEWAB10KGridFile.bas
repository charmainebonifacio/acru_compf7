Attribute VB_Name = "Step_6_CreateNEWAB10KGridFile"
'---------------------------------------------------------------------
' Date Created : July 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 13, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopyALLFiles
' Description  : This function copies the .OUT folders into a temp
'                in order to process the new composite file.
' Parameters   : String, String
' Returns      : Long
'---------------------------------------------------------------------
Function CopyALLFiles(ByRef sourcePath As String, ByVal targetPath As String) As Long

    Dim objFSO As Object
    Dim objFolder As Object, objFolderDest As Object
    Dim FileCount As Long, FileCopiedCount As Long
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    On Error Resume Next

    ' Need to formate string
    If Right(sourcePath, 1) = "\" Then sourcePath = Left(sourcePath, Len(sourcePath) - 1)
    If Right(targetPath, 1) = "\" Then targetPath = Left(targetPath, Len(targetPath) - 1)
    
    ' Check if source file exist
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(sourcePath) = False Then
        Debug.Assert "File Does Not Exist or Path Not Found"
        CopyALLFiles = 0
        Exit Function
    End If

    '-------------------------------------------------------------
    ' Copy folder to another destination.
    '-------------------------------------------------------------
    objFSO.CopyFolder sourcePath, targetPath
    
    '-------------------------------------------------------------
    ' Loop through all the .OUT files and count the # of files
    ' in each folder. Verify the same # of files exist on the
    ' target path.
    '-------------------------------------------------------------
    FileCount = 0
    Set objFolder = objFSO.GetFolder(sourcePath).Files
    For Each objFILE In objFolder
      If UCase(Right(objFILE.Path, (Len(objFILE.Path) - InStrRev(objFILE.Path, ".")))) = UCase("out") Then
        FileCount = FileCount + 1
      End If
    Next
    Debug.Print "There are " & FileCount & " .OUT Files to be copied."
    
    Set objFolderDest = objFSO.GetFolder(targetPath).Files
    FileCopiedCount = 0
    For Each objFILEDEST In objFolderDest
      If UCase(Right(objFILEDEST.Path, (Len(objFILEDEST.Path) - InStrRev(objFILEDEST.Path, ".")))) = UCase("out") Then
        FileCopiedCount = FileCopiedCount + 1
      End If
    Next
    
    logtxt = "SuccessFull Copied " & FileCopiedCount & " .OUT Files."
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
    CopyALLFiles = FileCopiedCount
    
End Function
'---------------------------------------------------------------------
' Date Created : July 16, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ProcessReferenceID
' Description  : This function opens the reference table and
'                calls another function to initialize two arrays.
' Parameters   : String
' Returns      : -
'---------------------------------------------------------------------
Function ProcessReferenceID(ByRef refIDArray() As String, ByVal sTXTDIR As String, _
ByVal sOUTDIR As String, ByVal oAB10KDIR As String) As Integer
    
    Dim refIndex As Long
    Dim TxtFile As String
    Dim SourceFilePath As String
    Dim objFSO As Object
    Dim missingFilesCount As Integer
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    missingFilesCount = 0
    
    For refIndex = LBound(refIDArray) To UBound(refIDArray)
        '-------------------------------------------------------------
        ' Array contains the name of the file to be processed.
        '-------------------------------------------------------------
        TxtFile = refIDArray(refIndex) & ".txt"
        logfile.WriteLine "Processing Grid File: " & TxtFile
        SourceFilePath = sTXTDIR & TxtFile
        If Right(sTXTDIR, 1) <> "\" Then SourceFilePath = sTXTDIR & "\" & TxtFile
        
        '-------------------------------------------------------------
        ' Check whether the file exists or not under the reference
        ' directory which contains all the renamed AB10K grid files.
        '-------------------------------------------------------------
        If objFSO.fileExists(SourceFilePath) = False Then
            logtxt = "STATUS: Source File Does Not Exist or Path Not Found"
            Debug.Print logtxt
            logfile.WriteLine logtxt
            Exit Function
        End If
        '-------------------------------------------------------------
        ' Proceed to process the old .TXT files.
        '-------------------------------------------------------------
        If Process3VarTXTFiles(SourceFilePath, sTXTDIR, sOUTDIR, oAB10KDIR) = False Then
            logtxt = "STATUS: Missing source .OUT files. Unable to proceed in creating the grid file."
            logfile.WriteLine logtxt
            missingFilesCount = missingFilesCount + 1
        End If
    Next refIndex
    
    ProcessReferenceID = missingFilesCount

End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Process3VarTXTFiles
' Description  : This function processes the AB10K grid files and
'                appends RAD, REL, SUN and WND values.
' Parameters   : String, String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------
Function Process3VarTXTFiles(ByVal oFileName As String, ByVal TXTDIR As String, _
ByVal OUTDIR As String, ByVal AB10KDIR As String) As Boolean

    Dim wbMaster As Workbook, MasterSht As Worksheet
    Dim wbOrig As Workbook, OrigSheet As Worksheet
    Dim TxtFile As String
    Dim RC As Long, CC As Long
    Dim origName As String
    Dim fDataType As Integer
    Dim textFileArray() As String
    Dim sourceFilesArray() As String
    Dim TimeseriesCheck As Boolean
    Dim TimeSeriesStart As Long, TimeSeriesEnd As Long

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Process3VarTXTFiles = True
    
    logfile.WriteLine "Appending RAD, REL, SUN and WND values to 10KM grid files."
    '-------------------------------------------------------------
    ' Loop through all the .txt files
    '-------------------------------------------------------------
    Set wbOrig = Workbooks.Open(oFileName)
    Set OrigSheet = wbOrig.Worksheets(1)
    '-------------------------------------------------------------
    ' Pass on the .TXT filename.
    '-------------------------------------------------------------
    origName = OrigSheet.Name
    TxtFile = reDefineName(origName)
    '-------------------------------------------------------------
    ' Setup Master file.
    '-------------------------------------------------------------
    TimeseriesCheck = ValidateCompositeTimeseries(wbOrig, OrigSheet, _
        RC, CC, textFileArray(), TimeSeriesStart, TimeSeriesEnd)
    If TimeSeriesStart = 0 Or TimeSeriesEnd = 0 Then GoTo Cancel:
    Set wbMaster = Workbooks.Add(1)
    Set MasterSht = wbMaster.Worksheets(1)
    MasterSht.Name = TxtFile
    
    '-------------------------------------------------------------
    ' Copy All Data To Master file.
    '-------------------------------------------------------------
    Call CopyCorrectTimeSeries(OrigSheet, MasterSht, TimeSeriesStart, TimeSeriesEnd, CC)
    wbOrig.Close SaveChanges:=False 'DO NOT SAVE ANY CHANGES
    
    '-------------------------------------------------------------
    ' Process the corresponding four files.
    ' Then save the final worksheet.
    '-------------------------------------------------------------
    If CallOUTFiles(OUTDIR, TxtFile, MasterSht, sourceFilesArray()) = True Then
        Call ProcessOUTFiles(MasterSht, OUTDIR, textFileArray(), sourceFilesArray())
        Call SaveTXT(wbMaster, MasterSht, AB10KDIR, TxtFile)
        logfile.WriteLine "Succesfully saved file." & TxtFile
    Else
        Process3VarTXTFiles = False
        logtxt = "One of the source files does not exist. No grid file was saved for ." & TxtFile
        logfile.WriteLine logtxt
    End If
    wbMaster.Close SaveChanges:=False

Cancel:
    Set wbOrig = Nothing
    Set OrigSheet = Nothing
    Set wbMaster = Nothing
    Set MasterSht = Nothing
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 5, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateDateColumn
' Description  : This function creates a date column based on the
'                Year, Month and Day columns of the timeseries which
'                can be found in Columns A:C.
' Parameters   : Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function ValidateCompositeTimeseries(ByRef wbTmp As Workbook, ByRef TmpSheet As Worksheet, _
ByRef RC As Long, ByRef CC As Long, ByRef tableRefArray() As String, ByRef SeriesStart As Long, _
ByRef SeriesEnd As Long) As Boolean

    Dim RowCount As Boolean, ColCount As Boolean
    
    TmpSheet.Activate
    ValidateCompositeTimeseries = True
    
    Call CreateDateColumn(TmpSheet, RC, CC)
    Call InitializeDateArray(tableRefArray())
    Call DataCheckValidation(TmpSheet, tableRefArray(), SeriesStart, SeriesEnd)
    
    If SeriesStart = 0 Then ColCount = False
    If SeriesEnd = 0 Then RowCount = False
    If RowCount = False And ColCount = False Then ValidateCompositeTimeseries = False
    
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 5, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateDateColumn
' Description  : This function creates a date column based on the
'                Year, Month and Day columns of the timeseries which
'                can be found in Columns A:C.
' Parameters   : Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function CreateDateColumn(DestSht As Worksheet, ByRef RC As Long, ByRef CC As Long)

    Dim LC As Long, LR As Long
    Dim ColCount As Integer, RowCount As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Activate Source Worksheet.
    DestSht.Activate
        
    Call FindLastRowColumn(LR, LC)
    RowCount = CInt(LR)
    ColCount = CInt(LC)

    Range("A1").Offset(0, ColCount).Select
    If ColCount >= 4 Then
        ActiveCell.FormulaR1C1 = "=DATE(RC[-" & ColCount & "],RC[-" & ColCount - 1 & "],RC[-" & ColCount - 2 & "])"
        Range(Cells(1, ColCount + 1), Cells(RowCount, ColCount + 1)).FillDown
    End If
    RC = RowCount
    CC = ColCount
    
End Function
'---------------------------------------------------------------------
' Date Created : July 5, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 11, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : InitializeDateArray
' Description  : This function initializes an arrays which contains
'                date reference.
' Parameters   : String Array
' Returns      : -
'---------------------------------------------------------------------
Function InitializeDateArray(ByRef tableRefArray() As String)

    Dim rACells As Range, rLoopCells As Range
    Dim refCellValue As String
    Dim refLat As Double, refLng As Double
    Dim indexArr As Integer, refIndex As Integer
    Dim LC As Long, LR As Long
        
    Call FindLastRowColumn(LR, LC)
    NewLR = LR - 1                ' No header!
    ReDim tableRefArray(NewLR, 1)
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Copy only the date column!
    If WorksheetFunction.CountA(Cells) > 0 Then
        Range(Cells(1, LC), Cells(LR, LC)).Select
    End If
    Set rACells = Selection
    On Error Resume Next 'In case of NO text constants.
    
    ' Set variable to all text constants
    Set rACells = rACells.SpecialCells(xlCellTypeConstants, xlTextValues)
    MsgBox rACells
        
    ' If could not find any text
    If rACells Is Nothing Then
        MsgBox "Could not find any text."
        On Error GoTo 0
        Exit Function
    End If
    
    indexArr = 0
    For Each rLoopCells In rACells
        refCellValue = rLoopCells.Value
        tableRefArray(indexArr, 0) = refCellValue
        tableRefArray(indexArr, 1) = Range(rLoopCells.Address).Row
        indexArr = indexArr + 1
    Next rLoopCells
        
End Function
'---------------------------------------------------------------------
' Date Created : June 28, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 15, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : DataCheckValidation
' Description  : This function checks the data timeseries for the
'                original AB10K Grid source file. It must start with
'                1950 and end in 2010.
' Parameters   : Worksheet, String, String, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function DataCheckValidation(TmpSht As Worksheet, ByRef tableRefArray() As String, _
ByRef SeriesStart As Long, ByRef SeriesEnd As Long)

    Dim DataInput As Date
    Dim DataValue As String
    Dim DateMonth As Long, DateDay As Long
    Dim refIndex As Integer
    Dim StartDate As Date, EndDate As Date
    Dim DefaultStart As String, DefaultEnd As String
    Dim TimeSeriesStartYear As String, TimeSeriesEndYear As String
    Dim TimeSeriesStart As Date, TimeSeriesEnd As Date
    Dim StartTimeSeries As Long, EndTimeSeries As Long
        
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    TmpSht.Activate
    
    ' Initialize Values
    StartTimeSeries = 0
    EndTimeSeries = 0
    TimeSeriesStartYear = "1950"
    TimeSeriesEndYear = "2010"
    DefaultStart = "01/01/" & TimeSeriesStartYear
    DefaultEnd = "12/31/" & TimeSeriesEndYear
    StartDate = DateValue(DefaultStart)
    EndDate = DateValue(DefaultEnd)
    
    For refIndex = LBound(tableRefArray) To UBound(tableRefArray)
        Debug.Print refIndex, tableRefArray(refIndex, 0)
        TimeSeries = DateValue(tableRefArray(refIndex, 0))
        If StartTimeSeries = 0 Then ' Find the First January 1
            DateMonth = DateDiff("m", StartDate, TimeSeries)
            DateDay = DateDiff("d", StartDate, TimeSeries)
            If DateMonth = 0 And DateDay = 0 Then
                StartTimeSeries = tableRefArray(refIndex, 1)
            End If
        End If
        If StartTimeSeries > 0 Then ' Check End Now
            DateMonth = DateDiff("m", EndDate, TimeSeries)
            DateDay = DateDiff("d", EndDate, TimeSeries)
            If DateMonth = 0 And DateDay = 0 Then
                EndTimeSeries = tableRefArray(refIndex, 1)
            End If
        End If
    Next refIndex
    
    If StartTimeSeries < EndTimeSeries Then
        Debug.Print "January 1 : ", StartTimeSeries
        Debug.Print "December 31: ", EndTimeSeries
    End If
    
    SeriesStart = StartTimeSeries
    SeriesEnd = EndTimeSeries
    
End Function
'---------------------------------------------------------------------
' Date Created : June 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 5, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SplitTimeSeries
' Description  : This function copies only the necessary data based on
'                the first row count and last row count.
' Parameters   : Worksheet, Worksheet, Long, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function CopyCorrectTimeSeries(SourceSht As Worksheet, TmpSheet As Worksheet, _
ByVal FirstRowCount As Long, ByVal LastRowCount As Long, ByVal ColCount As Long)
    
    Dim RngSelect
    Dim PasteSelect
    Dim tmpWB As Workbook
    Dim TmpSheet1 As Worksheet
    Dim TmpSheet2 As Worksheet
    Dim FirstRow As Long, FirstCol As Long
    Dim LastRow As Long, LastCol As Long

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
    ' Activate Source Worksheet.
    SourceSht.Activate

    FirstRow = FirstRowCount
    FirstCol = 1
    LastCol = ColCount
    LastRow = LastRowCount

    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    Call FindSpecificRange(SourceSht, FirstRow, FirstCol, LastRow, LastCol)
    RngSelect = Selection.Address
    Range(RngSelect).Copy

    '-------------------------------------------------------------
    ' Activate Appropriate Temp Worksheet.
    '-------------------------------------------------------------
    TmpSheet.Activate

    '-------------------------------------------------------------
    ' Call RowCheck function to check the last row.
    ' Then append the copied data into the Temp Worksheet.
    '-------------------------------------------------------------
    Range("A1").Select
    PasteSelect = Selection.Address
    Range(PasteSelect).Select
    TmpSheet.Paste
    
    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False
    
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : September 15, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CallOUTFiles
' Description  : This function redefines the filename and checks
'                whether the file exists or not. If it does, then
'                further processing is done. Otherwise, function
'                must end.
' Parameters   : String, String
' Returns      : -
'---------------------------------------------------------------------
Function CallOUTFiles(ByRef fileDir As String, ByVal fileID As String, _
MasterSht As Worksheet, ByRef sourceFilesArray() As String) As Boolean

    Dim fDataType As Integer
    Dim dataName As String
    Dim fileExistTrue As Boolean
    Dim dataFileName As String
    Dim sourceFileIndex As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    CallOUTFiles = True
    
    ReDim sourceFilesArray(3)
    sourceFileIndex = 0
    
    logfile.WriteLine "Checking which of the 4 .OUT files exists."
    
    For fDataType = 1 To 4
        '-------------------------------------------------------------
        ' Redefine the filename
        '-------------------------------------------------------------
        dataName = fileDataType(fDataType)
        dataFileName = dataName & fileID & ".out"
        
        Debug.Print "Prefix: " & dataName
        Debug.Print "File ID Reference: " & fileID
        Debug.Print "Filename: " & dataFileName
        
        logfile.WriteLine "Prefix: " & dataName
        logfile.WriteLine "File ID Reference: " & fileID
        logfile.WriteLine "Filename: " & dataFileName
        
        '-------------------------------------------------------------
        ' Check whether the file exists or not
        '-------------------------------------------------------------
        fileExistTrue = CheckFileExists(fileDir, dataFileName)
        Debug.Print "File existence status: " & fileExistTrue
        logfile.WriteLine "File existence status: " & fileExistTrue
        
        If fileExistTrue = True Then
            sourceFilesArray(sourceFileIndex) = dataFileName
            sourceFileIndex = sourceFileIndex + 1
        Else:
            CallOUTFiles = False
            Exit For
        End If
    Next fDataType
        
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 26, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : reDefineName
' Description  : This function changes the string and includes a
'                leading zeroes.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------
Function reDefineName(fName As String) As String

    Dim Temp As String
    Dim fLen As Integer
    fLen = Len(fName)
    Select Case fLen
        Case 1
            reDefineName = "000" & fName
        Case 2
            reDefineName = "00" & fName
        Case 3
            reDefineName = "0" & fName
        Case 4
            reDefineName = fName
    End Select
    
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ProcessOUTFiles
' Description  : This function opens the source files, creates a date
'                column for comparison.
' Parameters   : Worksheet, String, String Array, Date Array
' Returns      : -
'---------------------------------------------------------------------
Function ProcessOUTFiles(MasterSht As Worksheet, ByVal fileDir As String, _
ByRef textFileArray() As String, ByRef sourceFilesArray() As String)
    Dim wbSource As Workbook
    Dim SourceSheet As Worksheet
    Dim sThisFilePath As String, fileToOpen As String
    Dim fileIndex As Integer
    Dim TxtFile As String
    Dim arrText()
    Dim RC As Long, CC As Long
    Dim TimeseriesCheck As Boolean
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    '-------------------------------------------------------------
    ' Open Specific .OUT file and parse the data appropriately
    '-------------------------------------------------------------
    For fileIndex = LBound(sourceFilesArray) To UBound(sourceFilesArray)
        arrText = Array(Array(0, 1), Array(4, 1), Array(7, 1), Array(12, 1))
        fileToOpen = fileDir & sourceFilesArray(fileIndex)
        Debug.Print fileToOpen
        Workbooks.OpenText fileName:=fileToOpen, Origin:=xlMSDOS, StartRow:=1, _
            dataType:=xlFixedWidth, FieldInfo:=arrText, TrailingMinusNumbers:=True
        Set wbSource = ActiveWorkbook
        Set SourceSheet = wbSource.Worksheets(1)
        
        ' Harmonic Data Check
        Call HarmonicDataCheck(SourceSheet, fileIndex)
        '-------------------------------------------------------------
        ' Add date column
        '-------------------------------------------------------------
        Call CreateDateColumn(SourceSheet, RC, CC)
        '-------------------------------------------------------------
        ' Compare timeseries date for the first variable
        ' After that, copy the source data
        '-------------------------------------------------------------
        If fileIndex = 0 Then
            TimeseriesCheck = CompareTimeseries(SourceSheet, MasterSht, textFileArray, RC, CC)
            If TimeseriesCheck = True Then Call CopySourceData(SourceSheet, MasterSht, RC, CC)  ' Copy Source Data
        Else: Call CopySourceData(SourceSheet, MasterSht, RC, CC)  ' Copy Source Data
        End If
        wbSource.Close SaveChanges:=False
    Next fileIndex

Cancel:
    Set wbSource = Nothing
    Set SourceSheet = Nothing
End Function
'---------------------------------------------------------------------
' Date Created : July 23, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 23, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HarmonicDataCheck
' Description  : This function checks the data value. If REL > 99 then
'                REL = 99, REL < 5 then REL = 5, RAD < 0 then RAD = 0
'                (same thing with SUN and WIND)
' Parameters   : Worksheet, Integer
' Returns      : -
'---------------------------------------------------------------------
Function HarmonicDataCheck(SourceSht As Worksheet, ByVal fileIndex As Integer)

    Dim rACells As Range, rLoopCells As Range
    Dim refCellValue As String
    Dim refLat As Double, refLng As Double
    Dim indexArr As Integer, refIndex As Integer
    Dim valArray() As Double
    Dim LC As Long, LR As Long
    Dim ColCount As Integer, RowCount As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Activate Source Worksheet.
    SourceSht.Activate
        
    Call FindLastRowColumn(LR, LC)
    RowCount = CInt(LR)
    ColCount = CInt(LC)
    NewLR = LR - 1                ' No header!
    ReDim valArray(NewLR)
    
    ' Only the last column!
    If WorksheetFunction.CountA(Cells) > 0 Then
        Range(Cells(1, LC), Cells(LR, LC)).Select
    End If
    
    On Error Resume Next 'In case of NO text constants.
    
    ' Set variable to all text constants
    Set rACells = Selection
        
    ' If could not find any text
    If rACells Is Nothing Then
        MsgBox "Could not find any text."
        On Error GoTo 0
        Exit Function
    End If
    
    ' Copy all values into the array for easier comparison
    indexArr = 0
    For Each rLoopCells In rACells
        refCellValue = rLoopCells.Value
        valArray(indexArr) = CDbl(refCellValue)
        indexArr = indexArr + 1
    Next rLoopCells
    
    ' Check array value and change value if condition
    ' is met. Place this value next to the original
    ' column.
    Range("A1").Offset(0, ColCount).Select
    For Z = LBound(valArray) To UBound(valArray)
        Select Case fileIndex
            Case 0, 2, 3: ' RAD, SUN, WND
                If valArray(Z) < 0 Then valArray(Z) = 0
            Case 1:       ' REL
                If valArray(Z) > 99 Then valArray(Z) = 99
                If valArray(Z) < 5 Then valArray(Z) = 5
        End Select
        Range("A1").Offset(Z, ColCount).Value = valArray(Z)
    Next

End Function
'---------------------------------------------------------------------
' Date Created : June 28, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 23, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : -
' Description  : This function creates a date column based on the
'                Year, Month and Day columns of the timeseries which
'                can be found in Columns A:C.
' Parameters   : Worksheet, Long, String, String, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function CompareTimeseries(SourceSht As Worksheet, DestSht As Worksheet, _
ByRef textFileArray() As String, ByVal RC As Long, ByVal CC As Long) As Boolean

    Dim rACells As Range, rLoopCells As Range
    Dim rCellValue As String
    Dim refDate As Date, dateTextFile As Date
    Dim LC As Long
    Dim refIndex As Integer
    Dim CheckY As Long, CheckM As Long, CheckD As Long

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    CompareTimeseries = False
    
    ' Initialize Variables
    refIndex = 0
    CheckY = 1
    CheckM = 1
    CheckD = 1
    LC = CC + 1
    SourceSht.Activate
   
    If WorksheetFunction.CountA(Cells) > 0 Then
        Range(Cells(1, LC), Cells(RC, LC)).Select
    End If
    Set rACells = Selection
    
    On Error Resume Next 'In case of NO text constants.
    
    ' Set variable to all text constants
    Set rACells = Selection
        
    ' If could not find any text
    If rACells Is Nothing Then
        MsgBox "Could not find any text."
        On Error GoTo 0
        Exit Function
    End If
    
    For Each rLoopCells In rACells
        rCellValue = rLoopCells.Value
        Debug.Print rCellValue
        refDate = DateValue(rCellValue)
        dateTextFile = DateValue(textFileArray(refIndex, 0))
        CheckY = DateDiff("yyyy", dateTextFile, refDate)
        CheckM = DateDiff("m", dateTextFile, refDate)
        CheckD = DateDiff("d", dateTextFile, refDate)
        Debug.Print CheckY, CheckM, CheckD
        If (CheckY = 0 And CheckM = 0 And CheckD = 0) Then '"Found match"
            refIndex = refIndex + 1
        Else: Debug.Print "Not the right date."
        End If
    Next rLoopCells
    
    Debug.Print refIndex
    If RC = refIndex Then CompareTimeseries = True

End Function
'---------------------------------------------------------------------
' Date Created : July 2, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 5, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopySourceData
' Description  : This function copies data from the source files.
' Parameters   : Worksheet, Worksheet, Long, Long
' Returns      : -
'---------------------------------------------------------------------
Function CopySourceData(SourceSht As Worksheet, DestSht As Worksheet, _
ByVal RC As Long, ByVal CC As Long)
    
    Dim RngSelect
    Dim PasteSelect

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Activate Source Worksheet.
    SourceSht.Activate
      
    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    Call FindSourceRange(SourceSht, RC, CC)
    RngSelect = Selection.Address
    Range(RngSelect).Copy

    ' Activate Destination Worksheet.
    DestSht.Activate
    
    '-------------------------------------------------------------
    ' Call RowCheck function to check the last row.
    ' Then append the copied data into the Destination Worksheet.
    '-------------------------------------------------------------
    Call ColumnCheck(DestSht)
    PasteSelect = Selection.Address
    Range(PasteSelect).Select
    DestSht.Paste
    
    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False
    
End Function
'---------------------------------------------------------------------
' Date Created : June 28, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 28, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindSourceRange
' Description  : This function copies the values from the harmonic
'                analysis.
' Parameters   : Worksheet, Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function FindSourceRange(WKSheet As Worksheet, ByVal RC As Long, ByVal CC As Long)

    Dim FirstRow&, FirstCol&, LastRow&, LastCol&
    Dim myUsedRange As Range
        
    ' Activate the correct worksheet
    WKSheet.Activate
    
    ' Define variables
    FirstRow = 1
    FirstCol = CC
    LastRow = RC
    LastCol = CC
    
    ' Select Range using FirstRow, FirstCol, LastRow, LastCol
    Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
    myUsedRange.Select
    
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 26, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MainFunction
' Description  : This function will return a specific string depending
'                on the value passed into the function.
' Parameters   : Integer
' Returns      : String
'---------------------------------------------------------------------
Function fileDataType(ByVal dType As Integer) As String
    
    Dim Temp As String
    Select Case dType
        Case 1
            Debug.Print "RADIATION."
            Temp = "RAD_"
        Case 2
            Debug.Print "RELATIVE HUMIDITY."
            Temp = "REL_"
        Case 3
            Debug.Print "SUNSHINE HOURS."
            Temp = "SUN_"
        Case 4
            Debug.Print "WIND SPEED."
            Temp = "WND_"
    End Select
    
    fileDataType = Temp
    
End Function

