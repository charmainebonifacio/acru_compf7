Attribute VB_Name = "Step_1_Main"
Public objFSOlog As Object
Public logfile As TextStream
Public logtxt As String
Public appSTATUS As String
'---------------------------------------------------------------------------------------
' Date Created : July 6, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ACRU_COMPF7_MAIN
' Description  : This is the main function that ties in two other sub-main functions.
'                First, this function sets up the folders and validates the user input.
'                It then calls the ACRU_COMPF7_ProcessingZonalStat function to process
'                the .DBF files. After all the .OUT files have been created, it calls
'                on the ACRU_COMPF7_CompositeFile to create a new AB10K grid file and
'                a new composite file, of which both contains 7 variables.
'---------------------------------------------------------------------------------------
Function ACRU_COMPF7_MAIN()

    Dim start_time As Date, end_time As Date
    Dim ProcessingTime As Long
    Dim MessageSummary As String, SummaryTitle As String
    
    Dim UserSelectedFolder As String, DBFDIR As String
    Dim MAINFolder As String, compareIndex As Integer
    Dim PROGDIR As String, ABREFDIR As String
    Dim OUTDIR As String, OUTFDIR As String
    Dim ZSDIR As String, HADIR As String
    Dim BATDIR As String, CFDIR As String
    Dim TMPDIR As String, AB10KDIR As String
    Dim CopiedFiles As Long
    
    Dim MainOUT As String, ZSOUT As String, HAOUT As String
    Dim AB10KOUT As String, CFOUT As String, TMPOUT As String
    Dim BATOUT As String, ABREFIN As String
    Dim CheckABFolder As Boolean, CheckOUTFolder As Boolean
    Dim CheckZSFolder As Boolean, CheckHAFolder As Boolean
    Dim ResultCF As Integer
    Dim subARRAY() As String, outARRAY() As String
    Dim refIDArray() As String
    Dim refIndex As Integer
    
    ' Initialize Variables
    SummaryTitle = "Zonal Statistics Macro Diagnostic Summary"
    PROGDIR = "ACRU_COMPF7"
    ABREFDIR = "AB10KSource"
    OUTDIR = "Output"
    ZSDIR = "ZSOUT"
    HADIR = "HAOUT"
    BATDIR = "BATOUT"
    CFDIR = "CFOUT"
    TMPDIR = "TMPOUT"
    AB10KDIR = "AB10KOUT"
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    '---------------------------------------------------------------------
    ' I. USER INPUT
    '---------------------------------------------------------------------
    UserSelectedFolder = GetFolder
    Debug.Print UserSelectedFolder
    MAINFolder = ReturnFolderName(UserSelectedFolder)
    Debug.Print MAINFolder
    If Len(MAINFolder) = 0 Then GoTo Cancel:
    compareIndex = StrComp(PROGDIR, MAINFolder)
    If Not compareIndex = 0 Then GoTo Cancel:
    
    '---------------------------------------------------------------------
    ' II. USER INPUT VALIDATION
    ' Check for mandatory folders.
    '---------------------------------------------------------------------
    ABREFIN = ReturnSubFolder(UserSelectedFolder, ABREFDIR) ' Location of all the AB10KGrid Files
    Debug.Print ABREFIN
    CheckABFolder = CheckFolderExists(ABREFIN)

    MainOUT = ReturnSubFolder(UserSelectedFolder, OUTDIR)   ' Location of the Output folder
    Debug.Print MainOUT
    CheckOUTFolder = CheckFolderExists(MainOUT)
    
    ZSOUT = ReturnSubFolder(MainOUT, ZSDIR)
    Debug.Print ZSOUT
    CheckZSFolder = CheckFolderExists(ZSOUT)
    
    ' Setup Log File
    Dim logfilename As String, logtextfile As String, logext As String
    logext = ".txt"
    logfilename = "acru_compf7_log"
    logtextfile = SaveLogFile(MainOUT, logfilename, logext)
    
    Set objFSOlog = CreateObject("Scripting.FileSystemObject")
    Set logfile = objFSOlog.CreateTextFile(logtextfile, True)
    
    ' Maintain log starting from here
    logfile.WriteLine " [ Start of Program. ] "
    logfile.WriteLine "Selected directory: " & UserSelectedFolder
    logfile.WriteLine "Main directory: " & MAINFolder
    logfile.WriteLine "AB10Grid direcotry: " & ABREFIN
    logfile.WriteLine "Output direcotry: " & MainOUT
    logfile.WriteLine "Zonal statistcis directory: " & ZSOUT
    
    If CheckABFolder = False And CheckOUTFolder = False And CheckZSFolder = False Then GoTo Cancel

    '---------------------------------------------------------------------
    ' III. PROCESS ALL DBF FILES and LOOP for EACH SUBFOLDER in ZSOUT
    '---------------------------------------------------------------------
    Call GetSubFoldersArray(ZSOUT, subARRAY())
    For refIndex = LBound(subARRAY) To UBound(subARRAY)
        logtxt = "Directory Index: " & refIndex & "-" & subARRAY(refIndex)
        Debug.Print logtxt
        logfile.WriteLine logtxt
        If Len(subARRAY(refIndex)) = 0 Then GoTo Cancel
    Next refIndex
    Call WarningMessage
    start_time = Now()
    For refIndex = LBound(subARRAY) To UBound(subARRAY)
        DBFDIR = ReturnFolder(subARRAY(refIndex))
        Debug.Print DBFDIR
        logfile.WriteLine "Proceed to process .DBF files in " & DBFDIR
        Call ACRU_COMPF7_ProcessingZonalStat(DBFDIR, refIDArray(), refIndex)
    Next refIndex
    
    '---------------------------------------------------------------------
    ' IV. COPY ALL .OUT FILES to Temp Folder
    ' List the number of files copied. There should be at least one,
    ' depending on the sample size.
    '---------------------------------------------------------------------
    Call CreateNewFolder(MainOUT, TMPDIR) ' Create the temp location of all .OUT files
    logfile.WriteLine "Create the temp location of all .OUT files"
    TMPOUT = ReturnSubFolder(MainOUT, TMPDIR)
    HAOUT = ReturnSubFolder(MainOUT, HADIR)
    CheckHAFolder = CheckFolderExists(HAOUT)
    If CheckHAFolder = False Then GoTo Cancel
    logfile.WriteLine "HAOUT Folder exists."
    Call GetSubFoldersArray(HAOUT, outARRAY())
    For refIndex = LBound(outARRAY) To UBound(outARRAY)
        Debug.Print refIndex, outARRAY(refIndex)
        logfile.WriteLine refIndex & "-" & outARRAY(refIndex)
        If Len(outARRAY(refIndex)) = 0 Then GoTo Cancel
    Next refIndex
    For refIndex = LBound(outARRAY) To UBound(outARRAY)
        OUTFDIR = ReturnFolder(outARRAY(refIndex))
        CopiedFiles = CopyALLFiles(OUTFDIR, TMPOUT) ' Copy each folder contents first
    Next refIndex
    Debug.Print CopiedFiles
    logfile.WriteLine CopiedFiles & " of files were copied."
    If CopiedFiles = 0 Then GoTo Cancel ' No files were copied
    
    '---------------------------------------------------------------------
    ' IV. CREATE A COMPOSITE FILE for each file in SUBFOLDER in HAOUT
    '---------------------------------------------------------------------
    Call CreateNewFolder(MainOUT, AB10KDIR) ' Create the Composite File Directory
    AB10KOUT = ReturnSubFolder(MainOUT, AB10KDIR)
    Call CreateNewFolder(MainOUT, CFDIR)    ' Create the Composite File Directory
    CFOUT = ReturnSubFolder(MainOUT, CFDIR)
    ResultCF = ACRU_COMPF7_CompositeFile(refIDArray(), ABREFIN, TMPOUT, AB10KOUT, CFOUT)
    If ResultCF = 0 Then
        logtxt = "STATUS: There are NO missing source files."
        logfile.WriteLine logtxt
    Else ' Missing source files. No new composite files created.
        logtxt = "STATUS: There are missing source files. No new composite files created for the associated grid files."
        logfile.WriteLine logtxt
        'Call DeleteFolderAndContents(AB10KOUT)
        'logfile.WriteLine "Deleted folder and contents of AB10KOUT"
        'Call DeleteFolderAndContents(CFOUT)
        'logfile.WriteLine "Deleted folder and contents of CFOUT."
    End If
    '---------------------------------------------------------------------
    ' V. Clean up output directory by deleting TMPOUT and BATOUT folders.
    '---------------------------------------------------------------------
    BATOUT = ReturnSubFolder(MainOUT, BATDIR)
    Call DeleteFolderAndContents(BATOUT)
    logfile.WriteLine "Deleted folder and contents of BATOUT."
    Call DeleteFolderAndContents(TMPOUT)
    logfile.WriteLine "Deleted folder and contents of TMPOUT."
    logfile.WriteLine " [ End of Program. ] "
    
    ' End Program
    end_time = Now()
    ProcessingTime = DateDiff("n", CDate(start_time), CDate(end_time))
    MessageSummary = MacroTimer(ProcessingTime)
    MsgBox MessageSummary, vbOKOnly, SummaryTitle

    ' Close Log File
    logfile.Close
    Set logfile = Nothing
    Set objFSOlog = Nothing
    
Cancel:
    If CheckABFolder = False And CheckOUTFolder = False And CheckZSFolder = False Then
        logtxt = "Missing required folders. Please try again."
        MsgBox logtxt, vbOKOnly, SummaryTitle
        logfile.WriteLine logtxt
    End If
    If CheckHAFolder = False Then
        logfile.WriteLine "HAOUT Folder does not exist."
    End If
    If CopiedFiles = 0 Then
        logfile.WriteLine "No files were copied."
    End If
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 6, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ACRU_COMPF7_ProcessingZonalStat
' Description  : This function will process .DBF files into .DAT file. These .DAT files
'                will be used an input to the Harmonic Analysis program and will create
'                corresponding .OUT files.
'---------------------------------------------------------------------------------------
Function ACRU_COMPF7_ProcessingZonalStat(ByVal fileDir As String, ByRef refIDArray() As String, _
ByVal varIndex As Integer)
    
    Dim ZSFileType As String
    Dim VarZS As String
    Dim HarmonicFile As String
    Dim MeanVar As String
    Dim CodeID As String
    Dim HarmonicRun As Boolean
    
    MeanVar = "MEAN"
    CodeID = "AB_ID"
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    '---------------------------------------------------------------------
    ' I. Check For USER INPUT
    '---------------------------------------------------------------------
    appSTATUS = "[ In progress: Processing .dbf files. ]"
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    Call PROCESSDBFFILES(fileDir, VarZS)
    ZSFileType = ZonalStatFileFor(VarZS)
    Application.StatusBar = False
    
    '---------------------------------------------------------------------
    ' II. Validate USER INPUT
    '---------------------------------------------------------------------
    appSTATUS = "[ In progress: Compiling all data from zonal stats. ]"
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    Call SUMMARYWORKSHEET(fileDir, ZSFileType, CodeID, MeanVar, refIDArray(), varIndex)
    Application.StatusBar = False

    '---------------------------------------------------------------------
    ' III. Create .DAT files for each variable
    '---------------------------------------------------------------------
    appSTATUS = "[ In progress: Creating Master .DAT file. ]"
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    HarmonicFile = CREATEDATFILES(fileDir, ZSFileType)
    Application.StatusBar = False
    
    '---------------------------------------------------------------------
    ' IV. Create .OUT files for each row in the .DAT masterfile
    '---------------------------------------------------------------------
    appSTATUS = "[ In progress: Creating .OUT files using Harmonic Analysis. ]"
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    HarmonicRun = HARMONICANALYSIS(fileDir, HarmonicFile)
    If HarmonicRun = False Then
        logtxt = "Harmonic File does not exist. Check!"
        Debug.Print logtxt
        logfile.WriteLine logtxt
    End If
    Application.StatusBar = False

End Function
'---------------------------------------------------------------------------------------
' Date Created : June 6, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ACRU_COMPF7_CompositeFile
' Description  : This function will process the old AB10K grid files and .OUT files
'                in order to create the new composite file which contains 7 variables.
'---------------------------------------------------------------------------------------
Function ACRU_COMPF7_CompositeFile(ByRef refIDArray() As String, _
ByVal sourceTXTDIR As String, ByVal sourceOUTDIR As String, _
ByVal sourceAB10KDIR As String, ByVal sourceCFDIR As String) As Integer

    Dim TXTDIR As String
    Dim OUTDIR As String
    Dim AB10KDIR As String
    Dim CFDIR As String
    Dim ResultAB10K As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    ACRU_COMPF7_CompositeFile = True
    
    '---------------------------------------------------------------------
    ' I. Setup Folder strings to have "\" at the end... for efficiency
    ' purposes when it comes to putting two strings together
    ' sourceOUTDIR is the location of all the .OUT files
    ' sourceCFDIR is the final destination of all the new .TXT files
    '---------------------------------------------------------------------
    TXTDIR = ReturnFolder(sourceTXTDIR)
    OUTDIR = ReturnFolder(sourceOUTDIR)
    AB10KDIR = ReturnFolder(sourceAB10KDIR)
    CFDIR = ReturnFolder(sourceCFDIR)
    
    '---------------------------------------------------------------------
    ' II. Create a New AB10K Grid Files
    '---------------------------------------------------------------------
    appSTATUS = "[ In progress: New AB10K Grid Files: PRECIP, TMIN, TMAX, RAD, REL, SUN, and WND. ]"
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    ResultAB10K = ProcessReferenceID(refIDArray(), TXTDIR, OUTDIR, AB10KDIR)
    Application.StatusBar = False
    
    '---------------------------------------------------------------------
    ' III. Create a the final Composite Files
    '---------------------------------------------------------------------
    appSTATUS = "[ In progress: Creating new composite files. ]"
    Application.StatusBar = appSTATUS
    logfile.WriteLine appSTATUS
    Call ProcessCompositeFiles(AB10KDIR, CFDIR)
    Application.StatusBar = False

    ACRU_COMPF7_CompositeFile = ResultAB10K
    
End Function
