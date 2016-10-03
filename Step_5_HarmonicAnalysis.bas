Attribute VB_Name = "Step_5_HarmonicAnalysis"
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
'---------------------------------------------------------------------
' Date Acquired: July 4, 2013
' Source       : http://www.vbaexpress.com/forum/showthread.php?t=37457
'---------------------------------------------------------------------
' Date Edited  : July 4, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Shell_AndWait
' Description  : This function stops the application in its tracks --
'                it doesn't repond to keyboard, mouse, etc until
'                the shelled app is finished.
' Parameters   : String, ShellWindow, Long
' Returns      : Boolean
'---------------------------------------------------------------------
Public Enum ShellTiming
    SH_IGNORE = 0 'Ignore signal
    SH_INFINITE = -1& 'Infinite timeout
    SH_PROCESS_QUERY_INFORMATION = &H400
    SH_STILL_ACTIVE = &H103
    SH_SYNCHRONIZE = &H100000
End Enum
Public Enum ShellWait
    SH_WAIT_ABANDONED = &H80&
    SH_WAIT_FAILED = -1& 'Error on call
    SH_WAIT_OBJECT_0 = 0 'Normal completion
    SH_WAIT_TIMEOUT = &H102& 'Timeout period elapsed
End Enum
Public Enum ShellWindow
    SH_HIDE = 0
    SH_SHOWNORMAL = 1 'normal with focus
    SH_SHOWMINIMIZED = 2 'minimized with focus (default in VB)
    SH_SHOWMAXIMIZED = 3 'maximized with focus
    SH_SHOWNOACTIVATE = 4 'normal without focus
    SH_SHOW = 5 'normal with focus
    SH_MINIMIZE = 6 'minimized without focus
    SH_SHOWMINNOACTIVE = 7 'minimized without focus
    SH_SHOWNA = 8 'normal without focus
    SH_RESTORE = 9 'normal with focus
End Enum
Function Shell_AndWait(ByVal CommandLine As String, _
    Optional ExecMode As ShellWindow = SH_HIDE, _
    Optional Timeout As Long = SH_INFINITE) As Boolean
    Dim ProcessID As Long
    Dim hProcess As Long
    Dim nRet As Long
    Const fdwAccess = SH_SYNCHRONIZE
    If ExecMode < SH_HIDE Or ExecMode > SH_RESTORE Then ExecMode = SH_SHOWNORMAL
    ProcessID = Shell(CommandLine, CLng(ExecMode))
    hProcess = OpenProcess(fdwAccess, False, ProcessID)
    nRet = WaitForSingleObject(hProcess, CLng(Timeout))
    Shell_AndWait = (nRet <> 0)
End Function
'---------------------------------------------------------------------
' Date Created : July 8, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HARMONICANALYSIS
' Description  : This function finds the necesary directory and file
'                in order to create a batch file that will call the
'                fortran program 'Harmonic Analysis'. All .OUT files
'                will then be copied from the C:\Harmonic directory.
' Parameters   : String, String
' Returns      : Boolean
'---------------------------------------------------------------------
Function HARMONICANALYSIS(ByVal fileDir As String, ByVal HarmonicFile As String) As Boolean

    Dim PROGLoc As String
    Dim DATLoc As String
    Dim OUTDIR As String
    Dim HARMONICDIR As String
    Dim MOUTdir As String
    Dim BATPreFile As String
    Dim exeFile As String
    
    HARMONICDIR = "C:\Harmonic"
    PROGLoc = FindRootFolder(fileDir, OUTDIR, MOUTdir)
    DATLoc = FindHarmonicFile(HarmonicFile)
    If Len(PROGLoc) = 0 Or Len(DATLoc) = 0 Then
        HARMONICANALYSIS = False
        Exit Function
    End If
    
    exeFile = "ACRU_COMPF7_HarmonicAnalysis.exe"
    BATPreFile = VariableType(DATLoc)
    Call CallBATCHFile(MOUTdir, DATLoc, PROGLoc, exeFile, BATPreFile)
    If CopyOUTFile(MOUTdir, OUTDIR, HARMONICDIR) = True Then HARMONICANALYSIS = True
    
End Function
'---------------------------------------------------------------------
' Date Created : July 8, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 8, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CallBATCHFile
' Description  : This function parses the directory where all the .DBF
'                files are located and returns the root folder.
' Parameters   : String, String, String
' Returns      : String
'---------------------------------------------------------------------
Function FindRootFolder(ByVal fileDir As String, ByRef HADIR As String, _
ByRef MainOUT As String) As String
   
    Dim PROGDIR As String
    Dim MOUTdir As String
    Dim OUTDIR As String
    Dim VARdir As String
    Dim NewDIR As String
    Dim Divisor As Integer
    
    ' Always be Static!
    PROGDIR = "ACRU_COMPF7"
    MOUTdir = "Output\"
    OUTDIR = "ZSOUT\"
    
    Divisor = InStrRev(fileDir, "\")
    NewDIR = Left(fileDir, Divisor - 1)
    logtxt = "This is the Output directory: " & NewDIR
    Debug.Print logtxt
    logfile.WriteLine logtxt
    VARdir = Mid(NewDIR, Divisor - 3)
    logtxt = "Working on the following variable folder directory: " & VARdir
    Debug.Print logtxt
    logfile.WriteLine logtxt
    HADIR = VARdir
    
    Divisor = InStrRev(NewDIR, "\")
    NewDIR = Left(NewDIR, Divisor - Len(OUTDIR)) ' Remove the sub directory
    logtxt = "This is the main output folder in ACRU_COMPF7: " & NewDIR
    Debug.Print logtxt
    logfile.WriteLine logtxt
    MainOUT = NewDIR
    
    Divisor = InStrRev(NewDIR, "\")
    NewDIR = Left(NewDIR, Divisor - Len(MOUTdir)) ' Remove the main output directory
    logtxt = "This is the main output folder in ACRU_COMPF7: " & NewDIR
    Debug.Print logtxt
    logfile.WriteLine logtxt
    FindRootFolder = NewDIR
    
End Function
'---------------------------------------------------------------------
' Date Created : July 8, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 8, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindHarmonicFile
' Description  : This function passes on the .DAT file to be processed.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------
Function FindHarmonicFile(ByVal HarmonicFile As String) As String
   
    Dim PROGDIR As String
    Dim OUTDIR As String
    Dim VARdir As String
    Dim NewDIR As String
    Dim Divisor As Integer
    
    ' Always be Static!
    Divisor = InStrRev(HarmonicFile, "\")
    VARdir = Mid(HarmonicFile, Divisor + 1)
    logtxt = "Working on the following variable folder directory: " & VARdir
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
    FindHarmonicFile = VARdir
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 24, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 23, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : VariableType
' Description  : This function purpose is to name output files only.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function VariableType(ByVal VarName As String) As String
    
    Dim BATFile As String
    Dim UpperVarName As String
    Dim status As String
    
    UpperVarName = UCase(VarName)
    Select Case UpperVarName
        Case "ZSRAD.DAT"
            status = "Processed the zonal statistics files for RADIATION."
            BATFile = "RAD"
        Case "ZSREL.DAT"
            status = "Processed the zonal statistics files for RELATIVE HUMIDITY."
            BATFile = "REL"
        Case "ZSSUN.DAT"
            status = "Processed the zonal statistics files for SUNSHINE HOURS."
            BATFile = "SUN"
        Case "ZSWND.DAT"
            status = "Processed the zonal statistics files for WIND SPEED."
            BATFile = "WND"
    End Select
    
    logtxt = status
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
    VariableType = BATFile
    
End Function
'---------------------------------------------------------------------
' Date Created : July 4, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 8, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CallBATCHFile
' Description  : This function creates a batch file which will call
'                the necessary .EXE file and feed into it.
' Parameters   : String, ShellWindow, Boolean, Integer
' Returns      : Integer
'---------------------------------------------------------------------
Function CallBATCHFile(ByVal filePath As String, ByVal fileName As String, _
ByVal exeFilePath As String, ByVal exeFile As String, ByVal BATprefix As String)
    
    Dim FSO As Object
    Dim BATFile As String, TargetFolderPath As String
    Dim ProcessID As Boolean, FileNumber As Integer
    
    logfile.WriteLine "Creating a batch file"
    ' FIRST: Create main destination folder if it does not exists
    TargetFolderPath = CreateFolder(filePath & "BATOUT")

    ' THEN: Create batch file
    BATFile = TargetFolderPath & BATprefix & "HarmonicAnalysis.bat"
    logfile.WriteLine "Batch File: " & BATFile
    FileNumber = FreeFile() ' Get unused file number.
    Open BATFile For Output As #FileNumber ' Create file name.
    Print #FileNumber, "cd " & exeFilePath
    logfile.WriteLine "cd " & exeFilePath
    Print #FileNumber, exeFile & " " & fileName
    logfile.WriteLine exeFile & " " & fileName
    Close #FileNumber ' Close file.
    
    ProcessID = Shell_AndWait(BATFile)
    Debug.Print ProcessID
    Kill BATFile ' Delete the file right away
    
End Function
'---------------------------------------------------------------------
' Date Created : July 4, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 23, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopyOUTFile
' Description  : This function copies the newly saved .OUT files from
'                the C:\Harmonic directory back into the output
'                directory. The log files are copied onto the HAOUT
'                directory.
' Parameters   : String, String, String
' Returns      : Boolean
'---------------------------------------------------------------------
Function CopyOUTFile(ByRef FolderPath As String, ByVal OUTDIR As String, _
ByVal HARMONICDIR As String) As Boolean

    Dim objFolder As Object, objFSO As Object
    Dim SourceFilePath As String
    Dim MainFolderPath As String
    Dim TargetFolderPath As String
    Dim TxtFile As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    CopyOUTFile = True

    On Error Resume Next
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Check if source file exist
    SourceFilePath = HARMONICDIR
    If Right(SourceFilePath, 1) <> "\" Then SourceFilePath = SourceFilePath & "\"
    If objFSO.FolderExists(SourceFilePath) = False Then
        Debug.Assert "File Does Not Exist or Path Not Found"
        CopyOUTFile = False
        Exit Function
    End If

    ' FIRST: Create main output folder if it does not exists
    MainFolderPath = CreateFolder(FolderPath & "HAOUT")

    ' THEN: Create sub-folder if it does not exists
    TargetFolderPath = CreateFolder(MainFolderPath & OUTDIR)
    
    '-------------------------------------------------------------
    ' Loop through all the .OUT files and create a copy
    ' Then delete the original within C:\Harmonic
    '-------------------------------------------------------------
    Set objFolder = objFSO.GetFolder(SourceFilePath).Files
    
    For Each objFILE In objFolder
        If UCase(Right(objFILE.Path, (Len(objFILE.Path) - InStrRev(objFILE.Path, ".")))) = UCase("out") Then
            objFSO.CopyFile objFILE, TargetFolderPath
            objFSO.DeleteFile objFILE
        Else:
            TxtFile = UCase(Mid(objFILE.Path, InStrRev(objFILE.Path, "\") + 1))
            logfile.WriteLine "Checking status on: " & TxtFile
            Result = VariableTypeCheck(TxtFile)
            logfile.WriteLine "Check result: " & Result
            If Result = True Then
                objFSO.CopyFile objFILE, MainFolderPath
                logfile.WriteLine "Copying Logfile for: " & TxtFile
            End If
            objFSO.DeleteFile objFILE
            logfile.WriteLine "Deleting Logfile for: " & TxtFile
        End If
    Next
    
    logtxt = "Successfully Copied All .OUT Files."
    Debug.Print logtxt
    logfile.WriteLine logtxt

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
' Title        : VariableTypeCheck
' Description  : This function return true for Log File, false for
'                everything else.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------
Function VariableTypeCheck(ByVal VarName As String) As Boolean
    
    Dim FileCheck As Boolean
    Dim UpperVarName As String
    
    UpperVarName = UCase(VarName)
    Select Case UpperVarName
        Case "LOGRUN_ZSRAD.DAT"
            FileCheck = True
        Case "LOGRUN_ZSREL.DAT"
            FileCheck = True
        Case "LOGRUN_ZSSUN.DAT"
            FileCheck = True
        Case "LOGRUN_ZSWND.DAT"
            FileCheck = True
        Case "ZSRAD.DAT"
            FileCheck = False
        Case "ZSREL.DAT"
            FileCheck = False
        Case "ZSSUN.DAT"
            FileCheck = False
        Case "ZSWND.DAT"
            FileCheck = False
    End Select
    
    VariableTypeCheck = FileCheck
    
End Function

