Attribute VB_Name = "Step_4_CreateDATFile"
'---------------------------------------------------------------------
' Date Created : June 5, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CREATEDATFILES
' Description  : This function creates the .DAT file for each vari-
'                able. It looks at the last worksheet for reference.
'                It only copies the data, not the header.
' Parameters   : String, String
' Returns      : String
'---------------------------------------------------------------------
Function CREATEDATFILES(ByVal fileDir As String, ByVal ZSFileType As String) _
As String

    Dim CurrentWB As Workbook
    Dim CurrentSht As Worksheet
    Dim TempSht As Worksheet
    Dim TmpName As String, fileNameSaved As String
    Dim StationID As String
    Dim LastColIndex As Long, LastRowIndex As Long
    Dim LastRowInd As Integer, LastColInd As Integer
    Dim RowIndex As Integer, ColIndex As Integer
    
    Set CurrentWB = ActiveWorkbook
    TmpName = "Tmp"
    RowIndex = 2 ' Start from the second row
    ColIndex = 1

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Add a temporary worksheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = TmpName
    Set TempSht = ActiveWorkbook.Worksheets(Sheets.Count)
    Set CurrentSht = ActiveWorkbook.Worksheets(Sheets.Count - 1)

    ' Activate Summary Worksheet
    CurrentSht.Activate
    
    ' Define Variables
    LastColIndex = LastCol
    LastRowIndex = LastRow
    LastColInd = CInt(LastColIndex)
    LastRowInd = CInt(LastRowIndex)
    
    ' Create a master sheet without the headers and save as a .TXT file
    ' Mastersheet still includes first column
    Range(Cells(RowIndex, ColIndex), Cells(LastRowInd, LastColInd)).Select
    Selection.Copy
    TempSht.Activate
    Range("A1").Select
    TempSht.Paste
    
    ' Save Mastersheet according to the variable type: RAD, SUN, REL, WND
    fileNameSaved = SaveAsDAT(CurrentWB, TempSht, fileDir, ZSFileType)
    CREATEDATFILES = CopyDATFile(fileDir, fileNameSaved)

    ' Close Workbook without any saving the changes
    CurrentWB.Close SaveChanges:=False
    
End Function
'---------------------------------------------------------------------
' Date Created : June 5, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveAsDAT
' Description  : This function saves specific string as a .dat file.
'                It returns the filename string.
' Parameters   : Workbook, Worksheet, String, String
' Returns      : String
'---------------------------------------------------------------------
Function SaveAsDAT(wbTmp As Workbook, ByRef TmpSht As Worksheet, _
ByVal fileDir As String, ByVal ZSFileType As String) As String

    Dim saveFile As String
    Dim fileName As String
    
    ' Activate the appropriate Worksheet
    TmpSht.Select
    
    ' Check the Excel version
    If Val(Application.Version) < 9 Then Exit Function
    
    ' Save information as textfile
    fileName = ZSFileType & ".dat"
    saveFile = fileDir & fileName
    If Right(fileDir, 1) <> "\" Then saveFile = fileDir & "\" & fileName
    wbTmp.SaveAs saveFile, FileFormat:=xlText, CreateBackup:=False
    
    SaveAsDAT = fileName
    
End Function
'---------------------------------------------------------------------
' Date Created : July 4, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 8, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopyDATFile
' Description  : This function copies the newly saved .DAT file to the
'                C:\Harmonic directory.
' Parameters   : String, String
' Returns      : String
'---------------------------------------------------------------------
Function CopyDATFile(ByRef sourcePath As String, _
ByVal SourceFile As String) As String

    Dim SourceFilePath As String
    Dim TargetFolderPath As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    On Error Resume Next
    
    ' Check if source file exist
    If Right(sourcePath, 1) <> "\" Then sourcePath = sourcePath & "\"
    SourceFilePath = sourcePath & SourceFile
    If objFSO.fileExists(SourceFilePath) = False Then
        Debug.Assert "File Does Not Exist or Path Not Found"
    End If

    ' Create destination folder if it does not exists
    logtxt = "Creating the Harmonic directory if it does not exist..."
    Debug.Print logtxt
    logfile.WriteLine logtxt
    TargetFolderPath = "C:\HARMONIC\"
    If Right(TargetFolderPath, 1) <> "\" Then
        TargetFolderPath = TargetFolderPath & "\"
        If objFSO.FolderExists(TargetFolderPath) = False Then
            logtxt = "Harmonic folder doesn't exist."
            Debug.Print logtxt
            logfile.WriteLine logtxt
        End If
        MkDir (TargetFolderPath)
    Else
        logtxt = "Harmonic folder exist."
        Debug.Print logtxt
        logfile.WriteLine logtxt
    End If
    
    objFSO.CopyFile SourceFilePath, TargetFolderPath
    CopyDATFile = TargetFolderPath & SourceFile ' Pass on new directory and file name

End Function



