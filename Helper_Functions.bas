Attribute VB_Name = "Helper_Functions"
'---------------------------------------------------------------------
' Date Acquired : May 18, 2012
' Source : http://www.vbaexpress.com/kb/getarticle.php?kb_id=767
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : filePath
' Description  : This function takes a file path and returns only
'                the filepath, excluding the file name and extension.
' Parameters   : Variant
' Returns      : String
'---------------------------------------------------------------------
Function filePath(ByVal strPath As Variant) As String

    filePath = Left$(strPath, InStrRev(strPath, "\"))
    
End Function
'---------------------------------------------------------------------
' Date Created : June 26, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CheckFileExists
' Description  : This function checks if the file exists or not.
' Parameters   : String, String
' Returns      : Boolean
'---------------------------------------------------------------------
Function CheckFileExists(ByRef sDirectory As String, ByVal sFileName As String) As Boolean

    Dim objFSO As Object
    Dim sFile As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrHandler
    
    If Right(sDirectory, 1) <> "\" Then sDirectory = sDirectory & "\"
    sFile = sDirectory & sFileName
    If objFSO.fileExists(sFile) = True Then
        Debug.Print "File exists."
        CheckFileExists = True
    Else
        Debug.Print "File does not exists."
        CheckFileExists = False
    End If
    
ErrHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Resume Next
    End If
    Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CheckFolderExists
' Description  : This function checks if the folder exists or not.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckFolderExists(ByVal fileDir As String) As Boolean

    Dim objFSO  As Object
    
    Set objFSO = CreateObject("scripting.filesystemobject")

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' Check if Folder exists
    If objFSO.FolderExists(fileDir) = True Then
        Debug.Print "Folder exists."
        CheckFolderExists = True
    Else
        Debug.Print "Folder does not exists."
        CheckFolderExists = False
    End If
    
ErrHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Resume Next
    End If
    Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CheckSubFolder
' Description  : This function checks if the subfolder exists under the root folder.
' Parameters   : String, String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function CheckSubFolder(ByRef fileDir As String, ByVal subFolder As String) As Boolean

    Dim objFSO  As Object
    Dim MainRootFolderPath As String, SubFolderPath As String
    
    Set objFSO = CreateObject("scripting.filesystemobject")

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    CheckSubFolder = True
    
    On Error GoTo ErrHandler

    ' Check if Folder has been selected, if not, go to default
    MainRootFolderPath = fileDir
    If Right(MainRootFolderPath, 1) <> "\" Then MainRootFolderPath = MainRootFolderPath & "\"
    
    ' Check Sub Folder
    SubFolderPath = MainRootFolderPath & subFolder
    If Right(SubFolderPath, 1) <> "\" Then SubFolderPath = SubFolderPath & "\"
    If objFSO.FolderExists(SubFolderPath) = False Then
        Debug.Print "Sub folder doesn't exist"
        CheckSubFolder = False
    Else: Debug.Print "Sub folder exist"
    End If
    
ErrHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Resume Next
    End If
    Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReturnFolder
' Description  : This function returns the directory with "\" at the end of it.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function ReturnFolder(ByVal fileDir As String) As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    ' Check if Folder has been selected, if not, go to default
    If Right(fileDir, 1) <> "\" Then ReturnFolder = fileDir & "\"
    If Right(fileDir, 1) = "\" Then ReturnFolder = fileDir
    
ErrHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Resume Next
    End If
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReturnSubFolder
' Description  : This function returns the directory of the subfolder with "\" at the
'                at the end.
' Parameters   : String, String
' Returns      : String
'---------------------------------------------------------------------------------------
Function ReturnSubFolder(ByRef fileDir As String, ByVal subFolder As String) As String

    Dim MainRootFolderPath As String, SubFolderPath As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    ' Check if Folder has been selected, if not, go to default
    MainRootFolderPath = fileDir
    If Right(MainRootFolderPath, 1) <> "\" Then MainRootFolderPath = MainRootFolderPath & "\"
    
    ' Check Sub Folder
    SubFolderPath = MainRootFolderPath & subFolder
    If Right(SubFolderPath, 1) <> "\" Then SubFolderPath = SubFolderPath & "\"
    ReturnSubFolder = SubFolderPath

ErrHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Resume Next
    End If
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : GetSubFolders
' Description  : This function returns the list of sub directories within a directory.
' Parameters   : String
' Returns      : -
'---------------------------------------------------------------------------------------
Function GetSubFolders(RootPath As String)

    Dim objFSO As Object
    Dim fld As Object
    Dim sf As Object
    Dim Arr()
    Dim folderName As String
    Dim Counter As Integer
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set fld = objFSO.GetFolder(RootPath).SubFolders
    Counter = 0
    
    For Each sf In fld
        folderName = sf.Name
        Debug.Print folderName
    Next
    
    Set sf = Nothing
    Set fld = Nothing
    Set objFSO = Nothing
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : GetSubFoldersArray
' Description  : This function returns the list of sub directories into an array.
' Parameters   : String, String Array
' Returns      : -
'---------------------------------------------------------------------------------------
Function GetSubFoldersArray(ByVal RootPath As String, ByRef subDIRArray() As String)

    Dim objFSO As Object
    Dim fld As Object
    Dim sf As Object
    Dim Arr()
    Dim folderName As String
    Dim Counter As Integer
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set fld = objFSO.GetFolder(RootPath).SubFolders
    Counter = 0
    
    For Each sf In fld
        ReDim Preserve subDIRArray(Counter)
        folderName = sf.Path
        subDIRArray(Counter) = folderName
        Counter = Counter + 1
    Next
    
    Set sf = Nothing
    Set fld = Nothing
    Set objFSO = Nothing
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ReturnFolderName
' Description  : This function returns the name of the folders.
' Parameters   : String
' Returns      : Boolean
'---------------------------------------------------------------------------------------
Function ReturnFolderName(ByVal TxtFile As String) As String

    Dim fPath As String, fName As String
    Dim CommaLocation As Integer
    
    CommaLocation = InStrRev(TxtFile, "\")
    If CommaLocation = 0 Then ReturnFolderName = ""
    If CommaLocation > 0 Then
        fPath = Left(TxtFile, CommaLocation - 1)
        fName = Mid(TxtFile, CommaLocation + 1)
        ReturnFolderName = fName
    End If
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : July 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : GetFolder
' Description  : This function opens dialog box and then lets the user choose the folder
'                to open. It returns the string of the selected directory.
' Parameters   : -
' Returns      : String
'---------------------------------------------------------------------------------------
Function GetFolder() As String

    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .AllowMultiSelect = False
        .title = "Select a Folder"
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 8, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateFolder
' Description  : This function creates a new folder based on the string input parameter.
'                If the directory already exists, it will not create the same folder.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function CreateFolder(ByVal fileDir As String) As String

    Dim objFSO As Object
    Dim MainRootFolderPath As String, SubFolderPath As String
    Dim RootPath As String
    Dim RootFolder As String
    Set objFSO = CreateObject("scripting.filesystemobject")

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    ' Check if Folder has been selected, if not, go to default
    MainRootFolderPath = fileDir
    If Right(MainRootFolderPath, 1) <> "\" Then MainRootFolderPath = MainRootFolderPath & "\"
    
    If objFSO.FolderExists(MainRootFolderPath) = False Then
        Debug.Print "Folder doesn't exist. Creating new folder."
        MkDir (MainRootFolderPath)
    Else: Debug.Print "Folder already exist."
    End If
    
    CreateFolder = MainRootFolderPath
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Err.Clear
        Resume Next
    End If
    Set objFSO = Nothing
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
' Title        : SaveTXT
' Description  : This function saves the new AB10K grid as a .TXT file.
' Parameters   : Workbook, Worksheet, String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveTXT(ByRef wbTmp As Workbook, ByRef TmpSheet As Worksheet, _
ByVal fileDir As String, ByVal fName As String)

    Dim saveFile As String
    Dim fileName As String
    
    ' Activate the appropriate Worksheet
    TmpSheet.Activate
    
    ' Check the Excel version
    If Val(Application.Version) < 9 Then Exit Function
    
    ' Save information as textfile
    fileName = fName & ".txt"
    saveFile = fileDir & fileName
    If Right(fileDir, 1) <> "\" Then saveFile = fileDir & "\" & fileName
    
    wbTmp.SaveAs saveFile, FileFormat:=xlText, CreateBackup:=False
    
End Function
Function ReturnOutputFile(ByVal fileDir As String, ByVal fName As String) As String

    Dim saveFile As String
    Dim fileName As String
      
    ' Return information as textfile
    fileName = fName & ".txt"
    saveFile = fileDir & fileName
    If Right(fileDir, 1) <> "\" Then saveFile = fileDir & "\" & fileName
    
    ReturnOutputFile = saveFile
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 11, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateNewFolder
' Description  : This function creates a new folder within a root folder if it does not
'                exist.
' Parameters   : String, String
' Returns      : String
'---------------------------------------------------------------------------------------
Function CreateNewFolder(ByVal fileDir As String, ByVal RootFolder As String) As String

    Dim objFSO As Object
    Dim MainRootFolderPath As String, SubFolderPath As String
    Dim RootPath As String
    Set objFSO = CreateObject("scripting.filesystemobject")

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    ' Check if Folder has been selected, if not, go to default
    MainRootFolderPath = fileDir
    If Right(MainRootFolderPath, 1) <> "\" Then MainRootFolderPath = MainRootFolderPath & "\"
    
    '---------------------------------------------------------------------
    ' Create Root Folder if it does not exist
    '---------------------------------------------------------------------
    RootPath = MainRootFolderPath & RootFolder
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"
    
    If objFSO.FolderExists(RootPath) = False Then
        Debug.Print "Folder doesn't exist. Creating a new folder..."
        MkDir (RootPath)
    Else: Debug.Print "Folder exist"
    End If
    
    CreateOUTPUTFolder = RootPath
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Err.Number
        Err.Clear
        Resume Next
    End If
    Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : July 18, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : DeleteFolderAndContents
' Description  : This function Delete whole folder without removing the files first.
' Parameters   : String
' Returns      : -
'---------------------------------------------------------------------------------------
Function DeleteFolderAndContents(ByVal fileDir As String)

    Dim objFSO As Object
    Dim MyPath As String
    
    Set objFSO = CreateObject("scripting.filesystemobject")
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    
    MyPath = fileDir
    If Right(MyPath, 1) = "\" Then MyPath = Left(MyPath, Len(MyPath) - 1)
    
    If objFSO.FolderExists(MyPath) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Function
    End If
    objFSO.DeleteFolder MyPath

ErrHandler:
    Set objFSO = Nothing
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 26, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : June 26, 2012
' Edited By    : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : filePath
' Description  : This function takes a file path and returns only
'                the filepath, excluding the file name and extension.
' Parameters   : Variant
' Returns      : -
'---------------------------------------------------------------------------------------
Function FindLastRowColumn(ByRef LR As Long, ByRef LC As Long)

    Dim LastRowIndex As Long
    Dim LastColIndex As Long
    
    LastRowIndex = 1
    LastColIndex = 1
    
    If WorksheetFunction.CountA(Cells) > 0 Then
        LastRowIndex = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LastColIndex = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    
    LR = LastRowIndex
    LC = LastColIndex

End Function
'---------------------------------------------------------------------
' Date Created : August 3, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 16, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LastRow
' Description  : This function returns the last row count for the
'                activesheet.
' Parameters   : -
' Returns      : Long
'---------------------------------------------------------------------
Function LastRow() As Long

    Dim LastRowIndex As Long
    
    If WorksheetFunction.CountA(Cells) > 0 Then
       LastRowIndex = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
    
    LastRow = LastRowIndex
    
End Function
'---------------------------------------------------------------------
' Date Created : August 3, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 16, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LastCol
' Description  : This function returns the last column count for the
'                activesheet.
' Parameters   : -
' Returns      : Long
'---------------------------------------------------------------------
Function LastCol() As Long

    Dim LastColIndex As Long
    
    If WorksheetFunction.CountA(Cells) > 0 Then
       LastColIndex = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    
    LastCol = LastColIndex
    
End Function
'---------------------------------------------------------------------
' Date Created : August 3, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : August 3, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RangeAddress
' Description  : This function finds the input string and returns its
'                address.
' Parameters   : String
' Returns      : Address String
'---------------------------------------------------------------------
Function RangeAddress(ByVal InputString As String)

    Dim Found As Range
    Dim DynamicAddress
    
    Set Found = Rows(1).Find(What:=InputString, SearchDirection:=xlNext, SearchOrder:=xlByColumns)
    With Found
        DynamicAddress = Found.Address
        Debug.Print DynamicAddress
    End With

    RangeAddress = DynamicAddress
    
End Function
'---------------------------------------------------------------------
' Date Created : June 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 12, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindValuesRange
' Description  : This function selects the values range only.
' Parameters   : -
'---------------------------------------------------------------------
Function FindSpecificRange(DestSheet As Worksheet, ByVal FirstRowN As Long, _
ByVal FirstColN As Long, ByVal LastRowN As Long, ByVal LastColN As Long)

    Dim FirstRow&, FirstCol&, LastRow&, LastCol&
    Dim myUsedRange As Range
        
    ' Activate the correct worksheet
    DestSheet.Activate
    
    FirstRow = FirstRowN
    FirstCol = FirstColN
    LastRow = LastRowN
    LastCol = LastColN
    
    ' Select Range using FirstRow, FirstCol, LastRow, LastCol
    With ActiveWorkbook.ActiveSheet
        .Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol)).Select
    End With
    
End Function
'---------------------------------------------------------------------
' Date Created: June 4, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited: June 13, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RowCheck
' Description  : This function will check if the worksheet is empty.
'                If worksheet is empty, function exits. Otherwise,
'                it checks for the last row. If found, selects one row
'                after the last known row.
' Parameters   : Worksheet
'---------------------------------------------------------------------
Function RowCheck(WKSheet As Worksheet)

    Dim LastRow As Long
    
    ' Activate correct worksheet
    WKSheet.Activate
    
    '-------------------------------------------------------------
    ' For Empty/New Workbook
    '-------------------------------------------------------------
    If WorksheetFunction.CountA(Cells) = 0 Then
        Range("A1").Select
        Exit Function
    End If

    '-------------------------------------------------------------
    ' Check for the last used row. Select the row after the
    ' last known row.
    '-------------------------------------------------------------
    LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    ActiveSheet.Range("A" & LastRow + 1).Select

End Function
'---------------------------------------------------------------------
' Date Acquired : May 28, 2012
' Source : http://msdn.microsoft.com/en-us/library/ff198177.aspx
'---------------------------------------------------------------------
' Date Edited  : June 13, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : FindRange
' Description  : This function selects the range using the last used
'                row and column even with missing values in
'                between many columns.
' Parameters   : Worksheet
'---------------------------------------------------------------------
Function FindRange(WKSheet As Worksheet)

    Dim FirstRow&, FirstCol&, LastRow&, LastCol&
    Dim myUsedRange As Range
        
    ' Activate the correct worksheet
    WKSheet.Activate
    
    ' Define variables
    FirstRow = Cells.Find(What:="*", SearchDirection:=xlNext, SearchOrder:=xlByRows).Row
    FirstCol = Cells.Find(What:="*", SearchDirection:=xlNext, SearchOrder:=xlByColumns).Column
    LastRow = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    LastCol = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    
    ' Select Range using FirstRow, FirstCol, LastRow, LastCol
    Set myUsedRange = Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol))
    myUsedRange.Select
    
End Function
'---------------------------------------------------------------------
' Date Created : July 20, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 20, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : LastRowIndex
' Description  : This function selects the range depending on the row
'                that contains the header information.
' Parameters   : -
'---------------------------------------------------------------------
Function LastRowIndex(ByVal Index As Integer)

    Dim RowIndex As Integer
    Dim Found As Range
    Dim LastCell
    
    Range("A1").Select
    Set Found = Cells.Find(What:="Date/Time", After:=ActiveCell)
    LastCell = Found.Address
    Range(LastCell).Select
    RowIndex = Selection.Row
    
    '-------------------------------------------------------------
    ' Process metadata. The first retains the header information.
    '-------------------------------------------------------------
    If Index = 1 Then
        RowIndex = RowIndex - 1
    End If
        
    '-------------------------------------------------------------
    ' Delete metadata according to RowIndex.
    '-------------------------------------------------------------
    ActiveSheet.Rows("1:" & RowIndex).Select
    Selection.Delete

End Function

'---------------------------------------------------------------------
' Date Created : June 4, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 5, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RowCheck
' Description  : This function will check if the worksheet is empty.
'                If worksheet is empty, function exits. Otherwise,
'                it checks for the last row. If found, selects one row
'                after the last known row.
' Parameters   : Worksheet
'---------------------------------------------------------------------
Function ColumnCheck(WKSheet As Worksheet)

    Dim LastCol As Long
    
    ' Activate correct worksheet
    WKSheet.Activate
    
    '-------------------------------------------------------------
    ' For Empty/New Workbook
    '-------------------------------------------------------------
    If WorksheetFunction.CountA(Cells) = 0 Then
        Range("A1").Select
        Exit Function
    End If

    '-------------------------------------------------------------
    ' Check for the last used row. Select the row after the
    ' last known row.
    '-------------------------------------------------------------
    If WorksheetFunction.CountA(Cells) > 0 Then
       LastCol = Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    End If
    ActiveSheet.Range("A1").Offset(0, LastCol).Select

End Function
