Attribute VB_Name = "Step_2_PreprocessDBFFiles"
'---------------------------------------------------------------------
' Date Created : June 3, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 16, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : PROCESSDBFFILES
' Description  : This function processes all the .dbf files within a
'                file directory. The .dbf files must be numbered
'                chronologically. User must select the first file that
'                denotes the first month of the year (ie. January).
'                The function returns the file directory which will
'                be used in later functions.
' Parameters   : -
' Returns      : String
'---------------------------------------------------------------------
Function PROCESSDBFFILES(ByVal fileDir As String, ByRef ZSFileType As String) As String

    Dim objFolder As Object, objFSO As Object
    Dim wbSource As Workbook, SourceSheet As Worksheet
    Dim wbDest As Workbook, DestSheet As Worksheet
    Dim FileCounter As Long
    Dim sThisFilePath As String, sFile As String
    Dim var As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Initialize Variables
    FileCount = 0
    
    '-------------------------------------------------------------
    ' Add New WorkBook with 12 worksheets - in order.
    '-------------------------------------------------------------
    Call AddNameWorksheet
    Set wbDest = ActiveWorkbook
    Set DestSheet = wbDest.Worksheets(1)

    '-------------------------------------------------------------
    ' Check the files... which should be obvious at this point.
    '-------------------------------------------------------------
    sThisFilePath = fileDir
    If (Right(sThisFilePath, 1) <> "\") Then sThisFilePath = sThisFilePath & "\"
    sFile = Dir(sThisFilePath & "*.dbf") ' Only .dbf files not .dbf.xml!
    ZSFileType = Left(sFile, 5) ' Only the first five characters are important
    Debug.Print ZSFileType
    var = ZonalStatFileFor(ZSFileType)
    
    '-------------------------------------------------------------
    ' Loop through all the .DBF files
    '-------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(sThisFilePath).Files
    logfile.WriteLine "Looping through all the .DBF files"
    For Each objFILE In objFolder
        If UCase(Right(objFILE.Path, (Len(objFILE.Path) - InStrRev(objFILE.Path, ".")))) = UCase("dbf") Then
            ' File count of current folder which corresponds to monthly index
            FileCounter = FileCounter + 1
            Debug.Print FileCounter, " of files processed."
            logfile.WriteLine FileCounter & " of files processed."
            
            ' Open file and set it as source worksheet
            Set DestSheet = wbDest.Worksheets(FileCounter)
            Set wbSource = Workbooks.Open(objFILE.Path)
            Set SourceSheet = wbSource.Worksheets(1)
            SourceSheet.Activate
            
            ' Copy / Paste Data
            Call DatabaseFile(SourceSheet, DestSheet)
            
            ' Ignore Clipboard Alerts
            Application.CutCopyMode = True
            
            ' Save Changes to the Processed Files
            wbSource.Close SaveChanges:=False

        Else
            logtxt = objFILE & " is not a valid file to process."
            Debug.Print logtxt
            logfile.WriteLine logtxt
        End If
    Next

    Set wbSource = Nothing
    Set SourceSheet = Nothing
    Set wbDest = Nothing
    Set DestSheet = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 24, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ZonalStatFileFor
' Description  : This function purpose is to name output files only.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function ZonalStatFileFor(ByVal VarName As String) As String
    
    Dim DATFile As String
    Dim status As String
    Select Case VarName
        Case "zs_rd"
            status = "Processed the zonal statistics files for RADIATION."
            DATFile = "ZSRAD"
        Case "zs_rh"
            status = "Processed the zonal statistics files for RELATIVE HUMIDITY."
            DATFile = "ZSREL"
        Case "zs_sh"
            status = "Processed the zonal statistics files for SUNSHINE HOURS."
            DATFile = "ZSSUN"
        Case "zs_ws"
            status = "Processed the zonal statistics files for WIND SPEED."
            DATFile = "ZSWND"
    End Select
    
    logtxt = status
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
    ZonalStatFileFor = DATFile
    
End Function
'---------------------------------------------------------------------
' Date Created : June 3, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : DatabaseFile
' Description  : This function copies the data from the zonal stats
'                .dbf file.
' Parameters   : Worksheet, Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function DatabaseFile(SourceSht As Worksheet, DestSht As Worksheet)
    
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
    Call FindRange(SourceSht)
    RngSelect = Selection.Address
    Range(RngSelect).Copy

    ' Activate Destination Worksheet.
    DestSht.Activate
    
    '-------------------------------------------------------------
    ' Call RowCheck function to check the last row.
    ' Then append the copied data into the Destination Worksheet.
    '-------------------------------------------------------------
    Call RowCheck(DestSht)
    PasteSelect = Selection.Address
    Range(PasteSelect).Select
    DestSht.Paste
    
    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False
    
End Function
'---------------------------------------------------------------------
' Date Created : June 3, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : AddNameWorksheet
' Description  : This function creates 12 worksheets and names these
'                worksheet from 1-12 for simplicity purposes.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------
Function AddNameWorksheet()

    Dim wbDest As Workbook
    Dim FileCount As Integer
    Dim MonthElement(1 To 12) As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Initialize Variables
    Set wbDest = Workbooks.Add(1)
    FileCount = 0
    
    For x = LBound(MonthElement) To UBound(MonthElement)
        FileCount = FileCount + 1
        MonthElement(x) = CStr(FileCount)
        Debug.Print MonthElement(x)
        If x = 1 Then ActiveSheet.Name = MonthElement(x)
        If x > 1 Then Sheets.Add(After:=Sheets(Sheets.Count)).Name = MonthElement(x)
    Next x
    
End Function
