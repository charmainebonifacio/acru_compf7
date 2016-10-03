Attribute VB_Name = "Step_3_CreateSummary"
'---------------------------------------------------------------------
' Date Created : June 4, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 18, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SUMMARYWORKSHEET
' Description  : This function will create a summary worksheet that
'                contains all the mean values from January to December.
'                Major assumption that all zonal stats used the same
'                configurations which means the columns will remain
'                constant for all 12 months.
' Parameters   : String, String, String, String, String Array, Integer
' Returns      : -
'---------------------------------------------------------------------
Function SUMMARYWORKSHEET(ByVal FileDirectory As String, ByVal ZSFileType As String, _
ByVal CodeID As String, ByVal MeanVar As String, ByRef refIDArray() As String, _
ByVal varIndex As Integer)

    Dim DestSht As Worksheet
    Dim SourceSht As Worksheet
    Dim CurrentSht As Worksheet
    Dim WorksheetCount As Integer
    Dim LastRowThisWorksheet As Long
    Dim CellAddress As String
    Dim CallConvertFunction As Boolean
    
    ' Initialize Variables
    TestInt = 0
    CallConvertFunction = False
    
    ' Create the Summary Worksheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Summary"
    Set DestSht = ActiveWorkbook.Worksheets(13)
    Set CurrentSht = ActiveWorkbook.Worksheets(1)
    
    Application.ScreenUpdating = False
    
    ' Copy Values to the Summary Worksheet
    For Each SourceSht In ActiveWorkbook.Worksheets
        WorksheetCount = WorksheetCount + 1
        SourceSht.Activate
        If SourceSht.Name = "Summary" Then Exit For
        If WorksheetCount = 1 Then
            Call SelectColumn(SourceSht, DestSht, CodeID)
            Call CreateReferenceID(refIDArray())
        End If
        ' Otherwise, grab Mean Values =12, Start from row 2 until last row
        Call SelectColumn(SourceSht, DestSht, MeanVar)
        Selection.NumberFormat = "0.0"
    Next SourceSht
    
    ' Change Headers for clarity! Otherwise, there are 12 "MEAN" columns
    Call ColumnHeaderProcessor(DestSht, MeanVar)
    
    ' Convert values into proper units depending on the variable
    ' Only convert units for RAD, SUN and WND values
    If Not varIndex = 1 Then
        CallConvertFunction = ConversionSheet(varIndex, CodeID)
    End If
    logtxt = "The Conversion Sheet Function was made: " & CallConvertFunction
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
    ' Save Summary Workbook according to the variable type: RAD, SUN, REL, WND
    Call SaveFileAs(ActiveWorkbook, FileDirectory, ZSFileType)
    logtxt = "Successfully created a summary worksheet."
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
End Function
'---------------------------------------------------------------------
' Date Created : June 4, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 24, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SelectColumn
' Description  : This function will find the column that denotes
'                either the CODEID or the mean values within a
'                specific worksheet. All data within the column
'                will be copied and pasted on to another worksheet.
' Parameters   : Worksheet, Worksheet, String
' Returns      : -
'---------------------------------------------------------------------
Function SelectColumn(SourceSht As Worksheet, DestSht As Worksheet, _
ByVal StringToFind As String)

    Dim Found As Range
    Dim FoundCell

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Activate Source Worksheet.
    SourceSht.Activate

    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    Set Found = Rows(1).Find(What:=StringToFind, After:=Range("A1"), _
                LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    With Found
        FoundCell = Found.Address
        Debug.Print FoundCell
        Range(FoundCell).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
    End With

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
    
End Function
'---------------------------------------------------------------------
' Date Created : July 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 12, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreateReferenceID
' Description  : This function creates the reference ID array. This
'                way the tool avoids having to look at a file for
'                to reference which files need to be edited.
' Parameters   : String Array
' Returns      : -
'---------------------------------------------------------------------
Function CreateReferenceID(ByRef refIDArray() As String)

    Dim rACells As Range, rLoopCells As Range
    Dim refIDCellValue As Integer
    Dim rowID As Integer
    Dim LC As Long, LR As Long, NewLR As Long
    Dim refIndex As Integer
    
    Call FindLastRowColumn(LR, LC)
    NewLR = LR - 2  ' Because of the header on row one!
    ReDim refIDArray(NewLR)
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Initialize Variables
    If WorksheetFunction.CountA(Cells) > 0 Then
        Range(Cells(2, 1), Cells(LR, 1)).Select
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

    'Initializing values in the array to the present AB_ID."
    rowID = 0
    For Each rLoopCells In rACells
        refIDCellValue = rLoopCells.Value
        refIDArray(rowID) = refIDCellValue
        rowID = rowID + 1
    Next rLoopCells

    'Printing values in the Reference ID Array."
    logfile.WriteLine "Printing values in the Reference ID ARRAY"
    For refIndex = LBound(refIDArray) To UBound(refIDArray)
        Debug.Print refIndex, refIDArray(refIndex)
        logfile.WriteLine refIndex & "-" & refIDArray(refIndex)
    Next refIndex
        
End Function
'---------------------------------------------------------------------
' Date Created : June 4, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ColumnHeaderProcessor
' Description  : This function processes the summary worksheet by
'                cleaning up the header names into the appropriate
'                months.
' Parameters   : Worksheet, String
' Returns      : -
'---------------------------------------------------------------------
Function ColumnHeaderProcessor(DestSht As Worksheet, ByVal StringToFind As String)

    Dim Found As Range
    Dim ColIndex As Integer
    Dim FoundCell
    Dim MonthElement(1 To 12) As String

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    ' Activate Destination Worksheet.
    DestSht.Activate
    
    ' Set Variable
    MonthElement(1) = "JAN"
    MonthElement(2) = "FEB"
    MonthElement(3) = "MAR"
    MonthElement(4) = "APR"
    MonthElement(5) = "MAY"
    MonthElement(6) = "JUN"
    MonthElement(7) = "JUL"
    MonthElement(8) = "AUG"
    MonthElement(9) = "SEP"
    MonthElement(10) = "OCT"
    MonthElement(11) = "NOV"
    MonthElement(12) = "DEC"
    
    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    Set Found = Rows(1).Find(What:=StringToFind, After:=Range("A1"), LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    With Found
        FoundCell = Found.Address
        Debug.Print FoundCell
        Range(FoundCell).Select
        For ColIndex = 1 To 12
            Range(FoundCell).Offset(0, ColIndex - 1).Select
            ActiveCell.Value = MonthElement(ColIndex)
        Next ColIndex
    End With

End Function
'---------------------------------------------------------------------
' Date Created : July 17, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 18, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ConversionSheet
' Description  : This function converts the existing units for RAD,
'                SUN and WND variables using the summary worksheet
'                and creates a new worksheet with the output result.
' Parameters   : Integer
' Returns      : Boolean
'---------------------------------------------------------------------
Function ConversionSheet(ByVal varIndex As Integer, ByVal CodeID As String) As Boolean

    Dim CurrentWB As Workbook
    Dim CurrentSht As Worksheet
    Dim TempSht As Worksheet
    Dim TmpName As String
    Dim RowIndex As Long, ColIndex As Long
    
    ConversionSheet = False
    Set CurrentWB = ActiveWorkbook
    TmpName = "Conversion"

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Add a temporary worksheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = TmpName
    Set TempSht = ActiveWorkbook.Worksheets(Sheets.Count)
    Set CurrentSht = ActiveWorkbook.Worksheets(Sheets.Count - 1)

    ' Activate Summary Worksheet
    CurrentSht.Activate
    Call FindLastRowColumn(RowIndex, ColIndex)

    Dim OUTPUTVAL() As Double
    Dim INPUTVAL() As Double
    Dim DEN As Double
    
    Dim MonthElement() As String
    ReDim MonthElement(0 To ColIndex - 1)
    MonthElement(0) = "AB_ID"
    MonthElement(1) = "JAN"
    MonthElement(2) = "FEB"
    MonthElement(3) = "MAR"
    MonthElement(4) = "APR"
    MonthElement(5) = "MAY"
    MonthElement(6) = "JUN"
    MonthElement(7) = "JUL"
    MonthElement(8) = "AUG"
    MonthElement(9) = "SEP"
    MonthElement(10) = "OCT"
    MonthElement(11) = "NOV"
    MonthElement(12) = "DEC"

    ReDim INPUTVAL(1 To RowIndex - 1)
    ReDim OUTPUTVAL(1 To RowIndex - 1)
                  
    For Z = LBound(MonthElement) To UBound(MonthElement)
        CurrentSht.Activate
        Select Case varIndex
            Case 0, 2:
                DEN = Switch(MonthElement(Z) = "JAN", 31, _
                            MonthElement(Z) = "FEB", 28.25, _
                            MonthElement(Z) = "MAR", 31, _
                            MonthElement(Z) = "APR", 30, _
                            MonthElement(Z) = "MAY", 31, _
                            MonthElement(Z) = "JUN", 30, _
                            MonthElement(Z) = "JUL", 31, _
                            MonthElement(Z) = "AUG", 31, _
                            MonthElement(Z) = "SEP", 30, _
                            MonthElement(Z) = "OCT", 31, _
                            MonthElement(Z) = "NOV", 30, _
                            MonthElement(Z) = "DEC", 31, _
                            MonthElement(0) = CodeID, 1)
            Case 3:
                DEN = 24
        End Select
        For i = LBound(INPUTVAL) To UBound(INPUTVAL)
            INPUTVAL(i) = Range("A1").Offset(i, Z).Value
            If Z = 0 Then OUTPUTVAL(i) = INPUTVAL(i)
            If Z > 0 And varIndex = 0 Or varIndex = 2 Then OUTPUTVAL(i) = (INPUTVAL(i) / DEN)
            If Z > 0 And varIndex = 3 Then OUTPUTVAL(i) = (INPUTVAL(i) * DEN)
            Debug.Print Z, i, DEN
        Next i
        
        TempSht.Activate
        Range("A1").Offset(0, Z).Value = MonthElement(Z)
        For i = LBound(OUTPUTVAL) To UBound(OUTPUTVAL)
            Range("A1").Offset(i, Z).Value = OUTPUTVAL(i)
            If Z > 0 Then Range("A1").Offset(i, Z).NumberFormat = "0.0"
        Next i
    Next Z

    ConversionSheet = True
    
End Function
'---------------------------------------------------------------------
' Date Created : June 6, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveFileAs
' Description  : This function saves specific string as a .XLSX file.
' Parameters   : Workbook, String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveFileAs(wbTmp As Workbook, ByVal fileDir As String, ByVal ZSFileType As String)

    Dim saveFile As String
    Dim fileName As String
    Dim FileFormatValue As Long

    ' Check the Excel version
    If Val(Application.Version) < 9 Then Exit Function
    
    ' FileFormat refers to .xlsx
    FileFormatValue = 51
    
    ' Save information as Temp, which can then be renamed later..
    fileName = "Summary_" & ZSFileType & ".xlsx"
    saveFile = fileDir & fileName
    If Right(fileDir, 1) <> "\" Then saveFile = fileDir & "\" & fileName
    wbTmp.SaveAs saveFile, FileFormat:=FileFormatValue, CreateBackup:=False
    
End Function
