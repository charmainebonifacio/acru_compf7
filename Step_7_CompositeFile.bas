Attribute VB_Name = "Step_7_CompositeFile"
'---------------------------------------------------------------------
' Date Created : July 2, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 18, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : ProcessCompositeFiles
' Description  : This function processes the original AB10K grid files
'                and formats the information to serve as an ACRU input.
'                The .TXT file will not include a newline at the end
'                of the file.
' Parameters   : String, String
' Returns      : -
'---------------------------------------------------------------------
Function ProcessCompositeFiles(ByVal strPath As String, ByVal strOutPath As String)

    Dim objFolder As Object, objFSO As Object
    Dim stream As TextStream
    Dim wbOrig As Workbook, OrigSheet As Worksheet
    Dim wbMaster As Workbook, MasterSht As Worksheet
    Dim TxtFile As String, LastLine As String, fileName As String
    Dim LastRow As Long, LastCol As Long, NewLastRow As Long
    Dim YearData() As String
    Dim MonthData() As String
    Dim DayData() As String
    Dim Precip() As String
    Dim Tmax() As String
    Dim Tmin() As String
    Dim SolRad() As String
    Dim RelHum() As String
    Dim SunHours() As String
    Dim WindSpd() As String
    Dim OutputText() As String
    Dim precipspace As Integer
    Dim tmaxspace As Integer
    Dim tminspace As Integer
    Dim solradspace As Integer
    Dim relhumspace As Integer
    Dim sunhrspace As Integer
    Dim windspace As Integer
    Dim FileCount As Integer
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    
    '-------------------------------------------------------------
    ' Loop through all the .txt files within the folder
    '-------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strPath)
    FileCount = 0
    
    logfile.WriteLine "Looping through all AB10K .TXT files."
    For Each objFILE In objFolder.Files
        logtxt = "Checking file..." & objFILE
        Debug.Print logtxt
        logfile.WriteLine logtxt

        If UCase(Right(objFILE.Path, (Len(objFILE.Path) - InStrRev(objFILE.Path, ".")))) = UCase("txt") Then
            FileCount = FileCount + 1
            Set wbOrig = Workbooks.Open(objFILE.Path)
            Set OrigSheet = wbOrig.Worksheets(1)
            TxtFile = OrigSheet.Name
            OrigSheet.Activate
            
            fileName = ReturnOutputFile(strOutPath, "comp_" & TxtFile)
            Set stream = objFSO.CreateTextFile(fileName, True)
            
            Call FindLastRowColumn(LastRow, LastCol)
            NewLastRow = LastRow - 1
            
            ReDim YearData(0 To NewLastRow)
            ReDim MonthData(0 To NewLastRow)
            ReDim DayData(0 To NewLastRow)
            ReDim Precip(0 To NewLastRow)
            ReDim Tmax(0 To NewLastRow)
            ReDim Tmin(0 To NewLastRow)
            ReDim SolRad(0 To NewLastRow)
            ReDim RelHum(0 To NewLastRow)
            ReDim SunHours(0 To NewLastRow)
            ReDim WindSpd(0 To NewLastRow)
            ReDim OutputText(0 To NewLastRow)
                  
            For i = LBound(YearData) To UBound(YearData)
                YearData(i) = Range("A1").Offset(i, 0).Value
                MonthData(i) = Format(Range("B1").Offset(i, 0).Value, "00")
                DayData(i) = Format(Range("C1").Offset(i, 0).Value, "00")
                Precip(i) = Format(Range("D1").Offset(i, 0).Value, "00.0")
                Tmax(i) = Format(Range("E1").Offset(i, 0).Value, "00.0")
                Tmin(i) = Format(Range("F1").Offset(i, 0).Value, "00.0")
                SolRad(i) = Format(Range("G1").Offset(i, 0).Value, "0.00")
                RelHum(i) = Format(Range("H1").Offset(i, 0).Value, "0.00")
                SunHours(i) = Format(Range("I1").Offset(i, 0).Value, "0.00")
                WindSpd(i) = Format(Range("J1").Offset(i, 0).Value, "0.00")
                
                ' Define Spacing
                precipspace = 1
                tmaxspace = 2
                tminspace = 2
                solradspace = 2
                relhumspace = 1
                sunhrspace = 2
                windspace = 1
                
                If Len(Precip(i)) = 5 Then precipspace = 0 ' More than 3 significant values ie flood event
                If Len(Tmax(i)) = 5 Then tmaxspace = 1 ' Negative Values
                If Len(Tmin(i)) = 5 Then tminspace = 1 ' Negative Values
                If Len(SolRad(i)) > 4 Then solradspace = 1
                If Len(RelHum(i)) < 4 Then relhumspace = 2
                If Len(SunHours(i)) > 4 Then sunhrspace = 1
                
                ' Output Text Values
                If Not i = UBound(YearData) Then
                    OutputText(i) = Space(6) & YearData(i) & MonthData(i) & DayData(i) & Space(precipspace) & Precip(i) & _
                                Space(tmaxspace) & Tmax(i) & Space(tminspace) & Tmin(i) & Space(7) & "-99.900" & _
                                Space(49) & Space(solradspace) & SolRad(i) & Space(relhumspace) & RelHum(i) & _
                                Space(sunhrspace) & SunHours(i) & Space(windspace) & WindSpd(i) & vbCrLf
                Else
                    OutputText(i) = Space(6) & YearData(i) & MonthData(i) & DayData(i) & Space(precipspace) & Precip(i) & _
                                Space(tmaxspace) & Tmax(i) & Space(tminspace) & Tmin(i) & Space(7) & "-99.900" & _
                                Space(49) & Space(solradspace) & SolRad(i) & Space(relhumspace) & RelHum(i) & _
                                Space(sunhrspace) & SunHours(i) & Space(windspace) & WindSpd(i)
                End If
                stream.Write OutputText(i)
            Next i
            wbOrig.Close SaveChanges:=False
            stream.Close
            logfile.WriteLine "Created composite file for: " & objFILE
        End If
    Next
    logfile.WriteBlankLines (1)
    logtxt = "There are " & FileCount & " files processed."
    Debug.Print logtxt
    logfile.WriteLine logtxt
    
Cancel:
    Set wbOrig = Nothing
    Set OrigSheet = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set stream = Nothing
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
            Temp = "000" & fName
            
        Case 2
            Temp = "00" & fName
            
        Case 3
            Temp = "0" & fName
    End Select
    
    reDefineName = Temp
    
End Function

