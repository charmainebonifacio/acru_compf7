Attribute VB_Name = "Step_0_Start"
'---------------------------------------------------------------------
' Date Created : June 6, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : September 2, 2016
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Start_Here
' Description  : The purpose of function is to initialize the userform.
'---------------------------------------------------------------------
Sub Start_Here()
   
    Dim button1 As String, button2 As String, button3 As String
    Dim button4 As String, button5 As String, button6 As String
    Dim strLabel1 As String, strLabel2 As String
    Dim strLabel3 As String, strLabel4 As String
    Dim strLabel5 As String, strLabel6 As String
    Dim strLabel7 As String, strLabel8 As String
    Dim frameLabel1 As String, frameLabel2 As String, frameLabel3 As String
    Dim userFormCaption As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Label Strings
    userFormCaption = "KIENZLE LAB TOOLS"
    button1 = "CREATE 7VAR COMPOSITE FILE"
    frameLabel2 = "TOOL GUIDE"
    frameLabel3 = "HELP SECTION"
    
    strLabel1 = "THE ACRU-HYDRO-CLIMATOLOGICAL DATA FILE GENERATOR"
    strLabel2 = "STEP 1."
    strLabel3 = "Run Zonal Statistics (as Table) Python Script." & vbLf
    strLabel4 = "STEP 2."
    strLabel5 = "For more information, hover mouse over button."
    
    ' UserForm Initialize
    ACRU_COMPF7_Form.Caption = userFormCaption
    ACRU_COMPF7_Form.Frame2.Caption = frameLabel2
    ACRU_COMPF7_Form.Frame5.Caption = frameLabel3
    ACRU_COMPF7_Form.Frame2.Font.Bold = True
    ACRU_COMPF7_Form.Frame5.Font.Bold = True
    ACRU_COMPF7_Form.Label1.Caption = strLabel1
    ACRU_COMPF7_Form.Label1.Font.Size = 17
    ACRU_COMPF7_Form.Label1.Font.Bold = True
    ACRU_COMPF7_Form.Label1.TextAlign = fmTextAlignCenter
    
    ACRU_COMPF7_Form.Label2 = strLabel2
    ACRU_COMPF7_Form.Label2.Font.Size = 13
    ACRU_COMPF7_Form.Label2.Font.Bold = True
    ACRU_COMPF7_Form.Label3 = strLabel3
    ACRU_COMPF7_Form.Label3.Font.Size = 11
    ACRU_COMPF7_Form.Label4 = strLabel4
    ACRU_COMPF7_Form.Label4.Font.Size = 13
    ACRU_COMPF7_Form.Label4.Font.Bold = True
    ACRU_COMPF7_Form.CommandButton1.Caption = button1
    ACRU_COMPF7_Form.CommandButton1.Font.Size = 11
    
    ' Help File
    ACRU_COMPF7_Form.Label5 = strLabel5
    ACRU_COMPF7_Form.Label5.Font.Size = 8
    ACRU_COMPF7_Form.Label5.Font.Italic = True
    
    Application.StatusBar = "Macro has been initiated."
    ACRU_COMPF7_Form.Show

End Sub
'---------------------------------------------------------------------------------------
' Date Created : July 18, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HELPFILE
' Description  : This function will feed the help tip section depending on the button
'                that has been activated.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function HELPFILE(ByVal Notification As Integer) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case 1
            NotifyUser = "TITLE: COMPLETE 7VARIABLE COMPOSITE FILE" & vbLf
            NotifyUser = NotifyUser & "DESCRIPTION: This macro will append " & _
                "RAD, REL, SUN, and WND values to the original composite files. This will " & _
                "create the composite files, which will contain 7-variables." & vbLf
            NotifyUser = NotifyUser & "INPUT: Find the location of the ACRU_COMPF7 folder" & vbLf
            NotifyUser = NotifyUser & "OUTPUT: (4).DAT Files, Harmonic Analysis .OUT Files, Composite .TXT Files" & vbLf
    End Select
    
    HELPFILE = NotifyUser
    
End Function
'---------------------------------------------------------------------
' Date Created : September 22, 2014
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : September 22, 2014
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveLogFile
' Description  : This function saves file as .TXT.
'                When new file is named after an existing file, the
'                same name is used with an number attached to it.
' Parameters   : String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveLogFile(ByVal fileDir As String, _
ByVal fileName As String, ByVal fileExt As String) As String

    Dim saveFile As String
    Dim formatDate As String
    Dim saveDate As String
    Dim saveName As String
    Dim sPath As String

    ' Date
    formatDate = Format(Date, "MM/dd/yyyy")
    saveDate = Replace(formatDate, "/", "")
    
    ' Save information as Temp, which can then be renamed later..
    sPath = fileDir
    If Right(fileDir, 1) <> "\" Then sPath = fileDir & "\"
    saveName = fileName & "_" & saveDate & fileExt
    
    ' Rename existing file
    i = 1
    If CheckFileExists(sPath, saveName) = True Then
        If Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) <> "" Then
            Do Until Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) = ""
                i = i + 1
            Loop
            saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        Else: saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        End If
    Else: saveFile = sPath & fileName & "_" & saveDate & fileExt
    End If
    
    SaveLogFile = saveFile
    
End Function

