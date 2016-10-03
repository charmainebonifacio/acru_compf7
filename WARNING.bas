Attribute VB_Name = "WARNING"
'---------------------------------------------------------------------------------------
' Date Created : June 12, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 8, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : WarningMessage
' Description  : This function will notify user that tool is currently processing the
'                user request to split the timeseries into two new timeseries.
' Parameters   : -
' Returns      : -
'---------------------------------------------------------------------------------------
Function WarningMessage()

    Dim WarningPrompt As String
    Dim WindowTitle As String
    Dim DefaultTimer As Long
    
    DefaultTimer = 100 ' < Set Timer
    
    WindowTitle = "The Processing Zonal Statistics Tool"
    WarningPrompt = "ATTENTION." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "The macro is currently processing your " & _
        "request. First, it will process Zonal Statistics (.DBF) files into one summary (.DAT) file. " & _
        "Second, it will run the Harmonic Analysis program in order to create corresponding (.OUT) files " & _
        "for each selected Alberta 10KGrid. Last, these values for RADIATION, RELATIVE HUMIDITY, " & _
        "SUNSHINE HOURS and WIND RUN will be appended the original AB10K Grid files." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    WarningPrompt = WarningPrompt & "Please click [OK] to continue." & vbCrLf
    WarningPrompt = WarningPrompt & vbCrLf
    
    TimedMsgBox WarningPrompt, DefaultTimer, WindowTitle ' Call New MsgBox
    
End Function
'---------------------------------------------------------------------------------------
' Date Created : June 13, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : June 13, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : MacroTimer
' Description  : This function will notify user how much time has elapsed to complete
'                the entire procedure.
' Parameters   : Long
' Returns      : String
'---------------------------------------------------------------------------------------
Function MacroTimer(ByVal TimeElapsed As Long) As String

    Dim NotifyUser As String
    
    NotifyUser = "MACRO RUN IS SUCCESSFUL!"
    NotifyUser = NotifyUser & vbCrLf
    NotifyUser = NotifyUser & "The macro has processed .DBF files and created new composite files." & vbCrLf
    NotifyUser = NotifyUser & "The macro run took a total of " & TimeElapsed & " minutes."

    MacroTimer = NotifyUser
    
End Function
