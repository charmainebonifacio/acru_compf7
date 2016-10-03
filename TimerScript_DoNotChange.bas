Attribute VB_Name = "TimerScript_DoNotChange"
Option Explicit
'---------------------------------------------------------------------------------------
' Date Acquired: April 16, 2013
' http://www.vbforums.com/showthread.php?424001-Timed-message-box-Resolved
'---------------------------------------------------------------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Const WM_CLOSE      As Long = 16
Private CurMBTitle          As String
Private Sub TimeOutMB(hwnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)
 
    SendMessage FindWindow(vbNullString, CurMBTitle), WM_CLOSE, 0&, 0&
 
End Sub
'---------------------------------------------------------------------------------------
' Date Acquired: April 16, 2013
' http://www.vbforums.com/showthread.php?424001-Timed-message-box-Resolved
'---------------------------------------------------------------------------------------
' Date Edited  : April 17, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : TimedMsgBox
' Description  : This function will notify user that tool is currently processing the
'                user request to download, process and merge files.
' Parameters   : String, Long, String
' Returns      : -
'---------------------------------------------------------------------------------------
Sub TimedMsgBox(ByVal Prompt As String, ByVal Timeout As Long, ByVal title As String)
    
    Dim ResponseValue As Integer
    Dim TimerId   As Long
    CurMBTitle = title
    
    Timeout = Timeout * 1000
    TimerId = SetTimer(0, 0, Timeout, AddressOf TimeOutMB)
    Debug.Print TimerId
    
    ResponseValue = MsgBox(Prompt, vbExclamation + vbOKOnly, CurMBTitle)
    ' Check pressed button
    If ResponseValue = vbOK Then
         Debug.Print "User acknowledge warning window."
    Else: Debug.Print "Closed window."
    End If
    
    KillTimer 0, TimerId
End Sub
 

