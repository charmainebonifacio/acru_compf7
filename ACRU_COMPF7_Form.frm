VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ACRU_COMPF7_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   OleObjectBlob   =   "ACRU_COMPF7_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ACRU_COMPF7_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Date Created : August 23, 2012
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : August 23, 2012
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
'---------------------------------------------------------------------------------------

Private Sub CommandButton1_Click()
' Download Environment Data

    If Val(Application.Version) < 12 Then
        MsgBox "You are using Microsoft Excel 2003 and older."
    Else
        ACRU_COMPF7_Form.Hide
        Debug.Print "Microsoft Excel 2007 or higher."
        Call ACRU_COMPF7_MAIN
    End If

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
' Description  : Placed section for processing Help Tips.
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
    ACRU_COMPF7_Form.Label6.Visible = False
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim strLabel6 As String
    ACRU_COMPF7_Form.Label6.Visible = True
    ACRU_COMPF7_Form.Label6.BackColor = RGB(255, 255, 153)
    strLabel6 = HELPFILE(1)
    ACRU_COMPF7_Form.Label6 = strLabel6
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ACRU_COMPF7_Form.Label6.Visible = False
End Sub
