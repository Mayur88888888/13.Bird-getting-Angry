VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGame 
   Caption         =   "UserForm1"
   ClientHeight    =   8640.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16512
   OleObjectBlob   =   "frmGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === frmGame Code ===

Private Sub cmdLaunch_Click()
    Call LaunchBird
End Sub

Private Sub cmdReset_Click()
    Call ResetGame
End Sub

Private Sub cmdNextLevel_Click()
    Static Level As Integer
    Level = Level + 1
    If Level > 3 Then Level = 1
    LoadLevel Level
End Sub


Private Sub imgPig_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Initialize()
    Call ResetGame
End Sub

