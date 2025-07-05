Attribute VB_Name = "modLevels"
' === modLevels ===

Sub LoadLevel(LevelNumber As Integer)
    ' Clear previous items
    ClearLevelObjects
    
    Select Case LevelNumber
        Case 1
            ' Place 1 pig and 2 blocks
            AddPig 600, 300
            AddBlock 450, 320
            AddBlock 420, 320
            
        Case 2
            ' A higher pig and stacked blocks
            AddBlock 480, 250
            AddBlock 480, 230
            AddPig 480, 210
            
        Case 3
            ' More complex arrangement
            AddBlock 470, 250
            AddBlock 410, 250
            AddBlock 410, 250
            AddBlock 490, 230
            AddPig 490, 210
            
        Case Else
            MsgBox "Level not found!"
    End Select
End Sub
Sub AddPig(X As Single, Y As Single)
    Dim pig As Control
    Set pig = frmGame.Controls.Add("Forms.Image.1", "Pig_" & frmGame.Controls.Count, True)
    pig.Left = X
    pig.Top = Y
    pig.Width = 40
    pig.Height = 40
    pig.Picture = LoadPicture("E:\My Git Projects\13. Bird getting Angry\V1\imgpig.bmp") ' update path
    pig.Tag = "Pig"
End Sub

Sub AddBlock(X As Single, Y As Single)
    Dim block As Control
    Set block = frmGame.Controls.Add("Forms.Image.1", "Block_" & frmGame.Controls.Count, True)
    block.Left = X
    block.Top = Y
    block.Width = 40
    block.Height = 20
    block.Picture = LoadPicture("E:\My Git Projects\13. Bird getting Angry\V1\block.bmp") ' update path
    block.Tag = "Block"
End Sub

Sub ClearLevelObjects()
    Dim ctrl As Control
    Dim i As Integer
    For i = frmGame.Controls.Count - 1 To 0 Step -1
        Set ctrl = frmGame.Controls.Item(i)
        If ctrl.Tag = "Pig" Or ctrl.Tag = "Block" Then
            frmGame.Controls.Remove ctrl.Name
        End If
    Next i
End Sub

