Attribute VB_Name = "modGameLogic"
' === modGameLogic ===

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Gravity As Double
Public BirdVelocity As Double
Public BirdAngle As Double
Public IsFiring As Boolean

Sub LaunchBird()
'frmGame.Show

If frmGame.txtAngle.Text = "" Or frmGame.txtAngle.Text = "" Then
frmGame.txtAngle.Text = 10 And frmGame.txtAngle.Text = 20
End If

    Dim t As Double
    Dim x0 As Double, y0 As Double
    Dim X As Double, Y As Double
    Dim vx As Double, vy As Double
    
    Gravity = 9.8
    BirdAngle = CDbl(frmGame.txtAngle.Text) * (WorksheetFunction.Pi() / 180)
    BirdVelocity = CDbl(frmGame.txtPower.Text)
    
    vx = BirdVelocity * Cos(BirdAngle)
    vy = BirdVelocity * Sin(BirdAngle)
    
    x0 = frmGame.imgBird.Left
    y0 = frmGame.imgBird.Top
    
    
    
    
    IsFiring = True
    
    For t = 0 To 5 Step 0.1
        If Not IsFiring Then Exit For
        
        X = vx * t
        Y = vy * t - 0.5 * Gravity * t ^ 2
        
        frmGame.imgBird.Left = x0 + X * 5
        frmGame.imgBird.Top = y0 - Y * 5
        
        DoEvents
        Sleep 30
        
        ' Check for collision
        If IsColliding(frmGame.imgBird, frmGame.imgPig) Then
            frmGame.imgPig.Visible = False
            Exit For
        End If
    Next t
End Sub

Sub ResetGame()
    With frmGame
        .imgBird.Left = 50
        .imgBird.Top = 300
        .imgPig.Left = 600
        .imgPig.Top = 300
        .imgPig.Visible = True
       '.txtAngle.Text = "45"
       ' .txtPower.Text = "50"
       
       If frmGame.txtAngle.Text = "" Or frmGame.txtAngle.Text = "" Then
frmGame.txtAngle.Text = "10"
frmGame.txtPower.Text = "20"
End If
       
    End With
     LoadLevel 1
End Sub

