Attribute VB_Name = "modCollision"
' === modCollision ===

Function IsColliding(ctrl1 As Control, ctrl2 As Control) As Boolean
    IsColliding = Not (ctrl1.Left + ctrl1.Width < ctrl2.Left Or _
                       ctrl1.Left > ctrl2.Left + ctrl2.Width Or _
                       ctrl1.Top + ctrl1.Height < ctrl2.Top Or _
                       ctrl1.Top > ctrl2.Top + ctrl2.Height)
End Function


