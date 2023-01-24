Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Sub JumpButton_Click()
    ' Initialize the game
    Application.DisplayAlerts = False
    Dim Player As Shape
    Set Player = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50)
    Player.Fill.ForeColor.RGB = RGB(255, 0, 0)
    Dim Wall1 As Shape
    Set Wall1 = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 100, 100, 50, 150)
    Wall1.Fill.ForeColor.RGB = RGB(0, 255, 0)
    Dim Wall2 As Shape
    Set Wall2 = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 200, 50, 100)
    Wall2.Fill.ForeColor.RGB = RGB(0, 255, 0)
    Dim GameOver As Boolean
    Dim speed As Integer
    Dim HighScore As Integer
    speed = 0
    HighScore = 0
    Cells(1, 1) = HighScore
    GameOver = False

    ' Start the game loop
    Do
        ' Get the player and wall positions
        Dim PlayerTop As Integer
        PlayerTop = Player.Top
        Dim PlayerBottom As Integer
        PlayerBottom = Player.Top + Player.Height
        Dim Wall1Top As Integer
        Wall1Top = Wall1.Top
        Dim Wall1Bottom As Integer
        Wall1Bottom = Wall1.Top + Wall1.Height
        Dim Wall2Top As Integer
        Wall2Top = Wall2.Top
        Dim Wall2Bottom As Integer
        Wall2Bottom = Wall2.Top + Wall2.Height
        Cells(1, 1) = HighScore
        ' Check for collision with walls
If Player.Left + Player.Width >= Wall1.Left And Player.Left <= Wall1.Left + Wall1.Width And Player.Top + Player.Height >= Wall1.Top And Player.Top <= Wall1.Top + Wall1.Height Then
    GameOver = True
    MsgBox "Game Over!"
    Exit Sub
End If
If Player.Left + Player.Width >= Wall2.Left And Player.Left <= Wall2.Left + Wall2.Width And Player.Top + Player.Height >= Wall2.Top And Player.Top <= Wall2.Top + Wall2.Height Then
    GameOver = True
    MsgBox "Game Over!"
    Exit Sub
End If

        ' Move the walls
        Wall1.Left = Wall1.Left - (0.1 + speed)
        Wall2.Left = Wall2.Left - (0.1 + speed)
        
        If Wall1.Left < Player.Left Then
            Wall1.Left = Player.Left + 800
            Wall1.Top = Player.Top - 3
            speed = speed + 0.8
            HighScore = HighScore + 1
        End If
        If Wall2.Left < Player.Left Then
            Wall2.Left = Player.Left + 500
            Wall2.Top = Player.Top + 55
            speed = speed + 0.8
            HighScore = HighScore + 1
        End If


         If GetAsyncKeyState(38) < 0 Then
            Player.Top = Player.Top - 3
        End If
        'Check for Down arrow key press
        If GetAsyncKeyState(40) < 0 Then
            Player.Top = Player.Top + 3
        End If

        ' Allow Excel to process other events before looping
        DoEvents
    Loop Until GameOver = True
End Sub

