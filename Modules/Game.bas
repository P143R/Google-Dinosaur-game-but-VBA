Attribute VB_Name = "Game"

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" _
(ByVal vKey As Long) As Integer

Dim Game As Worksheet
Dim setting As Worksheet

Dim dino2 As Shape
Dim dino0 As Shape
Dim dino1 As Shape
Dim dinoLose As Shape

Dim cactus0 As Shape
Dim cactus1 As Shape
Dim cactus2 As Shape

Dim ground0 As Shape
Dim ground1 As Shape
Dim ground2 As Shape

Dim score As Shape
Dim press As Shape

Dim animate As Boolean
Dim isJumping As Boolean

Dim dinoIndex As LongLong
Dim scoreNumber As Double

Dim cactusInterval As LongLong
Dim speed As LongLong
Dim sinValue As Double
Dim deltaTime As Double
Dim multiplierValue As Double
Dim currentFrameTime As Double
Dim FrameTime As Integer

Dim frameLimit As Integer

'Reset/Default settings
Sub reset()
    setter
    'Dinos
    dino2.Top = 291
    dino2.Left = 1450
    dino0.Top = dino2.Top
    dino0.Left = dino2.Left
    dino1.Top = dino2.Top
    dino1.Left = dino2.Left
    dinoLose.Top = dino2.Top
    dinoLose.Left = dino2.Left
    'Score and press to start
    score.Left = Application.Width - 200 + ground0.Width
    press.Left = Application.Width / 2 - 100 + ground0.Width
    'Grounds
    ground0.Top = 375
    ground0.Left = ground0.Width
    ground1.Top = ground0.Top
    ground1.Left = ground0.Width * 2 - 50
    ground2.Top = ground0.Top
    ground2.Left = ground0.Width * 3 - 50
    'Cactus
    cactus0.Top = 315
    cactus1.Top = cactus0.Top
    cactus2.Top = cactus0.Top
    cactus0.Left = ground0.Width * 3
    cactus1.Left = ground0.Width * 4
    cactus2.Left = ground0.Width * 5
    
    'Scroll
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollRow = 1
    
    'Score
    scoreNumber = 0
    score.TextFrame.Characters.Text = "00000"
    'Visibility of shapes
    dino2.Visible = True
    dino0.Visible = False
    dino1.Visible = False
    dinoLose.Visible = False
    press.Visible = True
    'Variable values
    multiplierValue = 4.5
    speed = 100 * multiplierValue
    dinoIndex = 0
    sinValue = 0
    frameLimit = 60
    FrameTime = 100
    currentFrameTime = 0
    
End Sub
'Setting shapes
Sub setter()
    Set Game = ThisWorkbook.Sheets("Game")
    Set dino2 = Game.Shapes("dino2")
    Set dino0 = Game.Shapes("dino0")
    Set dino1 = Game.Shapes("dino1")
    Set dinoLose = Game.Shapes("dinoLose")
    
    Set cactus0 = Game.Shapes("cactus0")
    Set cactus1 = Game.Shapes("cactus1")
    Set cactus2 = Game.Shapes("cactus2")
    
    Set score = Game.Shapes("score")
    Set ground0 = Game.Shapes("ground0")
    Set ground1 = Game.Shapes("ground1")
    Set ground2 = Game.Shapes("ground2")
    Set press = Game.Shapes("pressSpace")
End Sub
'isJumping function
Sub jump()
    isJumping = True
End Sub
'Dino jump function
Sub handleJump()
    If isJumping Then
        dino2.Visible = True
        dino0.Visible = False
        dino1.Visible = False
        sinValue = sinValue + 1 * deltaTime * multiplierValue
        dino2.Top = 291 - dino2.Height * Sin(sinValue) * 2
        If dino2.Top >= 291 Then
            isJumping = False
            dino2.Top = 291
            sinValue = 0
            dino2.Visible = False
            dino0.Visible = True
        End If
    End If
End Sub
Sub update()
    startTime = 0
    While animate
        DoEvents
            startTime = Timer
            spacebarChecker
            handleJump
            dinoAnimation
            moveGround
            moveCactus
            scoreCounter
            increaseSpeed
            isDinoDead
            deltaTime = Timer - startTime
            While deltaTime < (1 / frameLimit)
            DoEvents
            deltaTime = Timer - startTime
            Wend
            Range("AE2").Value = Round(1 / deltaTime, 2)
    Wend
    dinoLose.Visible = True
    dino2.Visible = False
    press.Visible = True
End Sub
Sub spacebarChecker()
    If GetAsyncKeyState(&H20) Then
        jump
    End If
End Sub
'Start the game
Sub start()
    reset
    isJumping = False
    animate = True
    press.Visible = False
    dinoLose.Visible = False
    dino2.Visible = False
    dino0.Visible = True
    update
End Sub
Sub dinoAnimation()
    If isJumping = False Then
        If currentFrameTime >= FrameTime Then
            Game.Shapes("dino" + CStr(dinoIndex)).Visible = False
            dinoIndex = (dinoIndex + 1) Mod 2
            Game.Shapes("dino" + CStr(dinoIndex)).Visible = True
            currentFrameTime = currentFrameTime - FrameTime
        End If
        currentFrameTime = currentFrameTime + 1 * deltaTime * speed
    End If
End Sub
'Stops the game
Sub killDino()
    animate = False
End Sub
Sub scoreCounter()
    scoreNumber = scoreNumber + 1 * deltaTime * multiplierValue
    If scoreNumber < 10 Then
        score.TextFrame.Characters.Text = "0000" + CStr(Round(scoreNumber, 0))
    ElseIf scoreNumber >= 10 And scoreNumber < 100 Then
        score.TextFrame.Characters.Text = "000" + CStr(Round(scoreNumber, 0))
    ElseIf scoreNumber >= 100 And scoreNumber < 1000 Then
        score.TextFrame.Characters.Text = "00" + CStr(Round(scoreNumber, 0))
    ElseIf scoreNumber >= 1000 And scoreNumber < 10000 Then
        score.TextFrame.Characters.Text = "0" + CStr(Round(scoreNumber, 0))
    ElseIf scoreNumber >= 10000 And scoreNumber < 100000 Then
        score.TextFrame.Characters.Text = CStr(Round(scoreNumber, 0))
    End If
End Sub

Sub moveGround()
    ground0.Left = ground0.Left - speed * deltaTime
    ground1.Left = ground1.Left - speed * deltaTime
    ground2.Left = ground2.Left - speed * deltaTime
    If ground0.Left <= 50 Then
        ground0.Left = ground0.Width * 3
    ElseIf ground1.Left <= 50 Then
        ground1.Left = ground0.Width * 3
    ElseIf ground2.Left <= 50 Then
        ground2.Left = ground0.Width * 3
    End If
End Sub

Sub moveCactus()
    cactus0.Left = cactus0.Left - speed * deltaTime
    cactus1.Left = cactus1.Left - speed * deltaTime
    cactus2.Left = cactus2.Left - speed * deltaTime
    If cactus0.Left <= 50 Then
        cactus0.Left = ground0.Width * 3
    ElseIf cactus1.Left <= 50 Then
        cactus1.Left = ground0.Width * 3
    ElseIf cactus0.Left <= 50 Then
        cactus2.Left = ground0.Width * 3
    End If
End Sub

Sub isDinoDead()
    If (cactus0.Left <= (dino2.Left + dino2.Width) And (cactus0.Left + cactus0.Width) >= dino2.Left And cactus0.Top >= (dino2.Top - dino2.Height) And (cactus0.Top - cactus0.Height) <= dino2.Top) Then
        killDino
    ElseIf (cactus1.Left <= (dino2.Left + dino2.Width) And (cactus1.Left + cactus1.Width) >= dino2.Left And cactus1.Top >= (dino2.Top - dino2.Height) And (cactus1.Top - cactus1.Height) <= dino2.Top) Then
        killDino
    ElseIf (cactus2.Left <= (dino2.Left + dino2.Width) And (cactus2.Left + cactus2.Width) >= dino2.Left And cactus2.Top >= (dino2.Top - dino2.Height) And (cactus2.Top - cactus2.Height) <= dino2.Top) Then
        killDino
    End If
End Sub
Sub increaseSpeed()
    speed = speed * 1.0005
End Sub


