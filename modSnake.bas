Attribute VB_Name = "modSnake"
Public Type SnakeData
    Pos(2) As Integer   'the grid co-ordinates - not pixels
    Direction As Byte   'what direction is this element going in (see frmSnake.GetSnakeRect)
End Type

Public Type RouteCol
    FoundRoute(20 ^ 2) As SnakeData
    RLen As Integer
End Type

Public Type ScoresData
    Name As String      'player name
    Score As Integer    'player score
End Type

'you can set the different difficulty levels in the procedure
'SetDifficulty
Public Type DifficultyType
    RoomSize As Integer
    FoodAmount As Integer
    StartingSpeed As Integer
    MaxSpeed As Integer
End Type

Public Const ScoresFile = "Scores.ini" 'the name of the ini file for storing the game settings and scores
Public Const GridSize = 7  'the size is in pixels
Public Const MinRoomSize = 15
Public Const MaxRoomSize = 65
Public Const X = 0 'used for referencing position arrays
Public Const Y = 1 'used for referencing position arrays
Public Const MinFoodAmount = 1 'at any one time
Public Const MaxFoodAmount = 10 'at any one time
Public Const PointsPerFood = 5 'the amount your score goes up per food item
Public Const EasyStartSpeed = 200   'measured in ticks (1 second = 1000 ticks)
Public Const MaxSpeed = 10 'the minimum number of ticks to wait - ie the fastest the snake can move
Public Const MinSpeed = 250 'the slowest the snake can go
Public Const MaxAISpeed = 0 'The speed of the AI. We don't want the AI to go faster tan this no matter how powerfull the processor
Public Const SnakeStartSize = 5     'in grid units
Public Const SnakeColour = vbGreen
Public Const FoodColour = vbRed
Public Const Growth = 4 ' "Growth * (Score / PointsPerFood / GridSize)" : see procedure MoveSnake
Public Const Easy = 1 'used to reference the difficulty settings
Public Const Normal = 2
Public Const Hard = 3


Public Snake(MaxRoomSize ^ 2) As SnakeData 'maximum snake size is the size of the room
Public SnakeSize As Integer 'the current snake size
Public CurDirection As Byte 'the current direction of the computer
Public TempDirection As Byte 'records what key has been pressed during the delay peroid (see Speed)
Public Food(MaxFoodAmount, 2) As Integer 'an array holding where each item of food is on the grid
Public FoodAmount As Integer 'the number of food items currently displayed
Public FoodPerTick As Integer
Public Score As Long 'the current score
Public Speed As Integer 'the current snake speed (measured in ticks)
Public StartSpeed As Integer 'the speed the player will start playing at
Public TopSpeed As Integer 'the maximum speed the player can go
Public GrowAmount As Integer 'how much the snake has to grow (when it eats food)
Public StartFrameTick As Long 'used to calculate how long it took to make each frame. It it then subtracted from the speed to give an accurate frame rate
Public AICheat As Boolean 'let the AI cheat (it never dies)
Public AISpeed As Integer 'the speed of the AI
Public AIScore As Boolean
Public Scores(10) As ScoresData 'holds the player scores
Public Difficulty(3) As DifficultyType  'holds the different difficulty levels. These are set in the procedure SetDifficulty
Public DifficultyLevel As Byte 'the current level of difficulty
Public RoomSize As Byte 'the size of the room the snake will be in. Measured in blocks.

'Aritificial Intelligence
'store the route to the food
Dim Route(MaxRoomSize ^ 2) As SnakeData
Dim RouteLen As Integer
Dim RouteCollection(MaxRoomSize ^ 2) As RouteCol 'stores all possible routes to the food
Dim NumOfRoutes As Integer


Public Sub MoveSnake()
'This moves the snake in the currect direction

Dim Counter As Integer
Dim Index As Integer
Dim NewPos(2) As Integer

'start counting how long it take to exicute this code
StartFrameTick = GetTickCount

'update the current direction
CurDirection = TempDirection

'if collision, then game over
If Not AICheat Then
    If IsCollision(CurDirection) Then
        Call frmSnake.GameOver
    End If
Else
    'reset if the snake size is 2/3 the grid size
    If SnakeSize >= (((RoomSize ^ 2) / 3) * 2) Then
        Call frmSnake.GameOver
    End If
End If


'move the snake
NewPos(X) = Snake(0).Pos(X)
NewPos(Y) = Snake(0).Pos(Y)

'get the new co-ordinates for new position
Select Case CurDirection
Case KLeft
    NewPos(X) = NewPos(X) - 1
Case KUp
    NewPos(Y) = NewPos(Y) - 1
Case KRight
    NewPos(X) = NewPos(X) + 1
Case KDown
    NewPos(Y) = NewPos(Y) + 1
End Select

'check for food
'the array Food starts at 1 instead of 0 because if food amount is
'zero, there is no food
For Counter = 1 To FoodAmount
    If (NewPos(X) = Food(Counter, X)) And (NewPos(Y) = Food(Counter, Y)) Then
        'eat food
        Score = Score + PointsPerFood
        
        'remove that food item
        For Index = Counter To FoodAmount - 1
            Food(Index, X) = Food(Index + 1, X)
            Food(Index, Y) = Food(Index + 1, Y)
        Next Index
        
        'delete to last item in the index before decrementing the
        Food(FoodAmount, X) = 0
        Food(FoodAmount, Y) = 0
        FoodAmount = FoodAmount - 1
        Call NewFood(1)
        'if no more food, then create more
        'If FoodAmount < FoodPerTick Then
        '    Call NewFood(FoodPerTick - FoodAmount)
        'End If

        'increase the speed if necessary
        If ((Speed - (PointsPerFood * 2)) > MaxSpeed) And (Speed <> MaxSpeed) Then
            Speed = Speed - (PointsPerFood * 2)
        Else
            Speed = MaxSpeed
        End If
        
        'increase the snake length
        GrowAmount = GrowAmount + (Growth * (Score / PointsPerFood / GridSize))
        
        'food found, exit
        Exit For
    End If
Next Counter


'bubble the co-ordinates and direction from the end to start
For Counter = SnakeSize + 1 To 1 Step -1
    Snake(Counter) = Snake(Counter - 1)
Next Counter

'let the snake grow!
If GrowAmount > 0 Then
    'increase the size of the snake once per turn
    'until growth is complete.
    SnakeSize = SnakeSize + 1
    GrowAmount = GrowAmount - 1
End If

'move the head of the snake in the direction it's heading
Snake(0).Pos(X) = NewPos(X)
Snake(0).Pos(Y) = NewPos(Y)
Snake(0).Direction = CurDirection

'if AI is active, don't bother with speed
If (frmSnake.timAI.Enabled) And (Speed <> AISpeed) Then
    Speed = AISpeed
End If

'subtract the length of time it took to exicute the code from the
'frame limiter (Speed)
'pause before drawing the frame (speed)
If frmSnake.timAI.Enabled Then
    'computer playing
    Call Pause(AISpeed - (GetTickCount - StartFrameTick))
Else
    'human player
    Call Pause(Speed - (GetTickCount - StartFrameTick))
End If

'start counting how long it takes to create a frame
StartFrameTick = GetTickCount

'draw the grid + snake + food
Call frmSnake.DrawFrame
End Sub

Private Function IsCollision(MyDirection) As Boolean ', Optional MyX As Integer = Snake(0).Pos(X), Optional MyY As Integer = Snake(0).Pos(Y)) As Boolean
'returns whether or not there was a collision with either the
'grid walls or the snake body.

Dim Counter As Integer
Dim ColPos(2) As Integer
Dim Collision As Boolean

'set the collision position to snake head
ColPos(X) = Snake(0).Pos(X)
ColPos(Y) = Snake(0).Pos(Y)
Collision = False

'first check for out of bounds collision
Select Case MyDirection
Case KLeft
    'hit the left wall
    ColPos(X) = ColPos(X) - 1
    If ColPos(X) < 0 Then
        Collision = True
    End If
    
Case KUp
    'hit the top wall
    ColPos(Y) = ColPos(Y) - 1
    If ColPos(Y) < 0 Then
        Collision = True
    End If
    
Case KRight
    'hit the right wall
    ColPos(X) = ColPos(X) + 1
    If ColPos(X) >= RoomSize Then
        Collision = True
    End If
    
Case KDown
    'hit the bottom wall
    ColPos(Y) = ColPos(Y) + 1
    If ColPos(Y) >= RoomSize Then
        Collision = True
    End If
End Select

'next check for collision with snake if a collision has not been found
If Not Collision Then
    Collision = IsSnake(ColPos(X), ColPos(Y))
End If
'For Counter = 0 To SnakeSize
'    If (ColPos(X) = Snake(Counter).Pos(X)) And (ColPos(Y) = Snake(Counter).Pos(Y)) Then
'        'collision found, exit loop (no point to keep searching)
'        Collision = True
'        Exit For
'    End If
'Next Counter

'return result
IsCollision = Collision
End Function

Public Sub GetDirection()
'This will change the setting of the Direction variable according
'to which arrow key was pressed, only if the current form is active

Dim Counter As Byte
Dim KeyState As Integer

'if the form is not active then don't change direction
If Not AmIActive(frmSnake) Then
    Exit Sub
End If

'find out if the arrow keys have been pressed
For Counter = KLeft To KDown 'from 37 to 40, moving clockwise
    KeyState = GetAsyncKeyState(CLng(Counter))
    
    If KeyState <> 0 Then
        'the key was or is pressed
        
        'check to make sure that the new direction is not opposite to
        'the current direction.
        'Note :  there is a difference of 2 between the values
        'KUp, KDown and KLeft, KRight. I'm using this to make sure
        'the snake cannot move in on itself.
        If PositVal((CurDirection - KLeft) - (Counter - KLeft)) <> 2 Then
            'the directions are not opposite
            'set the new direction and exit
            TempDirection = Counter
            Exit For
        End If
    End If
Next Counter
End Sub

Public Sub ResetSnake()
'reset the initial conditions for the snake

Dim Counter As Integer
Dim Pos(2) As Integer

'the snake will always start in the centre of the screen, moving left
Pos(X) = ((RoomSize / 2) - (SnakeStartSize / 2))
Pos(Y) = (RoomSize / 2)
For Counter = 0 To SnakeStartSize
    Snake(Counter).Pos(X) = Pos(X) + Counter
    Snake(Counter).Pos(Y) = Pos(Y)
    Snake(Counter).Direction = KLeft
Next Counter
SnakeSize = SnakeStartSize

'set the speed
Speed = StartSpeed

Call NewFood(FoodPerTick)

End Sub

Public Sub NewFood(Amount)
'this will create new food in random places around the grid

Dim NewPos(2) As Integer
Dim Counter As Integer
Dim Index As Integer
Dim Found As Boolean
Dim IsEmpty As Boolean

'amount has to be greater than zero to generate a new food item
If (Amount < 1) Or ((Amount + FoodAmount) > MaxFoodAmount) Then
    Exit Sub
End If

Randomize

IsEmpty = False
For Counter = (FoodAmount + 1) To (FoodAmount + Amount)
    Do While Not IsEmpty
        'create a new random value
        NewPos(X) = Int((RoomSize - 1) * Rnd)
        NewPos(Y) = Int((RoomSize - 1) * Rnd)
        
        'check the new position against existing food co-ordinates
        IsEmpty = True
        For Index = 1 To FoodAmount
            If (Food(Index, X) = NewPos(X)) And (Food(Index, Y) = NewPos(Y)) Then
                'collision with existing value
                IsEmpty = False
                Exit For
            End If
        Next Index
        
        If IsEmpty Then
            'check co-ordinates against where the snake is
            IsEmpty = Not IsSnake(NewPos(X), NewPos(Y))
        End If
    Loop
    
    'once a valid new co-ordinate has been generated then add it
    'to the food array.
    FoodAmount = FoodAmount + 1
    Food(Counter, X) = NewPos(X)
    Food(Counter, Y) = NewPos(Y)
    
    'create the next food item
    IsEmpty = False
Next Counter
End Sub

Public Sub GetAIDirection()
'This procedure will find where the snake is supposed to go and
'set the current direction

'Dim Counter As Integer
'Dim SnakeAt As Integer 'where the snake is along the route
'
'For Counter = 0 To RouteLen
'    'find where the snake head is on the route to the food
'    If (Snake(0).Pos(X) = Route(Counter).Pos(X)) And (Snake(0).Pos(Y) = Route(Counter).Pos(Y)) Then
'        'found position
'        SnakeAt = Counter
'        Exit For
'    End If
'Next Counter

'If SnakeAt = 0 Then
    'if no match found, find new route
    Call FindRoute
'End If

'set the direction for the AI
'TempDirection = Route(SnakeAt).Direction
End Sub

Public Sub FindRoute()
'Find the most direct route to the food.

Dim TryDir As Byte
Dim MyDir As Byte
Dim Able(4) As Boolean '1=left, 2=up, 3=right, 4=down
Dim Counter As Integer

'the snake cannot go is reverse
Select Case Snake(0).Direction
Case KLeft
    Able(3) = True
Case KUp
    Able(4) = True
Case KRight
    Able(1) = True
Case KDown
    Able(2) = True
End Select

'find the most obvious direction
MyDir = TryDirection(Able(1), Able(2), Able(3), Able(4)) 'Snake(0).Direction '

For Counter = 1 To 4
    If IsCollision(MyDir) Then 'Not DirPossible(MyDir) Then
        'get a new direction
        
        'set a flag that the current direction is invalid
        Select Case MyDir
        Case KLeft
            Able(1) = True
        Case KUp
            Able(2) = True
        Case KRight
            Able(3) = True
        Case KDown
            Able(4) = True
        End Select
        
        MyDir = TryDirection(Able(1), Able(2), Able(3), Able(4)) '(MyDir + 1 Mod KDown) + KLeft '
    Else
        'exit for - you have found a valid direction
        Exit For
    End If
Next Counter

'move in the found direction
TempDirection = MyDir
End Sub

Public Function DirPossible(TheDir As Byte) As Boolean
'This will return whether or not it is possible for the snake to move
'in the given direction

Dim Counter As Integer
Dim CurPos As Point
Dim HitSnake As Boolean

'project the next position
'get the new co-ordinates for new position
CurPos.X = Snake(0).Pos(X)
CurPos.Y = Snake(0).Pos(Y)
Select Case TheDir 'Snake(0).Direction
Case KLeft
    CurPos.X = CurPos.X - 1
Case KUp
    CurPos.Y = CurPos.Y - 1
Case KRight
    CurPos.X = CurPos.X + 1
Case KDown
    CurPos.Y = CurPos.Y + 1
End Select

'check to see if the next position is a part of the snake
HitSnake = IsSnake(CurPos.X, CurPos.Y)

'it is possible for the snake to move in the current direction
DirPossible = Not HitSnake
End Function

Public Function TryDirection(Optional Left As Boolean, Optional Up As Boolean, Optional Right As Boolean, Optional Down As Boolean) As Byte
'this will test the specified directions to see if it is possible to
'move the snake anywhere. If it cannot go in the preferred direction,
'select another, always trying to move as close to the food as possible.

Dim Counter As Integer
Dim DistA As Integer
Dim DistB As Integer
Dim TempDir As Byte
Dim Nearist As Integer

'find the nearist food item
Nearist = FindNearistFood

'vertical
If ((Food(Nearist, X) < Snake(0).Pos(X)) And (Not Left)) Then 'Or (Right Or Up Or Down) Then
    'try and move left if able
    TryDirection = KLeft
    Exit Function
End If

If ((Food(Nearist, X) > Snake(0).Pos(X)) And (Not Right)) Then 'Or (Left Or Up Or Down) Then
    'move right
    TryDirection = KRight
    Exit Function
End If

'horizontal'
If ((Food(Nearist, Y) < Snake(0).Pos(Y)) And (Not Up)) Then 'Or (Left Or Right Or Down) Then
    'try moving up
    TryDirection = KUp
    Exit Function
End If

If ((Food(Nearist, Y) > Snake(0).Pos(Y)) And (Not Down)) Then 'Or (Left Or Up Or Right) Then
    'try moving down
    TryDirection = KDown
    Exit Function
End If

'if it cannot move in the direction it wants to go, change direction

TryDirection = Snake(0).Direction

'do not change direction, if the current direction is valid
If Not IsCollision(TryDirection) Then
    Exit Function
End If

'rotate direction
For Counter = 1 To 4
    'change direction
    TryDirection = ((TryDirection + 1 - KLeft) Mod 4) + KLeft
    
    'if the snake is allowed to move in it's current direction, then
    Select Case TryDirection
    Case KLeft And (Not Left)
        Exit For
    Case KUp And (Not Up)
        Exit For
    Case KRight And (Not Right)
        Exit For
    Case KDown And (Not Down)
        Exit For
    End Select
Next Counter

Exit Function

'***********
'The following code is faulty - do not use
'***********

'if further along in the current direction, the snake head will
'encounter another part of the snake, then reverse direction.
'If the snake head will still encounter another part of the snake,
'then pick the direction with the furthest distance between the
'snake head and the snake body.

'start searching in the current direction
TempDir = TryDirection
DistA = SnakeAhead(Snake(0).Pos(X), Snake(0).Pos(Y), TempDir)
'reverse direction and check it
Temp = ((TempDir + 2 - KLeft) Mod 4) + KLeft
DistB = SnakeAhead(Snake(0).Pos(X), Snake(0).Pos(Y), TempDir)

'go in which ever direction is has the no snake ahead of it
Select Case 0
Case DistA
    Exit Function
Case DistB
    TryDirection = TempDir
    Exit Function
End Select

'if there is a part of the snake in both directions, then choose the one
'with the largest gap.
If DistB > DistA Then
    TryDirection = TempDir
End If
End Function

Public Function IsSnake(ByVal MyX As Integer, ByVal MyY As Integer) As Boolean
'This function returns True is the specified co-ordinates are part
'of the snake.

Dim Counter As Integer

For Counter = 0 To SnakeSize
    'if the co-ordinates are part of the snake then return True and exit
    If (MyX = Snake(Counter).Pos(X)) And (MyY = Snake(Counter).Pos(Y)) Then
        'co-ordinate match part of the snake
        IsSnake = True
        Exit Function
    End If
Next Counter
End Function

Public Function SnakeAhead(MyX As Integer, MyY As Integer, Direction As Byte) As Integer
'This will return the first position of the snake it finds from the
'current point, in the current direction. eg, If the snake body is
'3 units away from the current point, in the current direction, then
'the function will return 3. If the snake body is not found then the
'function returns 0.

Dim Counter As Integer
Dim Start As Integer
Dim Finish As Integer
Dim GotPoint As Integer
Dim Forwards As Boolean
Dim NewPos(2) As Integer

'set the starting and finishing points
Select Case Direction
Case KLeft
    'starting point is the Y axis
    Start = MyY
    'finishing point is the left wall
    Finish = 0
Case KUp
    'starting point is the X axis
    Start = MyX
    'finishing point is the top wall
    Finish = 0
Case KRight
    'starting point is the Y axis
    Start = MyY
    'finishing point is the right wall
    Finish = RoomSize
Case KDown
    'starting point is the X axis
    Start = MyX
    'finishing point is the bottom wall
    Finish = RoomSize
Case Else
    'invalid direction
    Exit Function
End Select


If Finish < Start Then
    'search backwards
    Call FlipVal(Start, Finish)
    Forwards = False
Else
    'search normally
    Forwards = True
End If

'set the starting points
NewPos(X) = MyX
NewPos(Y) = MyY

For Counter = Start To Finish
    If (Direction Mod 2) = 0 Then
        'direction is verticle
        NewPos(X) = NewPos(X) + 1
    Else
        'direction is horizontal
        NewPos(Y) = NewPos(Y) + 1
    End If
    
    'if searching forwards, then stop on first position found
    If Forwards Then
        'stop at first found
        If IsSnake(NewPos(X), NewPos(Y)) Then
            GotPoint = Counter
            Exit For
        End If
    Else
        'stop at last
        If IsSnake(NewPos(X), NewPos(Y)) Then
            GotPoint = Counter
        End If
    End If
Next Counter

'return value
SnakeAhead = GotPoint
End Function

Public Sub FlipVal(ByRef Val1 As Integer, ByRef Val2 As Integer)
'This procedure will swap the values passed to it.

Dim Temp As Integer

Temp = Val1
Val1 = Val2
Val2 = Temp
End Sub

Public Function FindNearistFood() As Integer
'This returns the food item number of the nearist food item
'to the snake head. It calculates the hypothenuse to the snake head
'from X and Y.

Dim Hyp() As Single
Dim Counter As Integer
Dim Shortest As Single
Dim Index As Integer

'resize array to FoodAmount
ReDim Hyp(FoodAmount)

Shortest = RoomSize * 2
For Counter = 1 To FoodAmount
    '(hyp ^ 2) = (X ^ 2) + (Y ^ 2)
    'where X and Y are in relation to the position of the snake head
    'and the hypothenuse is always positive
    
    Hyp(Counter) = Sqr(Abs((Food(Counter, X) - Snake(0).Pos(X)) ^ 2) + Abs((Food(Counter, Y) - Snake(0).Pos(Y)) ^ 2))
    
    'record shortest distance
    If Hyp(Counter) < Shortest Then
        Shortest = Hyp(Counter)
        Index = Counter
    End If
Next Counter

'return the index of the food item closest to the snake head
FindNearistFood = Index
End Function

Public Sub SetDifficulty()
'This procedure is used to set the different difficulty levels in the
'game. These are the default values.

'food amount
Const EasyFoodAmount = 8
Const NormalFoodAmount = 4
Const HardFoodAmount = 2

'room size (in blocks)
Const EasyRoomSize = 50
Const NormalRoomSize = 50
Const HardRoomSize = 40

'the starting speed
Const EasySSpeed = 200
Const NormalSSpeed = 150
Const HardSSpeed = 100

'maximum speed
Const EasyMSpeed = 50
Const NormalMSpeed = 40
Const HardMSpeed = 30

Dim Counter As Byte

For Counter = Easy To Hard
    Select Case Counter
    Case Easy
        Difficulty(Counter).FoodAmount = EasyFoodAmount
        Difficulty(Counter).MaxSpeed = EasyMSpeed
        Difficulty(Counter).RoomSize = EasyRoomSize
        Difficulty(Counter).StartingSpeed = EasySSpeed
    Case Normal
        Difficulty(Counter).FoodAmount = NormalFoodAmount
        Difficulty(Counter).MaxSpeed = NormalMSpeed
        Difficulty(Counter).RoomSize = NormalRoomSize
        Difficulty(Counter).StartingSpeed = NormalSSpeed
    Case Hard
        Difficulty(Counter).FoodAmount = HardFoodAmount
        Difficulty(Counter).MaxSpeed = HardMSpeed
        Difficulty(Counter).RoomSize = HardRoomSize
        Difficulty(Counter).StartingSpeed = HardSSpeed
    End Select
Next Counter
End Sub

Public Sub LoadScores()
'This procedure loads the high scores from the .ini file into the
'labels.

Dim Counter As Integer
Dim ErrNum As Long
Dim FileNum As Integer
Dim FilePath As String
Dim FileLine As String

'set the difficulty settings into the array
Call SetDifficulty

'get .ini path
FilePath = AddFile(App.Path, ScoresFile)

'check for errors
On Error Resume Next

FileNum = FreeFile
Open FilePath For Input As #FileNum
    ErrNum = Err

'if error found, exit
If ErrNum <> 0 Then
    'exit procedure
    Close #FileNum
    On Error GoTo 0
    
    'if the file does not exist, then create it
    'Error #53 = "File Not Found"
    If ErrNum = 53 Then
        Call SaveScores
    End If
    
    Exit Sub
End If

'continue loading file
Counter = 0
While Not EOF(FileNum)
    'load each score item
    Line Input #FileNum, FileLine
    
    Select Case LCase(GetBefore(FileLine))
    Case "player"
        'a new players score
        Counter = Counter + 1
    Case "name"
        Scores(Counter - 1).Name = GetAfter(FileLine)
    Case "score"
        Scores(Counter - 1).Score = Val(GetAfter(FileLine))
    Case "aicheats"
        AICheat = GetAfter(FileLine)
    Case "aihiscore"
        AIScore = GetAfter(FileLine)
    Case "aispeed"
        AISpeed = Val(GetAfter(FileLine))
    Case "difficulty"
        DifficultyLevel = Val(GetAfter(FileLine))
        'Call ActivateSettings
    Case "roomsize"
        RoomSize = Val(GetAfter(FileLine))
        Call frmSnake.ResizeGameArea
    Case "foodamount"
        FoodAmount = Val(GetAfter(FileLine))
    Case "playerstartspeed"
        StartSpeed = Val(GetAfter(FileLine))
    Case "playertopspeed"
        Val (GetAfter(FileLine))
    End Select
Wend

'file loaded, close file access
Close #FileNum

'resume normal error handling
On Error GoTo 0

'if the game variables were zero, then default to the difficulty level
If (RoomSize = 0) Or (FoodAmount = 0) Or ((playerstartspeed = 0) And (playertopspeed = 0)) Then
    Call SetDifficulty
End If

'if the form is visible, then show the scores
If frmSHighScores.Visible Then
    Call frmSHighScores.ShowScores
End If
End Sub

Public Sub SaveScores()
'This will save the score to the .ini file from the labels

Dim Counter As Integer
Dim ErrNum As Long
Dim FileNum As Integer
Dim FilePath As String

'check for errors
On Error Resume Next
FileNum = FreeFile
FilePath = AddFile(App.Path, ScoresFile)
Open FilePath For Output As #FileNum
    'record error if one occured
    ErrNum = Err
    
    If ErrNum <> 0 Then
        'exit procedure
        Close #FileNum
        On Error GoTo 0
        Exit Sub
    End If
    
    'use marker for score details
    Print #FileNum, App.ProductName; "  v" & App.Major & "." & App.Minor & "." & App.Revision
    Print #FileNum, "Author : Eric O'Sullivan"
    Print #FileNum, "Email : DiskJunky@hotmail.com"
    Print #FileNum, ""
    Print #FileNum, ""
    Print #FileNum, "[HI SCORES]"

    For Counter = 1 To 10
        'write file
        Print #FileNum, "Player=" & Counter
        Print #FileNum, "Name=" & Scores(Counter - 1).Name 'GetName(lblScore(Counter - 1).Caption)
        Print #FileNum, "Score=" & Scores(Counter - 1).Score 'LTrim(GetScore(lblScore(Counter - 1).Caption))
        Print #FileNum, ""
    Next Counter
    
    Print #FileNum, ""
    Print #FileNum, "[ARTIFICIAL INTELLIGENCE]"
    Print #FileNum, "AICheats=" & AICheat
    Print #FileNum, "AIHiScore=" & AIScore
    Print #FileNum, "AISpeed=" & AISpeed
    Print #FileNum, ""
    Print #FileNum, "[GAME SETTINGS]"
    Print #FileNum, "Difficulty=" & DifficultyLevel
    Print #FileNum, "RoomSize=" & RoomSize
    Print #FileNum, "FoodAmount=" & FoodAmount
    Print #FileNum, "PlayerStartSpeed=" & StartSpeed
    Print #FileNum, "PlayerTopSpeed=" & TopSpeed
Close #FileNum

'resume normal error hondling
On Error GoTo 0
End Sub

Public Function GetBefore(Sentence As String) As String
'This procedure returns all the character of a
'string before the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Before As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If (Counter <> Len(Sentence)) And (Counter <> 0) Then
    Before = Left(Sentence, (Counter - 1))
Else
    Before = ""
End If

GetBefore = Before
End Function


Public Function GetAfter(Sentence As String) As String
'This procedure returns all the character of a
'string after the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Rest As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If Counter <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - Counter))
Else
    Rest = ""
End If

GetAfter = Rest
End Function

Public Function GetPath(Address As String) As String
'This function returns the path from a string containing the full
'path and filename of a file.

Dim Counter As Integer
Dim LastPos As Integer

'find the position of the last "\" mark in the string
LastPos = 1
For Counter = 1 To Len(Address)
    If Mid(Address, Counter, 1) = "\" Then
        LastPos = Counter
    End If
Next Counter

'return everything before the last "\" mark
GetPath = Left(Address, (LastPos - 1))
End Function

Public Function AddFile(Path As String, File As String) As String
'This procedure adds a file name to a path.

If Right(Path, 2) = ":\" Then
    Path = Path & File
Else
    Path = Path & "\" & File
End If

AddFile = Path
End Function

Public Sub EnterScore(ByVal Name As String, ByVal TheScore As Integer)
'This will check to see if the score is able to go into the high score
'list and enter it in the appropiate place

Dim Counter As Integer
Dim ScorePos As Byte

'check the last score to see if able to enter
If Scores(9).Score > TheScore Then
    'exit, score not high enough
    Exit Sub
End If

'find the appropiate place for score
ScorePos = 1
For Counter = 10 To 1 Step -1
    If Scores(Counter - 1).Score > TheScore Then
        'found place
        ScorePos = Counter + 1
        Exit For
    End If
Next Counter

'move all score below the score point down one and update ranking
For Counter = 9 To ScorePos Step -1
    Scores(Counter) = Scores(Counter - 1)
Next Counter

'enter the score
Scores(ScorePos - 1).Name = Name
Scores(ScorePos - 1).Score = TheScore

'if the form is visible, then update the scores
If frmSHighScores.Visible Then
    Call frmSHighScores.ShowScores
End If

Call SaveScores
End Sub

Public Sub ActivateSettings()
'This will set the game settings according to the current difficulty
'level

FoodAmount = Difficulty(DifficultyLevel).FoodAmount
RoomSize = Difficulty(DifficultyLevel).RoomSize
StartSpeed = Difficulty(DifficultyLevel).StartingSpeed
TopSpeed = Difficulty(DifficultyLevel).MaxSpeed

'resize the form to new play area and start new game
Call frmSnake.ResizeGameArea
Call frmSnake.GameOver
Call frmSnake.NewGame
End Sub

