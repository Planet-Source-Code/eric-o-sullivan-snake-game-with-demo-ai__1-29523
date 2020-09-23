VERSION 5.00
Begin VB.Form frmSnake 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5295
   FillColor       =   &H00800000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmSnake.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timAI 
      Interval        =   1
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer timDirection 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer timPlay 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblShowSize 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Snake Size :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblShowSpeed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Speed Increase :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblShowScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   5520
      Width           =   615
   End
   Begin VB.Line lnBreak 
      X1              =   120
      X2              =   5160
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGamePlay 
         Caption         =   "&Play"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameAI 
         Caption         =   "&A.I."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameStart 
         Caption         =   "&Start A.I."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGameCheat 
         Caption         =   "A.I. Cheats"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGameView 
         Caption         =   "&View High Scores"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGameSettings 
         Caption         =   "&Game Settings..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGameBreakExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Blocks((MaxRoomSize ^ 2) / 2) As Point 'the maximum number of rectangles in which to draw the snake
Dim BlockNum As Integer 'the number of times the snake turns

'a point on the grid
Private Type Point
    X As Integer
    Y As Integer
End Type

Dim BackBmp As BitmapStruc

Public Sub GameOver()
'Stop current human or AI game.

Dim PlayerName As String

timPlay.Enabled = False

If timAI.Enabled Then
    'start a new AI game after one second
    If (AIScore) And (Not AICheat) Then
        'let the ai enter it's score
        Call EnterScore("Snake A.I.", Score)
    End If
    Call Pause(1000)
    Call NewGame
Else
    'see if the user can put their name in the hi score list
    If Score >= Scores(9).Score Then 'frmSHighScores.GetScore(frmSHighScores.lblScore(9).Caption) Then
        'get the users' name and enter to the high score list
        PlayerName = InputBox("Congratulations! Please enter your name", "New High Score!", "Player1")
        Call EnterScore(PlayerName, Score)
    End If
    
    Call Pause(1000) 'stop execution for 1 second
    
    'play demo
    timAI.Enabled = True
End If
End Sub

Public Sub NewGame()
'Starts a new game

Dim Counter As Integer

'start counting how long it takes to create a frame
StartFrameTick = GetTickCount

Score = 0
lblScore.Caption = 0
Speed = StartSpeed

'erase any data stored in food array
For Counter = 1 To FoodAmount
    Food(Counter, X) = 0
    Food(Counter, Y) = 0
Next Counter
FoodAmount = 0
GrowAmount = 0

'erase any data about the snake
For Counter = 0 To SnakeSize
    Snake(Counter).Direction = 0
    Snake(Counter).Pos(X) = 0
    Snake(Counter).Pos(X) = 0
Next Counter
SnakeSize = 0

FoodPerTick = Difficulty(DifficultyLevel).FoodAmount

'create new values for initial conditions
Call ResetSnake

'set the default direction
TempDirection = KLeft

'play the game unless the AI is playing
If Not timAI.Enabled Then
    timPlay.Enabled = True
End If

Call UpdateDisplay
End Sub

Private Sub cmdPlay_Click()
timAI.Enabled = False
timDirection.Enabled = True
Call NewGame
End Sub

Private Sub Form_Load()
'load the game settings first
Call LoadScores
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'exit program on all unload queries
End
End Sub

Private Sub mnuGameAI_Click()
'show ai settings
Load frmAI
frmAI.Show
End Sub

Private Sub mnuGameCheat_Click()
'set/unset whether or no the ai will cheat
AICheat = Not AICheat
mnuGameCheat.Checked = AICheat
End Sub

Private Sub mnuGameExit_Click()
'end the game
End
End Sub

Private Sub mnuGamePlay_Click()
'let the human player, play
Call cmdPlay_Click
End Sub

Private Sub mnuGameSettings_Click()
'show the game settings
Load frmSnakeSettings
frmSnakeSettings.Visible = True
End Sub

Private Sub mnuGameStart_Click()
'enable AI if not already enabled
If Not timAI.Enabled Then
    timAI.Enabled = True
End If
Call NewGame
End Sub

Private Sub mnuGameView_Click()
'show the high scores
Load frmSHighScores
frmSHighScores.Visible = True
End Sub

Private Sub mnuHelpAbout_Click()
'view the about screen
Load frmAboutScreen
frmAboutScreen.Show
End Sub

Private Sub timAI_Timer()
'This timer will activate the AI. The AI will play a game of snake

'start a new game with the AI
If timDirection.Enabled Then
    timDirection.Enabled = False
    Call NewGame
End If

'find where to go
Call FindRoute

Call MoveSnake

Call UpdateDisplay
End Sub

Private Sub timDirection_Timer()

'if the program does not have focus or the AI is active,
'then do not get the direction to move the snake in.
If AmIActive(frmSnake) Then
    'find this move's direction
    Call GetDirection
End If
End Sub

Private Sub timPlay_Timer()
'move the snake
Call MoveSnake

'update the scores and stuff
Call UpdateDisplay
End Sub

Public Sub DrawFrame()
'This will draw the entire game area onto an off-screen bitmap
'before blitting to the screen. All direct graphical data is used
'from this procedure.

Dim Bmp As BitmapStruc
Dim Counter As Integer
Dim Result As Long
Dim Pos(2) As Point
Dim Font As FontStruc

DoEvents

'set the bitmap dimensions (in pixels)
Bmp.Area.Right = (GridSize * RoomSize) + 1
Bmp.Area.Bottom = (GridSize * RoomSize) + 1

'create the bitmap in memory
Call CreateNewBitmap(Bmp.hDcMemory, Bmp.hDcBitmap, Bmp.hDcPointer, Bmp.Area, frmSnake, frmSnake.FillColor, InPixels)

'if the background size does not match new settings, create a new
'background bitmap after deleting the old one
If (BackBmp.Area.Right <> Bmp.Area.Right) Or (BackBmp.Area.Bottom <> Bmp.Area.Bottom) Then
    'delete the old bitmap and create a new one
    Call DeleteBitmap(BackBmp.hDcMemory, BackBmp.hDcBitmap, BackBmp.hDcPointer)
    
    BackBmp.Area = Bmp.Area
    
    Call CreateNewBitmap(BackBmp.hDcMemory, BackBmp.hDcBitmap, BackBmp.hDcPointer, BackBmp.Area, frmSnake, frmSnake.FillColor, InPixels)
    
    'draw the grid lines onto the bitmap
    For Counter = 0 To (GridSize * RoomSize) + 2 Step GridSize
        'horizontal lines
        Call DrawLine(BackBmp.hDcMemory, 0, Counter, BackBmp.Area.Right, Counter, frmSnake.ForeColor, 1, InPixels)
        
        'vertical lines
        Call DrawLine(BackBmp.hDcMemory, Counter, 0, Counter, BackBmp.Area.Bottom, frmSnake.ForeColor, 1, InPixels)
    Next Counter
End If

'copy the background onto the new bitmap
Result = BitBlt(Bmp.hDcMemory, 0, 0, Bmp.Area.Right, Bmp.Area.Bottom, BackBmp.hDcMemory, 0, 0, SRCCOPY)

'if ai is active, then display message
If timAI.Enabled Then
    'set font from forms font settings
    Font.Alignment = vbCentreAlign
    Font.Name = frmSnake.FontName
    Font.Bold = frmSnake.FontBold
    Font.Colour = vbYellow 'frmSnake.ForeColor
    Font.Italic = frmSnake.FontItalic
    Font.StrikeThru = frmSnake.FontStrikethru
    Font.PointSize = frmSnake.FontSize
    Font.Underline = frmSnake.FontUnderline
    
    
    Call MakeText(Bmp.hDcMemory, "Press F2 to play", ((RoomSize * GridSize) - GetTextHeight(frmSnake.hDc)) / 2, 0, GetTextHeight(frmSnake.hDc), (RoomSize * GridSize), Font, InPixels)
End If

'draw the food
For Counter = 1 To FoodAmount
    Pos(1).X = (Food(Counter, X) * GridSize)
    Pos(1).Y = (Food(Counter, Y) * GridSize)
    Call DrawRect(Bmp.hDcMemory, FoodColour, Pos(1).X + 1, Pos(1).Y + 1, (Pos(1).X + 1) + (GridSize - 1), (Pos(1).Y + 1) + (GridSize - 1))
Next Counter

'draw the snake
Call GetSnakeRect
For Counter = 1 To BlockNum - 1
    Pos(1).X = (Blocks(Counter).X * GridSize) + 1
    Pos(1).Y = (Blocks(Counter).Y * GridSize) + 1
    Pos(2).X = (Blocks(Counter + 1).X * GridSize) + GridSize
    Pos(2).Y = (Blocks(Counter + 1).Y * GridSize) + GridSize
    
    'swap values if necessary to draw all elements of the snake
    If Pos(1).X > Pos(2).X Then
        Pos(1).X = Pos(1).X + GridSize - 1
        Pos(2).X = Pos(2).X - GridSize + 1
    End If
    
    If Pos(1).Y > Pos(2).Y Then
        Pos(1).Y = Pos(1).Y + GridSize - 1
        Pos(2).Y = Pos(2).Y - GridSize + 1
    End If
    
    Call DrawRect(Bmp.hDcMemory, SnakeColour, Pos(1).X, Pos(1).Y, Pos(2).X, Pos(2).Y)
Next Counter


'draw the final result
Result = BitBlt(frmSnake.hDc, 0, 0, Bmp.Area.Right, Bmp.Area.Bottom, Bmp.hDcMemory, 0, 0, SRCCOPY)

'delete the bitmap
Call DeleteBitmap(Bmp.hDcMemory, Bmp.hDcBitmap, Bmp.hDcPointer)
End Sub

Private Sub GetSnakeRect()
'This procedure will find out how many times the snake changes
'direction and calculate the minimum amount of rectangles to
'draw the snake with. This saves cpu processing time for the
'graphics and api calls.
'The ultimate aim of this is that, when drawing the snake, there will
'be gridlines all around each element of the snake except where it
'touches the next or previous element of the snake. This means that
'you can see exactly what path the snake has taken.

Dim Counter As Integer
Dim LastDirection As Byte

BlockNum = 0
For Counter = 0 To SnakeSize
    'find out how many times the snake turns
    If (LastDirection <> Snake(Counter).Direction) Or (Counter = SnakeSize) Then
        'the snake turned or tail of snake, record position
        LastDirection = Snake(Counter).Direction
        BlockNum = BlockNum + 1
        Blocks(BlockNum).X = Snake(Counter).Pos(X)
        Blocks(BlockNum).Y = Snake(Counter).Pos(Y)
    End If
Next Counter
End Sub

Private Sub UpdateDisplay()
'update the scores display

If (Score <> Val(lblScore.Caption)) Then
    lblScore.Caption = Score
End If

'Debug.Print (StartSpeed - Speed) / 1000, Val(lblSpeed.Caption), Speed, AISpeed
If Not timAI.Enabled Then
    If ((StartSpeed - Speed) / 1000) <> Val(lblSpeed.Caption) Then
        'the ai moves at a faster speed than the human player. Account
        'for this variance.
        lblSpeed.Caption = Format(((StartSpeed - Speed) / 1000), "0.00") & " Seconds"
    End If
Else
    If ((StartSpeed - AISpeed) / 1000) <> Val(lblSpeed.Caption) Then
        lblSpeed.Caption = Format(((StartSpeed - AISpeed) / 1000), "0.00") & " Seconds"
    End If
End If

If SnakeSize <> Val(lblSize.Caption) Then
    lblSize.Caption = SnakeSize
End If
End Sub

Public Sub ResizeGameArea()
'This will resize the ford according to the current RoomSize

Dim XPixel As Integer

Const LineDist = 120

'set the game area width
Me.Width = (((GridSize * RoomSize) + 1) * Screen.TwipsPerPixelX) + (Me.Width - Me.ScaleWidth)

'adjust the controls accordingly
'----------------

'the line
lnBreak.X1 = LineDist
lnBreak.X2 = Me.ScaleWidth - LineDist
lnBreak.Y1 = Me.ScaleWidth + LineDist
lnBreak.Y2 = lnBreak.Y1

'command button
cmdPlay.Top = lnBreak.Y1 + LineDist

'the labels
XPixel = Screen.TwipsPerPixelX
lblShowScore.Left = (ScaleWidth / 2) - lblShowScore.Width - XPixel
lblScore.Left = (ScaleWidth / 2)
lblScore.Top = cmdPlay.Top
lblShowScore.Top = lblScore.Top

lblShowSpeed.Left = (ScaleWidth / 2) - lblShowSpeed.Width - XPixel
lblSpeed.Left = (ScaleWidth / 2)
lblSpeed.Top = (lblScore.Top + lblScore.Height) + 30
lblShowSpeed.Top = lblSpeed.Top

lblShowSize.Left = (ScaleWidth / 2) - lblShowSize.Width - XPixel
lblSize.Left = (ScaleWidth / 2)
lblSize.Top = (lblSpeed.Top + lblSpeed.Height) + 30
lblShowSize.Top = lblSize.Top
'---------------

'set the forms height, accounting for the title bar and the game display
Me.Height = (lblSize.Top + lblSize.Height + 30) + (Me.Height - Me.ScaleHeight)

'refresh the display
frmSnake.Cls
End Sub
