VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSnakeSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Settings"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmSnakeSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   6240
      Width           =   615
   End
   Begin VB.Frame framSettings 
      Caption         =   "Settings"
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
      Begin ComctlLib.Slider sldRoom 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   327682
         Min             =   15
         Max             =   65
         SelStart        =   15
         TickFrequency   =   5
         Value           =   15
      End
      Begin ComctlLib.Slider sldMSpeed 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   327682
         Min             =   10
         Max             =   250
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin ComctlLib.Slider sldSSpeed 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   327682
         Min             =   10
         Max             =   250
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin ComctlLib.Slider sldFood 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   327682
         LargeChange     =   3
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblRoomMax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "65"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   4560
         Width           =   375
      End
      Begin VB.Label lblRoomMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   375
      End
      Begin VB.Label lblSpeedMax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "250"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   18
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label lblSpeedMax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "250"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   17
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblSpeedMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label lblSpeedMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblRoomSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Size (blocks)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label lblMaxSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Speed Delay (milliseconds)"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label lblInitialSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Speed Delay (milliseconds)"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblMaxFood 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblMinFood 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblFoodAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Food Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame framDiff 
      Caption         =   "Difficulty"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin ComctlLib.Slider sldDiff 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   873
         _Version        =   327682
         Min             =   1
         Max             =   3
         SelStart        =   2
         TickStyle       =   1
         Value           =   2
      End
      Begin VB.Label lblHard 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hard"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblNormal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   195
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblEasy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Easy"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmSnakeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SetSliders(ByVal Level As Integer)
'This will set the sliders according to the difficulty
'level.

sldFood.Value = Difficulty(Level).FoodAmount
sldRoom.Value = Difficulty(Level).RoomSize
sldSSpeed.Value = Difficulty(Level).StartingSpeed
sldMSpeed.Value = Difficulty(Level).MaxSpeed

Call UpdateValues
End Sub

Public Sub UpdateValues()
'This procedure will set the internal variables that affect the
'game, like FoodAmount, RoomSize etc.

FoodAmount = sldFood.Value
RoomSize = sldRoom.Value
StartingSpeed = sldSSpeed.Value
TopSpeed = sldMSpeed.Value

Call ActivateSettings
End Sub

Private Sub SetInitialValues()
'This will set the maximum and minimum values of the sliders and
'the corresponding labels

Dim Counter As Byte

'set food range values
sldFood.Max = MaxFoodAmount
sldFood.Min = MinFoodAmount
sldFood.Value = FoodAmount
lblMinFood.Caption = MinFoodAmount
lblMaxFood.Caption = MaxFoodAmount

'set room size range values
sldRoom.Max = MaxRoomSize
sldRoom.Min = MinRoomSize
sldRoom.Value = RoomSize
lblRoomMin.Caption = MinRoomSize
lblRoomMax.Caption = MaxRoomSize

'set starting and finishing speed range values
sldSSpeed.Max = MinSpeed
sldMSpeed.Max = MinSpeed
sldSSpeed.Min = MaxSpeed
sldMSpeed.Min = MaxSpeed
sldSSpeed.Value = StartSpeed
sldMSpeed.Value = TopSpeed
For Counter = 0 To 1
    lblSpeedMax(Counter).Caption = MinSpeed
    lblSpeedMin(Counter).Caption = MaxSpeed
Next Counter
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetDifficulty
Call SetInitialValues
Call SetSliders(DifficultyLevel)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SaveScores
Unload Me
End Sub

Private Sub sldDiff_Change()
'set the slider values for each difficulty setting
DifficultyLevel = sldDiff.Value
Call SetSliders(DifficultyLevel)
End Sub

Private Sub sldFood_Change()
'save food amount
FoodAmount = sldFood.Value
End Sub

Private Sub sldMSpeed_Change()
'save the players top speed
TopSpeed = sldMSpeed.Value
End Sub

Private Sub sldRoom_Change()
'save the game room size in blocks
RoomSize = sldRoom.Value
End Sub

Private Sub sldSSpeed_Change()
'set the starting player speed
StartSpeed = sldSSpeed.Value
End Sub
