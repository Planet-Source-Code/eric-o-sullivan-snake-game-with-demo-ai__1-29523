VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artificial Intelligence"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmAI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame fraCheat 
      Caption         =   "Cheat"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   1215
      Width           =   5055
      Begin VB.CheckBox chkScore 
         Caption         =   "AI In Hi-Score List"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkCheat 
         Caption         =   "AI Cheats"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraSpeed 
      Caption         =   "AI Speed Increase"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin ComctlLib.Slider sldSpeed 
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         _Version        =   327682
         Max             =   50
         SelStart        =   5
         TickFrequency   =   5
         Value           =   5
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.20 Seconds"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label lblFast 
         BackStyle       =   0  'Transparent
         Caption         =   "Fast"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblSlow 
         BackStyle       =   0  'Transparent
         Caption         =   "Slow"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SettingValues As Boolean

Private Sub chkCheat_Click()
'If the AI is allowed to cheat, then check the box, else uncheck the
'box.

If SettingValues Then
    Exit Sub
End If

AICheat = Not AICheat

'don't let the ai enter it's score
If AICheat Then
    AIScore = False
    chkScore.Value = 0
End If

chkScore.Enabled = Not AICheat
End Sub

Private Sub chkScore_Click()
'this sets whether or not the ai can save its' score in the hi-score
'list.

If SettingValues Or AICheat Then
    Exit Sub
End If

AIScore = Not AIScore
End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetValues
End Sub

Private Sub SetValues()
'This procedure will set all values on the controls to the most current
'settings.

'set a flag that the controls are not to exicute code during this
'procedure.
SettingValues = True

'set slider
sldSpeed.Max = StartSpeed
sldSpeed.Min = MaxAISpeed
sldSpeed.Value = AISpeed
sldSpeed.TickFrequency = (PointsPerFood * 2)
sldSpeed.SmallChange = (PointsPerFood * 2)
sldSpeed.LargeChange = (PointsPerFood * 2) * 5
Call SliderDisplay

'check the box
If AICheat Then
    'check the box
    chkCheat.Value = 1
Else
    chkCheat.Value = 0
End If

'hi scores
If AICheat Then
    'if the ai is cheating then, disable control
    chkScore.Value = 0
    chkScore.Enabled = False
    AIScore = False
Else
    'ai is not cheating, but check the box if necessary
    chkScore.Enabled = True
    If AIScore Then
        'check the box
        chkScore.Value = 1
    Else
        chkScore.Value = 0
    End If
End If

'let the controls exicute code
SettingValues = False
End Sub

Private Sub SliderDisplay()
'This will set the caption of the label displaying the speed of the ai
'to reflect an accurate time

AISpeed = sldSpeed.Value
lblSpeed.Caption = Format(((StartSpeed - AISpeed) / 1000), "0.000") & " Seconds"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'save settings
Call SaveScores
End Sub

Private Sub sldSpeed_Change()
'update speed

AISpeed = sldSpeed.Value
Call SliderDisplay
End Sub

Private Sub sldSpeed_Scroll()
Call sldSpeed_Change
End Sub
