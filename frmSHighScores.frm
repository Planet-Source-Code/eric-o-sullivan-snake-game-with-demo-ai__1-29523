VERSION 5.00
Begin VB.Form frmSHighScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake High Scores"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "frmSHighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   2040
      TabIndex        =   31
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
   Begin VB.Line lnBreak 
      X1              =   120
      X2              =   5760
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   720
      TabIndex        =   30
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   720
      TabIndex        =   29
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   28
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   720
      TabIndex        =   27
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   26
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   25
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   24
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   23
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   22
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Snake A.I."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   21
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   0
      TabIndex        =   20
      Top             =   3480
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   0
      TabIndex        =   19
      Top             =   3120
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   0
      TabIndex        =   17
      Top             =   2400
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   0
      TabIndex        =   15
      Top             =   1680
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   655
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1 -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   655
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   4800
      TabIndex        =   10
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4800
      TabIndex        =   9
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   8
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   5
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   3
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
'exit the form
frmSHighScores.Visible = False
End Sub

Public Sub ShowScores()
'This will update the display on the screen

Dim Counter As Byte

'display each name and score in the corresponding label
For Counter = 0 To 9
    lblName(Counter).Caption = Scores(Counter).Name
    lblScore(Counter).Caption = Scores(Counter).Score
Next Counter
End Sub

Private Sub cmdReset_Click()
'confirm reset. if Yes, then erase all the scores

Dim Result As Integer
Dim Counter As Byte

'confirm reset
Result = MsgBox("This will erase all scores data." & vbCrLf & "Are you sure?", vbYesNo + vbExclamation, "Erase Scores ")
If Result = vbYes Then
    'erase all the scores
    For Counter = 0 To 9
        Scores(Counter).Name = "No Name"
        Scores(Counter).Score = 0
    Next Counter
    
    Call ShowScores
    Call SaveScores
End If
End Sub

Private Sub Form_Load()
Call ShowScores
End Sub
