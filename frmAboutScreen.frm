VERSION 5.00
Begin VB.Form frmAboutScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAboutScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timText 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
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
      Left            =   1740
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Line lnSpacer 
      X1              =   120
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmAboutScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This screen was first created on the 17/11/2001 and was intended
'for use in several future programs. The idea was that I should only
'have to create this screen once and be able to integrate it into any
'other project seemlessly. I wanted to do this instead of creating a
'new about screen for every project where I wanted one.
'
'A note on this About Screen :
'This screen requires the module APIGraphics (APIGraphics.bas) to
'operate the display.
'
'Eric O'Sullivan
'email DiskJunky@hotmail.com
'============================================================

Dim AllText As String
Dim Start As Boolean

Private Sub cmdOk_Click()
'exit screen
Unload Me
End Sub

Private Sub Form_Load()
Call SetText
End Sub

Private Sub timText_Timer()
'This timer will scroll the animated text

Const Wait = 50 'wait 15 ticks before drawing the next frame

Dim Font As FontStruc
Dim Bmp As BitmapStruc
Dim Mask As BitmapStruc
Dim BmpSize As Rect
Dim Result As Integer
Dim TextHeight As Integer
Dim StartingTick As Long

Static Surphase As BitmapStruc
Static Scroll As Integer

'find out how much time it takes to draw a frame
StartingTick = GetTickCount

'set the bitmap dimensions and create them
BmpSize.Right = picText.ScaleWidth
BmpSize.Bottom = picText.ScaleHeight

Call RectToPixels(BmpSize)

Mask.Area = BmpSize
Surphase.Area = BmpSize
Bmp.Area = BmpSize

'set font variables
Font.Alignment = vbCentreAlign
Font.Name = picText.FontName
Font.Bold = picText.FontBold
Font.Colour = vbWhite 'picText.ForeColor
Font.Italic = picText.FontItalic
Font.StrikeThru = picText.FontStrikethru
Font.PointSize = picText.FontSize
Font.Underline = picText.FontUnderline

'test code - not currently used
'Call MakeText(picText.hDc, "Hello World!", 0, 0, 40, 180, Font, InPixels)

TextHeight = GetTextHeight(picText.hDc) * LineCount(AllText)

Scroll = Scroll - Screen.TwipsPerPixelY
If (Scroll < -(TextHeight * Screen.TwipsPerPixelY)) Or (Not Start) Then
    Scroll = picText.ScaleHeight '+ (TextHeight * Screen.TwipsPerPixelY)
    Start = True
End If

'only create the surphase if necessary
If Surphase.hDcMemory = 0 Then
    Call CreateNewBitmap(Surphase.hDcMemory, Surphase.hDcBitmap, Surphase.hDcPointer, Surphase.Area, frmAboutScreen, picText.ForeColor, InPixels)
    
    'create the surphase
    'text fade in
    Call Gradient(Surphase.hDcMemory, picText.ForeColor, picText.FillColor, 0, (Surphase.Area.Bottom - ((TextHeight / LineCount(AllText)) * 2)), Surphase.Area.Right, (TextHeight / LineCount(AllText) * 2), GradHorizontal, InPixels)
    'text fade out
    Call Gradient(Surphase.hDcMemory, picText.FillColor, picText.ForeColor, 0, 0, Surphase.Area.Right, (TextHeight / LineCount(AllText)) * 2, GradHorizontal, InPixels)
End If
Call CreateNewBitmap(Mask.hDcMemory, Mask.hDcBitmap, Mask.hDcPointer, Mask.Area, frmAboutScreen, vbBlack, InPixels)
Call CreateNewBitmap(Bmp.hDcMemory, Bmp.hDcBitmap, Bmp.hDcPointer, Bmp.Area, frmAboutScreen, vbWhite, InPixels)

'draw the text onto the mask in black
Call MakeText(Mask.hDcMemory, AllText, (Scroll / Screen.TwipsPerPixelY), 0, TextHeight, Bmp.Area.Right, Font, InPixels)

'copy the surphase onto the background
Result = BitBlt(Bmp.hDcMemory, 0, 0, Bmp.Area.Right, Bmp.Area.Bottom, Surphase.hDcMemory, 0, 0, SRCCOPY)

'place the mask onto the background
Result = BitBlt(Bmp.hDcMemory, 0, 0, Bmp.Area.Right, Bmp.Area.Bottom, Mask.hDcMemory, 0, 0, SRCAND)

'copy the result to the screen
Result = BitBlt(frmAboutScreen.hDc, 0, 0, Bmp.Area.Right, Bmp.Area.Bottom, Bmp.hDcMemory, 0, 0, SRCCOPY)

'remove the bitmaps created
Call DeleteBitmap(Bmp.hDcMemory, Bmp.hDcBitmap, Bmp.hDcPointer)
Call DeleteBitmap(Mask.hDcMemory, Mask.hDcBitmap, Mask.hDcPointer)

'wait X ticks minus the time it took to draw the frame
Call Pause(Wait - (GetTickCount - StartingTick))
End Sub

Private Sub SetText()
'This procedure is used to setting the text displayed in the picture box

'" & vbCrLf & "

'please note that ProductName can be set by going to
'Project, Project Properties,Make tab. You should see a list box about
'half way down on the left side. Scroll down until you come to
'Product Name and enter some text into the text box on the right
'side of the list box.
AllText = App.ProductName & vbCrLf & "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "" & vbCrLf & "This program was made by" & vbCrLf & "Eric O'Sullivan." & vbCrLf & "" & vbCrLf & "Copyright 2001" & vbCrLf & "All rights reserved" & vbCrLf & "" & vbCrLf & "For more information, email" & vbCrLf & "DiskJunky@hotmail.com"
End Sub

Public Function LineCount(Text As String) As Integer
'This function will return the number of lines in the text

Dim Temp As Integer
Dim Counter As Integer
Dim LastPos As Integer

LastPos = 1

Do
    Temp = LastPos
    LastPos = InStr(LastPos + Len(vbCrLf), Text, vbCrLf)
    
    If Temp <> LastPos Then
        'a line was found
        Counter = Counter + 1
    End If
Loop Until LastPos = 0 'LastPos will =0 when InStr cannot find any more occurances of vbCrlf

LineCount = Counter
End Function
