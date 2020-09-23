VERSION 5.00
Begin VB.Form FrmBB 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Brainballs"
   ClientHeight    =   6000
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5970
   Icon            =   "FrmBB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Scroller 
      Interval        =   100
      Left            =   6180
      Top             =   1320
   End
   Begin VB.PictureBox PicBalls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   75
      Index           =   2
      Left            =   6240
      Picture         =   "FrmBB.frx":08CA
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   4
      Top             =   1020
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2700
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicBalls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   225
      Index           =   1
      Left            =   6300
      Picture         =   "FrmBB.frx":0967
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicBalls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   225
      Index           =   0
      Left            =   6300
      Picture         =   "FrmBB.frx":1019
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox PicFront 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00FFFFFF&
      Height          =   6000
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Menu MnuSet 
      Caption         =   "Settings"
      Begin VB.Menu MnuStyles 
         Caption         =   "Style"
         Begin VB.Menu MnuStyle 
            Caption         =   "Balls"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu MnuStyle 
            Caption         =   "Blocks"
            Index           =   1
         End
         Begin VB.Menu MnuStyle 
            Caption         =   "Bombs"
            Index           =   2
         End
      End
      Begin VB.Menu MnuLvl 
         Caption         =   "Level"
         Begin VB.Menu MnuLevel 
            Caption         =   "Easy"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MnuLevel 
            Caption         =   "Normal"
            Index           =   3
         End
         Begin VB.Menu MnuLevel 
            Caption         =   "Hard"
            Index           =   4
         End
         Begin VB.Menu MnuLevel 
            Caption         =   "Extreme"
            Index           =   5
         End
      End
   End
   Begin VB.Menu MnuReset 
      Caption         =   "Reset"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuEnd 
      Caption         =   "End Game"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FrmBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Brainballs
'© Scythe 2003

'I wrote it after i saw Fosters Cascade on PSC
'His version is nice but a little bit to easy

'btw. this game is completly coded by me
'no routines taken from Foster

'It could be much smaller
'but about 40k (compiled with vb5)is small enough.

'It also should get some gfx improvements but i´m to lazy

Option Explicit

'To clear the Screen or Parts real Fast
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const BLACKNESS = &H42            ' dest = BLACK

'Fast Print
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'Get Scroller Data
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Dim Balls(24, 24) As Long   'The Screen u see
Dim Hits(24, 24) As Long    'What balls are detroyed
Dim HTS As Long             'Number of hits
Dim HitCol As Long          'Color of the clicked ball
Dim Points As Long          'Score
Dim GameStyle As Long       'Gamestyle (0 Balls / 1 Blocks)
Dim GameLevel As Long       'Level
Dim GameRun As Boolean      'Is the game running
Dim BombsTrue As Long       'Game with Bombs

Private Type HiSc           'HighScores
 Score As Long
 User As String
End Type

Dim AppPath As String       'Holds the Applications path
Dim Scores(2, 3, 4) As HiSc 'Holds the Scores

'Needed in most routines for For/Next
Dim i As Long
Dim f As Long
Dim g As Long


'Start the Prog
Private Sub Form_Load()
 'Lets get the App´s path
 'if we saved at root App.Path returens a "\" at the end of the string
 AppPath = App.Path
 If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

 'Set the Gamelevel to EASY
 GameLevel = 2

 'Load HighScores
 LoadScore

 'Show Startscreen
 StartNew
End Sub

'Set all back to start a new game
Private Sub StartNew()

 'No Game no points
 Points = 0
 HTS = 0
 GameRun = False

 'Calculate a new Field
 CalcField

 'Draw HeigScores & Click to start
 BitBlt PicFront.hdc, 150, 150, 100, 35, 0, 0, 0, BLACKNESS
 TextOut PicFront.hdc, 170, 160, "Click to Start", 14
 BitBlt PicFront.hdc, 120, 198, 160, 110, 0, 0, 0, BLACKNESS
 TextOut PicFront.hdc, 130, 205, "Top 5 Scores for this Level", 27
 For i = 0 To 4
  TextOut PicFront.hdc, 130, 225 + i * 15, Str(Scores(GameStyle + BombsTrue, GameLevel - 2, i).Score), Len(Str(Scores(GameStyle + BombsTrue, GameLevel - 2, i).Score))
  TextOut PicFront.hdc, 190, 225 + i * 15, Scores(GameStyle + BombsTrue, GameLevel - 2, i).User, Len(Scores(GameStyle + BombsTrue, GameLevel - 2, i).User)
 Next i

 'How to Play
 BitBlt PicFront.hdc, 60, 313, 277, 80, 0, 0, 0, BLACKNESS
 Select Case GameStyle
 Case 0
  TextOut PicFront.hdc, 70, 320, "Click balls with same clored neighbour to remove them.", 54
  TextOut PicFront.hdc, 95, 335, "More balls at once will give a heigher score", 44
  If BombsTrue = 2 Then
   TextOut PicFront.hdc, 78, 350, "Bombs will explode if you remove all the neighbours", 51
  End If
  TextOut PicFront.hdc, 95, 370, "© Scythe 2003      scythe@scythe-tools.de", 41
 Case 1
  TextOut PicFront.hdc, 65, 320, "Click blocks with same clored neighbour to remove them.", 55 '
  TextOut PicFront.hdc, 95, 335, "More blocks at once will give a heigher score", 45
  TextOut PicFront.hdc, 63, 365, "Original idea by Fosters posted as Cascade 2003 on PSC", 54
 End Select

 'Show it
 PicFront.Refresh

 'Set the Menus
 MnuSet.Visible = True
 MnuEnd.Visible = False
 MnuReset.Visible = False
 Scroller.Enabled = True

End Sub

Private Sub CalcField()
 'Fill the Field with random Balls/Blocks
 Randomize (Timer)


 For i = 0 To 24
  For f = 0 To 24
   'Gamelevel = Number of different Blocks
   Balls(i, f) = Rnd * GameLevel
  Next f
 Next i


 'Set Bombs
 If BombsTrue = 2 Then
  '10 Bombs
  For g = 0 To 9
   i = Rnd * 24
   f = Rnd * 20
   Balls(i, f) = 6 + Rnd * 2
  Next g
 End If

 'Show it
 DrawField
End Sub


Private Sub DrawField()
 Dim x As Long
 Dim y As Long

 'Move thru the field
 For i = 0 To 24
  For f = 0 To 24
   'Real x & y coordinates
   x = i * 16
   y = f * 16
   'Is there a ball (-1 = No Ball there)
   If Balls(i, f) = -1 Then
    'Clear this part
    BitBlt PicFront.hdc, x, y, 15, 15, 0, 0, 0, BLACKNESS
   Else
    'Blit a ball/Block
    BitBlt PicFront.hdc, x, y, 15, 15, PicBalls(GameStyle).hdc, Balls(i, f) * 15, 0, vbSrcCopy
   End If
  Next f
 Next i
 'Schow the result
 PicFront.Refresh
End Sub


'Master Gameroutine
Private Sub PicFront_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

 'Enter your name so dot do anything
 If Text1.Visible = True Then Exit Sub

 'First click so remove Text and start the game
 If GameRun = False Then
  GameRun = True
  'Hide Settings Menu and show Game Menus
  MnuSet.Visible = False
  MnuEnd.Visible = True
  MnuReset.Visible = True
  Scroller.Enabled = False
  DoEvents
  'Clear the whole Pic
  BitBlt PicFront.hdc, 0, 0, 400, 400, 0, 0, 0, BLACKNESS
  'redraw it
  DrawField
  'Out here
  Exit Sub
 End If

 'Get the real Coordinates
 x = Int((x) / 16)
 y = Int((y) / 16)

 'What color did we click
 HitCol = Balls(x, y)


 'We hit a ball
 If HitCol > -1 And HitCol < 6 Then

  'No hits at this moment
  HTS = 0
  'Clear the hits field
  For i = 0 To 24
   For f = 0 To 24
    Hits(i, f) = 0
   Next f
  Next i

  'Now search for all Balls we hit
  FindNext x, y

  'Now search for Bombs and lets detonate if we found
  If BombsTrue = 2 Then
   ScanForBombs
  End If

 Else
  Exit Sub
 End If


 'Calculate and Show new score
 'This is the only thing i took from Cascade
 'the formular for Points
 Points = Points + HTS * ((HTS + 1) * 0.5)
 Me.Caption = "Brainballs   Your Score = " & Points

MoveIt:
 'Move all possible Balls down
 MoveDown

 'Now check for the 2 styles/Games
 If GameStyle = 0 Then
  RollDown
 Else
  MoveLeft
 End If

 If BombsTrue = 2 Then
  For i = 0 To 24
   If Balls(i, 24) > 5 Then
    Balls(i, 24) = -1
    GoTo MoveIt
   End If
  Next i
 End If

 'Show the new field
 DrawField

 'Check if we did the last possible move
 HTS = 0
 For i = 0 To 24
  For f = 0 To 24
   If Balls(i, f) <> -1 Then
    HTS = 1
    If SurBall(i, f) = True Then
     Exit Sub
    End If
   End If
  Next f
 Next i

 'draw a new field if we removed all balls
 If HTS = 0 Then
  CalcField
 Else
  'No moves left but there are still balls
  MsgBox "Game Over"

  'Let the user enter his HeighScore if he has a new
  For f = 0 To 4
   If Points > Scores(GameStyle + BombsTrue, GameLevel - 2, f).Score Then
    BitBlt PicFront.hdc, 160, 165, 80, 35, 0, 0, 0, BLACKNESS
    TextOut PicFront.hdc, 160, 160, "Enter your Name", 15
    PicFront.Refresh
    Text1.Visible = True
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.SetFocus
    GameRun = False
    MnuSet.Visible = False
    MnuEnd.Visible = False
    MnuReset.Visible = False
    Exit Sub
   End If
  Next f

  'Start a new game
  StartNew
 End If
End Sub

'Fill the holes under the balls
Private Sub MoveDown()
 Dim Tmp As Long

 'Move Down
 For i = 0 To 24
  For f = 24 To 0 Step -1

   'we found a hole
   If Balls(i, f) = -1 Then
    'So inc our tem counter
    Tmp = Tmp + 1

    'if the temcounter <>0 we hve holes under this ball
   ElseIf Tmp <> 0 Then
    'move it down
    'tmp tells how far
    Balls(i, f + Tmp) = Balls(i, f)
    Balls(i, f) = -1
   End If
  Next f
  If Tmp > 0 Then
   For f = Tmp - 1 To 0 Step -1
    If Balls(i, f) = -1 Then Exit For
    Balls(i, f) = -1
   Next f
  End If
  Tmp = 0
 Next i
End Sub

'This routine is only if we play with balls
Private Sub RollDown()

 'Before    After
 'o
 'O         O
 'O         Oo
 'OO        OO
 'First try from to roll from left then from right

 Dim Roll As Boolean

Do
Roll = False
For i = 0 To 24
 For f = 0 To 23
  'is there a hole we can fill
  If Balls(i, f) = -1 And Balls(i, f + 1) = -1 Then
   If i > 0 Then
    'Move the ball to the right and down
    If Balls(i - 1, f) <> -1 Then
     For g = f + 1 To 24
      If Balls(i, g) <> -1 Then Exit For
     Next g
     g = g - 1
     'set the new Position
     Balls(i, g) = Balls(i - 1, f)
     'Remove the old Ball
     Balls(i - 1, f) = -1
     'Set true so we check again
     Roll = True
     'Try for a new ball
     GoTo nextball
    End If
   End If
   'Right side
   If i < 24 Then
    If Balls(i + 1, f) <> -1 Then
     For g = f + 1 To 24
      If Balls(i, g) <> -1 Then Exit For
     Next g
     g = g - 1
     Balls(i, g) = Balls(i + 1, f)
     Balls(i + 1, f) = -1
     Roll = True
     GoTo nextball
    End If
   End If
  End If
 Next f
Next i
nextball:
'Loop until we didnt find any ball to roll
Loop Until Roll = False
End Sub

'Only for Blocks (Cascade)
Private Sub MoveLeft()

 'Check if there is an empty line
 'move everything to the left if there is
 For g = 0 To 23
  If Balls(g, 24) = -1 Then
   For i = g + 1 To 24
    For f = 0 To 24
     Balls(i - 1, f) = Balls(i, f)
    Next f
   Next i

   'Clear the last row
   For f = 0 To 24
    Balls(24, f) = -1
   Next f

   'check if we moved a blank field to the left
   For f = g To 24
    If Balls(f, 24) <> -1 Then
     g = g - 1
     Exit For
    End If
   Next f
  End If
 Next g
End Sub

'Is this a poosible hit
Private Function SurBall(ByVal x As Long, ByVal y As Long) As Boolean
 Dim BallCol As Long

 BallCol = Balls(x, y)

 If x > 0 Then
  If Balls(x - 1, y) = BallCol Then
   SurBall = True
   Exit Function
  End If
 End If
 If x < 24 Then
  If Balls(x + 1, y) = BallCol Then
   SurBall = True
   Exit Function
  End If
 End If
 If y > 0 Then
  If Balls(x, y - 1) = BallCol Then
   SurBall = True
   Exit Function
  End If
 End If
 If y < 24 Then
  If Balls(x, y + 1) = BallCol Then
   SurBall = True
  End If
 End If

End Function

'Find any ball we hit
Private Sub FindNext(ByVal x As Long, ByVal y As Long)

 'Test for left
 If x > 0 Then
  'If the ball to the left has the right color and
  'we havent selectet it already
  If Balls(x - 1, y) = HitCol And Hits(x - 1, y) = 0 Then
   'Select this ball
   Hits(x - 1, y) = 1
   'Remove the ball from the map
   Balls(x - 1, y) = -1
   'Add an hit for score
   HTS = HTS + 1
   'Search again
   FindNext x - 1, y
   'Check for a Bomb
  ElseIf Balls(x - 1, y) > 5 Then
   Hits(x - 1, y) = 2
  End If
 End If

 If x < 24 Then
  If Balls(x + 1, y) = HitCol And Hits(x + 1, y) = 0 Then
   Hits(x + 1, y) = 1
   Balls(x + 1, y) = -1
   HTS = HTS + 1
   FindNext x + 1, y
  ElseIf Balls(x + 1, y) > 5 Then
   Hits(x + 1, y) = 2
  End If
 End If
 If y > 0 Then
  If Balls(x, y - 1) = HitCol And Hits(x, y - 1) = 0 Then
   Hits(x, y - 1) = 1
   Balls(x, y - 1) = -1
   HTS = HTS + 1
   FindNext x, y - 1
  ElseIf Balls(x, y - 1) > 5 Then
   Hits(x, y - 1) = 2
  End If
 End If

 If y < 24 Then
  If Balls(x, y + 1) = HitCol And Hits(x, y + 1) = 0 Then
   Hits(x, y + 1) = 1
   Balls(x, y + 1) = -1
   HTS = HTS + 1
   FindNext x, y + 1
  ElseIf Balls(x, y + 1) > 5 Then
   Hits(x, y + 1) = 2
  End If
 End If

End Sub

Private Sub ScanForBombs()
 Dim BombFound As Boolean

Do
BombFound = False
'Scan all possible positions for bombs
'1-23 because a bomb on 0 cant be sourounded from balls
For i = 1 To 23
 For f = 1 To 23
  If Hits(i, f) = 2 Then
   Explode ScanAround(i, f)
   BombFound = True
  End If
 Next f
Next i
Loop Until BombFound = False

End Sub

Private Function ScanAround(x As Long, y As Long) As Boolean
 Dim Tmp As Long

 Hits(x, y) = 3
 ScanAround = True

 If x > 0 Then
  Tmp = Hits(x - 1, y)
  'Nothing there
  If Tmp = 0 Then
   ScanAround = False
   'Found a second Bomb
  ElseIf Tmp = 2 Then
   ScanAround = ScanAround(x - 1, y)
  End If
 Else
  ScanAround = False
 End If
 If ScanAround = False Then Exit Function

 If x < 24 Then
  Tmp = Hits(x + 1, y)
  'Nothing there
  If Tmp = 0 Then
   ScanAround = False
   'Found a second Bomb
  ElseIf Tmp = 2 Then
   ScanAround = ScanAround(x + 1, y)
  End If
 Else
  ScanAround = False
 End If
 If ScanAround = False Then Exit Function

 If y > 1 Then
  Tmp = Hits(x, y - 1)
  'Nothing there
  If Tmp = 0 Then
   ScanAround = False
   'Found a second Bomb
  ElseIf Tmp = 2 Then
   ScanAround = ScanAround(x, y - 1)
  End If
 Else
  ScanAround = False
 End If
 If ScanAround = False Then Exit Function

 If y < 24 Then
  Tmp = Hits(x, y + 1)
  'Nothing there
  If Tmp = 0 Then
   ScanAround = False
   'Found a second Bomb
  ElseIf Tmp = 2 Then
   ScanAround = ScanAround(x - 1, y)
  End If
 Else
  ScanAround = False
 End If
End Function


Private Sub Explode(DoIt As Boolean)
 Dim Second As Boolean
 Dim h As Long
 Dim r As Long

Do
Second = False
If DoIt = False Then
 'Bombs are not sourounded by color so dont explode
 For i = 0 To 24
  For f = 0 To 24
   If Hits(i, f) = 3 Then Hits(i, f) = 0
  Next f
 Next i
Else
 For i = 0 To 24
  For f = 0 To 24
   If Hits(i, f) = 3 Then
    Hits(i, f) = 0
    'Up Down
    If Balls(i, f) = 7 Then
     Balls(i, f) = -1
     For g = 0 To 24
      'A second bomb we hit ?
      If Balls(i, g) > 5 Then
       Hits(i, g) = 3
       Second = True
      ElseIf Balls(i, g) <> -1 Then
       HTS = HTS + 1
       Balls(i, g) = -1
      End If
     Next g
     'Left Right
    ElseIf Balls(i, f) = 6 Then
     Balls(i, f) = -1
     For g = 0 To 24
      'A second bomb we hit ?
      If Balls(g, f) > 5 Then
       Hits(g, f) = 3
       Second = True
      ElseIf Balls(g, f) <> -1 Then
       HTS = HTS + 1
       Balls(g, f) = -1
      End If
     Next g
     'Normal Bomb create a round explosion
    ElseIf Balls(i, f) = 8 Then
     Balls(i, f) = -1
     For g = f - 3 To f + 2
      If g < f - 1 Then h = h + 1
      If g > f + 1 Then h = h - 1
      For r = i - h To i + h
       If r < 0 Or r > 24 Or g < 0 Or g > 24 Then
        r = r
       Else
        'A second bomb we hit ?
        If Balls(r, g) > 5 Then
         Hits(r, g) = 3
         Second = True
        ElseIf Balls(r, g) <> -1 Then
         HTS = HTS + 1
         Balls(r, g) = -1
        End If
       End If
      Next r
     Next g
    End If
   End If
  Next f
 Next i
End If
Loop Until Second = False
End Sub
'Complete Restart
Private Sub MnuEnd_Click()
 Form_Load
End Sub

'Change the level
Private Sub MnuLevel_Click(Index As Integer)
 MnuLevel(GameLevel).Checked = False
 MnuLevel(Index).Checked = True
 GameLevel = Index
 StartNew
End Sub

'New Field
Private Sub MnuReset_Click()
 CalcField
 Points = 0
End Sub

'Change Style/game
Private Sub MnuStyle_Click(Index As Integer)
 MnuStyle(GameStyle + BombsTrue).Checked = False
 MnuStyle(Index).Checked = True
 If Index = 2 Then
  GameStyle = 0
  BombsTrue = 2
 Else
  GameStyle = Index
  BombsTrue = 0
 End If

 StartNew
End Sub


'Load/Save/Enter HighScore


Private Sub SaveScore()

 Open AppPath & "BB.SC" For Output As #1


  For i = 0 To 2
   For f = 0 To 3
    For g = 0 To 4
     Print #1, Str(Scores(i, f, g).Score)
     Print #1, Scores(i, f, g).User
    Next g
   Next f
  Next i
 Close
End Sub
Private Sub LoadScore()
 Dim Tmp As String
 'No scorefile so set some data
 If Dir$(AppPath & "BB.SC") = "" Then
  For i = 0 To 2
   For f = 0 To 3
    For g = 0 To 4
     Scores(i, f, g).Score = 1000
     Scores(i, f, g).User = "Nobody"
    Next g
   Next f
  Next i
 Else
  'load scorefile
  Open AppPath & "BB.SC" For Input As #1
   For i = 0 To 2
    For f = 0 To 3
     For g = 0 To 4
      Line Input #1, Tmp
      Scores(i, f, g).Score = Val(Tmp)
      Line Input #1, Scores(i, f, g).User
     Next g
    Next f
   Next i
  Close
 End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
 'Enter your Name for the Highscore
 If KeyCode = 13 Then
  For f = 0 To 4
   If Points > Scores(GameStyle + BombsTrue, GameLevel - 2, f).Score Then
    Text1.Visible = True
    For g = 4 To f + 1 Step -1
     Scores(GameStyle + BombsTrue, GameLevel - 2, g) = Scores(GameStyle + BombsTrue, GameLevel - 2, g - 1)
    Next g
    Scores(GameStyle + BombsTrue, GameLevel - 2, f).Score = Points
    Scores(GameStyle + BombsTrue, GameLevel - 2, f).User = Text1.Text
    SaveScore
    Exit For
   End If
  Next f
  MnuSet.Visible = True
  MnuEnd.Visible = False
  MnuReset.Visible = False
  Text1.Visible = False
  StartNew
 End If
End Sub

Private Sub Scroller_Timer()
 'A simple pseudoscroller
 'It reads the scrolltext from a Picture

 'I could make a complete ABC and calc the text from a string
 'but its smaller to save only the text i need (this time)
 For i = 0 To 24
  'Start on pixel over the pic and end one under it
  'to get a border
  For f = -1 To 5
   g = GetPixel(PicBalls(2).hdc, HTS + i, f)
   If g <> 0 Then g = 1
   BitBlt PicFront.hdc, i * 16, f * 16 + 32, 15, 15, PicBalls(GameStyle).hdc, g * 15, 0, vbSrcCopy
  Next f
 Next i
 HTS = HTS + 1
 If HTS > PicBalls(2).Width - 24 Then HTS = 0
 PicFront.Refresh
End Sub
