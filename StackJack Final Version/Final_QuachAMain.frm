VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   Caption         =   "StackJack"
   ClientHeight    =   7035
   ClientLeft      =   300
   ClientTop       =   2805
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Final_QuachAMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   12675
   Begin VB.Timer tmrGame 
      Left            =   960
      Top             =   1920
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   5
      Left            =   10800
      ScaleHeight     =   391
      ScaleMode       =   0  'User
      ScaleWidth      =   105
      TabIndex        =   19
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Discard Card"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox picUserCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   480
      ScaleHeight     =   151
      ScaleMode       =   0  'User
      ScaleWidth      =   119.621
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   3
      Left            =   6960
      ScaleHeight     =   391
      ScaleMode       =   0  'User
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   4
      Left            =   8880
      ScaleHeight     =   391
      ScaleMode       =   0  'User
      ScaleWidth      =   105
      TabIndex        =   13
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   2
      Left            =   5040
      ScaleHeight     =   391
      ScaleMode       =   0  'User
      ScaleWidth      =   111
      TabIndex        =   11
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picStack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   1
      Left            =   3120
      ScaleHeight     =   391
      ScaleMode       =   0  'User
      ScaleWidth      =   111
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblMultiplier 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Caption         =   "Consecutive clears:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lblCardsLeft 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cards Left:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   12000
      TabIndex        =   18
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   9960
      TabIndex        =   17
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   8040
      TabIndex        =   16
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   15
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblTotalPoints 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Caption         =   "Points:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Count:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Count:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10920
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Count:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Count:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Count:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Current Card"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "High Scores"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuRules 
         Caption         =   "&Rules"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Quach
'Date: June 6 / 2017
'Purpose: To create and successfully run the game of stackjack
Option Explicit
'Maximum amount of cards in deck
Const MAXCARD = 52
'Minimum amount of cards in deck
Const MINCARD = 1
Const SHRINK = 8
'Keeps track of the card index
Dim CardIndex(MINCARD To MAXCARD) As Integer
'Handles the user's response at end game
Dim NewGame As Integer
'Keeps track of the different scores
Dim UserScore As Long
Dim ColumnScore(1 To MAXGAMEBOARDS) As Integer
'Keeps track of the time and the pause status of the game
Dim PauseCounter As Integer
Dim StartTime As Date

Private Sub Form_Load()

    Dim K As Integer
    Dim X As Integer
    Dim CardsDLLInit As Long
       
    Randomize
    
    '26 bytes + 4 bytes (long) + 8 bytes (date) - Length of one individual record
    RecordLen = 38
    
    'Creates the file if no high score file exists
    If Dir$(App.Path & FNAME) = "" Then

        CreateFile

    End If
    
    'Reads the current highscores on file
    ReadFile
    
    'Intializes the game
     Intialize picStack(), picUserCard, cmdDiscard, mnuNewGame, UserScore, ColumnScore(), lblTotalPoints, lblCount, NewGame, tmrGame, lblTime, lblMultiplier
        
    'Intializes the DLL
    CardsDLLInit = cdtInit(CardWidth, CardHeight)

    'Generates the height and width of each card according to the game box (in pixels)
    CardHeight = picUserCard.ScaleHeight
    CardWidth = picUserCard.ScaleWidth - SHRINK
    
    'Generates the deck of cards for game use
    GenerateDeck CardIndex(), MAXCARD, MINCARD
    DisplayUserCard CardIndex(), picUserCard, lblCardsLeft
    
    'Centers the form
    CenterForm Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Dim Ret As Long
    
    'Frees up the memory
    Ret = cdtTerm()
    
End Sub

Private Sub mnuAbout_Click()

    PauseGame frmMain
    'Displays the about form
    frmAbout.Show vbModal

End Sub
Private Sub mnuExit_Click()

    PauseGame frmMain

    'Exits the program if the user selects yes
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you would like to exit?", vbInformation + vbYesNo, "Exit")

    If Response = vbYes Then
        Beep
        End
    Else
        'Resumes the game
        UnPauseGame frmMain, tmrGame
    End If

End Sub

Private Sub mnuHighScore_Click()
    
    PauseGame frmMain
    'Displays the high score form
    frmHighScores.Show vbModal

End Sub

Private Sub mnuNewGame_Click()
 
    'Executes the appropriate steps for a new game
    Intialize picStack(), picUserCard, cmdDiscard, mnuNewGame, UserScore, ColumnScore(), lblTotalPoints, lblCount, NewGame, tmrGame, lblTime, lblMultiplier
    GenerateDeck CardIndex(), MAXCARD, MINCARD
    DisplayUserCard CardIndex(), picUserCard, lblCardsLeft
 
End Sub

Private Sub mnuRules_Click()

    PauseGame frmMain
    'Displays the rules form
    frmRules.Show vbModal

End Sub

Private Sub picStack_Click(Index As Integer)
    
    'Starts the game
    GameStart = True
    
    'Starts the game & timer once the user has clicked
    If GameStart = True And tmrGame.Enabled = False And GamePause = False Then
        StartTime = Now
        PauseTime = Now + #12:00:01 AM#
        tmrGame.Enabled = True
    End If

    Dim AdditionAmount As Integer
    Dim UserScoreAdditionAmount As Integer
    
    'Checks to see what the user clicked
    DiscardClick = False

    If CardCount <= 52 Then
        'Displays the card that was just placed by the user
        DisplayStackCard CardIndex(), picStack(Index), Index
        'Updates the user's score
        UserScoreAdditionAmount = UserScoreCheck(CardIndex())
        UserScore = UserScore + UserScoreAdditionAmount
        'Adds the appropriate amount of score to the appropriate hand and displays it above the appropriate hand
        AdditionAmount = ColumnScoreCheck(CardIndex(), ColumnScore(), Index)
        ColumnScore(Index) = ColumnScore(Index) + AdditionAmount
        'Checks if the hand has gone bust or if the hand has scored (21 points)
        ScoreCheck ColumnScore(), Index, picStack(), UserScore, UserScoreAdditionAmount
    End If
    
    'Updates the card count
    CardCount = CardCount + 1
    'Checks to see if all hands have bust
    GameOverCheck picStack()
    'Updates the user deck
    DisplayUserCard CardIndex(), picUserCard, lblCardsLeft
    'Displays the scores
    DisplayScore UserScore, ColumnScore(), lblTotalPoints, lblCount(), Index, picStack()
    
    'End game procedures
    If GameOver = True Then
        EndGameProcedures cmdDiscard, mnuNewGame, NewGame, UserScore, picStack(), tmrGame
        If NewGame = vbYes Then
            Call mnuNewGame_Click
        End If
    End If
End Sub
Private Sub cmdDiscard_Click()
    
    GameStart = True

    'Starts the game & timer once the user has clicked
    If GameStart = True And tmrGame.Enabled = False And GamePause = False Then
        StartTime = Now
        PauseTime = Now + #12:00:01 AM#
        tmrGame.Enabled = True
    End If
    
    'Keeps track of the users click to make sure the correct amount of points are subtracted later on
    DiscardClick = True
    'Increments the card counter by one
    CardCount = CardCount + 1
    'Resets the multiplier (discarding a card ends the consective multiplier)
    Multiplier = 1
    lblMultiplier.Caption = VBA.Str$(Multiplier) & "x"
    
    'Checks if the user has run out of cards
    GameOverCheck picStack()
    'Displays the next card in the deck
    DisplayUserCard CardIndex(), picUserCard, lblCardsLeft
    
    'Discards when there is still a deck
    'Checks if card count is 53 because when the user discards the last card, the card count becomes 53
    If CardCount <= 53 Then
      'Checks to see if the procedure was called in discard
      ScoreSubtract UserScore
    End If

    DisplayScore UserScore, ColumnScore(), lblTotalPoints, lblCount(), 0, picStack()
    
    'End Game procedures
    If GameOver = True Then
        EndGameProcedures cmdDiscard, mnuNewGame, NewGame, UserScore, picStack, tmrGame
        If NewGame = vbYes Then
            Call mnuNewGame_Click
        End If
    End If
         
End Sub

Private Sub tmrGame_Timer()

    'Pauses the timer if the user clicks away to another part of the application
    If GamePause = True Then
        tmrGame.Enabled = False
    ElseIf GamePause = False And GameStart = True Then
        'Displays the time that the game has been in progress for
        UserTime = PauseTime - StartTime
        lblTime.Caption = VBA.Format$(UserTime, "hh:mm:ss")
        PauseTime = PauseTime + #12:00:01 AM#
    End If

End Sub
