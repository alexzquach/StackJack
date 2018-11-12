Attribute VB_Name = "mdlMainCode"
'Programmer: Alex Quach
'Date: June 6 / 2017
'Purpose: To create and successfully run the game of stackjack
Option Explicit
Global CardHeight As Long
Global CardWidth As Long

'Global variables make it more accesible throughout the module and main form (multiple variables that are
'in use excessively and are difficult to call when procedures call other procedures)

'Keeps track of the number of gameboards
Global Const MAXGAMEBOARDS = 5

'Keeps track of how many cards have been used up so far, 1 to 52
Global CardCount As Single

'GameOver keeps track of the games current status
Global GameOver As Boolean
'HandCheck keeps track of how many hands have bust
Global HandCheck As Integer

'Keeps track of the multiplier for consecutive clears
Global Multiplier As Integer

'Keeps track if the user clicked the hand or discard button
Global DiscardClick As Boolean

'Keeps track of the time the user has elapsed during the game
Global UserTime As Date
Global PauseTime As Date
Global GamePause As Boolean
Global GameStart As Boolean
'All variables in code module are not needed in the main form (used for checking scores, hands and columns)

'Keeps track of the amount of points the user gets for clearing a hand
Const USERPOINTS = 500

'Keeps track of the last hand cleared
Dim LastCardClear As Integer


'Keeps track of the height per stack
Dim StackedHeight(1 To MAXGAMEBOARDS) As Long

'Determines if the user has placed an ace in the hand
Dim AceInHand(1 To MAXGAMEBOARDS) As Boolean
'Keeps track of how many times the ace value can be changed in comparison to number of aces in hand
'The aces can only be changed in value for as many times as there are aces (e.g 2 aces = 2 changes in value)
Dim SubtractCounter(1 To MAXGAMEBOARDS) As Integer
Dim MultipleAcesInHand(1 To MAXGAMEBOARDS) As Integer

Public Sub Intialize(picBox1 As Variant, PicBox2 As PictureBox, cmdButton1 As CommandButton, mnuButton1 As Menu, UScore As Long, CScore() As Integer, lblTotal As Label, lblColumn As Variant, NewGame As Integer, tmrTimer As Timer, lblTime As Label, lblMultiplier As Label)
 
    Dim K As Integer
 
    'Intializes all the necessary controls and variables
    CardCount = 1
    UScore = 0
    'Intializes the user time to 0 seconds
    UserTime = Now - Now
    GamePause = False
    GameStart = False
    'New game menu and discard button
    cmdButton1.Enabled = Not (cmdButton1.Enabled)
    mnuButton1.Enabled = Not (mnuButton1.Enabled)
    'Timer (1000 ms = 1 s)
    tmrTimer.Enabled = False
    tmrTimer.Interval = 1000
    lblTime.Caption = VBA.Format$(0, "hh:mm:ss")
    lblTotal.Caption = VBA.Str$(UScore)
    'Checks to see what hte user has clicked
    DiscardClick = False
    'Checks how many hands are bust
    HandCheck = 0
    'Keeps track of the most recent card used to clear a hand
    LastCardClear = 0
    'Keeps track of the score multipler for bonus scoring
    Multiplier = 1
    lblMultiplier.Caption = VBA.Str$(Multiplier) & "x"
    'Handles the current status of the game
    GameOver = False
    'NewGame handles the user's response to the new game message box at end game
    NewGame = 0
    For K = 1 To MAXGAMEBOARDS
        'All three variables are used to check for ace scoring
        AceInHand(K) = False
        SubtractCounter(K) = 0
        MultipleAcesInHand(K) = 0
        'StackedHeight contains the position of each card in each hand
        StackedHeight(K) = 0
        'Hands
        picBox1(K).Enabled = Not (picBox1(K).Enabled)
        picBox1(K).Cls
        'Labels above hands
        lblColumn(K).ForeColor = &HFFFFFF
        'Column scores
        CScore(K) = 0
        lblColumn(K).Caption = VBA.Str$(CScore(K))
    Next K
    
End Sub
Public Sub GenerateDeck(CardIndex() As Integer, Max As Integer, Low As Integer)
    Dim K As Integer, Num1 As Integer, Num2 As Integer, TempSwapHold As Integer
    
    'Generates the deck of cards using indices to keep track of the card
    For K = 1 To Max
        CardIndex(K) = K - 1
    Next K
    
    'Randomizes the deck of cards using the card indices generated
    For K = 1 To 32000
        Num1 = Int(Rnd() * (Max - Low + 1)) + Low
        Num2 = Int(Rnd() * (Max - Low + 1)) + Low

        TempSwapHold = CardIndex(Num1)
        CardIndex(Num1) = CardIndex(Num2)
        CardIndex(Num2) = TempSwapHold
    Next K

End Sub

Public Sub DisplayUserCard(CardIndex() As Integer, picBox As PictureBox, lblCardsLeft As Label)
    
    'Clears the deck in order to display a new one instead of displaying a new card on top of
    'an old card
    picBox.Cls
        
    Dim Ret As Long
        
    'Displays the card if there are still cards in the deck and if the game is running
    If CardCount < 53 And GameOver = False Then
        Ret = cdtDrawExt(picBox.hDC, 0, 0, CardWidth, CardHeight, CardIndex(CardCount), C_FACES, 0)
        'Changes the colour depending on how many cards are left in the deck
        Select Case CardCount
            Case 1 To 13
                lblCardsLeft.ForeColor = &HC000&
            Case 14 To 26
                lblCardsLeft.ForeColor = &HFFFF&
            Case 27 To 39
                lblCardsLeft.ForeColor = &H8080FF
            Case 40 To 52
                lblCardsLeft.ForeColor = &HFF&
        End Select
        lblCardsLeft.Caption = VBA.Str$(52 - CardCount)
    Else
        'Displays the "back of the cards" if the game is over or if there are no more cards (deck is flipped)
        Ret = cdtDrawExt(picBox.hDC, 0, 0, CardWidth, CardHeight, Flowers, C_BACKS, 0)
    End If
    'Refreshes the picture box since autoredraw is set to true
    picBox.Refresh

End Sub

Public Sub DisplayStackCard(CardIndex() As Integer, picBox As PictureBox, GameBoardNum As Integer)

    Dim Ret As Long
    
    'Displays the card if there are still cards
    Ret = cdtDrawExt(picBox.hDC, 0, StackedHeight(GameBoardNum), CardWidth, CardHeight, CardIndex(CardCount), C_FACES, 0)
    'Stacks the cards on top of each other and increases the height accordingly to
    'ensure the stack remains properly stacked
    StackedHeight(GameBoardNum) = StackedHeight(GameBoardNum) + 25

    'If the user places an ace, then there is an ace in that hand
    If CardIndex(CardCount) >= 0 And CardIndex(CardCount) <= 3 Then
        AceInHand(GameBoardNum) = True
        MultipleAcesInHand(GameBoardNum) = MultipleAcesInHand(GameBoardNum) + 1
    End If
    
    'Refreshes the picture box since autoredraw is set to true
    picBox.Refresh
    

End Sub


Public Function UserScoreCheck(CardIndex() As Integer) As Integer

    Dim TempScore As Integer

    'Checks to see what card was placed on the hand and assigns an appropriate amount of points corresponding
    'to the card's value (user score)
    Select Case CardIndex(CardCount)
        Case 0 To 3, 40 To 51
            '50 user points for a face card or an ace
            TempScore = 50
        Case Else
            '40 user points for any other card
            TempScore = 40
    End Select
    
    UserScoreCheck = TempScore

End Function

Public Sub DisplayScore(UScore As Long, CScore() As Integer, lblTPoints As Label, lblCount As Variant, ColumnNum As Integer, picColumn As Variant)

    'Displays appropriate user scores depending on the status of the game
    If GameOver = True And HandCheck = 5 Then
        lblTPoints.Caption = VBA.Str$(0)
    Else
        lblTPoints.Caption = VBA.Str$(UScore)
    End If
    
    'Checks if the procedure was called when the user discards a card
    If DiscardClick = False Then
        'Changes the colour of the column text if the user has busted the column
        If picColumn(ColumnNum).Enabled = False And ColumnNum <> 0 Then
            lblCount(ColumnNum).ForeColor = &HFF&
        End If
        
        lblCount(ColumnNum).Caption = VBA.Str$(CScore(ColumnNum))
    End If
    
End Sub

Public Function ColumnScoreCheck(CardIndex() As Integer, CScore() As Integer, GameBoardNum As Integer) As Integer

    Dim TempScore As Integer
    
    Select Case CardIndex(CardCount)
        Case 40 To 51
            'All face cards are worth 10 points in the hand/column
            TempScore = 10
        Case 0 To 3
            'Assigns 11 to the ace card
            TempScore = 11
        Case Else
            'Assigns points according to actual card value
            TempScore = VBA.Int(CardIndex(CardCount) / 4) + 1
    End Select
    
    'Checks if the ace in the hand makes the points go over
    If AceInHand(GameBoardNum) = True Then
        AceCheck CScore(), TempScore, GameBoardNum
    End If
    
    ColumnScoreCheck = TempScore

End Function

Public Sub AceCheck(CScore() As Integer, TempScore As Integer, GameBoardNum As Integer)
    
    'Subtracts 10 from the hand if the score goes over because the ace is worth 11
    If (CScore(GameBoardNum) + TempScore) > 21 And SubtractCounter(GameBoardNum) < MultipleAcesInHand(GameBoardNum) Then
        CScore(GameBoardNum) = CScore(GameBoardNum) - 10
        'Ensures the ace value is not changed more than once per hand
        SubtractCounter(GameBoardNum) = SubtractCounter(GameBoardNum) + 1
    End If
    
End Sub

Public Sub ScoreCheck(ColumnScore() As Integer, GameBoardNum As Integer, picBox As Variant, UScore As Long, AdditionAmount As Integer)

    If ColumnScore(GameBoardNum) = 21 Then
        'Resets the hand once the user has scored 21
        ColumnScore(GameBoardNum) = 0
        picBox(GameBoardNum).Cls
        StackedHeight(GameBoardNum) = 0
        AceInHand(GameBoardNum) = False
        MultipleAcesInHand(GameBoardNum) = 0
        SubtractCounter(GameBoardNum) = 0
        'Gives the user 500 points multiplied by the multplier for clearing the hand and subtracts the additional
        'points receieved from placing the card
        UScore = UScore + ConsecutiveCheck()
        UScore = UScore - AdditionAmount
        'Keeps track of the card used in the last clear
        LastCardClear = CardCount
    ElseIf ColumnScore(GameBoardNum) > 21 Then
        'Busts the hand, increases the number of bust hands by one, and subtracts the proper amount of pounts
        'Subtracts the addition amount from user score because you are not supposed to add if the card busts the column
        UScore = UScore - AdditionAmount
        picBox(GameBoardNum).Enabled = False
        HandCheck = HandCheck + 1
        ScoreSubtract UScore
        'Resets the multiplier to 1 if the user busts the hand after clearing another
        Multiplier = 1
        'Calls the form name since this is a "local" procedure (only used in code module)
        frmMain.lblMultiplier.Caption = VBA.Str$(Multiplier) & "x"
    Else
        'Resets the multiplier to 1 if the user does not clear the hand
        Multiplier = 1
        'Calls the form name since this is a "local" procedure (only used in code module)
        frmMain.lblMultiplier.Caption = VBA.Str$(Multiplier) & "x"
    End If

End Sub

Public Sub ScoreSubtract(UScore As Long)

    'Subtracts accordingly if the user clicked discard
    If DiscardClick = True Then
        If UScore > 150 Then
            UScore = UScore - 150
        Else
            UScore = 0
        End If
    End If
    
    'DiscardClick being false means that this procedure was called when the user busts a hand
    'and proceeds to subtracts accordingly for the hand busting
    If DiscardClick = False Then
        'Subtracts 700 if column goes bust
        If UScore > 700 Then
            UScore = UScore - 700
        Else
            UScore = 0
        End If
    End If
    
End Sub

Public Sub GameOverCheck(picBox1 As Variant)
    
    'Checks to see if the game is over (If CardCount = 53, that means the user has run out of cards)
    '(If HandCheck = 5, that means the user has bust all 5 hands)
    If CardCount = 53 Or HandCheck = 5 Then
        GameOver = True
    End If

End Sub

Public Sub EndGameProcedures(cmdButton As CommandButton, mnuButton As Menu, NGame As Integer, UScore As Long, picBox As Variant, tmrTimer As Timer)
       
       Dim K As Integer
       Dim TempName As String
       
       'Fixes the appropriate controls
        cmdButton.Enabled = False
        mnuButton.Enabled = True
        tmrTimer.Enabled = False
        GameStart = False
        'Checks if the user has won or lost the game now that the game is over
        If HandCheck = 5 Then
            'Resets the user score because if they bust, they lose the game and have a score of 0
            UScore = 0
            'Asks the user if they would like to start a new game when the game is finished
            NGame = MsgBox("Oh no, you lose!  You have busted all your hands!  Your final score is: " & VBA.Str$(UScore) & vbCrLf & "Would you like to start a new game?", vbCritical + vbYesNo, "Game over")
        'If the user did not lose, then that means the user used all the cards in the deck and the user has won
        Else
            For K = 1 To MAXGAMEBOARDS
                'Checks to see what hands need to be disabled in order to prepare for the new game
                If picBox(K).Enabled = True Then
                    picBox(K).Enabled = False
                End If
            Next K
            
            If UScore = 0 Then
                'Asks the user if they would like to start a new game when the game is finished
                NGame = MsgBox("Oh no, you lose! You used the entire deck and your final score is: " & VBA.Str$(UScore) & vbCrLf & "Would you like to start a new game?", vbCritical + vbYesNo, "Game over")
            Else
                MsgBox "Congratulations, you won the game by using the entire deck!  Your final score is: " & VBA.Str$(UScore), vbInformation, "Winner!"
                'Checks to see if there is a new highscore by comparing to the last person on the
                'highscore board (highscores are sorted)
                If UScore > HighScore(HIGHSCOREMAX).UHighScore Then

                    'Allows the user to input their name if they
                    'achieved a high score. Enters anonymous if the user has entered nothing.
                    TempName = VBA.Trim$(InputBox$("You have gotten a high score!" & vbCrLf & "Please enter your name!", "High score Winner!"))
                    If TempName = "" Then

                        TempName = "Anonymous"

                    End If

                    'Changes the user highscores and displays the highscores to the user
                    ChangeHighScore UScore, TempName
                    frmHighScores.Show vbModal
                    
                Else
                    'Asks the user if they would like to start a new game when the game is finished (if the user has not achieved a highscore)
                    NGame = MsgBox("Would you like to start a new game?", vbYesNo + vbInformation, "Winner!")
                End If
      
            End If
        End If
            
End Sub

Public Function ConsecutiveCheck() As Integer
    
    Dim TempAdditionAmount As Integer
    
    'Checks if they cleared the hand consecutively
    If (CardCount - LastCardClear) = 1 Then
        Multiplier = Multiplier + 1
    'If not consecutive, resets the consecutive variable to one
    Else
        Multiplier = 1
    End If
    
    'Calls the form name since this is a "local" procedure (only used in code module)
    frmMain.lblMultiplier.Caption = VBA.Str$(Multiplier) & "x"
    TempAdditionAmount = Multiplier * USERPOINTS
    ConsecutiveCheck = TempAdditionAmount

End Function

Public Sub CenterForm(frmForm As Form)

    'Centers the form
    frmForm.Move (Screen.Width - frmForm.ScaleWidth) / 2, (Screen.Height - frmForm.ScaleHeight) / 2

End Sub

Public Sub PauseGame(frmMain As Form)

    'Pauses the game
    GamePause = True
    If GameStart = True Then
        frmMain.Caption = frmMain.Caption & " (Paused)"
    End If

End Sub

Public Sub UnPauseGame(frmMain As Form, tmrGame As Timer)
    'Unpauses the game
    If GamePause = True And GameStart = True Then
        GamePause = False
        tmrGame.Enabled = True
        frmMain.Caption = "StackJack"
    Else
        GamePause = False
        tmrGame.Enabled = False
    End If
End Sub
