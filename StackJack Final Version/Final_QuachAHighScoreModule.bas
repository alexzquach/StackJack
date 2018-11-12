Attribute VB_Name = "mdlHighScoreCode"
'Programmer: Alex Quach
'Date: June 6 / 2017
'Purpose: To maintain the highscores
Option Explicit
Global Const HIGHSCOREMAX = 5
'File name of the file that stores the highscores
Global Const FNAME = "\HIGHSCORES.rec"
'Highscore record type
Type HighScoreRec
    UName As String * 26
    UHighScore As Long
    HighTime As Date
End Type
Global RecordLen As Integer
'Stores the highscores
Global HighScore(1 To HIGHSCOREMAX) As HighScoreRec
Public Sub CreateFile()

    'Creates the file name if the file is not already created
    Dim K As Integer
    
    Open App.Path & FNAME For Random As #1 Len = RecordLen
    
    'Writes default information to the file
    For K = 1 To HIGHSCOREMAX
        'Intializes the default information
        With HighScore(K)
            .UName = "Anonymous"
            'Sets the default score to the lowest possible score (unachievable since the user loses if their score is 0)
            .UHighScore = 0
            'Sets the default time to the highest time possible
            .HighTime = #11:59:59 PM#
        End With
        Put #1, K, HighScore(K)
    Next K
    
    Close #1

End Sub

Public Sub ReadFile()

    Dim X As Integer

    'Reads the current high score information in the file
    Open App.Path & "\" & FNAME For Random As #1 Len = RecordLen
    For X = 1 To HIGHSCOREMAX
        Get #1, X, HighScore(X)
    Next X
    Close #1
    

End Sub

Public Sub DisplayFile(grdData As Control)

    Const ROWADJUST = 7
    Const UROWS = 6
    Const UCOLS = 3

    Dim K As Integer
    Dim DColumn As Integer
    Dim DRow As Integer
    Dim W As Integer

    'Sets the dimensions, headings of the grid (3 high score categories, max of 5 highscores per file), and
    'displays the data in the grid control as well
    With grdData

        .Rows = UROWS
        .Cols = UCOLS
        .FixedCols = 0
        
        For DRow = 0 To .Rows - 1
            .RowHeight(DRow) = 450 + ROWADJUST
        Next DRow
           
        'Main form font: courier new size 12
        .ColWidth(0) = frmMain.TextWidth("WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW")
        .ColWidth(1) = frmMain.TextWidth(" User Score ")
        .ColWidth(2) = frmMain.TextWidth(" User Time (24 hour format) ")
        
        .Row = 0
        .Col = 0
        .Text = "Name"
        .Col = 1
        .Text = "User Score"
        .Col = 2
        .Text = "User Time (24 hour format)"
        'Displays the highscore data
        For K = 1 To HIGHSCOREMAX
            .Row = K
            'Resets the column alignments
            For DColumn = 0 To UCOLS - 1
                'Left center alignment
                .ColAlignment(DColumn) = 1
            Next DColumn
            .Col = 0
            'Displays the names
            If VBA.Len(VBA.Trim$(HighScore(K).UName)) > 25 Then
                .Text = VBA.Trim$(VBA.Left$(HighScore(K).UName, 25)) & "..."
            Else
                .Text = VBA.Trim$(HighScore(K).UName)
            End If
            .Col = 1
            .Text = VBA.Trim$(VBA.Str$(HighScore(K).UHighScore))
            .Col = 2
            .Text = VBA.Format$((HighScore(K).HighTime), "hh:mm:ss")
        Next K
           
    End With

End Sub
Public Sub WriteFile()

    'Writes to the record file
    Dim X As Integer

    On Error GoTo ErrorHandler
    Kill FNAME

    Open App.Path & "\" & FNAME For Random As #1 Len = RecordLen
    For X = 1 To HIGHSCOREMAX
        Put #1, X, HighScore(X)
    Next X
    Close #1
    Exit Sub
    
ErrorHandler:
    Resume Next

End Sub

Public Sub ChangeHighScore(ByVal S As Long, N As String)

    Dim K As Integer
    Dim TempName As String
    Dim TempNum As Integer
    Dim TempTime1 As Date

    For K = 1 To HIGHSCOREMAX
        With HighScore(K)
            If S > .UHighScore Then
                'Swaps the highest score
                TempNum = S
                S = .UHighScore
                .UHighScore = TempNum
                'Swaps the highest name
                TempName = N
                N = .UName
                .UName = TempName
                'Swaps the highest time
                TempTime1 = UserTime
                UserTime = .HighTime
                .HighTime = TempTime1
            'Checks if there is a tiebreaker and if the scores are not default (unobtainable scores)
            ElseIf S = .UHighScore And .UHighScore <> 0 Then
                TempTime1 = UserTime
                'Swaps if there is a tiebreaker using time as the decider (names in alphabetical order are not considered
                'making the first person that has that specific time receive the position on the high score board)
                If TempTime1 < .HighTime Then
                    'Swaps the highest score
                    TempNum = S
                    S = .UHighScore
                    .UHighScore = TempNum
                    'Swaps the highest name
                    TempName = N
                    N = .UName
                    .UName = TempName
                    'Swaps the time
                    UserTime = .HighTime
                    .HighTime = TempTime1
                End If
            End If
        End With
    Next K
    'Writes to the record file
    WriteFile
    
End Sub
      

