VERSION 5.00
Begin VB.Form frmRules 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rules (Version 2.0)"
   ClientHeight    =   7260
   ClientLeft      =   4185
   ClientTop       =   3660
   ClientWidth     =   7020
   Icon            =   "Final_QuachARules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstRules 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      ItemData        =   "Final_QuachARules.frx":0442
      Left            =   720
      List            =   "Final_QuachARules.frx":0444
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "StackJack Rules"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Quach
'Date: June 2 / 2017
'Purpose: To display the full rules of version 2.0 of the game to the user
Option Explicit
Private Sub cmdOkay_Click()
    'Unloads the rules form
    Unload Me
End Sub

Private Sub Form_Load()
    'Displays the rules and general information of the game for version 2.0
    With lstRules
        .AddItem "General information: Upon loading the "
        .AddItem "game, you are intially shown one card "
        .AddItem "at the top of the deck. To start the "
        .AddItem "game and game timer, choose "
        .AddItem "which of the five hands you wish to add "
        .AddItem "the card to. The object of the game is "
        .AddItem "to get each hand as close to 21 as "
        .AddItem "possible without going over. If you "
        .AddItem "don't want to play a card, you can "
        .AddItem "click on 'Discard Card' to discard "
        .AddItem "it and get the next one. Once all the "
        .AddItem "cards in the deck have been used the "
        .AddItem "game ends and you win."
        .AddItem "Note: The timer will be paused when you"
        .AddItem "are inside a menu."
        .AddItem ""
        .AddItem "-----------------Rules-----------------"
        .AddItem "Rule #1: If a column adds to a value "
        .AddItem "over 21, it goes bust and can no "
        .AddItem "longer be used for the remainder of "
        .AddItem "the game."
        .AddItem "Rule #2: You lose 700 points when a "
        .AddItem "column goes bust. "
        .AddItem "Rule #3: If all five columns go over, "
        .AddItem "then you lose and the game is over. "
        .AddItem "Your score becomes zero. "
        .AddItem "Rule #4: Every time a card is "
        .AddItem "sucessfully placed, it gives 40 points "
        .AddItem "to you. An additional 10 points "
        .AddItem "(50 points total) is awarded when a "
        .AddItem "face card (Jack, Queen, King) or an Ace "
        .AddItem "is placed."
        .AddItem "Rule #5: Hand (column scores) scores"
        .AddItem "are based off of the cards face value. "
        .AddItem "An ace can be 11 or 1, depending on the "
        .AddItem "situation. All face cards are worth"
        .AddItem "10, while all other cards are based"
        .AddItem "on number value. (e.g. 10 is worth 10, "
        .AddItem "2 is worth 2, 3 is worth 3). No "
        .AddItem "additional points are given for "
        .AddItem "different suits."
        .AddItem "Rule #6: If a hand adds to a value "
        .AddItem "of exactly 21, it is cleared and the "
        .AddItem "program continues to function properly."
        .AddItem "You get 500 points times the number of"
        .AddItem "consecutive hand clears."
        .AddItem "Rule #7: If a card is discarded, 150 "
        .AddItem "points are deducted from the current "
        .AddItem "score."
        .AddItem "Rule #8: There is no negative scoring. "
        .AddItem "The lowest amount of points is 0."
        .AddItem "Rule #9: You only win when all the "
        .AddItem "cards in the deck are used up and "
        .AddItem "your score is above 0."
    End With
    
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Resumes the game
    UnPauseGame frmMain, frmMain.tmrGame
    
End Sub

Private Sub lstRules_Click()
    'Deselects the listbox so a blue highlight is not there
    lstRules.ListIndex = -1
End Sub
