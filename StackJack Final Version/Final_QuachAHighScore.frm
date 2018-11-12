VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHighScores 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF80FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   3735
   ClientLeft      =   1920
   ClientTop       =   3750
   ClientWidth     =   10800
   Icon            =   "Final_QuachAHighScore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   10800
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Okay"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   10575
   End
   Begin MSFlexGridLib.MSFlexGrid grdData 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4948
      _Version        =   393216
      BackColor       =   65535
      Enabled         =   0   'False
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Quach
'Date: June 2 / 2017
'Prupose: To display the highscores
Option Explicit

Private Sub cmdOkay_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    'Centers the form
    CenterForm Me
    'Reads the highscore file and displays the data
    ReadFile
    DisplayFile grdData
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Resumes the game
    UnPauseGame frmMain, frmMain.tmrGame
End Sub
