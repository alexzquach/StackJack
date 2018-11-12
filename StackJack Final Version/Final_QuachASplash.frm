VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   ClientHeight    =   4185
   ClientLeft      =   3795
   ClientTop       =   3570
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00000000&
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.Timer tmrLoad 
         Interval        =   3000
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Warning: The unauthorized reproduction or distribution of this game is strictly prohibited. "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   6735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Copyright © Alex Quach 2017 "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "Version 2.0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   360
         Picture         =   "Final_QuachASplash.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "By Alex Quach"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "StackJack"
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
         Left            =   2880
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Quach
'Date: June 2 / 2017
'Purpose: To display a splash screen that the user sits through for 3 seconds
'in order to add the feeling of video game "realism"
Option Explicit

Private Sub Form_Load()

    'Centers the form
    CenterForm Me

End Sub

Private Sub tmrLoad_Timer()

    'Loads the main form after a certain time interval
    tmrLoad.Enabled = False
    Unload frmSplash
    frmMain.Show
    
End Sub
