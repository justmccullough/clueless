VERSION 5.00
Begin VB.Form dlgGuessResults 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guess Results"
   ClientHeight    =   3210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5220
   Icon            =   "dlgGuessResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1013
      Picture         =   "dlgGuessResults.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Image imgSuspect 
      Height          =   1905
      Index           =   3
      Left            =   3360
      Picture         =   "dlgGuessResults.frx":1CAC
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgSuspect 
      Height          =   1905
      Index           =   1
      Left            =   3360
      Picture         =   "dlgGuessResults.frx":40A1
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgSuspect 
      Height          =   1905
      Index           =   4
      Left            =   3360
      Picture         =   "dlgGuessResults.frx":6539
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgSuspect 
      Height          =   1905
      Index           =   2
      Left            =   3360
      Picture         =   "dlgGuessResults.frx":895A
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgSuspect 
      Height          =   1905
      Index           =   5
      Left            =   3360
      Picture         =   "dlgGuessResults.frx":ACCD
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgSuspect 
      Height          =   1905
      Index           =   0
      Left            =   3360
      Picture         =   "dlgGuessResults.frx":D097
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   5
      Left            =   1800
      Picture         =   "dlgGuessResults.frx":F5BE
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   4
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":11654
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   0
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":13828
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   2
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":1576C
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   3
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":17534
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   7
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":1934F
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   6
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":1B434
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgRoom 
      Height          =   1905
      Index           =   1
      Left            =   1920
      Picture         =   "dlgGuessResults.frx":1D525
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgWeapon 
      Height          =   1905
      Index           =   0
      Left            =   480
      Picture         =   "dlgGuessResults.frx":1F3E0
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgWeapon 
      Height          =   1920
      Index           =   1
      Left            =   480
      Picture         =   "dlgGuessResults.frx":21156
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgWeapon 
      Height          =   1905
      Index           =   2
      Left            =   480
      Picture         =   "dlgGuessResults.frx":22F7A
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgWeapon 
      Height          =   1860
      Index           =   3
      Left            =   480
      Picture         =   "dlgGuessResults.frx":24EAE
      Top             =   240
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Image imgWeapon 
      Height          =   1905
      Index           =   4
      Left            =   480
      Picture         =   "dlgGuessResults.frx":2697A
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgWeapon 
      Height          =   1875
      Index           =   5
      Left            =   480
      Picture         =   "dlgGuessResults.frx":284EA
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "dlgGuessResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Private Sub cmdContinue_Click()
Unload Me
Dim i As Integer
For i = 0 To 5
    imgSuspect(i).Visible = False
    imgWeapon(i).Visible = False
Next
For i = 0 To 7
    imgRoom(i).Visible = False
Next
frmMain.SwitchPlayer
End Sub
