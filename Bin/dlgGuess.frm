VERSION 5.00
Begin VB.Form dlgGuess 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guess"
   ClientHeight    =   4500
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7395
   Icon            =   "dlgGuess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMove 
      Height          =   735
      Left            =   4800
      Picture         =   "dlgGuess.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame frameGuess 
      BackColor       =   &H00008000&
      Caption         =   "Guess"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   4800
      TabIndex        =   15
      Top             =   240
      Width           =   2415
      Begin VB.TextBox txtWeapon 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtSuspect 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtRoom 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame frameWeapons 
      BackColor       =   &H00008000&
      Caption         =   "Weapons"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   4455
      Begin VB.OptionButton optWeapon 
         BackColor       =   &H00008000&
         Caption         =   "Wrench"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   5
         Left            =   3000
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optWeapon 
         BackColor       =   &H00008000&
         Caption         =   "Lead Pipe"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   4
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optWeapon 
         BackColor       =   &H00008000&
         Caption         =   "Rope"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optWeapon 
         BackColor       =   &H00008000&
         Caption         =   "Revolver"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optWeapon 
         BackColor       =   &H00008000&
         Caption         =   "Knife"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optWeapon 
         BackColor       =   &H00008000&
         Caption         =   "Candlestick"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frameSuspect 
      BackColor       =   &H00008000&
      Caption         =   "Suspects"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Colonel Mustard"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Mrs. White"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Professor Plum"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   2
         Left            =   2940
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Mrs. Peacock"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   5
         Left            =   3000
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Miss Scarlet"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   4
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Mr. Green"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdGuess 
      Height          =   735
      Left            =   4800
      Picture         =   "dlgGuess.frx":1B6D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "dlgGuess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit
Dim iCharacterIndex, iWeaponIndex As Integer

Private Sub cmdGuess_Click()
Dim bRoomMatch, bWeaponMatch, bSuspectMatch As Boolean
Dim i As Integer

If txtRoom.Text = sRoom Then
    bRoomMatch = True
    For i = 0 To 7
        If frmMain.chkRoom(i).Caption = txtRoom.Text Then
            frmMain.chkRoom(i).Value = vbUnchecked
        Else
            frmMain.AddToString frmMain.chkRoom(i).Caption, 0
        End If
    Next
Else
    Select Case txtRoom.Text
        Case "Kitchen"
            dlgGuessResults.imgRoom(0).Visible = True
        Case "Library"
            dlgGuessResults.imgRoom(1).Visible = True
        Case "Hall"
            dlgGuessResults.imgRoom(2).Visible = True
        Case "Study"
            dlgGuessResults.imgRoom(3).Visible = True
        Case "Conservatory"
            dlgGuessResults.imgRoom(4).Visible = True
        Case "Ballroom"
            dlgGuessResults.imgRoom(5).Visible = True
        Case "Billiard Room"
            dlgGuessResults.imgRoom(6).Visible = True
        Case "Dining Room"
            dlgGuessResults.imgRoom(7).Visible = True
    End Select
    bRoomMatch = False
    frmMain.AddToString txtRoom.Text, 0
End If

If txtWeapon.Text = sWeapon Then
    bWeaponMatch = True
    For i = 0 To 5
        If frmMain.chkWeapon(i).Caption = txtWeapon.Text Then
            frmMain.chkWeapon(i).Value = vbUnchecked
        Else
            frmMain.AddToString frmMain.chkWeapon(i).Caption, 2
        End If
    Next
Else
    dlgGuessResults.imgWeapon(iWeaponIndex).Visible = True
    bWeaponMatch = False
    frmMain.AddToString txtWeapon.Text, 2
End If

If txtSuspect.Text = sPerp Then
    bSuspectMatch = True
    For i = 0 To 5
        If frmMain.chkSuspect(i).Caption = txtSuspect.Text Then
            frmMain.chkSuspect(i).Value = vbUnchecked
        Else
            frmMain.AddToString frmMain.chkSuspect(i).Caption, 1
        End If
    Next
Else
    dlgGuessResults.imgSuspect(iCharacterIndex).Visible = True
    bSuspectMatch = False
    frmMain.AddToString txtSuspect.Text, 1
End If

If bRoomMatch = True And bWeaponMatch = True And bSuspectMatch = True Then
    Unload dlgGuess
    frmMain.Hide
    frmWon.Show vbModal
    Exit Sub
End If
        
frmMain.SetMoves 0
Unload Me
dlgGuessResults.Show vbModal

End Sub

Private Sub cmdMove_Click()

If sCurrentRoom = "Kitchen" Then
    frmMain.imgCharacter(iCurrentIndex).Top = 6360
    frmMain.imgCharacter(iCurrentIndex).Left = 8160
    sCurrentRoom = "Study"
ElseIf sCurrentRoom = "Study" Then
    frmMain.imgCharacter(iCurrentIndex).Top = 2520
    frmMain.imgCharacter(iCurrentIndex).Left = 2880
    sCurrentRoom = "Kitchen"
ElseIf sCurrentRoom = "Conservatory" Then
    frmMain.imgCharacter(iCurrentIndex).Top = 6360
    frmMain.imgCharacter(iCurrentIndex).Left = 2880
    sCurrentRoom = "Dining Room"
ElseIf sCurrentRoom = "Dining Room" Then
    frmMain.imgCharacter(iCurrentIndex).Top = 2520
    frmMain.imgCharacter(iCurrentIndex).Left = 7200
    sCurrentRoom = "Conservatory"
End If

Unload Me
Dim iMoves As Integer
iMoves = frmMain.GetMoves
frmMain.SetMoves iMoves - 1
End Sub

Private Sub Form_Load()
txtRoom.Text = sCurrentRoom

If sCurrentRoom = "Conservatory" Or sCurrentRoom = "Kitchen" Or sCurrentRoom = "Study" Or sCurrentRoom = "Dining Room" Then
    cmdMove.Visible = True
Else
    cmdMove.Visible = False
End If
    
End Sub

Private Sub optCharacter_Click(Index As Integer)
iCharacterIndex = Index
txtSuspect.Text = optCharacter(Index).Caption
End Sub

Private Sub optWeapon_Click(Index As Integer)
iWeaponIndex = Index
txtWeapon.Text = optWeapon(Index).Caption
End Sub

Private Sub SetArrays()
players(iCurrentPlayer).Rooms = Rooms
players(iCurrentPlayer).Suspects = Suspects
players(iCurrentPlayer).Weapons = Weapons
End Sub
