VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmSelectCharacter 
   BackColor       =   &H00008000&
   Caption         =   "Please select your character"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7680
   FillColor       =   &H00008000&
   ForeColor       =   &H00000000&
   Icon            =   "frmSelectCharacter.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7530
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPlayerName 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2213
      TabIndex        =   1
      Top             =   1140
      Width           =   3255
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2453
      Picture         =   "frmSelectCharacter.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Frame frmCharacters 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   600
      TabIndex        =   2
      Top             =   3060
      Width           =   6255
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Mr. Green"
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
         Height          =   615
         Index           =   0
         Left            =   713
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Miss Scarlet"
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
         Height          =   615
         Index           =   4
         Left            =   2633
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Mrs. Peacock"
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
         Height          =   615
         Index           =   5
         Left            =   4613
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Professor Plum"
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
         Height          =   615
         Index           =   2
         Left            =   4613
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Mrs. White"
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
         Height          =   615
         Index           =   3
         Left            =   713
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optCharacter 
         BackColor       =   &H00008000&
         Caption         =   "Colonel Mustard"
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
         Height          =   615
         Index           =   1
         Left            =   2633
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.Image imgCharacter 
         Height          =   480
         Index           =   1
         Left            =   3060
         Picture         =   "frmSelectCharacter.frx":1CAC
         Top             =   840
         Width           =   480
      End
      Begin VB.Image imgCharacter 
         Height          =   480
         Index           =   2
         Left            =   5040
         Picture         =   "frmSelectCharacter.frx":22C6
         Top             =   840
         Width           =   480
      End
      Begin VB.Image imgCharacter 
         Height          =   480
         Index           =   3
         Left            =   1140
         Picture         =   "frmSelectCharacter.frx":28B8
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image imgCharacter 
         Height          =   480
         Index           =   4
         Left            =   3060
         Picture         =   "frmSelectCharacter.frx":2CBC
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image imgCharacter 
         Height          =   480
         Index           =   5
         Left            =   5040
         Picture         =   "frmSelectCharacter.frx":32CD
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image imgCharacter 
         Height          =   480
         Index           =   0
         Left            =   1140
         Picture         =   "frmSelectCharacter.frx":38F1
         Top             =   840
         Width           =   480
      End
   End
   Begin MCI.MMControl mmcSelectCharacter 
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1085
      _Version        =   393216
      DeviceType      =   "Wave Audio"
      FileName        =   "C:\Documents and Settings\Owner\My Documents\Vbasic\Clueless\scream.wav"
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   113
      TabIndex        =   11
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your character"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   -367
      TabIndex        =   10
      Top             =   1920
      Width           =   8415
   End
End
Attribute VB_Name = "frmSelectCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Dim iRoom, iPerp, iWeapon As Integer
Dim bCharacterSelected As String
Dim iCount As Integer

Dim SelectedCharacters(5) As Integer

Private Sub cmdContinue_Click()
Dim i As Integer
Dim x As Integer

    
'check to see if a character is selected
For i = 0 To 5
    If optCharacter(i).Value = True Then
        For x = 0 To 5
            If optCharacter(i).Index <> SelectedCharacters(x) Then
                bCharacterSelected = True
            Else
                MsgBox "That character has already been selected"
                bCharacterSelected = False
                Exit Sub
            End If
        Next x
    End If
Next i

'if there is not a character selected show error message
If bCharacterSelected = False Then
    MsgBox "You need to make a selection.", vbInformation, "Clueless"
    Exit Sub
ElseIf txtPlayerName.Text = "" Then
    MsgBox "Please enter your name"
    Exit Sub
End If

Dim z As Integer
Dim y As Integer
For z = 0 To 5
    If optCharacter(z).Value = True Then
        For x = 0 To 5
            If SelectedCharacters(iCount) = -1 Then
                SelectedCharacters(iCount) = optCharacter(z).Index
                Exit For
            End If
        Next
    End If
Next z

If iCount <= iNumberofPlayers - 1 Then
    players(iCount).Init SelectedCharacters(iCount), txtPlayerName.Text
    iCount = iCount + 1
    If iNumberofPlayers <> iCount Then
        ClearForm
        Exit Sub
    End If
End If

'initialize the mmc control to play the scream.wav file when the user clicks continue
Unload frmStartup
SelectItems

frmMain.Show
Unload frmSelectCharacter

With mmcSelectCharacter
    .DeviceType = "WaveAudio"
    .FileName = App.Path & "\media\scream.wav"
    .Command = "Open"
    .Command = "Play"
End With

If MsgBox("Someone has been murdered!" & vbCrLf & "Can you solve it?", vbYesNo + vbExclamation, "Uh Oh!") = vbNo Then
    MsgBox "Don't worry we tracked down the murderer. " & vbCrLf & "There was a $100,000,000 reward for his capture." & vbCrLf & "I guess you should have tried.", vbInformation, "Murder Solved"
    frmMain.Visible = False
    CloseProgram
End If
    


End Sub

Private Sub Form_Load()
Dim iCnt As Integer
For iCnt = 0 To 5
    SelectedCharacters(iCnt) = -1
Next iCnt
iCount = 0
Label1.Caption = "Please select Player " & iCount + 1 & "'s character"
lblPlayerName.Caption = "Please enter Player " & iCount + 1 & "'s name"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload frmStartup
End Sub

Private Sub optCharacter_Click(Index As Integer)
'Set the player's character equal to the option button selected
Select Case Index
    Case 0
        sCharacter = "Colonel Mustard"
    Case 1
        sCharacter = "Mrs. White"
    Case 2
        sCharacter = "Professor Plum"
    Case 3
        sCharacter = "Mrs. Peacock"
    Case 4
        sCharacter = "Miss Scarlet"
    Case 5
        sCharacter = "Mr. Green"
End Select
End Sub

Private Sub ClearForm()
txtPlayerName.Text = ""

Dim i As Integer
For i = 0 To 5
    If SelectedCharacters(i) <> -1 Then
        optCharacter(SelectedCharacters(i)).Visible = False
        imgCharacter(SelectedCharacters(i)).Visible = False
    End If
    optCharacter(i).Value = False
Next i

Label1.Caption = "Please select Player " & iCount + 1 & "'s character"
lblPlayerName.Caption = "Please enter Player " & iCount + 1 & "'s name"

End Sub

Public Sub SelectItems()

'Select Random perp, room, weapon
Randomize
iRoom = CInt((7 * Rnd))
iPerp = CInt((5 * Rnd))
iWeapon = CInt((5 * Rnd))

'Set the room based on the random value
Select Case iRoom
    Case 1
        sRoom = "Kitchen"
    Case 2
        sRoom = "Library"
    Case 3
        sRoom = "Hall"
    Case 4
        sRoom = "Study"
    Case 5
        sRoom = "Conservatory"
    Case 6
        sRoom = "Ballroom"
    Case 7
        sRoom = "Billiard Room"
    Case 8
        sRoom = "Dining Room"
End Select

'Set the perp based on the random value
Select Case iPerp
    Case 0
        sPerp = "Mr. Green"
    Case 1
        sPerp = "Colonel Mustard"
    Case 2
        sPerp = "Professor Plum"
    Case 3
        sPerp = "Mrs. White"
    Case 4
        sPerp = "Miss Scarlet"
    Case 5
        sPerp = "Mrs. Peacock"
End Select


'Set the weapon based on the random value
Select Case iWeapon
    Case 0
        sWeapon = "Candlestick"
    Case 1
        sWeapon = "Knife"
    Case 2
        sWeapon = "Revolver"
    Case 3
        sWeapon = "Rope"
    Case 4
        sWeapon = "Lead Pipe"
    Case 5
        sWeapon = "Wrench"
End Select

End Sub
