VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmStartup 
   BackColor       =   &H00008000&
   Caption         =   " Welcome to Clueless"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10755
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10200
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboNumberofUsers 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmStartup.frx":0CCA
      Left            =   8520
      List            =   "frmStartup.frx":0CCC
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdContinue 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6360
      Picture         =   "frmStartup.frx":0CCE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How many players?"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Image imgMrsWhite 
      Height          =   1905
      Left            =   3720
      Picture         =   "frmStartup.frx":1CB0
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Image imgProfessorPlum 
      Height          =   1905
      Left            =   2040
      Picture         =   "frmStartup.frx":40A5
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Image imgMrsPeacock 
      Height          =   1905
      Left            =   360
      Picture         =   "frmStartup.frx":6418
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Image imgMissScarlet 
      Height          =   1905
      Left            =   3720
      Picture         =   "frmStartup.frx":87E2
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Image imgColonelMustard 
      Height          =   1905
      Left            =   2040
      Picture         =   "frmStartup.frx":AC03
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Image imgGreen 
      Height          =   1905
      Left            =   360
      Picture         =   "frmStartup.frx":D09B
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   3150
      Left            =   720
      Picture         =   "frmStartup.frx":F5C2
      Top             =   480
      Width           =   9450
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpStartup 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   3855
      URL             =   "C:\Documents and Settings\Owner\My Documents\Vbasic\Clueless\happy.mid"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6800
      _cy             =   873
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Private Sub cmdContinue_Click()
If cboNumberofUsers.Text = "" Or cboNumberofUsers.ListIndex = -1 Then
    MsgBox "Please select the number of players."
    Exit Sub
Else
    iNumberofPlayers = CInt(cboNumberofUsers.Text)
End If

frmSelectCharacter.Show
frmStartup.Hide
End Sub

Private Sub Form_Load()
wmpStartup.URL = App.Path & "\media\happy.mid"

Dim i As Integer
For i = 1 To 6
    cboNumberofUsers.AddItem (i)
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload frmSplash
End Sub
