VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   Caption         =   "Clueless"
   ClientHeight    =   10530
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   14580
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   10530
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdRoll 
      BackColor       =   &H00008000&
      Default         =   -1  'True
      Height          =   735
      Left            =   2760
      Picture         =   "frmMain.frx":190C
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2295
   End
   Begin VB.TextBox txtSetFocus 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Text            =   "Set Focus Here"
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox txtPlayerName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame fraScoreCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   9570
      TabIndex        =   100
      Top             =   1200
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtPlayerScorecard 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
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
         Height          =   855
         Left            =   0
         TabIndex        =   125
         Top             =   0
         Width           =   4695
      End
      Begin VB.Frame fraSuspects 
         BackColor       =   &H00008000&
         Caption         =   "Suspects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   60
         TabIndex        =   117
         Top             =   840
         Width           =   4455
         Begin VB.CheckBox chkSuspect 
            Appearance      =   0  'Flat
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
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox chkSuspect 
            Appearance      =   0  'Flat
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
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CheckBox chkSuspect 
            Appearance      =   0  'Flat
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
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkSuspect 
            Appearance      =   0  'Flat
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
            Height          =   495
            Index           =   5
            Left            =   2160
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox chkSuspect 
            Appearance      =   0  'Flat
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
            Height          =   495
            Index           =   4
            Left            =   2160
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox chkSuspect 
            Appearance      =   0  'Flat
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
            Height          =   495
            Index           =   3
            Left            =   2160
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraWeapons 
         BackColor       =   &H00008000&
         Caption         =   "Weapons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   60
         TabIndex        =   110
         Top             =   3240
         Width           =   4455
         Begin VB.CheckBox chkWeapon 
            Appearance      =   0  'Flat
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
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkWeapon 
            Appearance      =   0  'Flat
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
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkWeapon 
            Appearance      =   0  'Flat
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
            Height          =   375
            Index           =   5
            Left            =   2280
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkWeapon 
            Appearance      =   0  'Flat
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
            Height          =   375
            Index           =   4
            Left            =   360
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkWeapon 
            Appearance      =   0  'Flat
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
            Height          =   375
            Index           =   2
            Left            =   360
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkWeapon 
            Appearance      =   0  'Flat
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
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame fraRooms 
         BackColor       =   &H00008000&
         Caption         =   "Rooms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   120
         TabIndex        =   101
         Top             =   5640
         Width           =   4455
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Dining Room"
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
            Height          =   495
            Index           =   7
            Left            =   2400
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Billiard Room"
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
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Ballroom"
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
            Height          =   495
            Index           =   5
            Left            =   2400
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Study"
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
            Height          =   495
            Index           =   3
            Left            =   2400
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Library"
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
            Height          =   495
            Index           =   1
            Left            =   2400
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Conservatory"
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
            Height          =   495
            Index           =   4
            Left            =   240
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Hall"
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
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkRoom 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            Caption         =   "Kitchen"
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
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdViewCard 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   10800
      Picture         =   "frmMain.frx":2817
      Style           =   1  'Graphical
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   360
      Width           =   2235
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Picture         =   "frmMain.frx":36BD
      Style           =   1  'Graphical
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton cmdDown 
      Height          =   735
      Left            =   7320
      Picture         =   "frmMain.frx":44DB
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   735
      Left            =   5520
      Picture         =   "frmMain.frx":5305
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdRight 
      Height          =   735
      Left            =   7320
      Picture         =   "frmMain.frx":6099
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdUp 
      Height          =   735
      Left            =   5520
      Picture         =   "frmMain.frx":6EDF
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.Image imgMainTitle 
      Height          =   1455
      Left            =   4320
      Picture         =   "frmMain.frx":7A72
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Image Image12 
      Height          =   525
      Left            =   2400
      Picture         =   "frmMain.frx":101F2
      Top             =   5880
      Width           =   450
   End
   Begin VB.Label lblCurrentPlayer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Player"
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
      Height          =   495
      Left            =   2640
      TabIndex        =   126
      Top             =   360
      Width           =   2535
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpMain 
      Height          =   495
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      URL             =   "C:\Documents and Settings\Owner\My Documents\Vbasic\Clueless\slow.mid"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
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
      _cx             =   2990
      _cy             =   873
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Height          =   855
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblAnswers 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1440
      TabIndex        =   98
      Top             =   7800
      Width           =   375
   End
   Begin VB.Image imgDice 
      Height          =   2250
      Index           =   5
      Left            =   240
      Picture         =   "frmMain.frx":10713
      Top             =   8160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgDice 
      Height          =   2250
      Index           =   4
      Left            =   240
      Picture         =   "frmMain.frx":10A75
      Top             =   8160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgDice 
      Height          =   2250
      Index           =   3
      Left            =   240
      Picture         =   "frmMain.frx":10D98
      Top             =   8160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgDice 
      Height          =   2250
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":11068
      Top             =   8160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgDice 
      Height          =   2250
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":112DE
      Top             =   8160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgCharacter 
      Height          =   480
      Index           =   1
      Left            =   3360
      Picture         =   "frmMain.frx":11510
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image26 
      Height          =   525
      Left            =   5280
      Picture         =   "frmMain.frx":11B2A
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image Image25 
      Height          =   525
      Left            =   4800
      Picture         =   "frmMain.frx":1204B
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image Image24 
      Height          =   525
      Left            =   6720
      Picture         =   "frmMain.frx":1256C
      Top             =   4440
      Width           =   450
   End
   Begin VB.Image Image23 
      Height          =   525
      Left            =   6720
      Picture         =   "frmMain.frx":12A8D
      Top             =   2520
      Width           =   450
   End
   Begin VB.Image Image22 
      Height          =   525
      Left            =   8160
      Picture         =   "frmMain.frx":12FAE
      Top             =   5880
      Width           =   450
   End
   Begin VB.Image Image21 
      Height          =   525
      Left            =   5280
      Picture         =   "frmMain.frx":134CF
      Top             =   5880
      Width           =   450
   End
   Begin VB.Image Image20 
      Height          =   525
      Left            =   4800
      Picture         =   "frmMain.frx":139F0
      Top             =   5880
      Width           =   450
   End
   Begin VB.Image Image14 
      Height          =   525
      Left            =   2400
      Picture         =   "frmMain.frx":13F11
      Top             =   3480
      Width           =   450
   End
   Begin VB.Image Image13 
      Height          =   525
      Left            =   3360
      Picture         =   "frmMain.frx":14432
      Top             =   6360
      Width           =   450
   End
   Begin VB.Image Image11 
      Height          =   525
      Left            =   3360
      Picture         =   "frmMain.frx":14953
      Top             =   2520
      Width           =   450
   End
   Begin VB.Image imgCharacter 
      Height          =   480
      Index           =   5
      Left            =   6720
      Picture         =   "frmMain.frx":14E74
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharacter 
      Height          =   480
      Index           =   4
      Left            =   3360
      Picture         =   "frmMain.frx":15498
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharacter 
      Height          =   480
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":15AA9
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharacter 
      Height          =   480
      Index           =   2
      Left            =   960
      Picture         =   "frmMain.frx":15EAD
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCharacter 
      Height          =   480
      Index           =   0
      Left            =   6720
      Picture         =   "frmMain.frx":1649F
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label117 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      OLEDropMode     =   1  'Manual
      TabIndex        =   95
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label108 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   94
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label107 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   93
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label106 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   4800
      OLEDropMode     =   1  'Manual
      TabIndex        =   92
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label104 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      OLEDropMode     =   1  'Manual
      TabIndex        =   91
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label103 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   90
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label101 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   89
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label100 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      OLEDropMode     =   1  'Manual
      TabIndex        =   88
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label99 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   87
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label98 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      OLEDropMode     =   1  'Manual
      TabIndex        =   86
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label95 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   85
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label94 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   84
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label93 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   4800
      OLEDropMode     =   1  'Manual
      TabIndex        =   83
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label92 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   82
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label91 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   81
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label90 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   80
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label89 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   79
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label88 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   78
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label87 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   77
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label85 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   76
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label84 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   75
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label83 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   74
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label82 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   73
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label81 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   72
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label80 
      BackColor       =   &H80000007&
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   71
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label79 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   70
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label78 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   69
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label77 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   68
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label76 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   67
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label75 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   66
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label74 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   8640
      OLEDropMode     =   1  'Manual
      TabIndex        =   65
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label73 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8640
      OLEDropMode     =   1  'Manual
      TabIndex        =   64
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label72 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      OLEDropMode     =   1  'Manual
      TabIndex        =   63
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label70 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   7680
      OLEDropMode     =   1  'Manual
      TabIndex        =   62
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label69 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      OLEDropMode     =   1  'Manual
      TabIndex        =   61
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label68 
      BackColor       =   &H80000009&
      Height          =   15
      Left            =   7080
      TabIndex        =   60
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label67 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   7200
      OLEDropMode     =   1  'Manual
      TabIndex        =   59
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label66 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   58
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label65 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   57
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label64 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   7680
      OLEDropMode     =   1  'Manual
      TabIndex        =   56
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label63 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      OLEDropMode     =   1  'Manual
      TabIndex        =   55
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label62 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   7200
      OLEDropMode     =   1  'Manual
      TabIndex        =   54
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label61 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      OLEDropMode     =   1  'Manual
      TabIndex        =   53
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label60 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   8160
      OLEDropMode     =   1  'Manual
      TabIndex        =   52
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label59 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8640
      OLEDropMode     =   1  'Manual
      TabIndex        =   51
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label58 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   8640
      OLEDropMode     =   1  'Manual
      TabIndex        =   50
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label57 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8160
      OLEDropMode     =   1  'Manual
      TabIndex        =   49
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label56 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   48
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label55 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   47
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label51 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   46
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label50 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6720
      OLEDropMode     =   1  'Manual
      TabIndex        =   45
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label49 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   6240
      OLEDropMode     =   1  'Manual
      TabIndex        =   44
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label44 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   43
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   41
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label40 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   39
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label38 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   38
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label37 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   37
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   36
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label35 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   35
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label34 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   34
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   33
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label24 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   25
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      OLEDropMode     =   1  'Manual
      TabIndex        =   24
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   23
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   2400
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   1440
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgDice 
      Height          =   2250
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":16AAD
      Top             =   8160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label43 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   42
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   40
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label31 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   32
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   31
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label29 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   30
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label28 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   29
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   28
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label26 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   27
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   19
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      OLEDropMode     =   1  'Manual
      TabIndex        =   22
      Top             =   3480
      Width           =   495
   End
   Begin VB.Image imgStart 
      Height          =   480
      Index           =   0
      Left            =   6360
      Picture         =   "frmMain.frx":16C8E
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgStart 
      Height          =   480
      Index           =   1
      Left            =   3360
      Picture         =   "frmMain.frx":16EDD
      Top             =   1080
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgStart 
      Height          =   960
      Index           =   2
      Left            =   960
      Picture         =   "frmMain.frx":1712C
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   960
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":175CF
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   480
      Index           =   4
      Left            =   3360
      Picture         =   "frmMain.frx":17A72
      Top             =   7800
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgStart 
      Height          =   480
      Index           =   5
      Left            =   6240
      Picture         =   "frmMain.frx":17CC1
      Top             =   7800
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgBillardsRoom 
      Height          =   1455
      Left            =   1440
      Picture         =   "frmMain.frx":17F10
      Top             =   3960
      Width           =   1920
   End
   Begin VB.Image imgHall 
      Height          =   1455
      Left            =   4320
      Picture         =   "frmMain.frx":198B8
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Image imgKitchen 
      Height          =   1455
      Left            =   1440
      Picture         =   "frmMain.frx":1B2DE
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Image imgDiningRoom 
      Height          =   1455
      Left            =   1440
      Picture         =   "frmMain.frx":1CEEE
      Top             =   6360
      Width           =   1920
   End
   Begin VB.Image imgDen 
      Height          =   1455
      Left            =   7200
      Picture         =   "frmMain.frx":1E708
      Top             =   6360
      Width           =   1920
   End
   Begin VB.Image imgLibrary 
      Height          =   1455
      Left            =   7200
      Picture         =   "frmMain.frx":20311
      Top             =   3960
      Width           =   1920
   End
   Begin VB.Image imgGreenhouse 
      Height          =   1455
      Left            =   7200
      Picture         =   "frmMain.frx":21974
      Top             =   1560
      Width           =   1920
   End
   Begin VB.Image imgBallRoom 
      Height          =   1455
      Left            =   4320
      Picture         =   "frmMain.frx":23886
      Top             =   6360
      Width           =   1920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Dim iCurrentx, iCurrenty As Integer
Dim bStart, bAllowedMove, bInRoom As Boolean
Dim sErrorMessage As String
Dim iRandomNum, iRoll As Integer
Dim iKeyPressed As Integer

Private Sub cmdDown_Click()
If players(iCurrentPlayer).Moves <= 0 Then
    MsgBox "You are out of moves.", vbInformation, "Clueless"
    Exit Sub
Else
    CheckInRoomMove "Down"
    CheckMove "Down"
    ShowErrorMessage sErrorMessage
    MoveCharacter "Down"
    If bAllowedMove = True Then
        players(iCurrentPlayer).Moves = players(iCurrentPlayer).Moves - 1
        If players(iCurrentPlayer).Moves = 0 Then
            SwitchPlayer
        End If
    End If
End If
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub cmdLeft_Click()
If players(iCurrentPlayer).Moves <= 0 Then
    MsgBox "You are out of moves.", vbInformation, "Clueless"
    Exit Sub
Else
    CheckInRoomMove "Left"
    CheckMove "Left"
    ShowErrorMessage sErrorMessage
    MoveCharacter "Left"
    If bAllowedMove = True Then
        players(iCurrentPlayer).Moves = players(iCurrentPlayer).Moves - 1
        If players(iCurrentPlayer).Moves = 0 Then
            SwitchPlayer
        End If
    End If
End If
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdRight_Click()
If players(iCurrentPlayer).Moves <= 0 Then
    MsgBox "You are out of moves.", vbInformation, "Clueless"
    Exit Sub
Else
    CheckInRoomMove "Right"
    CheckMove "Right"
    ShowErrorMessage sErrorMessage
    MoveCharacter "Right"
    If bAllowedMove = True Then
        players(iCurrentPlayer).Moves = players(iCurrentPlayer).Moves - 1
        If players(iCurrentPlayer).Moves = 0 Then
            SwitchPlayer
        End If
    End If
End If
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub cmdRoll_Click()
Randomize
iRandomNum = CInt((5 * Rnd))
Timer1.Interval = 100
players(iCurrentPlayer).Moves = iRandomNum + 1
players(iCurrentPlayer).Score = players(iCurrentPlayer).Score - 50
Timer1.Enabled = True
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub cmdUp_Click()
If players(iCurrentPlayer).Moves <= 0 Then
    MsgBox "You are out of moves.", vbInformation, "Clueless"
    Exit Sub
Else
    CheckInRoomMove "Up"
    CheckMove "Up"
    ShowErrorMessage sErrorMessage
    MoveCharacter "Up"
    If bAllowedMove = True Then
        players(iCurrentPlayer).Moves = players(iCurrentPlayer).Moves - 1
        If players(iCurrentPlayer).Moves = 0 Then
            SwitchPlayer
        End If
    End If
End If
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub cmdViewCard_Click()
If fraScoreCard.Visible = True Then
    fraScoreCard.Visible = False
Else
    fraScoreCard.Visible = True
End If
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub Form_Load()
bStart = True
Timer1.Enabled = False
wmpMain.URL = App.Path & "\media\slow.mid"

Dim i As Integer
For i = 0 To iNumberofPlayers - 1
    imgCharacter(players(i).CharacterIndex).Visible = True
    imgStart(players(i).CharacterIndex).Visible = True
Next

iCurrentPlayer = 0
iCurrentIndex = players(iCurrentPlayer).CharacterIndex
txtPlayerScorecard.Text = players(iCurrentPlayer).Name & "'s ScoreCard"
txtPlayerName.Text = players(iCurrentPlayer).Name
LoadScoreCard
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmMain.Visible = False Then Exit Sub
If MsgBox("Are you sure you want to exit?", vbYesNo + vbInformation, "Exit") = vbYes Then
    MsgBox "Don't worry we tracked down the murderer. " & vbCrLf & "There was a $100,000,000 reward for his capture." & vbCrLf & "I guess you should have tried harder.", vbInformation, "Murder Solved"
    CloseProgram
Else
    Cancel = True
End If
End Sub

Private Sub lblAnswers_Click()
MsgBox "Room: " & sRoom & vbCrLf & "Weapon: " & sWeapon & vbCrLf & "Perp: " & sPerp & vbCrLf & "You cheater!", vbInformation, sTitle
players(iCurrentPlayer).Score = players(iCurrentIndex).Score - 1000
players(iCurrentPlayer).Moves = 100
If frmMain.Visible = True Then
    txtSetFocus.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

iRoll = iRoll + 1

For i = 0 To 5
    imgDice(i).Visible = False
Next

Select Case iRoll
    Case 1
        imgDice(0).Visible = True
    Case 2
        imgDice(3).Visible = True
    Case 3
        imgDice(2).Visible = True
    Case 4
        imgDice(5).Visible = True
    Case 5
        imgDice(4).Visible = True
    Case 6
        Timer1.Interval = 200
        imgDice(2).Visible = True
    Case 7
        Timer1.Interval = 250
        imgDice(1).Visible = True
    Case 8
        Timer1.Interval = 300
        imgDice(5).Visible = True
    Case 9
        Timer1.Interval = 350
        imgDice(3).Visible = True
    Case 10
        Timer1.Interval = 400
        imgDice(4).Visible = True
    Case 11
        Timer1.Interval = 450
        imgDice(2).Visible = True
    Case 12
        Timer1.Interval = 500
        imgDice(1).Visible = True
    Case 13
        Timer1.Enabled = False
        iRoll = 0
        ShowDice iRandomNum
End Select
End Sub

Private Sub CheckMove(ByVal sDirection As String)

If bAllowedMove = False Then Exit Sub

GetCurrentLocation
sErrorMessage = ""

'this code is for when the right button is pressed
'if we are at the start most of the players cannot move right
If sDirection = "Right" Then
    If players(iCurrentPlayer).CharacterIndex <> 2 And players(iCurrentPlayer).CharacterIndex <> 3 And players(iCurrentPlayer).AtStart = True Then
        sErrorMessage = "Right"
        bAllowedMove = False
        Exit Sub
    End If
players(iCurrentPlayer).AtStart = False
    If iCurrentx = 3840 Then
        If iCurrenty = 3000 Or iCurrenty = 3480 Or iCurrenty = 5400 Or iCurrenty = 5880 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        ElseIf iCurrenty = 3960 Or iCurrenty = 4440 Or iCurrenty = 4920 Then
            sErrorMessage = "Center"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If

    ElseIf iCurrentx = 6720 Then
        If iCurrenty = 3000 Or iCurrenty = 3480 Or iCurrenty = 5400 Or iCurrenty = 5880 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        ElseIf iCurrenty = 4440 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Library"
            Exit Sub
        ElseIf iCurrenty = 2520 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Conservatory"
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If
        
    ElseIf iCurrentx = 8640 Then
        sErrorMessage = "Right"
        bAllowedMove = False
        bInRoom = False
        Exit Sub
    Else
        bAllowedMove = True
        bInRoom = False
        Exit Sub
    End If
    
'this part is for checking the move when the left button is pressed
'if we are at the start none of the players can move left
ElseIf sDirection = "Left" Then
    If players(iCurrentPlayer).AtStart = True Then
        sErrorMessage = "Left"
        bAllowedMove = False
        bInRoom = False
        Exit Sub
    End If
    
players(iCurrentPlayer).AtStart = False

    If iCurrentx = 1440 Then
        sErrorMessage = "Left"
        bAllowedMove = False
        bInRoom = False
        Exit Sub
    ElseIf iCurrentx = 3360 Then
        If iCurrenty = 3000 Or iCurrenty = 3480 Or iCurrenty = 5400 Or iCurrenty = 5880 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        ElseIf iCurrenty = 2520 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Kitchen"
            Exit Sub
        ElseIf iCurrenty = 6360 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Dining Room"
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If
        
    ElseIf iCurrentx = 6240 Then
        If iCurrenty = 3000 Or iCurrenty = 3480 Or iCurrenty = 5400 Or iCurrenty = 5880 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        ElseIf iCurrenty = 3960 Or iCurrenty = 4440 Or iCurrenty = 4920 Then
            sErrorMessage = "Center"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If
    Else
        bAllowedMove = True
        bInRoom = False
        Exit Sub
    End If
    
'this portion of code is for when the user selects the up button
'at the start only miss scarlet and mrs. peacock can move up
ElseIf sDirection = "Up" Then
    If players(iCurrentPlayer).CharacterIndex <> 4 And players(iCurrentPlayer).CharacterIndex <> 5 And players(iCurrentPlayer).AtStart = True Then
        sErrorMessage = "Up"
        bAllowedMove = False
        bInRoom = False
        Exit Sub
    End If
    
players(iCurrentPlayer).AtStart = False

    If iCurrenty = 1560 Then
        sErrorMessage = "Up"
        bAllowedMove = False
        bInRoom = False
        Exit Sub
    ElseIf iCurrenty = 3000 Then
        If iCurrentx = 6720 Or iCurrentx = 6240 Or iCurrentx = 3840 Or iCurrentx = 3360 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        ElseIf iCurrentx = 4800 Or iCurrentx = 5280 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Hall"
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If
    ElseIf iCurrenty = 5400 Then
        If iCurrentx = 6720 Or iCurrentx = 6240 Or iCurrentx = 3840 Or iCurrentx = 3360 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        ElseIf iCurrentx = 4320 Or iCurrentx = 4800 Or iCurrentx = 5280 Or iCurrentx = 5760 Then
            sErrorMessage = "Center"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If
    Else
        bAllowedMove = True
        bInRoom = False
        Exit Sub
    End If
ElseIf sDirection = "Down" Then
    If players(iCurrentPlayer).CharacterIndex <> 1 And players(iCurrentPlayer).CharacterIndex <> 0 And players(iCurrentPlayer).AtStart = True Then
        sErrorMessage = "Down"
        bAllowedMove = False
        bInRoom = False
        Exit Sub
    End If
    
players(iCurrentPlayer).AtStart = False

    If iCurrenty = 3480 Then
        If iCurrentx = 2400 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Billiard Room"
            Exit Sub
        ElseIf iCurrentx = 4320 Or iCurrentx = 4800 Or iCurrentx = 5280 Or iCurrentx = 5760 Then
            sErrorMessage = "Center"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        ElseIf iCurrentx = 3360 Or iCurrentx = 3840 Or iCurrentx = 6240 Or iCurrentx = 6720 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        Else
            sErrorMessage = "Door"
            bAllowedMove = False
            bInRoom = False
            Exit Sub
        End If
    ElseIf iCurrenty = 5880 Then
        If iCurrentx = 2400 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Dining Room"
            Exit Sub
        ElseIf iCurrentx = 4800 Or iCurrentx = 5280 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Ballroom"
            Exit Sub
        ElseIf iCurrentx = 8160 Then
            bAllowedMove = True
            bInRoom = True
            sCurrentRoom = "Study"
            Exit Sub
        ElseIf iCurrentx = 3360 Or iCurrentx = 3840 Or iCurrentx = 6240 Or iCurrentx = 6720 Then
            bAllowedMove = True
            bInRoom = False
            Exit Sub
        Else
            bAllowedMove = False
            sErrorMessage = "Door"
            bInRoom = False
            Exit Sub
        End If
    Else
        bAllowedMove = True
        bInRoom = False
        Exit Sub
    End If
End If
End Sub

Private Sub GetCurrentLocation()

'this sets the current x and current y equal to the top and left of the visible image
    iCurrentx = imgCharacter(iCurrentIndex).Left
    iCurrenty = imgCharacter(iCurrentIndex).Top
End Sub

Private Sub MoveCharacter(ByVal sDirection As String)

'this performs the actual moving of the visible pieces

If bAllowedMove = True Then
    
    If sDirection = "Right" Then
        imgCharacter(iCurrentIndex).Left = imgCharacter(iCurrentIndex).Left + 480
    ElseIf sDirection = "Left" Then
        imgCharacter(iCurrentIndex).Left = imgCharacter(iCurrentIndex).Left - 480
    ElseIf sDirection = "Down" Then
        imgCharacter(iCurrentIndex).Top = imgCharacter(iCurrentIndex).Top + 480
    ElseIf sDirection = "Up" Then
        imgCharacter(iCurrentIndex).Top = imgCharacter(iCurrentIndex).Top - 480
    End If
End If

If bInRoom = True Then
    dlgGuess.Show vbModal
End If
End Sub

Private Sub ShowErrorMessage(ByVal sError As String)

    If sError = "" Then Exit Sub
    
    If sError = "Door" Then
        MsgBox "You must use the door to enter this room", vbInformation, sTitle
    ElseIf sError = "Up" Or sError = "Down" Or sError = "Right" Or sError = "Left" Then
        MsgBox "You cannot go " & sError & " any further", vbInformation, sTitle
    ElseIf sError = "Center" Then
        MsgBox "You cannot enter the center of the board", vbInformation
    End If
End Sub

Private Sub CheckInRoomMove(ByVal sDirection As String)
' this function checks to see if the character is in the room and if the move is valid
bAllowedMove = True
    If bInRoom = False Then Exit Sub
    
    If sCurrentRoom = "Kitchen" Then
        If sDirection <> "Right" Then
            MsgBox "You must move right to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Hall" Then
        If sDirection <> "Down" Then
            MsgBox "You must move down to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Conservatory" Then
        If sDirection <> "Left" Then
            MsgBox "You must move left to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Billiard Room" Then
        If sDirection <> "Up" Then
            MsgBox "You must move up to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Library" Then
        If sDirection <> "Left" Then
            MsgBox "You must move left to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Dining Room" Then
        If sDirection <> "Up" And sDirection <> "Right" Then
            MsgBox "You must move either up or right to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Ballroom" Then
        If sDirection <> "Up" Then
            MsgBox "You must move up to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    ElseIf sCurrentRoom = "Study" Then
        If sDirection <> "Up" Then
            MsgBox "You must move up to exit the room", vbInformation, sTitle
            bAllowedMove = False
        End If
    End If
End Sub

Public Sub SetMoves(ByVal Moves As Integer)
players(iCurrentPlayer).Moves = Moves
End Sub

Public Function GetMoves() As Integer
GetMoves = players(iCurrentPlayer).Moves
End Function

Private Sub ShowDice(ByVal iRandomNum As Integer)

Select Case iRandomNum
    Case 0
        imgDice(0).Visible = True
    Case 1
        imgDice(1).Visible = True
    Case 2
        imgDice(2).Visible = True
    Case 3
        imgDice(3).Visible = True
    Case 4
        imgDice(4).Visible = True
    Case 5
        imgDice(5).Visible = True
End Select
End Sub

Private Sub GetPlayerArrays()
    Rooms = Split(players(iCurrentPlayer).Rooms, ";")
    Weapons = Split(players(iCurrentPlayer).Weapons, ";")
    Suspects = Split(players(iCurrentPlayer).Suspects, ";")
End Sub

Private Sub LoadScoreCard()
Dim i, x As Integer
ClearScoreCard
    GetPlayerArrays
    For i = 0 To UBound(Rooms)
        If Rooms(i) = "" Then Exit For
        For x = 0 To 7
            If frmMain.chkRoom(x).Caption = Rooms(i) Then
                frmMain.chkRoom(x).Value = vbChecked
            End If
        Next
    Next
    
    For i = 0 To UBound(Suspects)
        If Suspects(i) = "" Then Exit For
        For x = 0 To 5
            If frmMain.chkSuspect(x).Caption = Suspects(i) Then
                frmMain.chkSuspect(x).Value = vbChecked
            End If
        Next
    Next
    
    For i = 0 To UBound(Weapons)
        If Weapons(i) = "" Then Exit For
        For x = 0 To 5
            If frmMain.chkWeapon(x).Caption = Weapons(i) Then
                frmMain.chkWeapon(x).Value = vbChecked
            End If
        Next
    Next
End Sub

Public Sub AddToString(ByVal sString As String, ByVal iString As Integer)
'this function needs an iArray parameter
'0 = Rooms, 1=Suspects, 2=Weapons
GetStrings
Select Case iString
    Case 0
        If sRooms = "" Then
            sRooms = sString
        Else
            sRooms = sRooms & ";" & sString
        End If
    Case 1
        If sSuspects = "" Then
            sSuspects = sString
        Else
            sSuspects = sSuspects & ";" & sString
        End If
    Case 2
        If sWeapons = "" Then
            sWeapons = sString
        Else
            sWeapons = sWeapons & ";" & sString
        End If
End Select
SetStrings
End Sub

Private Sub GetStrings()
sRooms = players(iCurrentPlayer).Rooms
sSuspects = players(iCurrentPlayer).Suspects
sWeapons = players(iCurrentPlayer).Weapons
End Sub

Private Sub SetStrings()
players(iCurrentPlayer).Rooms = sRooms
players(iCurrentPlayer).Suspects = sSuspects
players(iCurrentPlayer).Weapons = sWeapons
End Sub

Public Sub SwitchPlayer()
If iCurrentPlayer = iNumberofPlayers - 1 Then
    iCurrentPlayer = 0
Else
    SetPlayerInfo
    iCurrentPlayer = iCurrentPlayer + 1
End If
iCurrentIndex = players(iCurrentPlayer).CharacterIndex
txtPlayerScorecard.Text = players(iCurrentPlayer).Name & "'s ScoreCard"
txtPlayerName.Text = players(iCurrentPlayer).Name
If iNumberofPlayers > 1 Then
    MsgBox "It is " & players(iCurrentPlayer).Name & "'s turn."
    fraScoreCard.Visible = False
End If
GetPlayerInfo
LoadScoreCard
End Sub

Private Sub SetPlayerInfo()
players(iCurrentPlayer).InRoom = bInRoom
SetStrings
End Sub

Private Sub GetPlayerInfo()
bInRoom = players(iCurrentPlayer).InRoom
GetStrings
End Sub

Private Sub ClearScoreCard()
Dim i As Integer
For i = 0 To 7
    chkRoom(i).Value = vbUnchecked
Next

For i = 0 To 5
    chkSuspect(i).Value = vbUnchecked
Next

For i = 0 To 5
    chkWeapon(i).Value = vbUnchecked
Next

End Sub

Private Sub txtPlayerName_KeyDown(KeyCode As Integer, Shift As Integer)
iKeyPressed = KeyCode
KeyCode = 0
Select Case iKeyPressed
    Case 37
        cmdLeft_Click
    Case 38
        cmdUp_Click
    Case 39
        cmdRight_Click
    Case 40
        cmdDown_Click
End Select
End Sub

Private Sub txtSetFocus_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37
        cmdLeft_Click
    Case 38
        cmdUp_Click
    Case 39
        cmdRight_Click
    Case 40
        cmdDown_Click
End Select
End Sub
