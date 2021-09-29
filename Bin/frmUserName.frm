VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmWon 
   BackColor       =   &H00008000&
   Caption         =   "Please enter your name"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5940
   Icon            =   "frmUserName.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   4095
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHighScores 
      Height          =   615
      Left            =   1920
      Picture         =   "frmUserName.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpUserName 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   5175
      URL             =   "C:\Documents and Settings\Owner\My Documents\Vbasic\Clueless\celebrate.mid"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
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
      _cx             =   9128
      _cy             =   873
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Won!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
   End
End
Attribute VB_Name = "frmWon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit
Dim rsUsers As New ADODB.Recordset


Private Sub cmdHighScores_Click()

If modClueless.GetRS(sUsersSQL, rsUsers) = False Then
    MsgBox "The database could not be accessed. Please try again later"
End If

rsUsers.AddNew
rsUsers.Fields("UserName") = players(iCurrentPlayer).Name
rsUsers.Fields("Score") = players(iCurrentPlayer).Score
iUserID = rsUsers.RecordCount

If modClueless.PutRS(rsUsers) = False Then
    MsgBox "The Save has failed."
End If

Set rsUsers = Nothing

    frmWon.Hide
    frmHighScores.Show
End Sub

Private Sub Form_Load()
frmMain.wmpMain.Close
wmpUserName.URL = App.Path & "\media\celebrate.mid"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rsUsers.State = adStateOpen Then
    rsUsers.Close
End If
Set rsUsers = Nothing
CloseProgram
End Sub
