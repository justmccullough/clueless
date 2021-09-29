VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllScores 
   BackColor       =   &H00008000&
   Caption         =   "View All Scores"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4080
   Icon            =   "frmAllScores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
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
      Left            =   780
      Picture         =   "frmAllScores.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2295
   End
   Begin MSComctlLib.ListView lstAllScores 
      Height          =   5895
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmAllScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Dim rsScores As New ADODB.Recordset


Private Sub cmdQuit_Click()
CloseProgram
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim li As ListItem
If GetRS(sScoresSQL, rsScores) = False Then
    MsgBox "The database could not be accessed."
End If

Dim records As Integer
records = CInt(rsScores.RecordCount)

Dim playernames(100) As String
Dim playerscores(100) As Integer
i = 0
While rsScores.EOF = False
    playernames(i) = rsScores.Fields("UserName")
    playerscores(i) = rsScores.Fields("Score")
    rsScores.MoveNext
    i = i + 1
Wend

Dim x, y As Integer
Dim tempstring As String
Dim tempint As Integer
For x = 0 To records
    For y = 0 To records
        If CInt(playerscores(y + 1)) > CInt(playerscores(y)) Then
            If playernames(y + 1) = "" Then GoTo here
            tempstring = playernames(y)
            tempint = playerscores(y)
            playernames(y) = playernames(y + 1)
            playerscores(y) = playerscores(y + 1)
            playernames(y + 1) = tempstring
            playerscores(y + 1) = tempint
here:
        End If
    Next
Next
lstAllScores.ColumnHeaders.Add , , "Player"
lstAllScores.ColumnHeaders.Add , , "Score"

For x = 0 To records
    If playernames(x) = "" Then GoTo skip
    Set li = lstAllScores.ListItems.Add(, , playernames(x))
    li.Text = playernames(x)
    li.SubItems(1) = playerscores(x)
skip:
Next

If rsScores.State = adStateOpen Then
    rsScores.Close
End If
Set rsScores = Nothing

End Sub
