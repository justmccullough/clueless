VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H00008000&
   Caption         =   "High Scores"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9570
   Icon            =   "frmHighScores.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   5310
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAllScores 
      Height          =   735
      Left            =   6300
      Picture         =   "frmHighScores.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Frame fraHighScores 
      BackColor       =   &H00008000&
      Caption         =   "High Scores"
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
      Height          =   4455
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      Begin VB.TextBox txtHighScores 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox txtHighScores 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2940
         Width           =   1215
      End
      Begin VB.TextBox txtHighScores 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtHighScores 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txtHighScores 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblHighUsers 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label lblHighUsers 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
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
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   2940
         Width           =   2415
      End
      Begin VB.Label lblHighUsers 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblHighUsers 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
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
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label lblHighUsers 
         BackStyle       =   0  'Transparent
         Caption         =   "l"
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
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdExit 
      Height          =   735
      Left            =   6300
      Picture         =   "frmHighScores.frx":1B46
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label lblResult 
      BackStyle       =   0  'Transparent
      Caption         =   "Sorry! You did not make the top 5 all time scores."
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
      Left            =   5880
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Dim rsScores As New ADODB.Recordset

Private Sub cmdAllScores_Click()
frmAllScores.Show
frmHighScores.Hide
End Sub

Private Sub cmdExit_Click()
CloseProgram
End Sub

Private Sub Form_Load()
Dim i, x As Integer
Dim iHighlightedUser, iScores(), iTempScore, iUserIDs() As Integer
Dim sUsers(), sTempUser, sSTR As String
Dim bFoundHighlightedUser As Boolean
bFoundHighlightedUser = False
i = 0
x = 0
Unload frmMain
iHighlightedUser = 0

On Error GoTo err

If GetRS(sHighScoresSQL, rsScores) = False Then
    MsgBox "Sorry.  The database cannot be accessed right now." & vbCrLf & "Please try again later", vbInformation, "Information Access Did Not Succeed"
End If

With rsScores
ReDim iScores(.RecordCount - 1)
ReDim sUsers(.RecordCount - 1)
ReDim iUserIDs(.RecordCount - 1)

While .EOF = False
    iUserIDs(x) = .Fields("ID")
    sUsers(x) = .Fields("UserName")
    iScores(x) = .Fields("Score")
    .MoveNext
    x = x + 1
Wend
End With
For i = 0 To UBound(sUsers)
    For x = 0 To UBound(sUsers) - 1
        If CInt(iScores(x)) < CInt(iScores(x + 1)) Then
            iTempScore = iScores(x)
            sTempUser = sUsers(x)
            iScores(x) = iScores(x + 1)
            sUsers(x) = sUsers(x + 1)
            iScores(x + 1) = iTempScore
            sUsers(x + 1) = sTempUser
skip:
        End If
    Next
Next

For i = 0 To UBound(sUsers)
    If iUserIDs(i) = iUserID Then
        iHighlightedUser = i
    Else
        iHighlightedUser = -1
    End If
    lblHighUsers(i).Caption = sUsers(i)
    txtHighScores(i).Text = iScores(i)
Next

If iHighlightedUser <> -1 Then
    lblHighUsers(iHighlightedUser).ForeColor = vbRed
    txtHighScores(iHighlightedUser).ForeColor = vbRed
    lblResult.Caption = "Congratulations! You made the Top 5 all-time scores."
Else
    lblResult.Caption = "Sorry! You did not make the Top 5 all-time scores."
End If

Exit Sub
err:
    MsgBox err.Number & vbCrLf & err.Description & vbCrLf & "High Scores Form Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rsScores.State = adStateOpen Then
    rsScores.Close
    Set rsScores = Nothing
End If
CloseProgram
End Sub
