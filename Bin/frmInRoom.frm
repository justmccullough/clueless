VERSION 5.00
Begin VB.Form frmInRoom 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   ForeColor       =   &H00008000&
   LinkTopic       =   "Form4"
   ScaleHeight     =   8880
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move to the Greenhouse"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move to the Dinning Room"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move to the Den"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
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
      Left            =   1680
      Picture         =   "frmInRoom.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move to the Kitchen"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "What would you like to do?"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
End
Attribute VB_Name = "frmInRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'move to Kitchen
'Colonel Mustard
If Form2.Text1.Text = 1 Then
    Form2.Image1(0).Top = 2520
    Form2.Image1(0).Left = 2880
'Mr. Green
ElseIf Form2.Text1.Text = 2 Then
    Form2.Image1(1).Top = 2520
    Form2.Image1(1).Left = 2880
'Proffesor Plum
ElseIf Form2.Text1.Text = 3 Then
    Form2.Image1(2).Top = 2520
    Form2.Image1(2).Left = 2880
'Mrs. Peacock
ElseIf Form2.Text1.Text = 4 Then
    Form2.Image1(5).Top = 2520
    Form2.Image1(5).Left = 2880
'Miss Scarlet
ElseIf Form2.Text1.Text = 5 Then
    Form2.Image1(4).Top = 2520
    Form2.Image1(4).Left = 2880
'Mrs. White
ElseIf Form2.Text1.Text = 6 Then
    Form2.Image1(3).Top = 2520
    Form2.Image1(3).Left = 2880
End If
Command1.Visible = False
End Sub

Private Sub Command2_Click()
Form5.Visible = True
Form5.Text3.Text = Form5.Label4.Caption
Form4.Visible = False
End Sub

Private Sub Command3_Click()
'move to Den
'Colonel Mustard
If Form2.Text1.Text = 1 Then
    Form2.Image1(0).Top = 6360
    Form2.Image1(0).Left = 8160
    Form4.Visible = False
'Mr. Green
ElseIf Form2.Text1.Text = 2 Then
    Form2.Image1(1).Top = 6360
    Form2.Image1(1).Left = 8160
    Form4.Visible = False
'Proffesor Plum
ElseIf Form2.Text1.Text = 3 Then
    Form2.Image1(2).Top = 6360
    Form2.Image1(2).Left = 8160
    Form4.Visible = False
'Mrs. Peacock
ElseIf Form2.Text1.Text = 4 Then
    Form2.Image1(5).Top = 6360
    Form2.Image1(5).Left = 8160
    Form4.Visible = False
'Miss Scarlet
ElseIf Form2.Text1.Text = 5 Then
    Form2.Image1(4).Top = 6360
    Form2.Image1(4).Left = 8160
    Form4.Visible = False
'Mrs. White
ElseIf Form2.Text1.Text = 6 Then
    Form2.Image1(3).Top = 6360
    Form2.Image1(3).Left = 8160
    Form4.Visible = False
End If
Command3.Visible = False
End Sub

Private Sub Command4_Click()
'move to dinning room
'Colonel Mustard
If Form2.Text1.Text = 1 Then
    Form2.Image1(0).Top = 6360
    Form2.Image1(0).Left = 2880
    Form4.Visible = False
'Mr. Green
ElseIf Form2.Text1.Text = 2 Then
    Form2.Image1(1).Top = 6360
    Form2.Image1(1).Left = 2880
    Form4.Visible = False
'Proffesor Plum
ElseIf Form2.Text1.Text = 3 Then
    Form2.Image1(2).Top = 6360
    Form2.Image1(2).Left = 2880
    Form4.Visible = False
'Mrs. Peacock
ElseIf Form2.Text1.Text = 4 Then
    Form2.Image1(5).Top = 6360
    Form2.Image1(5).Left = 2880
    Form4.Visible = False
'Miss Scarlet
ElseIf Form2.Text1.Text = 5 Then
    Form2.Image1(4).Top = 6360
    Form2.Image1(4).Left = 2880
    Form4.Visible = False
'Mrs. White
ElseIf Form2.Text1.Text = 6 Then
    Form2.Image1(3).Top = 6360
    Form2.Image1(3).Left = 2880
    Form4.Visible = False
End If
Command4.Visible = True
End Sub

Private Sub Command5_Click()
'Move to Greenhouse
'Colonel Mustard
If Form2.Text1.Text = 1 Then
    Form2.Image1(0).Top = 2520
    Form2.Image1(0).Left = 7200
    Form4.Visible = False
'Mr. Green
ElseIf Form2.Text1.Text = 2 Then
    Form2.Image1(1).Top = 2520
    Form2.Image1(1).Left = 7200
    Form4.Visible = False
'Proffesor Plum
ElseIf Form2.Text1.Text = 3 Then
    Form2.Image1(2).Top = 2520
    Form2.Image1(2).Left = 7200
    Form4.Visible = False
'Mrs. Peacock
ElseIf Form2.Text1.Text = 4 Then
    Form2.Image1(5).Top = 2520
    Form2.Image1(5).Left = 7200
    Form4.Visible = False
'Miss Scarlet
ElseIf Form2.Text1.Text = 5 Then
    Form2.Image1(4).Top = 2520
    Form2.Image1(4).Left = 7200
    Form4.Visible = False
'Mrs. White
ElseIf Form2.Text1.Text = 6 Then
    Form2.Image1(3).Top = 2520
    Form2.Image1(3).Left = 7200
    Form4.Visible = False
End If
Command5.Visible = False
End Sub
