VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   6000
      Top             =   1920
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   2040
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Image imgTitle 
      Height          =   3150
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (C) 2007  Justin McCullough
Option Explicit

Dim i As Integer

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    i = i + 1
    
Select Case i
    Case 1
        lblStatus.Caption = "Selecting the Room..."
    Case 2
        lblStatus.Caption = "Selecting the Perp..."
    Case 3
        lblStatus.Caption = "Selecting the Weapon..."
    Case 4
        lblStatus.Caption = "Killing the guy..."
    Case 5
        Unload Me
        frmStartup.Show
        Timer1.Enabled = False
End Select
End Sub
