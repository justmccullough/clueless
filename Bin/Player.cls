VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************************************
'Author: Justin McCullough

'Purpose:  This is a class module for a Clueless Player

'Created: 3/25/3007

'*****************************************************************

Option Explicit

Private Const iStartScore As Integer = 1000
Private mName As String
Private mCharacterIndex As Integer
Private mMoves As Integer
Private mScore As Integer
Private mAtStart As Boolean
Private mRoomsGuessed As String
Private mWeaponsGuessed As String
Private mSuspectsGuessed As String
Private mInRoom As Boolean

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(Name As String)
    mName = Name
End Property
Public Property Let CharacterIndex(iCharacter As Integer)
    mCharacterIndex = iCharacter
End Property

Public Property Get CharacterIndex() As Integer
    CharacterIndex = mCharacterIndex
End Property

Public Property Get Moves() As Integer
    Moves = mMoves
End Property

Public Property Let Moves(Moves As Integer)
    mMoves = Moves
End Property

Public Property Get Score() As Integer
    Score = mScore
End Property

Public Property Let Score(Score As Integer)
    mScore = Score
End Property

Public Property Get AtStart() As Boolean
    AtStart = mAtStart
End Property

Public Property Let AtStart(AtStart As Boolean)
    mAtStart = AtStart
End Property

Public Property Get InRoom() As Boolean
    InRoom = mInRoom
End Property

Public Property Let InRoom(InRoom As Boolean)
    mInRoom = InRoom
End Property

Public Property Get Rooms() As String
    Rooms = mRoomsGuessed
End Property

Public Property Let Rooms(Rooms As String)
     mRoomsGuessed = Rooms
End Property

Public Property Get Suspects() As String
    Suspects = mSuspectsGuessed
End Property

Public Property Let Suspects(Suspects As String)
     mSuspectsGuessed = Suspects
End Property

Public Property Get Weapons() As String
    Weapons = mWeaponsGuessed
End Property

Public Property Let Weapons(Weapons As String)
     mWeaponsGuessed = Weapons
End Property

Public Sub Init(ByVal iCharacterIndex As Integer, ByVal sPlayerName As String)
'this is the class constructor
    mScore = iStartScore
    mMoves = 0
    mCharacterIndex = iCharacterIndex
    mName = sPlayerName
    mAtStart = True
    mInRoom = False
End Sub

Public Sub Class_Initialize()

End Sub

