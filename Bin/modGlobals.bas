Attribute VB_Name = "modGlobals"
Option Explicit

'Copyright (C) 2007  Justin McCullough

'items to be guessed
Public sPerp, sWeapon, sRoom As String

'the players character
Public sCharacter As String

'current room the player is in
Public sCurrentRoom As String

'the players score
Public iScore As Integer

'title for message boxes
Public Const sTitle As String = "Clueless"

'users name
Public sName As String

'the SQL statement to retrieve the High Scores from the scores table in the CluelessDb
Public Const sHighScoresSQL As String = "SELECT * FROM Scores ORDER BY Score DESC"

'the SQL statement  to retrieve the users from the scores table in the CluelessDb
Public Const sUsersSQL As String = "SELECT UserName, Score FROM Scores"

'the SQL statement to retrieve all the scores from the scores table in the CluelessDb
Public Const sScoresSQL As String = "SELECT UserName, score FROM Scores ORDER BY Score DESC"

'the bookmarked ID of the user before we get the scores back from the databse
Public iUserID As Integer

'the connection string used for the CluelessDb
Public sConnection As String

'the number of users playing
Public iNumberofPlayers As Integer

'The global Connection for the CluelessDB
Public gConnection As New ADODB.Connection

'public array
Public players(5) As New Player

'the indexes of both the current player and the image of the currect player's character
Public iCurrentIndex, iCurrentPlayer As Integer

'the arrays containing rooms, suspects, weapons guessed
Public Rooms, Suspects, Weapons As Variant

'the strings to hold the delimted string
Public sRooms, sSuspects, sWeapons As Variant
