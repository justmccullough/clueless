Attribute VB_Name = "modClueless"
'Author: Justin McCullough
'Created: 3/21/2007
'Purpose: This is the main module of the Clueless game
    'it holds all of the data access routines and any other public routines for game play
'Copyright (C) 2007  Justin McCullough
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

    
Option Explicit
Dim db As New DAL.DataAccess

Public Sub CloseProgram()
'Purpose: To close out all the clueless forms
Unload dlgGuess
Unload dlgGuessResults
Unload frmHighScores
Unload frmMain
Unload frmSelectCharacter
Unload frmStartup
Unload frmWon
Unload frmAllScores
End Sub

Public Sub OpenConnection()
'Purpose: to open a connection to the CluelessDB
'no inputs or returns
'Create an object of the data access class


On Error GoTo err

'the name for the CluelessDB
Dim CluelessDb As String
CluelessDb = App.Path & "\Clueless.mdb"

sConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CluelessDb & ";Persist Security Info=False;Jet OLEDB:Database Password=mrsharris"

'set the connections connection string
gConnection.ConnectionString = sConnection

gConnection.Open

db.dBPath = CluelessDb

'open the connection
db.OpenConnection , , , , gConnection

Exit Sub
err:
    If err.Number <> 0 Then
        MsgBox "Error: " & err.Number & vbCrLf & "Error Description: " & err.Description
        Debug.Print err.Number
        Debug.Print err.Description
        Debug.Print "Open Connection"
        Debug.Print ""
    End If
End Sub

Public Function GetRS(ByVal sSQL As String, ByRef rs As ADODB.Recordset) As Boolean
'Purpose: To get a connected recordset
'Argurments: Requires a SQL statement and the name of the recordset you want to retrieve
'Returns: A boolean value whether it suceeded or not


On Error GoTo err
OpenConnection

'this is to start a transmission
db.BeginTrans

If db.GetDisconnectedRecordset(sSQL, rs, False, False) Then
    GetRS = True
    db.CommitTrans
Else
    db.RollbackTrans
    GetRS = False

End If

db.CloseConnection
gConnection.Close

Set db = Nothing
Set gConnection = Nothing

Exit Function
err:
    If err.Number <> 0 Then
        MsgBox "Error: " & err.Number & vbCrLf & "Error Description: " & err.Description
        Debug.Print err.Number
        Debug.Print err.Description
        Debug.Print "GetRS"
        Debug.Print ""
    End If
End Function

Public Function PutRS(ByRef rs As ADODB.Recordset) As Boolean
'Purpose: To put a updated recordset back into the database
'Arguments: Requires the name of the recordset your are trying to put
'Returns: A boolean value whether or not it suceeded

On Error GoTo err

OpenConnection

'this is to start a transmission
db.BeginTrans

If db.PutRecordset(rs) Then
    PutRS = True
    db.CommitTrans
Else
    PutRS = False
    db.RollbackTrans
End If

db.CloseConnection
gConnection.Close

Set db = Nothing
Set gConnection = Nothing

Exit Function
err:
    If err.Number <> 0 Then
        MsgBox "Error: " & err.Number & vbCrLf & "Error Description: " & err.Description
        Debug.Print err.Number
        Debug.Print err.Description
        Debug.Print "PutRS"
        Debug.Print ""
    End If

End Function

