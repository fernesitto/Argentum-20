VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapturaParticipante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim UserRef As t_UserReference
Public LastPosMap As Integer
Public LastPosX As Byte
Public LastPosY As Byte
Public Team As Byte


Public Sub SetPlayer(ByVal UserIndex As Integer)
    Call SetUserRef(UserRef, UserIndex)
End Sub

Public Function PlayerIndex() As Integer
    PlayerIndex = UserRef.ArrayIndex
End Function

Public Function IsValid() As Boolean
    IsValid = IsValidUserRef(UserRef)
End Function
