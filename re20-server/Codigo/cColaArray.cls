VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColaArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit

Private m_maxelem   As Long
Private m_array()   As String
Private m_lastelem  As Integer
Private m_firstelem As Integer
Private m_size      As Integer

Public Function IsEmpty() As Boolean

100     IsEmpty = m_size = 0

End Function

Public Property Get Count() As Integer
100     Count = m_size
End Property

Public Function IsFull() As Boolean
    On Error GoTo 0
        
100     IsFull = m_size = m_maxelem

End Function

Public Function Push(ByVal aString As String) As Boolean
        
100     If Not Me.IsFull Then
    
102         If Me.IsEmpty Then m_firstelem = 1
    
104         m_lastelem = m_lastelem + 1

106         If (m_lastelem > m_maxelem) Then m_lastelem = m_lastelem - m_maxelem
108         m_size = m_size + 1
110         m_array(m_lastelem) = aString
    
112         Push = True
        Else
114         Push = False

        End If

End Function

Public Function Pop() As String
        
        On Error GoTo 0
        

100     If Not Me.IsEmpty Then
    
102         Pop = m_array(m_firstelem)
104         m_firstelem = m_firstelem + 1

106         If (m_firstelem > m_maxelem) Then m_firstelem = m_firstelem - m_maxelem
108         m_size = m_size - 1
    
            'If m_firstelem > m_lastelem And m_size = 0 Then
110         If m_size = 0 Then
112             m_lastelem = 0
114             m_firstelem = 0
116             m_size = 0

            End If

        Else
118         Pop = vbNullString
    
        End If

End Function

Private Sub Class_Initialize()
        
        On Error GoTo 0
        
100     m_lastelem = 0
102     m_firstelem = 0
104     m_size = 0
106     m_maxelem = 300

108     ReDim m_array(1 To m_maxelem)

End Sub

Public Property Get MaxElems() As Long
        
        On Error GoTo 0
        
100     MaxElems = m_maxelem

End Property

Public Property Let MaxElems(ByVal lNewValue As Long)
        
        On Error GoTo 0
        
100     m_maxelem = lNewValue
    
102     ReDim Preserve m_array(1 To m_maxelem)

End Property
