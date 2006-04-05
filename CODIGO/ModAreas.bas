Attribute VB_Name = "ModAreas"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    Dim loopX As Long, loopY As Long
    Dim tempint As Integer
    
    MinLimiteX = (X \ 9 - 1) * 9
    MaxLimiteX = MinLimiteX + 26
    
    MinLimiteY = (Y \ 9 - 1) * 9
    MaxLimiteY = MinLimiteY + 26
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                
                If MapData(loopX, loopY).CharIndex > 0 Then
                    If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                        tempint = MapData(loopX, loopY).CharIndex
                        Call EraseChar(MapData(loopX, loopY).CharIndex)
                        charlist(tempint).Nombre = loopX & "-" & loopY
                    End If
                End If
                
                'Erase OBJs
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
            End If
        Next
    Next
    
    Call RefreshAllChars
End Sub

Public Sub ClearMap()
Dim loopX As Long, loopY As Long
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            'Erase NPCs
            If MapData(loopX, loopY).CharIndex > 0 Then
                If MapData(loopX, loopY).CharIndex <> UserCharIndex Then
                    Call EraseChar(MapData(loopX, loopY).CharIndex)
                End If
            End If
            
            'Erase OBJs
            MapData(loopX, loopY).ObjGrh.GrhIndex = 0
        Next
    Next
    
    Call RefreshAllChars
End Sub
