Attribute VB_Name = "Areas"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Private Const TamanoAreas As Byte = 11

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    
    Dim loopX As Long, loopY As Long, CharIndex As Integer, ObjIndex As Integer
    
    MinLimiteX = (X \ TamanoAreas - 1) * TamanoAreas
    MaxLimiteX = MinLimiteX + ((TamanoAreas * 3) - 1)
    
    MinLimiteY = (Y \ TamanoAreas - 1) * TamanoAreas
    MaxLimiteY = MinLimiteY + ((TamanoAreas * 3) - 1)
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                
                'Erase NPCs
                CharIndex = Char_MapPosExits(loopX, loopY)
 
                If (CharIndex > 0) Then
                    If (CharIndex <> UserCharIndex) Then
                        Call Char_Erase(CharIndex)
                    End If
                End If
               
                'Erase OBJs
                ObjIndex = Map_PosExitsObject(loopX, loopY)
                                
                If (ObjIndex > 0) Then
                    Call Map_DestroyObject(loopX, loopY)
                End If
                
            End If
            
        Next loopY
    Next loopX
    
    Call RefreshAllChars
    
End Sub
