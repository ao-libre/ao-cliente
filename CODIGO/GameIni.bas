Attribute VB_Name = "GameIni"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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

Public Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
    Dim N As Integer
    Dim GameIni As tGameIni
    N = FreeFile
    Open App.path & "\init\Inicio.con" For Binary As #N
    Get #N, , MiCabecera
    
    Get #N, , GameIni
    
    Close #N
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
On Local Error Resume Next

Dim N As Integer
N = FreeFile
Open App.path & "\init\Inicio.con" For Binary As #N
Put #N, , MiCabecera
Put #N, , GameIniConfiguration
Close #N
End Sub

