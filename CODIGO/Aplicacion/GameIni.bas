Attribute VB_Name = "Game"
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

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Enum ePath
    INIT
    Graficos
    Sounds
    Musica
    Mapas
    Lenguajes
    Extras
End Enum

Public Type tGameIni
    tip As Byte
End Type

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Integer
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
    bNoSoundEffects As Boolean
    bGuildNews  As Boolean ' 11/19/09
    bDie        As Boolean ' 11/23/09 - FragShooter
    bKill       As Boolean ' 11/23/09 - FragShooter
    byMurderedLevel As Byte ' 11/23/09 - FragShooter
    bActive     As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs   As Byte
    
    
    'New dx8
    ProyectileEngine As Boolean
    PartyMembers As Boolean
    TonalidadPJ As Boolean
    UsarSombras As Boolean
    ParticleEngine As Boolean
    vSync As Boolean
    Aceleracion As Byte
    LimiteFPS As Boolean
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Private Lector As ClsIniReader

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
        .CRC = Rnd * 100
        .MagicWord = Rnd * 10
    End With
    
End Sub

Public Function path(ByVal PathType As ePath) As String

    Select Case PathType
        
        Case ePath.INIT
            path = App.path & "\INIT\"
        
        Case ePath.Graficos
            path = App.path & "\Graficos\"
            
        Case ePath.Lenguajes
            path = App.path & "\Lenguajes\"
            
        Case ePath.Mapas
            'En caso que no haya un mundo seleccionado en la propiedad Mundo
            'Seleccionamos Alkon como mundo default
            'Esto hay que eliminarlo de aqui ya que no tiene por que estar aqui, esto es un parche rapido para evitar posibles errores
            'Cuando hay problemas de conexion
            If LenB(MundoSeleccionado) = 0 Then
                MundoSeleccionado = "Alkon"
            End If
            
            path = App.path & "\Mapas\" & "\" & MundoSeleccionado & "\"
            
        Case ePath.Musica
            path = App.path & "\MIDI\"
            
        Case ePath.Sounds
            path = App.path & "\WAV\"
            
        Case ePath.Extras
            path = App.path & "\Extras\"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera
    
    Set Lector = New ClsIniReader
    Lector.Initialize (path(INIT) & "Client.DAT")
    
    With ClientSetup
        
        .bDinamic = CBool(Lector.GetValue("VIDEO", "DynamicLoad"))
        .byMemory = CInt(Lector.GetValue("VIDEO", "DinamicMemory"))
        .bNoMusic = CBool(Lector.GetValue("AUDIO", "DisableMIDI"))
        .bNoSound = CBool(Lector.GetValue("AUDIO", "DisableWAV"))
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DisableResolutionChange"))
        .bNoSoundEffects = CBool(Lector.GetValue("AUDIO", "DisableSoundEffects"))
        .bGuildNews = CBool(Lector.GetValue("GUILD", "GuildNews"))
        .bDie = CBool(Lector.GetValue("FRAGSHOOTER", "Die"))
        .bKill = CBool(Lector.GetValue("FRAGSHOOTER", "Kill"))
        .byMurderedLevel = CBool(Lector.GetValue("FRAGSHOOTER", "MurderedLevel"))
        .bActive = CBool(Lector.GetValue("FRAGSHOOTER", "Active"))
        .bGldMsgConsole = CBool(Lector.GetValue("GUILD", "GuildMessages"))
        .bCantMsgs = CByte(Lector.GetValue("GUILD", "MaxGuildMessages"))
        
        ' Nuevos motores vbGORE
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "ProyectileEngine"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PartyMembers"))
        .TonalidadPJ = CBool(Lector.GetValue("VIDEO", "TonalidadPJ"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "Sombras"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "ParticleEngine"))
        .vSync = CBool(Lector.GetValue("VIDEO", "vSync"))
        .Aceleracion = CByte(Lector.GetValue("VIDEO", "RenderMode"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitFPS"))

    End With

fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
        End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub

