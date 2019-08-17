Attribute VB_Name = "Game"
'Argentum Online 0.13.9.2

' ***********************************************
'   Nueva carga de configuracion mediante .INI
' ***********************************************

Option Explicit

Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Enum ePath
    INIT
    Graficos
    Interfaces
    Skins
    Sounds
    Musica
    Mapas
    Lenguajes
    Extras
End Enum

Public Type tSetupMods

    ' VIDEO
    bDinamic    As Boolean
    byMemory    As Integer
    ProyectileEngine As Boolean
    PartyMembers As Boolean
    TonalidadPJ As Boolean
    UsarSombras As Boolean
    ParticleEngine As Boolean
    vSync As Boolean
    Aceleracion As Byte
    LimiteFPS As Boolean
    bNoRes      As Boolean
    
    ' AUDIO
    bMusic    As Boolean
    bSound    As Boolean
    bSoundEffects As Boolean
    
    ' GUILDS
    bGuildNews  As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs   As Byte
    
    ' FRAGSHOOTER
    bActive     As Boolean
    bDie        As Boolean
    bKill       As Boolean
    byMurderedLevel As Byte
    
    ' OTHER
    MostrarTips As Byte
End Type

Public ClientSetup As tSetupMods
Public MiCabecera As tCabecera

Private Lector As clsIniManager
Private Const CLIENT_FILE As String = "Config.ini"

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
        
        Case ePath.Skins
            path = App.path & "\Graficos\Skins\"
            
        Case ePath.Interfaces
            path = App.path & "\Graficos\Interfaces\"
            
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
            path = App.path & "\AUDIO\MIDI\"
            
        Case ePath.Sounds
            path = App.path & "\AUDIO\WAV\"
            
        Case ePath.Extras
            path = App.path & "\AUDIO\Extras\"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera
    
    Set Lector = New clsIniManager
    Lector.Initialize (Game.path(INIT) & CLIENT_FILE)
    
    With ClientSetup
        
        ' VIDEO
        .bDinamic = CBool(Lector.GetValue("VIDEO", "DYNAMIC_LOAD"))
        .byMemory = CInt(Lector.GetValue("VIDEO", "DINAMIC_MEMORY"))
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DISABLE_RESOLUTION_CHANGE"))
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "PROYECTILE_ENGINE"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PARTY_MEMBERS"))
        .TonalidadPJ = CBool(Lector.GetValue("VIDEO", "TONALIDAD_PJ"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "SOMBRAS"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "PARTICLE_ENGINE"))
        .vSync = CBool(Lector.GetValue("VIDEO", "VSYNC"))
        .Aceleracion = CByte(Lector.GetValue("VIDEO", "RENDER_MODE"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LIMIT_FPS"))
        
        ' AUDIO
        .bMusic = CBool(Lector.GetValue("AUDIO", "MIDI"))
        .bSound = CBool(Lector.GetValue("AUDIO", "WAV"))
        .bSoundEffects = CBool(Lector.GetValue("AUDIO", "SOUND_EFFECTS"))
        
        ' GUILD
        .bGuildNews = CBool(Lector.GetValue("GUILD", "NEWS"))
        .bGldMsgConsole = CBool(Lector.GetValue("GUILD", "MESSAGES"))
        .bCantMsgs = CByte(Lector.GetValue("GUILD", "MAX_MESSAGES"))
        
        ' FRAGSHOOTER
        .bDie = CBool(Lector.GetValue("FRAGSHOOTER", "DIE"))
        .bKill = CBool(Lector.GetValue("FRAGSHOOTER", "KILL"))
        .byMurderedLevel = CBool(Lector.GetValue("FRAGSHOOTER", "MURDERED_LEVEL"))
        .bActive = CBool(Lector.GetValue("FRAGSHOOTER", "ACTIVE"))
        
        ' OTHER
        .MostrarTips = CBool(Lector.GetValue("OTHER", "MOSTRAR_TIPS"))
    End With

fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
        End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub

Public Sub GuardarConfiguracion()
    On Local Error GoTo fileErr:
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Game.path(INIT) & CLIENT_FILE)
    
    With ClientSetup
        
        ' VIDEO
        Call Lector.ChangeValue("VIDEO", "DYNAMIC_LOAD", CInt(.bDinamic))
        Call Lector.ChangeValue("VIDEO", "DINAMIC_MEMORY", CInt(.byMemory))
        Call Lector.ChangeValue("VIDEO", "DISABLE_RESOLUTION_CHANGE", CInt(.bNoRes))
        Call Lector.ChangeValue("VIDEO", "PROYECTILE_ENGINE", CInt(.ProyectileEngine))
        Call Lector.ChangeValue("VIDEO", "PARTY_MEMBERS", CInt(.PartyMembers))
        Call Lector.ChangeValue("VIDEO", "TONALIDAD_PJ", CInt(.TonalidadPJ))
        Call Lector.ChangeValue("VIDEO", "SOMBRAS", CInt(.UsarSombras))
        Call Lector.ChangeValue("VIDEO", "PARTICLE_ENGINE", CInt(.ParticleEngine))
        Call Lector.ChangeValue("VIDEO", "VSYNC", CInt(.vSync))
        Call Lector.ChangeValue("VIDEO", "RENDER_MODE", .Aceleracion)
        Call Lector.ChangeValue("VIDEO", "LIMIT_FPS", CInt(.LimiteFPS))
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MIDI", CInt(.bMusic))
        Call Lector.ChangeValue("AUDIO", "WAV", CInt(.bSound))
        Call Lector.ChangeValue("AUDIO", "SOUND_EFFECTS", CInt(.bSoundEffects))
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "NEWS", CInt(.bGuildNews))
        Call Lector.ChangeValue("GUILD", "MESSAGES", CInt(.bGldMsgConsole))
        Call Lector.ChangeValue("GUILD", "MAX_MESSAGES", CInt(.bCantMsgs))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "DIE", CInt(.bDie))
        Call Lector.ChangeValue("FRAGSHOOTER", "KILL", CInt(.bKill))
        Call Lector.ChangeValue("FRAGSHOOTER", "MURDERED_LEVEL", CInt(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "ACTIVE", CInt(.bActive))
        
        ' OTHER
        Call Lector.ChangeValue("OTHER", "MOSTRAR_TIPS", CInt(.MostrarTips))
    End With
    
    Call Lector.DumpFile(Game.path(INIT) & CLIENT_FILE)
fileErr:

    'If Err.number <> 0 Then
    '    MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    '    End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    'End If
End Sub
