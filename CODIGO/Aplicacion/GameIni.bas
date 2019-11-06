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
    Aceleracion As Byte
    byMemory    As Integer
    ProyectileEngine As Boolean
    PartyMembers As Boolean
    TonalidadPJ As Boolean
    UsarSombras As Boolean
    ParticleEngine As Boolean
    vSync As Boolean
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
    MostrarBindKeysSelection As Byte
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
            path = App.path & "\Extras\"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Game.path(INIT) & CLIENT_FILE)
    
    With ClientSetup
        
        ' VIDEO
        .Aceleracion = Lector.GetValue("VIDEO", "RENDER_MODE")
        .byMemory = Lector.GetValue("VIDEO", "DINAMIC_MEMORY")
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DISABLE_RESOLUTION_CHANGE"))
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "PROYECTILE_ENGINE"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PARTY_MEMBERS"))
        .TonalidadPJ = CBool(Lector.GetValue("VIDEO", "TONALIDAD_PJ"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "SOMBRAS"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "PARTICLE_ENGINE"))
        .vSync = CBool(Lector.GetValue("VIDEO", "VSYNC"))
        
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
        .byMurderedLevel = CByte(Lector.GetValue("FRAGSHOOTER", "MURDERED_LEVEL"))
        .bActive = CBool(Lector.GetValue("FRAGSHOOTER", "ACTIVE"))
        
        ' OTHER
        .MostrarTips = CBool(Lector.GetValue("OTHER", "MOSTRAR_TIPS"))
        .MostrarBindKeysSelection = CBool(Lector.GetValue("OTHER", "MOSTRAR_BIND_KEYS_SELECTION"))
        
        Debug.Print "Modo de Renderizado: " & IIf(.Aceleracion = 1, "Mixto (Hardware + Software)", "Hardware")
        Debug.Print "byMemory: " & .byMemory
        Debug.Print "bNoRes: " & .bNoRes
        Debug.Print "ProyectileEngine: " & .ProyectileEngine
        Debug.Print "PartyMembers: " & .PartyMembers
        Debug.Print "TonalidadPJ: " & .TonalidadPJ
        Debug.Print "UsarSombras: " & .UsarSombras
        Debug.Print "ParticleEngine: " & .ParticleEngine
        Debug.Print "vSync: " & .vSync
        Debug.Print "bMusic: " & .bMusic
        Debug.Print "bSound: " & .bSound
        Debug.Print "bSoundEffects: " & .bSoundEffects
        Debug.Print "bGuildNews: " & .bGuildNews
        Debug.Print "bGldMsgConsole: " & .bGldMsgConsole
        Debug.Print "bCantMsgs: " & .bCantMsgs
        Debug.Print "bDie: " & .bDie
        Debug.Print "bKill: " & .byMurderedLevel
        Debug.Print "bActive: " & .bActive
        Debug.Print "MostrarTips: " & .MostrarTips
        Debug.Print vbNullString
        
    End With
  
fileErr:

    'If Err.number <> 0 Then
    '    MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    '    End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    'End If
End Sub

Public Sub GuardarConfiguracion()
    On Local Error GoTo fileErr:
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Game.path(INIT) & CLIENT_FILE)

    With ClientSetup
        
        ' VIDEO
        Call Lector.ChangeValue("VIDEO", "RENDER_MODE", .Aceleracion)
        Call Lector.ChangeValue("VIDEO", "DINAMIC_MEMORY", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DISABLE_RESOLUTION_CHANGE", CBool(.bNoRes))
        Call Lector.ChangeValue("VIDEO", "PROYECTILE_ENGINE", CBool(.ProyectileEngine))
        Call Lector.ChangeValue("VIDEO", "PARTY_MEMBERS", CBool(.PartyMembers))
        Call Lector.ChangeValue("VIDEO", "TONALIDAD_PJ", CBool(.TonalidadPJ))
        Call Lector.ChangeValue("VIDEO", "SOMBRAS", CBool(.UsarSombras))
        Call Lector.ChangeValue("VIDEO", "PARTICLE_ENGINE", CBool(.ParticleEngine))
        Call Lector.ChangeValue("VIDEO", "VSYNC", CBool(.vSync))
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MIDI", CBool(Audio.MusicActivated))
        Call Lector.ChangeValue("AUDIO", "WAV", CBool(Audio.SoundActivated))
        Call Lector.ChangeValue("AUDIO", "SOUND_EFFECTS", CBool(Audio.SoundEffectsActivated))
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "NEWS", CBool(.bGuildNews))
        Call Lector.ChangeValue("GUILD", "MESSAGES", CBool(DialogosClanes.Activo))
        Call Lector.ChangeValue("GUILD", "MAX_MESSAGES", CByte(DialogosClanes.CantidadDialogos))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "DIE", CBool(.bDie))
        Call Lector.ChangeValue("FRAGSHOOTER", "KILL", CBool(.bKill))
        Call Lector.ChangeValue("FRAGSHOOTER", "MURDERED_LEVEL", CByte(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "ACTIVE", CBool(.bActive))
        
        ' OTHER
        ' Lo comento por que no tiene por que setearse aqui esto.
        ' Al menos no al hacer click en el boton Salir del formulario opciones (Recox)
        ' Call Lector.ChangeValue("OTHER", "MOSTRAR_TIPS", CBool(.MostrarTips))
    End With
    
    Call Lector.DumpFile(Game.path(INIT) & CLIENT_FILE)
fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
        End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub
