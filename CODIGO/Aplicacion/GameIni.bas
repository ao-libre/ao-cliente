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
    MusicaMp3
    Mapas
    Lenguajes
    Extras
    Fonts
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
    MusicVolume As Byte
    SoundVolume As Byte
    
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
            
        Case ePath.Fonts
            path = App.path & "\Graficos\Fonts\"
            
        Case ePath.Lenguajes
            path = App.path & "\Lenguajes\"
            
        Case ePath.Mapas       
            'ESTO HAY QUE BORRARLO CUANDO SE PUEDA
            'AGREGAR SERVERS QUE NO ESTEN EN LA LISTA
            'Y EL MISMO PUEDA SETEAR MUNDO SINO BUGGEA :)   
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

        Case ePath.MusicaMp3
            path = App.path & "\AUDIO\MP3\"
            
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
        .bMusic = CBool(Lector.GetValue("AUDIO", "MUSIC"))
        .bSound = CBool(Lector.GetValue("AUDIO", "SOUND"))
        .bSoundEffects = CBool(Lector.GetValue("AUDIO", "SOUND_EFFECTS"))
        .MusicVolume = CByte(Lector.GetValue("AUDIO", "MUSIC_VOLUME"))
        .SoundVolume = CByte(Lector.GetValue("AUDIO", "SOUND_VOLUME"))
        
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
        Debug.Print "MusicVolume: " & .MusicVolume
        Debug.Print "SoundVolume: " & .SoundVolume
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
        Call Lector.ChangeValue("VIDEO", "RENDER_MODE", .Aceleracion)
        Call Lector.ChangeValue("VIDEO", "DINAMIC_MEMORY", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DISABLE_RESOLUTION_CHANGE", IIf(.bNoRes, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "PROYECTILE_ENGINE", IIf(.ProyectileEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "PARTY_MEMBERS", IIf(.PartyMembers, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "TONALIDAD_PJ", IIf(.TonalidadPJ, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "SOMBRAS", IIf(.UsarSombras, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "PARTICLE_ENGINE", IIf(.ParticleEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "VSYNC", IIf(.vSync, "True", "False"))
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MUSIC", IIf(Audio.MusicActivated, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "SOUND", IIf(Audio.SoundActivated, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "SOUND_EFFECTS", IIf(Audio.SoundEffectsActivated, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "MUSIC_VOLUME", Audio.MusicVolume)
        Call Lector.ChangeValue("AUDIO", "SOUND_VOLUME", Audio.SoundVolume)
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "NEWS", IIf(.bGuildNews, "True", "False"))
        Call Lector.ChangeValue("GUILD", "MESSAGES", IIf(DialogosClanes.Activo, "True", "False"))
        Call Lector.ChangeValue("GUILD", "MAX_MESSAGES", CByte(DialogosClanes.CantidadDialogos))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "DIE", IIf(.bDie, "True", "False"))
        Call Lector.ChangeValue("FRAGSHOOTER", "KILL", IIf(.bKill, "True", "False"))
        Call Lector.ChangeValue("FRAGSHOOTER", "MURDERED_LEVEL", CByte(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "ACTIVE", IIf(.bActive, "True", "False"))
        
        ' OTHER
        ' Lo comento por que no tiene por que setearse aqui esto.
        ' Al menos no al hacer click en el boton Salir del formulario opciones (Recox)
        ' Call Lector.ChangeValue("OTHER", "MOSTRAR_TIPS", CBool(.MostrarTips))
    End With
    
    Call Lector.DumpFile(Game.path(INIT) & CLIENT_FILE)
fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    End If
End Sub
