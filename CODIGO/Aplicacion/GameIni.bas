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
    Videos
End Enum

Public Type tSetupMods

    ' VIDEO
    byMemory        As Integer
    ProyectileEngine As Boolean
    PartyMembers    As Boolean
    TonalidadPJ     As Boolean
    UsarSombras     As Boolean
    ParticleEngine  As Boolean
    LimiteFPS       As Boolean
    bNoRes          As Boolean
    OverrideVertexProcess As Byte
    
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
    KeyboardBindKeysConfig As String
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

        Case ePath.Videos
            'Hacemos un Left para poder solo obtener la letra del HD
            'Por que por culpa del UAC no guarda los videos en la carpeta del juego...
            Dim VideosPath As String
            VideosPath = Left$(App.path, 2) & "\AO-Libre\Videos\"

            If Dir(VideosPath, vbDirectory) = "" Then
                MkDir VideosPath
            End If
            
            path = VideosPath
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Game.path(INIT) & CLIENT_FILE)
    
    With ClientSetup
    
        ' VIDEO
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DisableResolutionChange"))
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "ProjectileEngine"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PartyMembers"))
        .TonalidadPJ = CBool(Lector.GetValue("VIDEO", "TonalidadPJ"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "Sombras"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "ParticleEngine"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        
        ' AUDIO
        .bMusic = CBool(Lector.GetValue("AUDIO", "Music"))
        .bSound = CBool(Lector.GetValue("AUDIO", "Sound"))
        .bSoundEffects = CBool(Lector.GetValue("AUDIO", "SoundEffects"))
        .MusicVolume = CByte(Lector.GetValue("AUDIO", "MusicVolume"))
        .SoundVolume = CByte(Lector.GetValue("AUDIO", "SoundVolume"))
        
        ' GUILD
        .bGuildNews = CBool(Lector.GetValue("GUILD", "News"))
        .bGldMsgConsole = CBool(Lector.GetValue("GUILD", "Messages"))
        .bCantMsgs = CByte(Lector.GetValue("GUILD", "MaxMessages"))
        
        ' FRAGSHOOTER
        .bDie = CBool(Lector.GetValue("FRAGSHOOTER", "Die"))
        .bKill = CBool(Lector.GetValue("FRAGSHOOTER", "Kill"))
        .byMurderedLevel = CByte(Lector.GetValue("FRAGSHOOTER", "MurderedLevel"))
        .bActive = CBool(Lector.GetValue("FRAGSHOOTER", "Active"))
        
        ' OTHER
        .MostrarTips = CBool(Lector.GetValue("OTHER", "MOSTRAR_TIPS"))
        .MostrarBindKeysSelection = CBool(Lector.GetValue("OTHER", "MOSTRAR_BIND_KEYS_SELECTION"))
        .KeyboardBindKeysConfig = Lector.GetValue("OTHER", "BIND_KEYS")

        Debug.Print "byMemory: " & .byMemory
        Debug.Print "bNoRes: " & .bNoRes
        Debug.Print "ProyectileEngine: " & .ProyectileEngine
        Debug.Print "PartyMembers: " & .PartyMembers
        Debug.Print "TonalidadPJ: " & .TonalidadPJ
        Debug.Print "UsarSombras: " & .UsarSombras
        Debug.Print "ParticleEngine: " & .ParticleEngine
        Debug.Print "LimitarFPS: " & .LimiteFPS
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
        Debug.Print "MostrarBindKeysSelection: " & .MostrarBindKeysSelection
        Debug.Print "KeyboardBindKeysConfig: " & .KeyboardBindKeysConfig
        Debug.Print vbNullString
        
    End With
    
    Exit Sub
    
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
        Call Lector.ChangeValue("VIDEO", "DynamicMemory", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DisableResolutionChange", IIf(.bNoRes, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "ProjectileEngine", IIf(.ProyectileEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "PartyMembers", IIf(.PartyMembers, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "TonalidadPJ", IIf(.TonalidadPJ, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Sombras", IIf(.UsarSombras, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ParticleEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "LimitarFPS", IIf(.LimiteFPS, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "VertexProcessingOverride", .OverrideVertexProcess)
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "Music", IIf(Audio.MusicActivated, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "Sound", IIf(Audio.SoundActivated, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "SoundEffects", IIf(Audio.SoundEffectsActivated, "True", "False"))
        Call Lector.ChangeValue("AUDIO", "MusicVolume", Audio.MusicVolume)
        Call Lector.ChangeValue("AUDIO", "SoundVolume", Audio.SoundVolume)
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "News", IIf(.bGuildNews, "True", "False"))
        Call Lector.ChangeValue("GUILD", "Messages", IIf(DialogosClanes.Activo, "True", "False"))
        Call Lector.ChangeValue("GUILD", "MaxMessages", CByte(DialogosClanes.CantidadDialogos))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "Die", IIf(.bDie, "True", "False"))
        Call Lector.ChangeValue("FRAGSHOOTER", "Kill", IIf(.bKill, "True", "False"))
        Call Lector.ChangeValue("FRAGSHOOTER", "MurderedLevel", CByte(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "Active", IIf(.bActive, "True", "False"))
        
        ' OTHER
        ' Lo comento por que no tiene por que setearse aqui esto.
        ' Al menos no al hacer click en el boton Salir del formulario opciones (Recox)
        ' Call Lector.ChangeValue("OTHER", "MOSTRAR_TIPS", CBool(.MostrarTips))
    End With
    
    Call Lector.DumpFile(Game.path(INIT) & CLIENT_FILE)
    
    Exit Sub
    
fileErr:

    If Err.number <> 0 Then
        Call MsgBox("Ha ocurrido un error al guardar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    End If
    
End Sub
