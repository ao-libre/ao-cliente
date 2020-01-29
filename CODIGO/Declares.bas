Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

Public Inet As clsInet

'Recuperar Cuenta mediante la API
Public AccountMailToRecover As String
Public AccountNewPassword As String

' Desvanecimiento en Techos
Public ColorTecho As Byte
Public temp_rgb(3) As Long
Public renderText As String
Public renderFont As Integer
Public colorRender As Byte
Public render_msg(3) As Long
Public Sonidos As clsSoundMapas

'//Caminata fluida
Public Movement_Speed As Single

'Objetos publicos
Public DialogosClanes As clsGuildDlg
Public Dialogos As clsDialogs
Public Audio As clsAudio
Public Inventario As clsGraphicalInventory
Public InvBanco(1) As clsGraphicalInventory

'Inventarios de comercio con usuario
Public InvComUsu As clsGraphicalInventory  ' Inventario del usuario visible en el comercio
Public InvOroComUsu(2) As clsGraphicalInventory  ' Inventarios de oro (ambos usuarios)
Public InvOfferComUsu(1) As clsGraphicalInventory  ' Inventarios de ofertas (ambos usuarios)

Public InvComNpc As clsGraphicalInventory  ' Inventario con los items que ofrece el npc

'Inventarios de herreria
Public Const MAX_LIST_ITEMS As Byte = 4
Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public InvObjArtesano(1 To MAX_LIST_ITEMS) As clsGraphicalInventory

Public Const MAX_ITEMS_CRAFTEO As Byte = 4

Public CustomKeys As clsCustomKeys
Public CustomMessages As clsCustomMessages

Public incomingData As clsByteQueue
Public outgoingData As clsByteQueue

''
'The main timer of the game.
Public MainTimer As clsTimer

'Error code
Public Enum eSockError
   TOO_FAST = 24036
   REFUSED = 24061
   TIME_OUT = 24060
End Enum

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_DICE As String = "cupdice.Wav"

' Constantes de intervalo
Public Enum eIntervalos
    INT_MACRO_HECHIS = 2000
    INT_MACRO_TRABAJO = 900
    INT_ATTACK = 1500
    INT_ARROWS = 1400
    INT_CAST_SPELL = 1400
    INT_CAST_ATTACK = 1000
    INT_WORK = 700
    INT_USEITEMU = 450
    INT_USEITEMDCK = 125
    INT_SENTRPU = 2000
End Enum

Public MacroBltIndex As Integer

Public Const NUMATRIBUTES As Byte = 5

Public Const iCuerpoMuerto As Integer = 8

Public Enum eCabezas
    CASPER_HEAD = 500
    FRAGATA_FANTASMAL = 87
    
    HUMANO_H_PRIMER_CABEZA = 1
    HUMANO_H_ULTIMA_CABEZA = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no seleccionables
    HUMANO_H_CUERPO_DESNUDO = 21
    
    ELFO_H_PRIMER_CABEZA = 101
    ELFO_H_ULTIMA_CABEZA = 122
    ELFO_H_CUERPO_DESNUDO = 210
    
    DROW_H_PRIMER_CABEZA = 201
    DROW_H_ULTIMA_CABEZA = 221
    DROW_H_CUERPO_DESNUDO = 32
    
    ENANO_H_PRIMER_CABEZA = 301
    ENANO_H_ULTIMA_CABEZA = 319
    ENANO_H_CUERPO_DESNUDO = 53
    
    GNOMO_H_PRIMER_CABEZA = 401
    GNOMO_H_ULTIMA_CABEZA = 416
    GNOMO_H_CUERPO_DESNUDO = 222
    
    HUMANO_M_PRIMER_CABEZA = 70
    HUMANO_M_ULTIMA_CABEZA = 89
    HUMANO_M_CUERPO_DESNUDO = 39
    
    ELFO_M_PRIMER_CABEZA = 170
    ELFO_M_ULTIMA_CABEZA = 188
    ELFO_M_CUERPO_DESNUDO = 259
    
    DROW_M_PRIMER_CABEZA = 270
    DROW_M_ULTIMA_CABEZA = 288
    DROW_M_CUERPO_DESNUDO = 40
    
    ENANO_M_PRIMER_CABEZA = 370
    ENANO_M_ULTIMA_CABEZA = 384
    ENANO_M_CUERPO_DESNUDO = 60
    
    GNOMO_M_PRIMER_CABEZA = 470
    GNOMO_M_ULTIMA_CABEZA = 484
    GNOMO_M_CUERPO_DESNUDO = 260
End Enum

Public ColoresPJ(0 To 50) As Long

Public ColoresDano(51 To 56) As Long

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    Mundo As String
End Type

Public ServersLst() As tServerInfo

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Enum ePartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public NumEscudosAnims As Integer

Public ArmasHerrero() As tItemsConstruibles
Public ArmadurasHerrero() As tItemsConstruibles
Public ObjCarpintero() As tItemsConstruibles
Public CarpinteroMejorar() As tItemsConstruibles
Public HerreroMejorar() As tItemsConstruibles
Public ObjArtesano() As tItemArtesano

Public UsaMacro As Boolean
Public CnTd As Byte

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

Public TradingUserName As String

Public Tips() As String * 255

'Direcciones
Public Enum E_Heading
    nada = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 35
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const INV_OFFER_SLOTS As Byte = 20
Public Const INV_GOLD_SLOTS As Byte = 1

Public Const MAXSKILLPOINTS As Byte = 100

Public Const MAXATRIBUTOS As Byte = 38

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1
Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1

Public Enum eClass
    Mage = 1      'Mago
    Cleric = 2    'Clerigo
    Warrior = 3   'Guerrero
    Assasin = 4   'Asesino
    Thief = 5     'Ladron
    Bard = 6      'Bardo
    Druid = 7     'Druida
    Bandit = 8    'Bandido
    Paladin = 9   'Paladin
    Hunter = 10   'Cazador
    Worker = 11   'Trabajador
    Pirate = 12    'Pirata
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix = 2
    cBanderbill = 3
    cLindos = 4
    cArghal = 5
End Enum

Enum eRaza
    Humano = 1
    Elfo = 2
    ElfoOscuro = 3
    Gnomo = 4
    Enano = 5
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apunalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer = 2
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otOro = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otLibros = 12 'Hacer algo con esto, no en uso
    otBebidas = 13
    otLena = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otMuebles = 20
    otJoyas = 21 'Hacer algo con esto, no en uso
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otMonturas = 25
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otGemas = 29 'No en uso, hacer algo con las gemas :)
    otFlores = 30 'No en uso, hacer algo con las flores :)
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManuales = 35
    otArbolElfico = 36
    otMochilas = 37
    otYacimientoPez = 38
    otCualquiera = 1000
End Enum

Public Enum eMochilas
    Mediana = 1
    GRANDE = 2
End Enum

Public MaxInventorySlots As Byte

Public Const FundirMetal As Integer = 88

' Determina el color del nick
Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMP3ToMap          '/FORCEMP3MAP
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMP3All             '/FORCEMP3
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC y /RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
    ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    EnableDenounces         '/DENUNCIAS
    ShowDenouncesList       '/SHOW DENUNCIAS
    MapMessage              '/MAPMSG
    SetDialog               '/SETDIALOG
    Impersonate             '/IMPERSONAR
    Imitate                 '/MIMETIZAR
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest
    ExitDestroy             '/DE
    ToggleCentinelActivated '/CENTINELAACTIVADO
    SearchNpc               '/BUSCAR
    SearchObj               '/BUSCAR
    LimpiarMundo            '/LIMPIARMUNDO
End Enum

'
' Mensajes
'

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_22 As String = "!"

Public Enum eMessages
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    ResuscitationSafeOff
    ResuscitationSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    HaveKilledUser
    UserKill
    EarnExp
    GoHome
    CancelGoHome
    FinishHome
    
    '//Nuevos mensajes
    UserMuerto
    NpcInmune
    
    Hechizo_HechiceroMSG_NOMBRE
    Hechizo_HechiceroMSG_ALGUIEN
    Hechizo_HechiceroMSG_CRIATURA
 
    Hechizo_PropioMSG
    Hechizo_TargetMSG
End Enum

'Inventario
Type Inventory
    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Long
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer 'Budi
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Integer
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Type tItemsConstruibles
    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Long
End Type

Type tItemCrafteo
    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    Amount As Integer
End Type

Type tItemArtesano
    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    
    ItemsCrafteo() As tItemCrafteo
End Type

Public Nombres As Boolean

Public UserHechizos(1 To MAXHECHI) As Integer

Public Type PjCuenta
    Nombre      As String
    Head        As Integer
    Body        As Integer
    shield      As Byte
    helmet      As Byte
    weapon      As Byte
    Mapa        As Integer
    Class       As Byte
    Race        As Byte
    Map         As Integer
    Level       As Byte
    Gold        As Long
    Criminal    As Boolean
    Dead        As Boolean
    GameMaster  As Boolean
End Type

Public cPJ() As PjCuenta

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public AccountName As String
Public AccountPassword As String
Public AccountHash As String
Public NumberOfCharacters As Byte
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public UserParalizado As Boolean
Public UserInvisible As Boolean
Public UserNavegando As Boolean
Public UserEquitando As Boolean
Public UserEvento As Boolean
Public UserHogar As eCiudad

Public UserFuerza As Byte
Public UserAgilidad As Byte

Public UserWeaponEqpSlot As Byte
Public UserArmourEqpSlot As Byte
Public UserHelmEqpSlot As Byte
Public UserShieldEqpSlot As Byte

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean
Public MirandoCarpinteria As Boolean
Public MirandoHerreria As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 20
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 12
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Byte
Public PorcentajeSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer

Public UsingSkill As Integer

Public pingTime As Long

Public EsPartyLeader As Boolean

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearCuenta = 4
    CambiarContrasena = 5
    ObtenerDatosServer = 6
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience = 2
    eo_Body = 3
    eo_Head = 4
    eo_CiticensKilled = 5
    eo_CriminalsKilled = 6
    eo_Level = 7
    eo_Class = 8
    eo_Skills = 9
    eo_SkillPointsLeft = 10
    eo_Nobleza = 11
    eo_Asesino = 12
    eo_Sex = 13
    eo_Raza = 14
    eo_addGold = 15
    eo_Vida = 16
    eo_Poss = 17
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param CASA dentro de una casa de las que se compran, para evitar limpiar items
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    nada = 0
    BAJOTECHO = 1
    CASA = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server
Public UserMap As Integer
Public nameMap As String

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

'******Mouse Cursor*********
'Esto es para poder usar iconos de mouse .ani
'https://www.gs-zone.org/temas/cursor-ani.45555/#post-375757
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Public Const GLC_HCURSOR = (-12)
Public hSwapCursor As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
'******End Mouse Cursor****

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Long
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
End Type

Public EsperandoLevel As Boolean

' Tipos de mensajes
Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

' Indica los privilegios para visualizar los diferentes foros
Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

' Indica el tipo de foro
Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

' Limite de posts
Public Const MAX_STICKY_POST As Byte = 5
Public Const MAX_GENERAL_POST As Byte = 30
Public Const STICKY_FORUM_OFFSET As Byte = 50

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    StickyAuthor(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
    GeneralAuthor(1 To MAX_GENERAL_POST) As String
End Type

' 1 foro general y 2 faccionarios
Public Foros(0 To 2) As tForo

' Forum info handler
Public clsForos As clsForum

'FragShooter variables
Public FragShooterCapturePending As Boolean
Public FragShooterNickname As String
Public FragShooterKilledSomeone As Boolean


Public Traveling As Boolean

Public bShowGuildNews As Boolean
Public GuildNames() As String
Public GuildMembers() As String

Public Const OFFSET_HEAD As Integer = -34

Public Enum eSMType
    sResucitation
    sSafemode
    mSpells
    mWork
End Enum

Public Const SM_CANT As Byte = 4
Public SMStatus(SM_CANT) As Boolean

'Hardcoded grhs and items
Public Const GRH_INI_SM As Long = 4978

Public Const ORO_INDEX As Long = 12
Public Const ORO_GRH As Long = 511

Public Const LH_GRH As Long = 724
Public Const LP_GRH As Long = 725
Public Const LO_GRH As Long = 723

Public Const MADERA_GRH As Long = 550
Public Const MADERA_ELFICA_GRH As Long = 1999

Public picMouseIcon As Picture

Public Enum eMoveType
    Inventory = 1
    Bank
End Enum

'/////OPTIMIZACION DE STRINGS////////
Public NumHechizos As Byte
Public Hechizos() As tHechizos
 
Public Type tHechizos
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    ManaRequerida As Integer
    SkillRequerido As Byte
    EnergiaRequerida As Integer
 
    '//Mensajes
    HechiceroMsg As String
    PropioMsg As String
    TargetMsg As String
End Type

'MundoSeleccionado desde la propiedad Mundo en sinfo.dat / World selected from sinfo.dat file
Public MundoSeleccionado As String

' * Configuracion de estilos de controles
Public Const uAOButton_bEsquina As String = "bEsquina.bmp"
Public Const uAOButton_bFondo As String = "bFondo.bmp"
Public Const uAOButton_bHorizontal As String = "bHorizontal.bmp"
Public Const uAOButton_bVertical As String = "bVertical.bmp"
Public Const uAOButton_cCheckbox As String = "cCheckbox.bmp" ' Grande
Public Const uAOButton_cCheckboxSmall As String = "cCheckboxSmall.bmp" ' Chico
' * Configuracion de estilo de controles

Public JsonTips As Object

'Nivel Maximo
Public STAT_MAXELV As Byte
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer

Public UserParalizadoSegundosRestantes As Integer
Public UserInvisibleSegundosRestantes As Integer

Public QuantityServers As Integer
Public IpApiEnabled As Boolean
