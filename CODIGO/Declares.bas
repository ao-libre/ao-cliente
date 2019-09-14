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
Public lastKeys As clsArrayList
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
                
Public CustomKeys As clsCustomKeys
Public CustomMessages As clsCustomMessages

Public incomingData As CsBuffer
Public outgoingData As CsBuffer

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
    Ping As String
    Country As String
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
Public Const MAX_INVENTORY_SLOTS As Byte = 25
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
    Pirat = 12    'Pirata
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
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLena = 14
    otFogata = 15
    otEscudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otArbolElfico = 36
    otMochilas = 37
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
    GMMessage = 1                   '/GMSG
    showName = 2                    '/SHOWNAME
    OnlineRoyalArmy = 3             '/ONLINEREAL
    OnlineChaosLegion = 4           '/ONLINECAOS
    GoNearby = 5                    '/IRCERCA
    Comment = 6                     '/REM
    serverTime = 7                  '/HORA
    Where = 8                       '/DONDE
    CreaturesInMap = 9              '/NENE
    WarpMeToTarget = 10             '/TELEPLOC
    WarpChar = 11                   '/TELEP
    Silence = 12                    '/SILENCIAR
    SOSShowList = 13                '/SHOW SOS
    SOSRemove = 14                  'SOSDONE
    GoToChar = 15                   '/IRA
    invisible = 16                  '/INVISIBLE
    GMPanel = 17                    '/PANELGM
    RequestUserList = 18            'LISTUSU
    Working = 19                    '/TRABAJANDO
    Hiding = 20                     '/OCULTANDO
    Jail = 21                       '/CARCEL
    KillNPC = 22                    '/RMATA
    WarnUser = 23                   '/ADVERTENCIA
    EditChar = 24                   '/MOD
    RequestCharInfo = 25            '/INFO
    RequestCharStats = 26           '/STAT
    RequestCharGold = 27            '/BAL
    RequestCharInventory = 28       '/INV
    RequestCharBank = 29            '/BOV
    RequestCharSkills = 30          '/SKILLS
    ReviveChar = 31                 '/REVIVIR
    OnlineGM = 32                   '/ONLINEGM
    OnlineMap = 33                  '/ONLINEMAP
    Forgive = 34                    '/PERDON
    Kick = 35                       '/ECHAR
    Execute = 36                    '/EJECUTAR
    banChar = 37                    '/BAN
    UnbanChar = 38                  '/UNBAN
    NPCFollow = 39                  '/SEGUIR
    SummonChar = 40                 '/SUM
    SpawnListRequest = 41           '/CC
    SpawnCreature = 42              'SPA
    ResetNPCInventory = 43          '/RESETINV
    ServerMessage = 44              '/RMSG
    nickToIP = 45                   '/NICK2IP
    IPToNick = 46                   '/IP2NICK
    GuildOnlineMembers = 47         '/ONCLAN
    TeleportCreate = 48             '/CT
    TeleportDestroy = 49            '/DT
    RainToggle = 50                 '/LLUVIA
    SetCharDescription = 51         '/SETDESC
    ForceMIDIToMap = 52             '/FORCEMIDIMAP
    ForceWAVEToMap = 53             '/FORCEWAVMAP
    RoyalArmyMessage = 54           '/REALMSG
    ChaosLegionMessage = 55         '/CAOSMSG
    CitizenMessage = 56             '/CIUMSG
    CriminalMessage = 57            '/CRIMSG
    TalkAsNPC = 58                  '/TALKAS
    DestroyAllItemsInArea = 59      '/MASSDEST
    AcceptRoyalCouncilMember = 60   '/ACEPTCONSE
    AcceptChaosCouncilMember = 61   '/ACEPTCONSECAOS
    ItemsInTheFloor = 62            '/PISO
    MakeDumb = 63                   '/ESTUPIDO
    MakeDumbNoMore = 64             '/NOESTUPIDO
    dumpIPTables = 65               '/DUMPSECURITY
    CouncilKick = 66                '/KICKCONSE
    SetTrigger = 67                 '/TRIGGER
    AskTrigger = 68                 '/TRIGGER with no args
    BannedIPList = 69               '/BANIPLIST
    BannedIPReload = 70             '/BANIPRELOAD
    GuildMemberList = 71            '/MIEMBROSCLAN
    GuildBan = 72                   '/BANCLAN
    BanIP = 73                      '/BANIP
    UnbanIP = 74                    '/UNBANIP
    CreateItem = 75                 '/CI
    DestroyItems = 76               '/DEST
    ChaosLegionKick = 77            '/NOCAOS
    RoyalArmyKick = 78              '/NOREAL
    ForceMIDIAll = 79               '/FORCEMIDI
    ForceWAVEAll = 80               '/FORCEWAV
    RemovePunishment = 81           '/BORRARPENA
    TileBlockedToggle = 82          '/BLOQ
    KillNPCNoRespawn = 83           '/MATA
    KillAllNearbyNPCs = 84          '/MASSKILL
    LastIP = 85                     '/LASTIP
    ChangeMOTD = 86                 '/MOTDCAMBIA
    SetMOTD = 87                    'ZMOTD
    SystemMessage = 88              '/SMSG
    CreateNPC = 89                  '/ACC y /RACC
    ImperialArmour = 90             '/AI1 - 4
    ChaosArmour = 91                '/AC1 - 4
    NavigateToggle = 92             '/NAVE
    ServerOpenToUsersToggle = 93    '/HABILITAR
    TurnOffServer = 94              '/APAGAR
    TurnCriminal = 95               '/CONDEN
    ResetFactions = 96              '/RAJAR
    RemoveCharFromGuild = 97        '/RAJARCLAN
    RequestCharMail = 98            '/LASTEMAIL
    AlterPassword = 99             '/APASS
    AlterMail = 100                 '/AEMAIL
    AlterName = 101                  '/ANAME
    DoBackUp = 102                  '/DOBACKUP
    ShowGuildMessages = 103         '/SHOWCMSG
    SaveMap = 104                   '/GUARDAMAPA
    ChangeMapInfoPK = 105           '/MODMAPINFO PK
    ChangeMapInfoBackup = 106       '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted = 107   '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic = 108      '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi = 109       '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu = 110       '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand = 111         '/MODMAPINFO TERRENO
    ChangeMapInfoZone = 112         '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc = 113     '/MODMAPINFO ROBONPCm
    ChangeMapInfoNoOcultar = 114    '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar = 115    '/MODMAPINFO INVOCARSINEFECTO
    SaveChars = 116                 '/GRABAR
    CleanSOS = 117                  '/BORRAR SOS
    ShowServerForm = 118            '/SHOW INT
    night = 119                     '/NOCHE
    KickAllChars = 120              '/ECHARTODOSPJS
    ReloadNPCs = 121                '/RELOADNPCS
    ReloadServerIni = 122           '/RELOADSINI
    ReloadSpells = 123              '/RELOADHECHIZOS
    ReloadObjects = 124             '/RELOADOBJ
    Restart = 125                   '/REINICIAR
    ResetAutoUpdate = 126           '/AUTOUPDATE
    ChatColor = 127                 '/CHATCOLOR
    Ignored = 128                   '/IGNORADO
    CheckSlot = 129                 '/SLOT
    SetIniVar = 130                 '/SETINIVAR LLAVE CLAVE VALOR
    CreatePretorianClan = 131       '/CREARPRETORIANOS
    RemovePretorianClan = 132       '/ELIMINARPRETORIANOS
    EnableDenounces = 133           '/DENUNCIAS
    ShowDenouncesList = 134         '/SHOW DENUNCIAS
    MapMessage = 135                '/MAPMSG
    SetDialog = 136                 '/SETDIALOG
    Impersonate = 137               '/IMPERSONAR
    Imitate = 138                   '/MIMETIZAR
    RecordAdd = 139
    RecordRemove = 140
    RecordAddObs = 141
    RecordListRequest = 142
    RecordDetailsRequest = 143
    ExitDestroy = 144
    ToggleCentinelActivated = 145   '/CENTINELAACTIVADO
    SearchNpc = 146                 '/BUSCAR
    SearchObj = 147                 '/BUSCAR
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
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
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
    GrhIndex As Integer
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
    GrhIndex As Integer
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Integer
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
Public PrimeraVez As Boolean
Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
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
Public Flags() As Integer

Public UsingSkill As Integer

Public pingTime As Long

Public EsPartyLeader As Boolean

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearCuenta = 4
    CambiarContrasena = 5
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
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
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
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
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
Public Const GRH_INI_SM As Integer = 4978

Public Const ORO_INDEX As Integer = 12
Public Const ORO_GRH As Integer = 511

Public Const LH_GRH As Integer = 724
Public Const LP_GRH As Integer = 725
Public Const LO_GRH As Integer = 723

Public Const MADERA_GRH As Integer = 550
Public Const MADERA_ELFICA_GRH As Integer = 1999

Public picMouseIcon As Picture

Public Enum eMoveType
    Inventory = 1
    Bank
End Enum

Public Const MP3_INITIAL_INDEX As Integer = 1000

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


