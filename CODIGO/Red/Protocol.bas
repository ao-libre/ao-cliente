Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martin Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    logged = 1                  ' LOGGED
    RemoveDialogs = 2           ' QTDL
    RemoveCharDialog = 3        ' QDL
    NavigateToggle = 4          ' NAVEG
    Disconnect = 5              ' FINOK
    CommerceEnd = 6             ' FINCOMOK
    BankEnd = 7                 ' FINBANOK
    CommerceInit = 8            ' INITCOM
    BankInit = 9                ' INITBANCO
    UserCommerceInit = 10        ' INITCOMUSU
    UserCommerceEnd = 11         ' FINCOMUSUOK
    UserOfferConfirm = 12
    CommerceChat = 13
    UpdateSta = 14               ' ASS
    UpdateMana = 15             ' ASM
    UpdateHP = 16                ' ASH
    UpdateGold = 17              ' ASG
    UpdateBankGold = 18
    UpdateExp = 19               ' ASE
    ChangeMap = 20               ' CM
    PosUpdate = 21              ' PU
    ChatOverHead = 22            ' ||
    ConsoleMsg = 23              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat = 24               ' |+
    ShowMessageBox = 25          ' !!
    UserIndexInServer = 26       ' IU
    UserCharIndexInServer = 27   ' IP
    CharacterCreate = 28         ' CC
    CharacterRemove = 29         ' BP
    CharacterChangeNick = 30
    CharacterMove = 31           ' MP, +, * and _ '
    ForceCharMove = 32
    CharacterChange = 33         ' CP
    HeadingChange = 34
    ObjectCreate = 35            ' HO
    ObjectDelete = 36            ' BO
    BlockPosition = 37           ' BQ
    PlayMp3 = 38
    PlayMIDI = 39                ' TM
    PlayWave = 40                ' TW
    guildList = 41               ' GL
    AreaChanged = 42             ' CA
    PauseToggle = 43             ' BKW
    RainToggle = 44              ' LLU
    CreateFX = 45                ' CFX
    UpdateUserStats = 46         ' EST
    ChangeInventorySlot = 47     ' CSI
    ChangeBankSlot = 48          ' SBO
    ChangeSpellSlot = 49         ' SHS
    Atributes = 50               ' ATR
    BlacksmithWeapons = 51       ' LAH
    BlacksmithArmors = 52        ' LAR
    InitCarpenting = 53          ' OBR
    RestOK = 54                  ' DOK
    ErrorMsg = 55                ' ERR
    Blind = 56                   ' CEGU
    Dumb = 57                    ' DUMB
    ShowSignal = 58              ' MCAR
    ChangeNPCInventorySlot = 59  ' NPCI
    UpdateHungerAndThirst = 60   ' EHYS
    Fame = 61                    ' FAMA
    MiniStats = 62               ' MEST
    LevelUp = 63                 ' SUNI
    AddForumMsg = 64             ' FMSG
    ShowForumForm = 65           ' MFOR
    SetInvisible = 66            ' NOVER
    DiceRoll = 67                ' DADOS
    MeditateToggle = 68          ' MEDOK
    BlindNoMore = 69             ' NSEGUE
    DumbNoMore = 70              ' NESTUP
    SendSkills = 71              ' SKILLS
    TrainerCreatureList = 72     ' LSTCRI
    guildNews = 73               ' GUILDNE
    OfferDetails = 74            ' PEACEDE & ALLIEDE
    AlianceProposalsList = 75    ' ALLIEPR
    PeaceProposalsList = 76      ' PEACEPR
    CharacterInfo = 77           ' CHRINFO
    GuildLeaderInfo = 78         ' LEADERI
    GuildMemberInfo = 79
    GuildDetails = 80            ' CLANDET
    ShowGuildFundationForm = 81  ' SHOWFUN
    ParalizeOK = 82              ' PARADOK
    ShowUserRequest = 83         ' PETICIO
    ChangeUserTradeSlot = 84     ' COMUSUINV
    SendNight = 85               ' NOC
    Pong = 86
    UpdateTagAndStatus = 87
    
    'GM =  messages
    SpawnList = 88               ' SPL
    ShowSOSForm = 89             ' MSOS
    ShowMOTDEditionForm = 90     ' ZMOTD
    ShowGMPanelForm = 91         ' ABPANEL
    UserNameList = 92            ' LISTUSU
    ShowDenounces = 93
    RecordList = 94
    RecordDetails = 95
    
    ShowGuildAlign = 96
    ShowPartyForm = 97
    UpdateStrenghtAndDexterity = 98
    UpdateStrenght = 99
    UpdateDexterity = 100
    AddSlots = 101
    MultiMessage = 102
    StopWorking = 103
    CancelOfferItem = 104
    PalabrasMagicas = 105
    PlayAttackAnim = 106
    FXtoMap = 107
    AccountLogged = 108         'CHOTS | Accounts
    SearchList = 109
    QuestDetails = 110
    QuestListSend = 111
    CreateDamage = 112          ' CDMG
    UserInEvent = 113
    renderMsg = 114
    DeletedChar = 115
    EquitandoToggle = 116
    EnviarDatosServer = 117
    InitCraftman = 118
    EnviarListDeAmigos = 119
    SeeInProcess = 120
    ShowProcess = 121
    Proyectil = 122
End Enum

Private Enum ClientPacketID
    LoginExistingChar = 1           'OLOGIN
    ThrowDices = 2                  'TIRDAD
    LoginNewChar = 3                'NLOGIN
    Talk = 4                        ';
    Yell = 5                        '-
    Whisper = 6                     '\
    Walk = 7                        'M
    RequestPositionUpdate = 8       'RPU
    Attack = 9                      'AT
    PickUp = 10                     'AG
    SafeToggle = 11                 '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle = 12
    RequestGuildLeaderInfo = 13     'GLINFO
    RequestAtributes = 14           'ATR
    RequestFame = 15                'FAMA
    RequestSkills = 16              'ESKI
    RequestMiniStats = 17           'FEST
    CommerceEnd = 18                'FINCOM
    UserCommerceEnd = 19            'FINCOMUSU
    UserCommerceConfirm = 20
    CommerceChat = 21
    BankEnd = 22                    'FINBAN
    UserCommerceOk = 23             'COMUSUOK
    UserCommerceReject = 24         'COMUSUNO
    Drop = 25                       'TI
    CastSpell = 26                  'LH
    LeftClick = 27                  'LC
    DoubleClick = 28                'RC
    Work = 29                       'UK
    UseSpellMacro = 30              'UMH
    UseItem = 31                    'USA
    CraftBlacksmith = 32            'CNS
    CraftCarpenter = 33             'CNC
    WorkLeftClick = 34              'WLC
    CreateNewGuild = 35             'CIG
    sadasdA = 36
    EquipItem = 37                  'EQUI
    ChangeHeading = 38              'CHEA
    ModifySkills = 39               'SKSE
    Train = 40                      'ENTR
    CommerceBuy = 41                'COMP
    BankExtractItem = 42            'RETI
    CommerceSell = 43               'VEND
    BankDeposit = 44                'DEPO
    ForumPost = 45                  'DEMSG
    MoveSpell = 46                  'DESPHE
    MoveBank = 47
    ClanCodexUpdate = 48            'DESCOD
    UserCommerceOffer = 49          'OFRECER
    GuildAcceptPeace = 50           'ACEPPEAT
    GuildRejectAlliance = 51        'RECPALIA
    GuildRejectPeace = 52           'RECPPEAT
    GuildAcceptAlliance = 53        'ACEPALIA
    GuildOfferPeace = 54            'PEACEOFF
    GuildOfferAlliance = 55         'ALLIEOFF
    GuildAllianceDetails = 56       'ALLIEDET
    GuildPeaceDetails = 57          'PEACEDET
    GuildRequestJoinerInfo = 58     'ENVCOMEN
    GuildAlliancePropList = 59      'ENVALPRO
    GuildPeacePropList = 60         'ENVPROPP
    GuildDeclareWar = 61            'DECGUERR
    GuildNewWebsite = 62            'NEWWEBSI
    GuildAcceptNewMember = 63       'ACEPTARI
    GuildRejectNewMember = 64       'RECHAZAR
    GuildKickMember = 65            'ECHARCLA
    GuildUpdateNews = 66            'ACTGNEWS
    GuildMemberInfo = 67            '1HRINFO<
    GuildOpenElections = 68         'ABREELEC
    GuildRequestMembership = 69     'SOLICITUD
    GuildRequestDetails = 70        'CLANDETAILS
    Online = 71                     '/ONLINE
    Quit = 72                       '/SALIR
    GuildLeave = 73                 '/SALIRCLAN
    RequestAccountState = 74        '/BALANCE
    PetStand = 75                   '/QUIETO
    PetFollow = 76                  '/ACOMPANAR
    ReleasePet = 77                 '/LIBERAR
    TrainList = 78                  '/ENTRENAR
    Rest = 79                       '/DESCANSAR
    Meditate = 80                   '/MEDITAR
    Resucitate = 81                 '/RESUCITAR
    Heal = 82                       '/CURAR
    Help = 83                       '/AYUDA
    RequestStats = 84               '/EST
    CommerceStart = 85              '/COMERCIAR
    BankStart = 86                  '/BOVEDA
    Enlist = 87                     '/ENLISTAR
    Information = 88                '/INFORMACION
    Reward = 89                     '/RECOMPENSA
    RequestMOTD = 90                '/MOTD
    UpTime = 91                     '/UPTIME
    PartyLeave = 92                 '/SALIRPARTY
    PartyCreate = 93                '/CREARPARTY
    PartyJoin = 94                  '/PARTY
    Inquiry = 95                    '/ENCUESTA ( with no params )
    GuildMessage = 96               '/CMSG
    PartyMessage = 97               '/PMSG
    GuildOnline = 98                '/ONLINECLAN
    PartyOnline = 99                '/ONLINEPARTY
    CouncilMessage = 100            '/BMSG
    RoleMasterRequest = 101         '/ROL
    GMRequest = 102                 '/GM
    bugReport = 103                 '/_BUG
    ChangeDescription = 104         '/DESC
    GuildVote = 105                 '/VOTO
    Punishments = 106               '/PENAS
    ChangePassword = 107            '/CONTRASENA
    Gamble = 108                    '/APOSTAR
    InquiryVote = 109               '/ENCUESTA ( with parameters )
    LeaveFaction = 110              '/RETIRAR ( with no arguments )
    BankExtractGold = 111           '/RETIRAR ( with arguments )
    BankDepositGold = 112           '/DEPOSITAR
    Denounce = 113                  '/DENUNCIAR
    GuildFundate = 114              '/FUNDARCLAN
    GuildFundation = 115
    PartyKick = 116                 '/ECHARPARTY
    PartySetLeader = 117            '/PARTYLIDER
    PartyAcceptMember = 118         '/ACCEPTPARTY
    Ping = 119                      '/PING
    RequestPartyForm = 120
    ItemUpgrade = 121
    GMCommands = 122
    InitCrafting = 123
    Home = 124
    ShowGuildNews = 125
    ShareNpc = 126                  '/COMPARTIR
    StopSharingNpc = 127
    Consultation = 128
    moveItem = 129
    LoginExistingAccount = 130
    LoginNewAccount = 131
    CentinelReport = 132            '/CENTINELA
    Ecvc = 133
    Acvc = 134
    IrCvc = 135
    DragAndDropHechizos = 136       'HECHIZOS
    Quest = 137                     '/QUEST
    QuestAccept = 138
    QuestListRequest = 139
    QuestDetailsRequest = 140
    QuestAbandon = 141
    CambiarContrasena = 142
    FightSend = 143
    FightAccept = 144
    CloseGuild = 145                '/CERRARCLAN
    Discord = 146                   '/DISCORD
    DeleteChar = 147
    ObtenerDatosServer = 148
    CraftsmanCreate = 149
    AddAmigos = 150
    DelAmigos = 151
    OnAmigos = 152
    MsgAmigos = 153
    Lookprocess = 154
    SendProcessList = 155
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK = 0
    FONTTYPE_FIGHT = 1
    FONTTYPE_WARNING = 2
    FONTTYPE_INFO = 3
    FONTTYPE_INFOBOLD = 4
    FONTTYPE_EJECUCION = 5
    FONTTYPE_PARTY = 6
    FONTTYPE_VENENO = 7
    FONTTYPE_GUILD = 8
    FONTTYPE_SERVER = 9
    FONTTYPE_GUILDMSG = 10
    FONTTYPE_CONSEJO = 11
    FONTTYPE_CONSEJOCAOS = 12
    FONTTYPE_CONSEJOVesA = 13
    FONTTYPE_CONSEJOCAOSVesA = 14
    FONTTYPE_CENTINELA = 15
    FONTTYPE_GMMSG = 16
    FONTTYPE_GM = 17
    FONTTYPE_CITIZEN = 18
    FONTTYPE_CONSE = 19
    FONTTYPE_DIOS = 20
    FONTTYPE_CRIMINAL = 21
End Enum

Public FontTypes(21) As tFont

Public Sub Connect(ByVal Modo As E_MODO)
    '*********************************************************************
    'Author: Jopi
    'Conexion al servidor mediante la API de Windows.
    '*********************************************************************
        
    'Evitamos enviar multiples peticiones de conexion al servidor.
    frmConnect.btnConectarse.Enabled = False
        
    'Primero lo cerramos, para evitar errores.
    If frmMain.Client.State <> (sckClosed Or sckConnecting) Then
        frmMain.Client.CloseSck
        DoEvents
    End If
    
    EstadoLogin = Modo

    'Usamos la API de Windows
    Call frmMain.Client.Connect(CurServerIp, CurServerPort)

    'Vuelvo a activar el boton.
    frmConnect.btnConectarse.Enabled = True
End Sub

''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 204
        .Green = 255
        .Blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .Green = 102
        .Blue = 102
        .bold = 1
        .italic = 0
    End With

    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 255
        .Green = 255
        .Blue = 102
        .bold = 1
        .italic = 0
    End With

    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 255
        .Green = 204
        .Blue = 153
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 255
        .Green = 204
        .Blue = 153
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 255
        .Green = 0
        .Blue = 127
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 252
        .Green = 203
        .Blue = 130
    End With

    With FontTypes(FontTypeNames.FONTTYPE_VENENO)
        .Red = 128
        .Green = 255
        .Blue = 0
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 205
        .Green = 101
        .Blue = 236
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_SERVER)
        .Red = 250
        .Green = 150
        .Blue = 237
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .Red = 130
        .Green = 130
        .Blue = 255
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .Red = 255
        .Green = 60
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .Green = 200
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .Red = 255
        .Green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Red = 240
        .Green = 230
        .Blue = 140
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .Red = 30
        .Green = 255
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .Red = 78
        .Green = 78
        .Blue = 252
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .Red = 30
        .Green = 150
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CRIMINAL)
        .Red = 224
        .Green = 52
        .Blue = 17
        .bold = 1
    End With
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error Resume Next

    Dim Packet As Long: Packet = CLng(incomingData.PeekByte())
    
    'Debug.Print Packet
    
    Select Case Packet
            
        Case ServerPacketID.logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
            
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
            
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm

        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
            
        Case ServerPacketID.UpdateBankGold
            Call HandleUpdateBankGold

        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
            
        Case ServerPacketID.HeadingChange
            Call HandleHeadingChange
            
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition

        Case ServerPacketID.PlayMp3
            Call HandlePlayMP3
        
        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              ' LLU
            Call HandleRainToggle
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats

        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       ' LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        ' LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.InitCarpenting          ' OBR
            Call HandleInitCarpenting
            
        Case ServerPacketID.InitCraftman
            Call HandleInitCraftman
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.Fame                    ' FAMA
            Call HandleFame
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.DiceRoll                ' DADOS
            Call HandleDiceRoll
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews               ' GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      ' PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest

        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
            
        Case ServerPacketID.SendNight               ' NOC
            Call HandleSendNight
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
        
        Case ServerPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo
            
        Case ServerPacketID.PalabrasMagicas
            Call HandlePalabrasMagicas
            
        Case ServerPacketID.PlayAttackAnim
            Call HandleAttackAnim
            
        Case ServerPacketID.FXtoMap
            Call HandleFXtoMap
        
        'CHOTS | Accounts
        Case ServerPacketID.AccountLogged
            Call HandleAccountLogged
            
        Case ServerPacketID.SearchList              '/BUSCAR
            Call HandleSearchList

        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails

        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend

        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
            
        Case ServerPacketID.ShowDenounces
            Call HandleShowDenounces
            
        Case ServerPacketID.RecordDetails
            Call HandleRecordDetails
            
        Case ServerPacketID.RecordList
            Call HandleRecordList
            
        Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
            
        Case ServerPacketID.ShowGuildAlign
            Call HandleShowGuildAlign
        
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem
            
        Case ServerPacketID.CreateDamage            ' CDMG
            Call HandleCreateDamage
    
        Case ServerPacketID.UserInEvent
            Call HandleUserInEvent
            
        Case ServerPacketID.renderMsg
            Call HandleRenderMsg

        Case ServerPacketID.DeletedChar             ' BORRA USUARIO
            Call HandleDeletedChar

        Case ServerPacketID.EquitandoToggle         'Para las monturas
            Call HandleEquitandoToggle

        Case ServerPacketID.EnviarDatosServer       'Para obtener info del server en la lista de servers
            Call HandleEnviarDatosServer
            
        Case ServerPacketID.EnviarListDeAmigos
            Call HandleEnviarListDeAmigos

        Case ServerPacketID.SeeInProcess
            Call HandleSeeInProcess
            
        Case ServerPacketID.ShowProcess
            Call HandleShowProcess
            
        Case ServerPacketID.Proyectil
            Call HandleProyectil

        Case Else
            'ERROR : Abort!
            Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.Length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Call Err.Clear
        Call HandleIncomingData
    End If
    
End Sub

Public Sub HandleMultiMessage()

    '***************************************************
    'Author: Unknown
    'Last Modification: 11/16/2010
    ' 09/28/2010: C4b3z0n - Ahora se le saco la "," a los minutos de distancia del /hogar, ya que a veces quedaba "12,5 minutos y 30segundos"
    ' 09/21/2010: C4b3z0n - Now the fragshooter operates taking the screen after the change of killed charindex to ghost only if target charindex is visible to the client, else it will take screenshot like before.
    ' 11/16/2010: Amraphen - Recoded how the FragShooter works.
    ' 04/12/2019: jopiortiz - Carga de mensajes desde JSON.
    '***************************************************
    Dim BodyPart As Byte

    Dim Dano As Integer

    Dim SpellIndex As Integer

    Dim Nombre     As String
    
    With incomingData
        Call .ReadByte
    
        Select Case .ReadByte

            Case eMessages.NPCSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("TEXTO"), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(1), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(2), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(3), _
                        True, False, True)
        
            Case eMessages.NPCKillUser
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.BlockedWithShieldUser
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.BlockedWithShieldOther
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.UserSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.SafeModeOn
                Call frmMain.ControlSM(eSMType.sSafemode, True)
        
            Case eMessages.SafeModeOff
                Call frmMain.ControlSM(eSMType.sSafemode, False)
        
            Case eMessages.ResuscitationSafeOff
                Call frmMain.ControlSM(eSMType.sResucitation, False)
         
            Case eMessages.ResuscitationSafeOn
                Call frmMain.ControlSM(eSMType.sResucitation, True)
        
            Case eMessages.NobilityLost
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("TEXTO"), _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(3), _
                                        False, False, True)
        
            Case eMessages.CantUseWhileMeditating
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("TEXTO"), _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(3), _
                                        False, False, True)
        
            Case eMessages.NPCHitUser

                Select Case incomingData.ReadByte()

                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("TEXTO") & CStr(incomingData.ReadInteger() & "!!"), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(3), _
                            True, False, True)

                End Select
        
            Case eMessages.UserHitNPC
                Dim MsgHitNpc As String
                    MsgHitNpc = JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("TEXTO")
                    MsgHitNpc = Replace$(MsgHitNpc, "VAR_DANO", CStr(incomingData.ReadLong()))
                    
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        MsgHitNpc, _
                                        JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(3), _
                                        True, False, True)
        
            Case eMessages.UserAttackedSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        charlist(incomingData.ReadInteger()).Nombre & JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("TEXTO"), _
                                        JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(3), _
                                        True, False, True)
        
            Case eMessages.UserHittedByUser

                Dim AttackerName As String
            
                AttackerName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Dano = incomingData.ReadInteger()
            
                Select Case BodyPart

                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                        AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(1), _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(2), _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(3), _
                        True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(3), _
                            True, False, True)

                End Select
        
            Case eMessages.UserHittedUser

                Dim VictimName As String
            
                VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Dano = incomingData.ReadInteger()
            
                Select Case BodyPart

                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(3), _
                            True, False, True)

                End Select
        
            Case eMessages.WorkRequestTarget
                
                UsingSkill = incomingData.ReadByte()
            
                frmMain.MousePointer = 2 'vbCrosshair
            
                Select Case UsingSkill

                    Case Magia
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(3))
                
                    Case Pesca
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(3))
                
                    Case Robar
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(3))
                
                    Case Talar
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(3))
                
                    Case Mineria
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(3))
                
                    Case FundirMetal
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(3))
                
                    Case Proyectiles
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(3))

                End Select

            Case eMessages.HaveKilledUser

                Dim KilledUser As Integer
                Dim Exp        As Long
                Dim MensajeExp As String
            
                KilledUser = .ReadInteger
                Exp = .ReadLong
            
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("TEXTO") & charlist(KilledUser).Nombre & MENSAJE_22, _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(3), _
                    True, False)
                
                ' Para mejor lectura
                MensajeExp = JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("TEXTO") 'String original
                MensajeExp = Replace$(MensajeExp, "VAR_EXP_GANADA", Exp) 'Parte a reemplazar
                
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                    MensajeExp, _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(3), _
                                    True, False)
            
                'Sacamos un screenshot si esta activado el FragShooter:
                If ClientSetup.bKill And ClientSetup.bActive Then
                    
                    If Exp \ 2 > ClientSetup.byMurderedLevel Then
                        FragShooterNickname = charlist(KilledUser).Nombre
                        FragShooterKilledSomeone = True
                        FragShooterCapturePending = True
                    End If

                End If
            
            Case eMessages.UserKill

                Dim KillerUser As Integer
                    KillerUser = .ReadInteger
            
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                    charlist(KillerUser).Nombre & JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(3), _
                                    True, False)
            
                'Sacamos un screenshot si esta activado el FragShooter:
                If ClientSetup.bDie And ClientSetup.bActive Then
                    FragShooterNickname = charlist(KillerUser).Nombre
                    FragShooterKilledSomeone = False
                
                    FragShooterCapturePending = True

                End If
            
            Case eMessages.NPCKill
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        JsonLanguage.item("NPC_KILL").item("TEXTO"), _
                                        JsonLanguage.item("NPC_KILL").item("COLOR").item(1), _
                                        JsonLanguage.item("NPC_KILL").item("COLOR").item(2), _
                                        JsonLanguage.item("NPC_KILL").item("COLOR").item(3), _
                                        True, False)
            
            Case eMessages.EarnExp
                
                Dim ExpObtenida As Long: ExpObtenida = .ReadLong()
            
                Dim MENSAJE_HAS_GANADO_EXP As String
                    MENSAJE_HAS_GANADO_EXP = JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("TEXTO")
                    MENSAJE_HAS_GANADO_EXP = Replace$(MENSAJE_HAS_GANADO_EXP, "VAR_EXP_GANADA", ExpObtenida)
                                    
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        MENSAJE_HAS_GANADO_EXP, _
                                        JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(3), _
                                        True, False)
        
            Case eMessages.GoHome

                Dim Distance As Byte
                Dim Hogar    As String
                Dim tiempo   As Integer
                Dim msg      As String
                Dim msgGoHome As String
            
                Distance = .ReadByte
                tiempo = .ReadInteger
                Hogar = .ReadASCIIString
            
                If tiempo >= 60 Then
                    If tiempo Mod 60 = 0 Then
                        msg = tiempo / 60 & " " & JsonLanguage.item("MINUTOS").item("TEXTO") & "."
                    Else
                        msg = CInt(tiempo \ 60) & " " & JsonLanguage.item("MINUTOS").item("TEXTO") & " " & JsonLanguage.item("LETRA_Y").item("TEXTO") & " " & tiempo Mod 60 & " " & JsonLanguage.item("SEGUNDOS").item("TEXTO") & "."  'Agregado el CInt() asi el numero no es con , [C4b3z0n - 09/28/2010]

                    End If

                Else
                    msg = tiempo & " " & JsonLanguage.item("SEGUNDOS").item("TEXTO") & "."

                End If
                
                msgGoHome = JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("TEXTO") & msg
                msgGoHome = Replace$(msgGoHome, "VAR_DISTANCIA_MAPAS", Distance)
                msgGoHome = Replace$(msgGoHome, "VAR_MAPA_DESTINO", Hogar)
                
                Call ShowConsoleMsg(msgGoHome, _
                                    JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("COLOR").item(3), _
                                    True)
                Traveling = True

            Case eMessages.CancelGoHome
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("COLOR").item(3), _
                                    True)
                Traveling = False
                   
            Case eMessages.FinishHome
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_HOGAR").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(3))
                Traveling = False
            
            Case eMessages.UserMuerto
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
        
            Case eMessages.NpcInmune
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                    JsonLanguage.item("NPC_INMUNE").item("TEXTO"), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(1), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(2), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(3))
            
            Case eMessages.Hechizo_HechiceroMSG_NOMBRE
                SpellIndex = .ReadByte
                Nombre = .ReadASCIIString
         
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " " & Nombre & ".", 210, 220, 220)
         
            Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
                SpellIndex = .ReadByte
         
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " " & JsonLanguage.item("ALGUIEN").item("TEXTO") & ".", 210, 220, 220)
         
            Case eMessages.Hechizo_HechiceroMSG_CRIATURA
                SpellIndex = .ReadByte
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " la criatura.", 210, 220, 220)
         
            Case eMessages.Hechizo_PropioMSG
                SpellIndex = .ReadByte
                Call ShowConsoleMsg(Hechizos(SpellIndex).PropioMsg, 210, 220, 220)
         
            Case eMessages.Hechizo_TargetMSG
                SpellIndex = .ReadByte
                Nombre = .ReadASCIIString
                Call ShowConsoleMsg(Nombre & " " & Hechizos(SpellIndex).TargetMsg, 210, 220, 220)

        End Select

    End With

End Sub

''
' Handles the DeletedChar message.

Private Sub HandleDeletedChar()
'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    MsgBox ("El personaje se ha borrado correctamente. Por favor vuelve a iniciar sesion para ver el cambio")

    'Close connection
    Call CloseConnectionAndResetAllInfo
End Sub


Private Sub HandleLogged()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    #If AntiExternos Then
        Security.Redundance = incomingData.ReadByte()
    #End If

    ' Variable initialization
    UserClase = incomingData.ReadByte
    IntervaloInvi = incomingData.ReadLong
    
    EngineRun = True
    Nombres = True
    bRain = False
    
    'Set connected state
    Call SetConnected
    
    If bShowTutorial Then
        Call frmTutorial.Show(vbModeless)
    End If
    
    'Show tip
    If ClientSetup.MostrarTips = True Then
        frmtip.Visible = True
    End If

    'Show Keyboard configuration
    If ClientSetup.MostrarBindKeysSelection = True Then
        Call frmKeysConfigurationSelect.Show(vbModeless, frmMain)
        Call frmKeysConfigurationSelect.SetFocus
    End If
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    'Close and go to frmPanelAccount
    Call CloseConnectionAndResetAllInfo
    
End Sub

Private Sub CloseConnectionAndResetAllInfo()
    
    Call ResetAllInfo(False)
    
    If CheckUserData() Then
        frmMain.Visible = False
        Call Protocol.Connect(E_MODO.Normal)
    End If
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    
    'Hide form
    Unload frmComerciar
    
    'Reset vars
    Comerciando = False
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/01/20
'Los comerciantes te saludan con un sonido (Recox)
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = New clsGraphicalInventory
    Set InvComNpc = New clsGraphicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, MAX_INVENTORY_SLOTS, , , , , , , , True)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
            End With
        End If
    Next i
    
    ' Fill Npc inventory
    For i = 1 To MAX_NPC_INVENTORY_SLOTS
        If NPCInventory(i).OBJIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .OBJIndex, _
                .Amount, 0, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .name, .Incompatible)
            End With
        End If
    Next i
    
    'Set state and show form
    frmComerciar.Show , frmMain

    'Reproducimos el saludo del comerciante (Recox)
    Call Audio.PlayWave("comerciante" & RandomNumber(1, 9) & ".wav")
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim BankGold As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = New clsGraphicalInventory
    Set InvBanco(1) = New clsGraphicalInventory
    
    BankGold = incomingData.ReadLong
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    
    For i = 1 To MAX_INVENTORY_SLOTS
        With Inventario
            Call InvBanco(1).SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
        End With
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .OBJIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .name, .Incompatible)
        End With
    Next i
    
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.lblUserGld.Caption = BankGold
    
    frmBancoObj.Show , frmMain

    'Reproducimos el saludo del banquero (Recox)
    Call Audio.PlayWave("banquero" & RandomNumber(1, 7) & ".wav")
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    TradingUserName = incomingData.ReadASCIIString
    
    Set InvComUsu = New clsGraphicalInventory
    Set InvOfferComUsu(0) = New clsGraphicalInventory
    Set InvOfferComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(0) = New clsGraphicalInventory
    Set InvOroComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(2) = New clsGraphicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciarUsu.picInvComercio, MAX_INVENTORY_SLOTS)
    Call InvOfferComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    Call InvOfferComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
    Call InvOroComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.OBJIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .OBJIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
            End With
        End If
    Next i

    ' Inventarios de oro
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")


    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    With frmComerciarUsu
        ' Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg TradingUserName & JsonLanguage.item("MENSAJE_COMM_OFERTA_ACEPTA").item("TEXTO"), FontTypeNames.FONTTYPE_CONSE
    End With
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    Dim bWidth As Byte
    
    bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 83)
    
    frmMain.shpEnergia.Width = 83 - bWidth
    frmMain.shpEnergia.Left = 797 + (83 - frmMain.shpEnergia.Width)
    
    frmMain.shpEnergia.Visible = (bWidth <> 83)
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    
    Dim bWidth As Byte
    
    If UserMaxMAN > 0 Then _
        bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 90)
        
    frmMain.shpMana.Width = 90 - bWidth
    frmMain.shpMana.Left = 902 + (90 - frmMain.shpMana.Width)
    
    frmMain.shpMana.Visible = (bWidth <> 90)
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    
    Dim bWidth As Byte
    
    bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 90)
    
    frmMain.shpVida.Width = 90 - bWidth
    frmMain.shpVida.Left = 902 + (90 - frmMain.shpVida.Width)
    
    frmMain.shpVida.Visible = (bWidth <> 90)
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    
        UserEquitando = 0
        'Reseteo el Speed
        Call SetSpeedUsuario
    Else
        UserEstado = 0
    End If
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 09/21/10
'Last Modified By: C4b3z0n
'- 08/14/07: Tavo - Added GldLbl color variation depending on User Gold and Level
'- 09/21/10: C4b3z0n - Modified color change of gold ONLY if the player's level is greater than 12 (NOT newbie).
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    Call frmMain.SetGoldColor

    frmMain.GldLbl.Caption = UserGLD
End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
'***************************************************
'Autor: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()

    frmMain.UpdateProgressExperienceLevelBar (UserExp)
End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblStrg.ForeColor = getStrenghtColor()
    frmMain.lblDext.ForeColor = getDexterityColor()
    IntervaloDopas = incomingData.ReadLong
    TiempoDopas = (IntervaloDopas * 0.05) - 1
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblStrg.ForeColor = getStrenghtColor()
    IntervaloDopas = incomingData.ReadLong
    TiempoDopas = (IntervaloDopas * 0.05) - 1
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblDext.ForeColor = getDexterityColor()
    IntervaloDopas = incomingData.ReadLong
    TiempoDopas = (IntervaloDopas * 0.05) - 1
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    nameMap = incomingData.ReadASCIIString
    
    'TODO: Once on-the-fly editor is implemented check for map version before loading....
    'For now we just drop it
    Call incomingData.ReadInteger
    
    If FileExist(Game.path(Mapas) & "Mapa" & UserMap & ".map", vbNormal) Then
        Call SwitchMap(UserMap)
        If bRain And bLluvia(UserMap) = 0 Then
                Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
        End If
    Else
        'no encontramos el mapa en el hd
        MsgBox JsonLanguage.item("ERROR_MAPAS").item("TEXTO")
        
        Call CloseClient
    End If
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Map_RemoveOldUser
    
    '// Seteamos la Posicion en el Mapa
    Call Char_MapPosSet(incomingData.ReadByte(), incomingData.ReadByte())

    'Update pos label
    Call Char_UserPos
End Sub

Private Sub WriteChatOverHeadInConsole(ByVal CharIndex As Integer, ByVal ChatText As String, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    Dim NameRed As Byte
    Dim NameGreen As Byte
    Dim NameBlue As Byte
    
    With charlist(CharIndex)
        'Todo: Hacer que los colores se usen de Colores.dat
        'Haciendo uso de ColoresPj ya que el mismo en algun momento lo hace para DX
        If .priv = 0 Then
            If .Atacable Then
                NameRed = 236
                NameGreen = 89
                NameBlue = 57
            Else
                If .Criminal Then
                    NameRed = 247
                    NameGreen = 44
                    NameBlue = 0
                Else
                    NameRed = 218
                    NameGreen = 131
                    NameBlue = 225
                End If
            End If
         Else
            NameRed = 222
            NameGreen = 221
            NameBlue = 211
        End If

        Dim Pos As Integer
        Pos = InStr(.Nombre, "<")
            
        If Pos = 0 Then Pos = LenB(.Nombre) + 2
        
        Dim name As String
        name = Left$(.Nombre, Pos - 2)
       
        'Si el npc tiene nombre lo escribimos en la consola
        ChatText = Trim$(ChatText)
        If LenB(.Nombre) <> 0 And LenB(ChatText) > 0 Then
            Call AddtoRichTextBox(frmMain.RecTxt, name & "> ", NameRed, NameGreen, NameBlue, True, False, True, rtfLeft)
            Call AddtoRichTextBox(frmMain.RecTxt, ChatText, Red, Green, Blue, True, False, False, rtfLeft)
        End If

    End With
    
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim CharIndex As Integer
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    
    Red = Buffer.ReadByte()
    Green = Buffer.ReadByte()
    Blue = Buffer.ReadByte()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If Char_Check(CharIndex) Then
        Call Dialogos.CreateDialog(Trim$(chat), CharIndex, RGB(Red, Green, Blue))

        'Aqui escribimos el texto que aparece sobre la cabeza en la consola.
        Call WriteChatOverHeadInConsole(CharIndex, chat, Red, Green, Blue)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
        
        ' Para no perder el foco cuando chatea por party
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If
'    Call checkText(chat)
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 04/07/08 (NicoNZ)
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = Buffer.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Char_UserIndexSet(incomingData.ReadInteger())
                     
    'Update pos label
    Call Char_UserPos
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim fX As Integer
    Dim FXLoops As Integer
    Dim name As String
    Dim NickColor As Byte
    Dim Privileges As Integer
    
    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    weapon = Buffer.ReadInteger()
    shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()
    fX = Buffer.ReadInteger()
    FXLoops = Buffer.ReadInteger()
    name = Buffer.ReadASCIIString()
    NickColor = Buffer.ReadByte()
    Privileges = Buffer.ReadByte()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    With charlist(CharIndex)
        
        Call Char_SetFx(CharIndex, fX, FXLoops)

        .Nombre = name

        .Clan = mid$(.Nombre, getTagPosition(.Nombre))

        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0

        If Privileges <> 0 Then
            
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (Privileges And PlayerType.ChaosCouncil) <> 0 And (Privileges And PlayerType.User) = 0 Then
                Privileges = Privileges Xor PlayerType.ChaosCouncil
            End If
            
            If (Privileges And PlayerType.RoyalCouncil) <> 0 And (Privileges And PlayerType.User) = 0 Then
                Privileges = Privileges Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If Privileges And PlayerType.RoleMaster Then
                Privileges = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(Privileges) / Log(2)
        
        Else
            .priv = 0
            
        End If
        
    End With
    
    Call Char_Make(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)

errhandler:

    Dim Error As Long
        Error = Err.number
        
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    
    Call Char_SetName(CharIndex, incomingData.ReadASCIIString)
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call Char_Erase(CharIndex)
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()

    With charlist(CharIndex)
        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind

        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If

    End With
    
    Call Char_MovebyPos(CharIndex, X, Y)
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call Char_MovebyHead(UserCharIndex, Direccion)
    Call Char_MoveScreen(Direccion)
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 21/09/2010 - C4b3z0n
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
'***************************************************
    If incomingData.Length < 17 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    '// Char Body
    Call Char_SetBody(CharIndex, incomingData.ReadInteger())

    '// Char Head
    Call Char_SetHead(CharIndex, incomingData.ReadInteger)
    
    '// Char Weapon
    Call Char_SetWeapon(CharIndex, incomingData.ReadInteger())
        
    '// Char Shield
    Call Char_SetShield(CharIndex, incomingData.ReadInteger())
        
    '// Char Casco
    Call Char_SetCasco(CharIndex, incomingData.ReadInteger())
        
    '// Char Fx
    Call Char_SetFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
End Sub

Private Sub HandleHeadingChange()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 21/09/2010 - C4b3z0n
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim CharIndex As Integer

    CharIndex = incomingData.ReadInteger()

    Call Char_SetHeading(CharIndex, incomingData.ReadByte())
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X        As Byte
    Dim Y        As Byte
    Dim GrhIndex As Long
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    GrhIndex = incomingData.ReadLong()
        
    Call Map_CreateObject(X, Y, GrhIndex)
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X   As Byte
    Dim Y   As Byte
    Dim obj As Integer

    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    obj = Map_PosExitsObject(X, Y)
        
    If (obj > 0) Then
        Call Map_DestroyObject(X, Y)
    End If
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    Dim block As Boolean
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    block = incomingData.ReadBoolean()
    
    If block Then
        Map_SetBlocked X, Y, 1
    Else
        Map_SetBlocked X, Y, 0
    End If
End Sub

''
' Handles the PlayMP3 message.

Private Sub HandlePlayMP3()
'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 06/01/20
'Utilizamos PlayBackgroundMusic para el uso de MP3, simplifique la funcion (Recox)
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMp3 As Integer
    Dim Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMp3 = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call Audio.PlayBackgroundMusic(CStr(currentMp3), MusicTypes.Mp3)
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 06/01/20
'Utilizamos PlayBackgroundMusic para el uso de  MIDI, simplifique la funcion (Recox)
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMidi As Integer
    Dim Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call Audio.PlayBackgroundMusic(CStr(currentMidi), MusicTypes.Midi, Loops)
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    
    wave = incomingData.ReadByte()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
        
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildAdm
        'Clear guild's list
        .guildslist.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .guildslist.AddItem(GuildNames(i))
            End If
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X As Byte
    Dim Y As Byte
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
        
    Call CambioDeArea(X, Y)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa
End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.BAJOTECHO Or _
        MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.CASA Or _
        MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.ZONASEGURA)

    If bRain And bLluvia(UserMap) Then
        'Stop playing the rain sound
        Call Audio.StopWave(RainBufferIndex)
        RainBufferIndex = 0
        
        If bTecho Then
            Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
        Else
            Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
        End If
        
        frmMain.IsPlaying = PlayLoop.plNone

    End If
    
    bRain = Not bRain
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call Char_SetFx(CharIndex, fX, Loops)
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    
    frmMain.UpdateProgressExperienceLevelBar (UserExp)
    
    frmMain.GldLbl.Caption = UserGLD
    frmMain.lblLvl.Caption = UserLvl
    
    'Stats
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    Dim bWidth As Byte
    
    '***************************
    If UserMaxMAN > 0 Then _
        bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 90)
        
    frmMain.shpMana.Width = 90 - bWidth
    frmMain.shpMana.Left = 902 + (90 - frmMain.shpMana.Width)
    
    frmMain.shpMana.Visible = (bWidth <> 90)
    '***************************
    
    bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 90)
    
    frmMain.shpVida.Width = 90 - bWidth
    frmMain.shpVida.Left = 902 + (90 - frmMain.shpVida.Width)
    
    frmMain.shpVida.Visible = (bWidth <> 90)
    '***************************
    
    bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 83)
    
    frmMain.shpEnergia.Width = 83 - bWidth
    frmMain.shpEnergia.Left = 797 + (83 - frmMain.shpEnergia.Width)
    
    frmMain.shpEnergia.Visible = (bWidth <> 83)
    '***************************
    
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If

    Call frmMain.SetGoldColor
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    Dim OBJIndex As Integer
    Dim name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Long
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim Value As Single
    Dim Incompatible As Boolean
    
    slot = Buffer.ReadByte()
    OBJIndex = Buffer.ReadInteger()
    name = Buffer.ReadASCIIString()
    Amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadLong()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    MaxDef = Buffer.ReadInteger()
    MinDef = Buffer.ReadInteger()
    Value = Buffer.ReadSingle()
    Incompatible = Buffer.ReadBoolean()
    
    If Equipped Then
        Select Case OBJType
            Case eObjType.otWeapon
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = slot
            Case eObjType.otArmadura
                frmMain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = slot
            Case eObjType.otescudo
                frmMain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = slot
            Case eObjType.otcasco
                frmMain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = slot
        End Select
    Else
        Select Case slot
            Case UserWeaponEqpSlot
                frmMain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0
            Case UserArmourEqpSlot
                frmMain.lblArmor = "0/0"
                UserArmourEqpSlot = 0
            Case UserHelmEqpSlot
                frmMain.lblShielder = "0/0"
                UserHelmEqpSlot = 0
            Case UserShieldEqpSlot
                frmMain.lblHelm = "0/0"
                UserShieldEqpSlot = 0
        End Select
    End If
    
    Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, name, Incompatible)

    If frmComerciar.Visible Then
        Call InvComUsu.SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, name, Incompatible)
    End If

    If frmBancoObj.Visible Then
        Call InvBanco(1).SetItem(slot, OBJIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, name, Incompatible)
        frmBancoObj.NoPuedeMover = False
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    MaxInventorySlots = incomingData.ReadByte
    Call Inventario.DrawInventory
End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_WORK_FINISHED"), .Red, .Green, .Blue, .bold, .italic)
    End With
    If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/10
'
'***************************************************
    Dim slot As Byte
    Dim Amount As Long
    
    Call incomingData.ReadByte
    
    slot = incomingData.ReadByte
    
    With InvOfferComUsu(0)
        Amount = .Amount(slot)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.OBJIndex(slot), Amount)
            
            ' Borro el item
            Call .SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
        End If
    End With
    
    ' Si era el unico item de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then
        Call frmComerciarUsu.HabilitarConfirmar(False)
    End If
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg(JsonLanguage.item("MENSAJE_NO_COMM_OBJETO").item("TEXTO"), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .OBJIndex = Buffer.ReadInteger()
        .name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadLong()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger()
        .Valor = Buffer.ReadLong()
        .Incompatible = Buffer.ReadBoolean()
        
        If frmBancoObj.Visible Then
            Call InvBanco(0).SetItem(slot, .OBJIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .name, .Incompatible)
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
 
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
 
    'Remove packet ID
    Call Buffer.ReadByte
 
    Dim slot As Byte
    slot = Buffer.ReadByte()
    Dim str As String
 
    UserHechizos(slot) = Buffer.ReadInteger()
 
    If slot <= frmMain.hlst.ListCount Then
         str = DevolverNombreHechizo(UserHechizos(slot))
        If str <> vbNullString Then
            frmMain.hlst.List(slot - 1) = str
        Else
            Call frmMain.hlst.AddItem(JsonLanguage.item("NADA").item("TEXTO"))
        End If
    Else
        str = DevolverNombreHechizo(UserHechizos(slot))
        If str <> vbNullString Then
            Call frmMain.hlst.AddItem(str)
        Else
            Call frmMain.hlst.AddItem(JsonLanguage.item("NADA").item("TEXTO"))
        End If
    End If
 
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
 
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set Buffer = Nothing
 
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                For i = 1 To NUMATRIBUTES
                    .lblAtributos(i).Caption = UserAtributos(i)
                Next i
                
                .UpdateStats
            End If
        End With
    Else
        LlegaronAtrib = True
    End If
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmasHerrero(i)
            .name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadLong()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    'En otras palabras, a partir de ahora podes usar "Exit Sub" sin romper nada.
    Call incomingData.CopyBuffer(Buffer)
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
        Exit Sub
    Else
        Call frmHerrero.Show(vbModeless, frmMain)
        MirandoHerreria = True
    End If
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvLingosHerreria(i) = New clsGraphicalInventory
    Next i
    
    With frmHerrero
        ' Inicializo los inventarios
        Call InvLingosHerreria(1).Initialize(DirectD3D8, .picLingotes0, 3, , , , , , False)
        Call InvLingosHerreria(2).Initialize(DirectD3D8, .picLingotes1, 3, , , , , , False)
        Call InvLingosHerreria(3).Initialize(DirectD3D8, .picLingotes2, 3, , , , , , False)
        Call InvLingosHerreria(4).Initialize(DirectD3D8, .picLingotes3, 3, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1, True)
    End With
    
    For i = 1 To Count
        With ArmasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmasHerrero(k).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).name = .name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmasHerrero(k).name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmasHerrero(k).GrhIndex
                        HerreroMejorar(J).LinH = ArmasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            .name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadLong()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    'En otras palabras, a partir de ahora podes usar "Exit Sub" sin romper nada.
    Call incomingData.CopyBuffer(Buffer)
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
                Exit Sub
    Else
        Call frmHerrero.Show(vbModeless, frmMain)
        MirandoHerreria = True
    End If
    
    J = UBound(HerreroMejorar)
    
    For i = 1 To Count
        With ArmadurasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmadurasHerrero(k).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).name = .name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).OBJIndex = .OBJIndex
                        HerreroMejorar(J).UpgradeName = ArmadurasHerrero(k).name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmadurasHerrero(k).GrhIndex
                        HerreroMejorar(J).LinH = ArmadurasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmadurasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmadurasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the InitCarpenting message.

Private Sub HandleInitCarpenting()
'***************************************************
'Author: Jopi
'Last Modification: 01/01/20
'Este solia ser HandleCarpenterObjects(), pero fue modificado para fusionar los paquetes CarptenterObjects y ShowCarpenterForm.
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count
        With ObjCarpintero(i)
            .name = Buffer.ReadASCIIString()        'Get the object's name
            .GrhIndex = Buffer.ReadLong()
            .Madera = Buffer.ReadInteger()          'The wood needed
            .MaderaElfica = Buffer.ReadInteger()    'The elfic wood needed
            .OBJIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    'En otras palabras, a partir de ahora podes usar "Exit Sub" sin romper nada.
    Call incomingData.CopyBuffer(Buffer)
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
        Exit Sub
    Else
        Call frmCarpinteria.Show(vbModeless, frmMain)
        MirandoCarpinteria = True
    End If
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvMaderasCarpinteria(i) = New clsGraphicalInventory
    Next i
    
    With frmCarpinteria
        ' Inicializo los inventarios
        Call InvMaderasCarpinteria(1).Initialize(DirectD3D8, .picMaderas0, 2, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(DirectD3D8, .picMaderas1, 2, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(DirectD3D8, .picMaderas2, 2, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(DirectD3D8, .picMaderas3, 2, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With
    
    For i = 1 To Count
        With ObjCarpintero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ObjCarpintero(k).OBJIndex Then
                        J = J + 1
                
                        ReDim Preserve CarpinteroMejorar(J) As tItemsConstruibles
                        
                        CarpinteroMejorar(J).name = .name
                        CarpinteroMejorar(J).GrhIndex = .GrhIndex
                        CarpinteroMejorar(J).OBJIndex = .OBJIndex
                        CarpinteroMejorar(J).UpgradeName = ObjCarpintero(k).name
                        CarpinteroMejorar(J).UpgradeGrhIndex = ObjCarpintero(k).GrhIndex
                        CarpinteroMejorar(J).Madera = ObjCarpintero(k).Madera - .Madera * 0.85
                        CarpinteroMejorar(J).MaderaElfica = ObjCarpintero(k).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleInitCraftman()
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    '***************************************************

    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    Dim CountObjs As Integer
    Dim CountCrafteo As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long

    frmArtesano.ArtesaniaCosto = Buffer.ReadLong()

    CountObjs = Buffer.ReadInteger()

    ReDim ObjArtesano(CountObjs) As tItemArtesano

    For i = 1 To CountObjs
        With ObjArtesano(i)
            .name = Buffer.ReadASCIIString()
            .GrhIndex = Buffer.ReadLong()
            .OBJIndex = Buffer.ReadInteger()
            
            CountCrafteo = Buffer.ReadByte()
            ReDim .ItemsCrafteo(CountCrafteo) As tItemCrafteo
            
            For J = 1 To CountCrafteo
                .ItemsCrafteo(J).name = Buffer.ReadASCIIString()
                .ItemsCrafteo(J).GrhIndex = Buffer.ReadLong()
                .ItemsCrafteo(J).OBJIndex = Buffer.ReadInteger()
                .ItemsCrafteo(J).Amount = Buffer.ReadInteger()
            Next J
        End With
    Next i

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

    Call frmArtesano.Show(vbModeless, frmMain)

    For i = 1 To MAX_LIST_ITEMS
        Set InvObjArtesano(i) = New clsGraphicalInventory
    Next i

    With frmArtesano
        ' Inicializo los inventarios
        Call InvObjArtesano(1).Initialize(DirectD3D8, .picObj0, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(2).Initialize(DirectD3D8, .picObj1, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(3).Initialize(DirectD3D8, .picObj2, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(4).Initialize(DirectD3D8, .picObj3, MAX_ITEMS_CRAFTEO, , , , , , False)

        Call .HideExtraControls(CountObjs)
        Call .RenderList(1)
    End With

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call MsgBox(Buffer.ReadASCIIString())
    
    If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
        frmMain.Client.CloseSck
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    
    Call InitCartel(tmp, Buffer.ReadLong())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With NPCInventory(slot)
        .name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadLong()
        .OBJIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger()
        .Incompatible = Buffer.ReadBoolean()
    
        If frmComerciar.Visible Then
            Call InvComNpc.SetItem(slot, .OBJIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .name, .Incompatible)
        End If
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    frmMain.lblHambre = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblSed = UserMinAGU & "/" & UserMaxAGU

    Dim bWidth As Byte
    
    bWidth = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 83)
    
    frmMain.shpHambre.Width = 83 - bWidth
    frmMain.shpHambre.Left = 797 + (83 - frmMain.shpHambre.Width)
    
    frmMain.shpHambre.Visible = (bWidth <> 83)
    '*********************************
    
    bWidth = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 83)
    
    frmMain.shpSed.Width = 83 - bWidth
    frmMain.shpSed.Left = 797 + (83 - frmMain.shpSed.Width)
    
    frmMain.shpSed.Visible = (bWidth <> 83)
    
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .Promedio = incomingData.ReadLong()
    End With
    
    LlegoFama = True
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    
    Call frmMain.LightSkillStar(True)
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim ForumType As eForumMsgType
    Dim Title As String
    Dim Message As String
    Dim Author As String
    
    ForumType = Buffer.ReadByte
    
    Title = Buffer.ReadASCIIString()
    Author = Buffer.ReadASCIIString()
    Message = Buffer.ReadASCIIString()
    
    If Not frmForo.ForoLimpio Then
        clsForos.ClearForums
        frmForo.ForoLimpio = True
    End If

    Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, EsAnuncio(ForumType))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmForo.Privilegios = incomingData.ReadByte
    frmForo.CanPostSticky = incomingData.ReadByte
    
    If Not MirandoForo Then
        frmForo.Show , frmMain
    End If
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim timeRemaining As Integer
    
    CharIndex = incomingData.ReadInteger()
    UserInvisible = incomingData.ReadBoolean()
    Call Char_SetInvisible(CharIndex, UserInvisible)
    
    If CharIndex = UserCharIndex Then
        If UserInvisible And TiempoInvi <= 0 Then
            TiempoInvi = (IntervaloInvi * 0.05) - 1 ' Qu�???
        Else
            TiempoInvi = 0
        End If
    
    End If
    
End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
    UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
    UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()
         
    With frmCrearPersonaje
        .lblAtributos(eAtributos.Fuerza) = UserAtributos(eAtributos.Fuerza)
        .lblAtributos(eAtributos.Agilidad) = UserAtributos(eAtributos.Agilidad)
        .lblAtributos(eAtributos.Inteligencia) = UserAtributos(eAtributos.Inteligencia)
        .lblAtributos(eAtributos.Carisma) = UserAtributos(eAtributos.Carisma)
        .lblAtributos(eAtributos.Constitucion) = UserAtributos(eAtributos.Constitucion)
        
        .UpdateStats
    End With
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    UserMeditar = Not UserMeditar
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Now the server send the percentage of progress of the skills.
'***************************************************
    If incomingData.Length < 2 + NUMSKILLS * 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    UserClase = incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        PorcentajeSkills(i) = incomingData.ReadByte()
    Next i
    
    LlegaronSkills = True
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    Dim Upper_creatures As Long
    
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatures = UBound(creatures())
    
    For i = 0 To Upper_creatures
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'***************************************************
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList() As String
    Dim Upper_guildList As Long
    Dim i As Long
    Dim sTemp As String
    
    'Get news' string
    frmGuildNews.news = Buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Upper_guildList = UBound(guildList)
    
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    If ClientSetup.bGuildNews Or bShowGuildNews Then frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim vsGuildList() As String, Upper_vsGuildList As Long
    Dim i As Long
    
    vsGuildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_vsGuildList = UBound(vsGuildList())
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To Upper_vsGuildList
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList()     As String
    Dim Upper_guildList As Long
    Dim i               As Long
    
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    With frmPeaceProp
    
        .lista.Clear
    
        Upper_guildList = UBound(guildList())
    
        For i = 0 To Upper_guildList
            .lista.AddItem (guildList(i))
        Next i
    
        .ProposalType = TIPO_PROPUESTA.PAZ
        .Show vbModeless, frmMain
    
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True
        End If
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .Raza.Caption = ListaRazas(Buffer.ReadByte())
        .Clase.Caption = ListaClases(Buffer.ReadByte())
        
        If Buffer.ReadByte() = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"
        End If
        
        .Nivel.Caption = Buffer.ReadByte()
        .Oro.Caption = Buffer.ReadLong()
        .Banco.Caption = Buffer.ReadLong()
        
        Dim reputation As Long
        reputation = Buffer.ReadLong()
        
        .reputacion.Caption = reputation
        
        .txtPeticiones.Text = Buffer.ReadASCIIString()
        .guildactual.Caption = Buffer.ReadASCIIString()
        .txtMiembro.Text = Buffer.ReadASCIIString()
        
        Dim armada As Boolean
        Dim caos As Boolean
        
        armada = Buffer.ReadBoolean()
        caos = Buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = JsonLanguage.item("ARMADA").item("TEXTO")
        ElseIf caos Then
            .ejercito.Caption = JsonLanguage.item("LEGION").item("TEXTO")
        End If
        
        .Ciudadanos.Caption = CStr(Buffer.ReadLong())
        .criminales.Caption = CStr(Buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " " & JsonLanguage.item("CIUDADANO").item("TEXTO")
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " " & JsonLanguage.item("CRIMINAL").item("TEXTO")
            .status.ForeColor = vbRed
        End If
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim i As Long
    Dim List() As String

    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .guildslist.AddItem(GuildNames(i))
            End If
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear

        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtguildnews = Buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear

        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildBrief
        .imgDeclararGuerra.Visible = .EsLeader
        .imgOfrecerAlianza.Visible = .EsLeader
        .imgOfrecerPaz.Visible = .EsLeader
        
        .Nombre.Caption = Buffer.ReadASCIIString()
        .fundador.Caption = Buffer.ReadASCIIString()
        .creacion.Caption = Buffer.ReadASCIIString()
        .lider.Caption = Buffer.ReadASCIIString()
        .web.Caption = Buffer.ReadASCIIString()
        .Miembros.Caption = Buffer.ReadInteger()
        
        If Buffer.ReadBoolean() Then
            .eleccion.Caption = UCase$(JsonLanguage.item("ABIERTA").item("TEXTO"))
        Else
            .eleccion.Caption = UCase$(JsonLanguage.item("CERRADA").item("TEXTO"))
        End If
        
        .lblAlineacion.Caption = Buffer.ReadASCIIString()
        .Enemigos.Caption = Buffer.ReadInteger()
        .Aliados.Caption = Buffer.ReadInteger()
        .antifaccion.Caption = Buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i As Long
        
        codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = Buffer.ReadASCIIString()
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowGuildAlign message.

Private Sub HandleShowGuildAlign()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmEligeAlineacion.Show vbModeless, frmMain
End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    Dim timeRemaining As Integer
    UserParalizado = Not UserParalizado

    timeRemaining = incomingData.ReadInteger()
    UserParalizadoSegundosRestantes = IIf(timeRemaining > 0, (timeRemaining * 0.04), 0) 'Cantidad en segundos

    If UserParalizado And timeRemaining > 0 Then frmMain.timerPasarSegundo.Enabled = True

End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '******************************************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 17/06/2020
    '17/06/2020 BelerianD - Se agrego un parametro que faltaba, causando runtime al comerciar.
    '******************************************************************************************
    If incomingData.Length < 24 Then
        Call Err.Raise(incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If
    
    On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    With Buffer
    
        'Remove packet ID
        Call Buffer.ReadByte

        Dim OfferSlot As Byte: OfferSlot = .ReadByte()
        Dim OBJIndex  As Integer: OBJIndex = .ReadInteger()
        Dim Amount    As Long: Amount = .ReadLong()
        Dim GrhIndex  As Long: GrhIndex = .ReadLong()
        Dim OBJType   As Byte: OBJType = .ReadByte()
        Dim MaxHit    As Integer: MaxHit = .ReadInteger()
        Dim MinHit    As Integer: MinHit = .ReadInteger()
        Dim MaxDef    As Integer: MaxDef = .ReadInteger()
        Dim MinDef    As Integer: MinDef = .ReadInteger()
        Dim SalePrice As Long: SalePrice = .ReadLong()
        Dim name      As String: name = .ReadASCIIString()
        Dim Incompatible As Boolean: Incompatible = .ReadBoolean()
   
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

    If OfferSlot = GOLD_OFFER_SLOT Then
        Call InvOroComUsu(2).SetItem(1, OBJIndex, Amount, 0, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, SalePrice, name, Incompatible)
    Else
        Call InvOfferComUsu(1).SetItem(OfferSlot, OBJIndex, Amount, 0, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, SalePrice, name, Incompatible)
    End If
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & JsonLanguage.item("MENSAJE_COMM_OFERTA_CAMBIA").item("TEXTO"), FontTypeNames.FONTTYPE_VENENO)
    
errhandler:
    Dim Error As Long
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Call Err.Raise(Error)
        Call Err.Raise(Error)
End Sub

''
' Handles the SendNight message.

Private Sub HandleSendNight()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'
'***************************************************
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tBool As Boolean 'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?
    tBool = incomingData.ReadBoolean()
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    Dim Upper_creatureList As Long
    
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatureList = UBound(creatureList())
    
    For i = 0 To Upper_creatureList
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    Dim Upper_sosList As Long
    
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_sosList = UBound(sosList())
    
    For i = 0 To Upper_sosList
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim DenounceList() As String
    Dim Upper_denounceList As Long
    Dim DenounceIndex As Long
    
    DenounceList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_denounceList = UBound(DenounceList())
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        For DenounceIndex = 0 To Upper_denounceList
            Call AddtoRichTextBox(frmMain.RecTxt, DenounceList(DenounceIndex), .Red, .Green, .Blue, .bold, .italic)
        Next DenounceIndex
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim members() As String
    Dim Upper_members As Long
    Dim i As Long
    
    EsPartyLeader = CBool(Buffer.ReadByte())
       
    members = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_members = UBound(members())
    
    For i = 0 To Upper_members
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    
    frmParty.lblTotalExp.Caption = Buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub



''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'*************************************Su**************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmMain
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim userList() As String
    
    userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        
        Dim i As Long
        Dim Upper_userlist As Long
            Upper_userlist = UBound(userList())
            
        For i = 0 To Upper_userlist
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleRenderMsg()

    Call incomingData.ReadByte
    
    renderMsgReset
    renderText = incomingData.ReadASCIIString
    renderFont = incomingData.ReadInteger
    colorRender = 240
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Dim MENSAJE_PING As String
        MENSAJE_PING = JsonLanguage.item("MENSAJE_PING").item("TEXTO")
        MENSAJE_PING = Replace$(MENSAJE_PING, "VAR_PING", (timeGetTime - pingTime))
        
    Call AddtoRichTextBox(frmMain.RecTxt, _
                            MENSAJE_PING, _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(3), _
                            True, False, True)
    
    pingTime = 0
End Sub

''
' Handles the Pong message.

Private Sub HandleGuildMemberInfo()
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long

        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .lstClanes.AddItem(GuildNames(i))
            End If
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear

        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain
    End With
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    
    CharIndex = Buffer.ReadInteger()
    NickColor = Buffer.ReadByte()
    UserTag = Buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With charlist(CharIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        .Nombre = UserTag
        .Clan = mid$(.Nombre, getTagPosition(.Nombre))
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Writes the "LoginExistingAccount" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingAccount()
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 12/10/2018
'Writes the "LoginExistingAccount" message to the outgoing data buffer
'***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingAccount)
        
        Call .WriteASCIIString(AccountName)
        
        Call .WriteASCIIString(AccountPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

''
' Writes the "DeleteChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeleteChar()
'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 07/01/2020
'Writes the "DeleteChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DeleteChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(AccountHash)
    End With
End Sub

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 12/10/2018
'CHOTS: Accounts
'Writes the "LoginExistingChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(AccountHash)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

Public Sub WriteLoginNewAccount()
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 12/10/2018
'CHOTS: Accounts
'Writes the "LoginNewAccount" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewAccount)
        
        Call .WriteASCIIString(AccountName)
        
        Call .WriteASCIIString(AccountPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

''
' Writes the "ThrowDices" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteThrowDices()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ThrowDices" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ThrowDices)
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LoginNewChar" message to the outgoing data buffer
'***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(AccountHash)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(UserHead)
        
        Call .WriteByte(UserHogar)
    End With
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Talk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Yell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharName As String, ByVal chat As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 03/12/10
'Writes the "Whisper" message to the outgoing data buffer
'03/12/10: Enanoh - Ahora se envia el nick y no el charindex.
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteASCIIString(CharName)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Walk" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)
    End With
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestPositionUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Attack" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Attack)
    'Iniciamos la animacion de ataque
    charlist(UserCharIndex).attacking = True
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PickUp" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SafeToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
End Sub

''
' Writes the "ResuscitationSafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResuscitationToggle()
'**************************************************************
'Author: Rapsodius
'Creation Date: 10/10/07
'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
'**************************************************************
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

Public Sub WriteRequestPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "RequestPartyForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)

End Sub

''
' Writes the "ItemUpgrade" message to the outgoing data buffer.
'
' @param    ItemIndex The index to the item to upgrade.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/09/09
'Writes the "ItemUpgrade" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAtributes" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)
End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestFame" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestSkills" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMiniStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

''
' Writes the "UserCommerceConfirm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserCommerceConfirm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/10/07
'Writes the "UserCommerceOk" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceReject" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Drop" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CastSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoubleClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Work" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseSpellMacro" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UseItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal item As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftBlacksmith" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(item)
    End With
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal item As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CraftCarpenter" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(item)
    End With
End Sub

''
' Writes the "CraftsmanCreate" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftsmanCreate(ByVal item As Integer)
'***************************************************
'Author: WyroX
'Last Modification: 27/01/2020
'Writes the "CraftsmanCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftsmanCreate)
        
        Call .WriteInteger(item)
    End With
End Sub

''
' Writes the "ShowGuildNews" message to the outgoing data buffer.
'

Public Sub WriteShowGuildNews()
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'Writes the "ShowGuildNews" message to the outgoing data buffer
'***************************************************
 
     outgoingData.WriteByte (ClientPacketID.ShowGuildNews)
End Sub


''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkLeftClick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)
    End With
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal name As String, ByVal Site As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNewGuild" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    Dim Lower_codex As Long, Upper_codex As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(name)
        Call .WriteASCIIString(Site)
        
        Lower_codex = LBound(Codex())
        Upper_codex = UBound(Codex())
        
        For i = Lower_codex To Upper_codex
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
    End With
End Sub


''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EquipItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeHeading" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
    End With
    
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ModifySkills" message to the outgoing data buffer
'***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Train" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)
    End With
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceBuy" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceSell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal Amount As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDeposit" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String, ByVal ForumMsgType As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForumPost" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteByte(ForumMsgType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MoveSpell" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'Writes the "MoveBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ClanCodexUpdate" message to the outgoing data buffer
'***************************************************
    Dim temp As String
    Dim i As Long
    Dim Lower_codex As Long, Upper_codex As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        Lower_codex = LBound(Codex())
        Upper_codex = UBound(Codex())
        
        For i = Lower_codex To Upper_codex
            temp = temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
    End With
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, ByVal Amount As Long, ByVal OfferSlot As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceOffer" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(slot)
        Call .WriteLong(Amount)
        Call .WriteByte(OfferSlot)
    End With
End Sub

Public Sub WriteCommerceChat(ByVal chat As String)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Writes the "CommerceChat" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        
        Call .WriteASCIIString(chat)
    End With
End Sub


''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferPeace" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOfferAlliance" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAllianceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeaceDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAlliancePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildPeacePropList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildDeclareWar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildNewWebsite" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(URL)
    End With
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRejectNewMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildKickMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildUpdateNews" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)
    End With
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOpenElections" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestMembership" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Online" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

Public Sub WriteDiscord(ByVal chat As String)
'***************************************************
'Author: Lucas Daniel Recoaro (Recox)
'Last Modification: 05/17/06
'Writes the "Discord" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Discord)
        
        Call .WriteASCIIString(chat)
    End With
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 08/16/08
'Writes the "Quit" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestAccountState" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetStand" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PetFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

''
' Writes the "ReleasePet" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReleasePet()
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'Writes the "ReleasePet" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ReleasePet)
End Sub


''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Rest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Meditate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Resucitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

''
' Writes the "Consultation" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsultation()
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Writes the "Consultation" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Consultation)

End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Heal" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Help" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestStats" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankStart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Enlist" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Information" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Reward" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.UpTime)
End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyLeave" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)
End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyCreate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)
End Sub

''
' Writes the "PartyJoin" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyJoin" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)
End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Inquiry" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildRequestDetails" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal Clave As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 02/05/2012
'                   Nuevo centinela : maTih.-
'Writes the "CentinelReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteASCIIString(Clave)
    End With
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyOnline" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoleMasterRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BugReport" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Punishments" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 10/10/07
'Last Modified By: Rapsodius
'Writes the "ChangePassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        Call .WriteASCIIString(oldPass)
        Call .WriteASCIIString(newPass)
    End With
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Gamble" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)
    End With
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "InquiryVote" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)
    End With
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LeaveFaction" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankExtractGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankDepositGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)
    End With
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Denounce" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 03/21/2001
'Writes the "GuildFundate" message to the outgoing data buffer
'14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
'03/21/2001: Pato - Deleted de clanType param.
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)
End Sub

''
' Writes the "GuildFundation" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "GuildFundation" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundation)
        
        Call .WriteByte(clanType)
    End With
End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartySetLeader" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
' @param    username The user to accept into the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PartyAcceptMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildMemberList" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildMemberList)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'Writes the "InitCrafting" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        Call .WriteLong(cantidad)
        
        Call .WriteInteger(NroPorCiclo)
    End With
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
'***************************************************
'Author: Budi
'Last Modification: 01/06/10
'Writes the "Home" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Home)
    End With
End Sub



''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GMMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowName" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.showName)
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineChaosLegion" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoNearby" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoNearby)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Comment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.comment)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerTime" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.serverTime)
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Where" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Where)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreaturesInMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreaturesInMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpMeToTarget" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarpChar" message to the outgoing data buffer
'***************************************************
    
    'Para que te vas a tepear al mismo lugar? Te pinta spamear el FX del summon?
    'No mandemos paquetes al pedo.
    If X = UserPos.X And Y = UserPos.Y And Map = UserMap Then Exit Sub
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
    
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Silence" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Silence)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSShowList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SOSRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GoToChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "invisible" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GMPanel" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestUserList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Working" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Hiding" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal Reason As String, ByVal Time As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Jail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
        
        Call .WriteByte(Time)
    End With
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPC" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WarnUser" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "EditChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(EditOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInfo" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharStats" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharStats)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharGold" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharGold)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharInventory" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharBank" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharSkills" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReviveChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "OnlineGM" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'Writes the "OnlineMap" message to the outgoing data buffer
'26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Forgive" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Forgive)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Kick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Execute" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NPCFollow" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SummonChar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnCreature" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
    End With
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetNPCInventory" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub
''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String)
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "MapMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MapMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NickToIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.NickToIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "IPToNick" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    Dim Upper_ip As Long, Lower_ip As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        
        Lower_ip = LBound(Ip())
        Upper_ip = UBound(Ip())
        
        For i = Lower_ip To Upper_ip
            Call .WriteByte(Ip(i))
        Next i
    End With
    
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildOnlineMembers" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportCreate" message to the outgoing data buffer
'***************************************************
    With outgoingData
            Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TeleportCreate)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Radio)
    End With
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)
End Sub
''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExitDestroy()
'***************************************************
'Author: Cucsijuan
'Last Modification: 30/09/18
'Writes the "TeleportDestroy" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ExitDestroy)
End Sub
''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetCharDescription" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetCharDescription)
        
        Call .WriteASCIIString(Desc)
    End With
End Sub

''
' Writes the "ForceMP3ToMap" message to the outgoing data buffer.
'
' @param    Mp3Id The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMP3ToMap(ByVal Mp3Id As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMP3ToMap)
        
        Call .WriteByte(Mp3Id)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)
    End With
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEToMap" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CitizenMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CitizenMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CriminalMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CriminalMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TalkAsNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ItemsInTheFloor" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumb" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumb)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MakeDumbNoMore" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumpIPTables" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DumpIPTables)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CouncilKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CouncilKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetTrigger" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        
        Call .WriteByte(Trigger)
    End With
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'Writes the "AskTrigger" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BannedIPReload" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildBan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BanIP" message to the outgoing data buffer
'***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            Dim i As Long
            Dim Upper_ip As Long, Lower_ip As Long
            
            Lower_ip = LBound(Ip())
            Upper_ip = UBound(Ip())
        
            For i = Lower_ip To Upper_ip
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UnbanIP" message to the outgoing data buffer
'***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    Dim Upper_ip As Long, Lower_ip As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        
        Lower_ip = LBound(Ip())
        Upper_ip = UBound(Ip())
        
        For i = Lower_ip To Upper_ip
            Call .WriteByte(Ip(i))
        Next i
    End With
    
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)
        Call .WriteInteger(cantidad)
    End With
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DestroyItems" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosLegionKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RoyalArmyKick" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ForceMP3All" message to the outgoing data buffer.
'
' @param    Mp3Id The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMP3All(ByVal Mp3Id As Byte)
'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 05/17/06
'Writes the "ForceMP3All" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMP3All)
        
        Call .WriteByte(Mp3Id)
    End With
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceMIDIAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIAll)
        
        Call .WriteByte(midiID)
    End With
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ForceWAVEAll" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEAll)
        
        Call .WriteByte(waveID)
    End With
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemovePunishment" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TileBlockedToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LastIP" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMOTD" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetMOTD" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetMOTD)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SystemMessage" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        
        Call .WriteASCIIString(Message)
    End With
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer, ByVal WithRespawn As Boolean)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateNPC" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        
        Call .WriteInteger(NPCIndex)
        Call .WriteBoolean(WithRespawn)
        
    End With
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ImperialArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChaosArmour" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)
    End With
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnOffServer" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TurnCriminal" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetFactions" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResetFactions)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RequestCharMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharMail)
        
        Call .WriteASCIIString(UserName)
    End With
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterPassword" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterMail" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
    End With
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AlterName" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
    End With
End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DoBackup" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGuildMessages" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
    End With
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveMap" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 19/09/2010
'Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoOcultar)
        
        Call .WriteBoolean(PermitirOcultar)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvocar)
        
        Call .WriteBoolean(PermitirInvocar)
    End With
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
    End With
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
    End With
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)
    End With
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
    End With
End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoStealNpc)
        
        Call .WriteBoolean(forbid)
    End With
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SaveChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CleanSOS" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowServerForm" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowServerForm)
End Sub

''
' Writes the "ShowDenouncesList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenouncesList()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "ShowDenouncesList" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowDenouncesList)
End Sub

''
' Writes the "EnableDenounces" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnableDenounces()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'Writes the "EnableDenounces" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.EnableDenounces)
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Night" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.night)
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "KickAllChars" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadNPCs" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadServerIni" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadSpells" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ReloadObjects" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Restart" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Restart)
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ResetAutoUpdate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    Red The red component of the new chat color.
' @param    Green The green component of the new chat color.
' @param    Blue The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatColor" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChatColor)
        
        Call .WriteByte(Red)
        Call .WriteByte(Green)
        Call .WriteByte(Blue)
    End With
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Ignored" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Writes the "CheckSlot" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)
    End With
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 26/01/2007
'Writes the "Ping" message to the outgoing data buffer
'***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    pingTime = timeGetTime
End Sub

''
' Writes the "ShareNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShareNpc()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Writes the "ShareNpc" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ShareNpc)
End Sub

''
' Writes the "StopSharingNpc" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteStopSharingNpc()
'***************************************************
'Author: ZaMa
'Last Modification: 15/04/2010
'Writes the "StopSharingNpc" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)
End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 21/06/2009
'Writes the "SetIniVar" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetIniVar)
        
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'Writes the "CreatePretorianClan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreatePretorianClan)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal Map As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'Writes the "DeletePretorianClan" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePretorianClan)
        Call .WriteInteger(Map)
    End With
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With outgoingData
        If .Length = 0 Then _
            Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)
        
        Call SendData(sndData)
    End With
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    'No enviamos nada si no estamos conectados
    If Not frmMain.Client.State = sckConnected Then Exit Sub
    #If AntiExternos Then

        Dim data() As Byte

        data = StrConv(sdData, vbFromUnicode)
        Security.NAC_E_Byte data, Security.Redundance
        sdData = StrConv(data, vbUnicode)
        'sdData = Security.NAC_E_String(sdData, Security.Redundance)
    #End If
    'Send data!
    Call frmMain.Client.SendData(sdData)
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    Dialog The new dialog of the NPC.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetDialog(ByVal dialog As String)
'***************************************************
'Author: Amraphen
'Last Modification: 18/11/2010
'Writes the "SetDialog" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetDialog)
        
        Call .WriteASCIIString(dialog)
    End With
End Sub

''
' Writes the "Impersonate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImpersonate()
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Writes the "Impersonate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Impersonate)
End Sub

''
' Writes the "Imitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImitate()
'***************************************************
'Author: ZaMa
'Last Modification: 20/11/2010
'Writes the "Imitate" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Imitate)
End Sub

''
' Writes the "RecordAddObs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordAddObs" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordAddObs)
        
        Call .WriteByte(RecordIndex)
        Call .WriteASCIIString(Observation)
    End With
End Sub

''
' Writes the "RecordAdd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal Reason As String)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordAdd" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordAdd)
        
        Call .WriteASCIIString(Nickname)
        Call .WriteASCIIString(Reason)
    End With
End Sub

''
' Writes the "RecordRemove" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordRemove" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordRemove)
        
        Call .WriteByte(RecordIndex)
    End With
End Sub

''
' Writes the "RecordListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordListRequest()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordListRequest" message to the outgoing data buffer
'***************************************************
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RecordListRequest)
End Sub

''
' Writes the "RecordDetailsRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'Writes the "RecordDetailsRequest" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordDetailsRequest)
        
        Call .WriteByte(RecordIndex)
    End With
End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim NumRecords As Byte
    Dim i As Long
    
    NumRecords = Buffer.ReadByte
    
    'Se limpia el ListBox y se agregan los usuarios
    frmPanelGm.lstUsers.Clear
    For i = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem Buffer.ReadASCIIString
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Dim tmpStr As String
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
       
    With frmPanelGm
        .txtCreador.Text = Buffer.ReadASCIIString
        .txtDescrip.Text = Buffer.ReadASCIIString
        
        'Status del pj
        If Buffer.ReadBoolean Then
            .lblEstado.ForeColor = vbGreen
            .lblEstado.Caption = UCase$(JsonLanguage.item("EN_LINEA").item("TEXTO"))
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = UCase$(JsonLanguage.item("DESCONECTADO").item("TEXTO"))
        End If
        
        'IP del personaje
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtIP.Text = tmpStr
        Else
            .txtIP.Text = JsonLanguage.item("USUARIO").item("TEXTO") & JsonLanguage.item("DESCONECTADO").item("TEXTO")
        End If
        
        'Tiempo online
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtTimeOn.Text = tmpStr
        Else
            .txtTimeOn.Text = JsonLanguage.item("USUARIO").item("TEXTO") & JsonLanguage.item("DESCONECTADO").item("TEXTO")
        End If
        
        'Observaciones
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtObs.Text = tmpStr
        Else
            .txtObs.Text = JsonLanguage.item("MENSAJE_NO_NOVEDADES").item("TEXTO")
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub


''
' Writes the "Moveitem" message to the outgoing data buffer.
'
Public Sub WriteMoveItem(ByVal originalSlot As Integer, ByVal newSlot As Integer, ByVal moveType As eMoveType)
'***************************************************
'Author: Budi
'Last Modification: 05/01/2011
'Writes the "MoveItem" message to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.moveItem)
        Call .WriteByte(originalSlot)
        Call .WriteByte(newSlot)
        Call .WriteByte(moveType)
    End With
End Sub

Private Sub HandlePalabrasMagicas()

    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Spell As Integer
    Dim CharIndex As Integer
    
    Spell = incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If Char_Check(CharIndex) Then _
        Call Dialogos.CreateDialog(Hechizos(Spell).PalabrasMagicas, CharIndex, RGB(200, 250, 150))

End Sub

Private Sub HandleAttackAnim()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim CharIndex As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    'Set the animation trigger on true
    charlist(CharIndex).attacking = True 'should be done in separated sub?
End Sub

Private Sub HandleFXtoMap()

    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim X As Integer, Y As Integer, FxIndex As Integer, Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Loops = incomingData.ReadByte
    X = incomingData.ReadInteger
    Y = incomingData.ReadInteger
    FxIndex = incomingData.ReadInteger

    'Set the fx on the map
    With MapData(X, Y) 'TODO: hay que hacer una funcion separada que haga esto
        .FxIndex = FxIndex
    
        If .FxIndex > 0 Then
                        
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)
            .fX.Loops = Loops

        End If

    End With

End Sub

Private Sub HandleAccountLogged()

    If incomingData.Length < 30 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    AccountName = Buffer.ReadASCIIString
    AccountHash = Buffer.ReadASCIIString
    NumberOfCharacters = Buffer.ReadByte

    frmPanelAccount.Show

    If NumberOfCharacters > 0 Then
    
        ReDim cPJ(1 To NumberOfCharacters) As PjCuenta
        
        Dim LoopC As Long
        
        For LoopC = 1 To NumberOfCharacters
        
            With cPJ(LoopC)
                .Nombre = Buffer.ReadASCIIString
                .Body = Buffer.ReadInteger
                .Head = Buffer.ReadInteger
                .weapon = Buffer.ReadInteger
                .shield = Buffer.ReadInteger
                .helmet = Buffer.ReadInteger
                .Class = Buffer.ReadByte
                .Race = Buffer.ReadByte
                .Map = Buffer.ReadInteger
                .Level = Buffer.ReadByte
                .Gold = Buffer.ReadLong
                .Criminal = Buffer.ReadBoolean
                .Dead = Buffer.ReadBoolean
                
                If .Dead Then
                    .Head = eCabezas.CASPER_HEAD
                    .Body = iCuerpoMuerto
                    .weapon = 0
                    .helmet = 0
                    .shield = 0
                ElseIf (.Body = 397 Or .Body = 395 Or .Body = 399) Then
                    .Head = 0
                End If

                .GameMaster = Buffer.ReadBoolean
            End With
            
            Call mDx8_Dibujado.DrawPJ(LoopC)
            
        Next LoopC
        
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleSearchList()
On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
 
    Dim num   As Integer
    Dim Datos As String
    Dim obj   As Boolean
        
    'Remove packet ID
    Call Buffer.ReadByte
   
    num = Buffer.ReadInteger()
    obj = Buffer.ReadBoolean()
    Datos = Buffer.ReadASCIIString()
 
    Call frmBuscar.AddItem(num, obj, Datos)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error
 
End Sub

Public Sub WriteSearchObj(ByVal BuscoObj As String)
 
        With outgoingData
        
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SearchObj)
           
                Call .WriteASCIIString(BuscoObj)
                
        End With

End Sub
 
Public Sub WriteSearchNpc(ByVal BuscoNpc As String)
 
        With outgoingData
        
                Call .WriteByte(ClientPacketID.GMCommands)
                Call .WriteByte(eGMCommands.SearchNpc)
       
                Call .WriteASCIIString(BuscoNpc)
                
        End With

End Sub

Public Sub WriteEnviaCvc()

        With outgoingData
                Call .WriteByte(ClientPacketID.Ecvc)
        End With

End Sub

Public Sub WriteAceptarCvc()

        With outgoingData
                Call .WriteByte(ClientPacketID.Acvc)
        End With

End Sub

Public Sub WriteIrCvc()

        With outgoingData
                Call .WriteByte(ClientPacketID.IrCvc)
        End With

End Sub

Public Sub WriteDragAndDropHechizos(ByVal Ant As Integer, ByVal Nov As Integer)

    With outgoingData
        .WriteByte (ClientPacketID.DragAndDropHechizos)
        .WriteInteger (Ant)
        .WriteInteger (Nov)

    End With

End Sub

Public Sub WriteQuest()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete Quest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.Quest)
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestDetailsRequest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)
    
    Call outgoingData.WriteByte(QuestSlot)
End Sub
 
Public Sub WriteQuestAccept()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestAccept al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
End Sub
 
Private Sub HandleQuestDetails()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestDetails del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim tmpStr As String
    Dim tmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer
    
    With Buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
        
        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay NPCs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay OBJs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If
 
        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With
    
    'Determinamos que formulario se muestra, seg�n si recibimos la informaci�n y la quest est� empezada o no.
    If QuestEmpezada Then
        frmQuests.txtInfo.Text = tmpStr
    Else
        frmQuestInfo.txtInfo.Text = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
    End If
    
    Call incomingData.CopyBuffer(Buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
 
    If Error <> 0 Then _
        Err.Raise Error
End Sub
 
Public Sub HandleQuestListSend()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Recibe y maneja el paquete QuestListSend del servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.Length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim i As Integer
    Dim tmpByte As Byte
    Dim tmpStr As String
    
    'Leemos el id del paquete
    Call Buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = Buffer.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Buffer.ReadASCIIString
        
        'Agregamos los items
        For i = 1 To tmpByte
            frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
        Next i
    End If
    
    'Mostramos el formulario
    frmQuests.Show vbModeless, frmMain
    
    'Pedimos la informaci�n de la primer quest (si la hay)
    If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(Buffer)
 
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
 
    If Error <> 0 Then _
        Err.Raise Error
End Sub
 
Public Sub WriteQuestListRequest()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestListRequest al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Escribe el paquete QuestAbandon al servidor.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)
End Sub

Private Sub HandleCreateDamage()
 
    ' @ Crea dano en pos X e Y.
 
    With incomingData
        
        ' Leemos el ID del paquete.
        .ReadByte
     
        Call mDx8_Dibujado.Damage_Create(.ReadByte(), .ReadByte(), 0, .ReadLong(), .ReadByte())
     
    End With
 
End Sub

Public Sub WriteCambiarContrasena()
    
    With outgoingData
    
        'Mando el ID del paquete
        Call .WriteByte(ClientPacketID.CambiarContrasena)
        
        'Mando los datos de la cuenta a modificar.
        Call .WriteASCIIString(AccountMailToRecover)
        Call .WriteASCIIString(AccountNewPassword)
    
    End With

End Sub
Private Sub HandleUserInEvent()
    Call incomingData.ReadByte
    
    UserEvento = Not UserEvento
End Sub


Public Sub WriteFightSend(ByVal ListUser As String, ByVal GldRequired As Long)
    
    With outgoingData
        Call .WriteByte(ClientPacketID.FightSend)
        Call .WriteASCIIString(ListUser)
        Call .WriteLong(GldRequired)
    End With
    
End Sub

Public Sub WriteFightAccept(ByVal UserName As String)
    
    With outgoingData
        Call .WriteByte(ClientPacketID.FightAccept)
        Call .WriteASCIIString(UserName)
    End With
    
End Sub

Public Sub WriteCloseGuild()
'***************************************************
'Author: Matias ezequiel (maTih.-)
'***************************************************

    Call outgoingData.WriteByte(ClientPacketID.CloseGuild)

End Sub

''
' Writes the "LimpiarMundo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLimpiarMundo()
'***************************************************
'Author: Jopi
'Last Modification: 11/01/2020
'Writes the "LimpiarMundo" message to the outgoing data buffer
'***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LimpiarMundo)
    End With
End Sub
    
' Handles the EquitandoToggle message.
Private Sub HandleEquitandoToggle()
'***************************************************
'Author: Lorwik
'Last Modification: 06/04/2020
'06/04/2020: FrankoH298 - Recibimos el contador para volver a equiparnos la montura.
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserEquitandoSegundosRestantes = incomingData.ReadLong()
    UserEquitando = Not UserEquitando
    
    If Not UserEquitando And UserEquitandoSegundosRestantes > 0 Then frmMain.timerPasarSegundo.Enabled = True
    Call SetSpeedUsuario
End Sub

Public Sub WriteObtenerDatosServer()

    With outgoingData
        'Mando el ID del paquete
        Call .WriteByte(ClientPacketID.ObtenerDatosServer)
    End With

End Sub

' Handles the EnviarDatosMessage message.
Private Sub HandleEnviarDatosServer()
'***************************************************
'Author: Recox
'Last Modification: 14/03/20
'Obtiene datos del server para imprimir en la lista.
'***************************************************
    ' If incomingData.Length < 4 Then
    '     Err.Raise incomingData.NotEnoughDataErrCode
    '     Exit Sub
    ' End If

On Error GoTo errhandler
    
    Dim MundoServidor As String
    Dim NombreServidor As String
    Dim DescripcionServidor As String
    Dim IpPublicaServidor As String
    Dim PuertoServidor As Integer
    Dim NivelMaximoServidor As Integer
    Dim MaxUsersSimultaneosServidor As Integer
    Dim CantidadUsuariosOnline As Integer
    Dim ExpMultiplierServidor As Integer
    Dim OroMultiplierServidor As Integer
    Dim OficioMultiplierServidor As Integer
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    'Get data and update form
    MundoServidor = Buffer.ReadASCIIString()
    NombreServidor = Buffer.ReadASCIIString()
    DescripcionServidor = Buffer.ReadASCIIString()
    NivelMaximoServidor = Buffer.ReadInteger()
    MaxUsersSimultaneosServidor = Buffer.ReadInteger()
    CantidadUsuariosOnline = Buffer.ReadInteger()
    ExpMultiplierServidor = Buffer.ReadInteger()
    OroMultiplierServidor = Buffer.ReadInteger()
    OficioMultiplierServidor = Buffer.ReadInteger()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    Dim MsPingResult As Long
        MsPingResult = (GetTickCount - pingTime)
        
    pingTime = 0

    Dim CountryCode As String
    
    If IpApiEnabled Then
        
        'If is not numeric do a url transformation
        If CheckIfIpIsNumeric(IpPublicaServidor) = False Then
            IpPublicaServidor = GetIPFromHostName(IpPublicaServidor)
        End If

        CountryCode = GetCountryCode(IpPublicaServidor) & " - "
    End If

    Dim Descripcion As String
    Descripcion = CountryCode & _
                    NombreServidor & vbNewLine & _
                    DescripcionServidor & vbNewLine & _
                    "Mundo: " & MundoServidor & vbNewLine & _
                    "Online: " & CantidadUsuariosOnline & " / " & MaxUsersSimultaneosServidor & vbNewLine & _
                    "Ping: " & MsPingResult & vbNewLine & _
                    "Nivel Maximo Permitido : " & NivelMaximoServidor


    frmConnect.lblDescripcionServidor = Descripcion

    STAT_MAXELV = NivelMaximoServidor

errhandler:

    Dim Error As Long
        Error = Err.number

    On Error GoTo 0

    If Error <> 0 Then Call Err.Raise(Error)

End Sub

Public Sub WriteAddAmigo(ByVal UserName As String, ByVal Index As Byte)

    '***************************************************
    'Author: Abusivo#1215 (DISCORD)
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AddAmigos)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(Index)
    End With
End Sub
Public Sub WriteDelAmigo(ByVal Index As Byte)

    '***************************************************
    'Author: Abusivo#1215 (DISCORD)
    '***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.DelAmigos)
        Call .WriteByte(Index)
    End With
    
End Sub

Public Sub WriteOnAmigo()

    '***************************************************
    'Author: Abusivo#1215 (DISCORD)
    '***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.OnAmigos)
    End With
    
End Sub

Public Sub WriteMsgAmigo(ByVal msg As String)

    '***************************************************
    'Author: Abusivo#1215 (DISCORD)
    '***************************************************
    
    With outgoingData
        Call .WriteByte(ClientPacketID.MsgAmigos)
        Call .WriteASCIIString(msg)
    End With
    
End Sub

Private Sub HandleEnviarListDeAmigos()

    '***************************************************
    'Author: Abusivo#1215 (DISCORD)
    '***************************************************
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call Buffer.ReadByte

    Dim slot As Byte
    Dim nombreamigo As String
    
    slot = Buffer.ReadByte()
    
    nombreamigo = Buffer.ReadASCIIString()
    
    amigos(slot) = nombreamigo
    
    Call frmAmigos.ActualizarLista
    
    If frmAmigos.Visible Then
        frmAmigos.SetFocus
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

errhandler:

    Dim Error As Long
        Error = Err.number
        
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Call Err.Raise(Error)
    
End Sub

Public Sub WriteLookProcess(ByVal data As String)
'***************************************************
'Author: Franco Emmanuel Gimenez (Franeg95)
'Last Modification: 18/10/10
'Writes the "Lookprocess" message and write the nickname of another user to the outgoing data buffer
'***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Lookprocess)
        Call .WriteASCIIString(data)
    End With
End Sub
 
Public Sub WriteSendProcessList()
'***************************************************
'Author: Franco Emmanuel Gimenez (Franeg95)
'Last Modification: 18/10/10
'Writes the "SendProcessList" message and write the process list of another user to the outgoing data buffer
'***************************************************
    Dim ProcesosList As String
    Dim CaptionsList As String

    ProcesosList = ListarProcesosUsuario()
    ProcesosList = Replace(ProcesosList, " ", "|")

    CaptionsList = ListarCaptionsUsuario()
    CaptionsList = Replace(CaptionsList, "#", "|")
    
    With outgoingData
        Call .WriteByte(ClientPacketID.SendProcessList)
        Call .WriteASCIIString(CaptionsList)
        Call .WriteASCIIString(ProcesosList)
    End With
End Sub
 
Private Sub HandleSeeInProcess()

    Call incomingData.ReadByte
    Call WriteSendProcessList
End Sub

Private Sub HandleShowProcess()

    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo errhandler
    
    Dim tmpCaptions() As String, tmpProcessList() As String
    Dim Captions As String, ProcessList As String
    Dim i As Long
    
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)

    Call Buffer.ReadByte

    Captions = Buffer.ReadASCIIString()
    ProcessList = Buffer.ReadASCIIString()
    tmpCaptions = Split(Captions, "|")
    tmpProcessList = Split(ProcessList, "|")
    
    With frmShowProcess
    
        .lstCaptions.Clear
        .lstProcess.Clear
        
        For i = LBound(tmpCaptions) To UBound(tmpCaptions)
            Call .lstCaptions.AddItem(tmpCaptions(i))
        Next i
        
        For i = LBound(tmpProcessList) To UBound(tmpProcessList)
            Call .lstProcess.AddItem(tmpProcessList(i))
        Next i
        
        .Show , frmMain
        
    End With
    
    Call incomingData.CopyBuffer(Buffer)

errhandler:
    Dim Error As Long
    Error = Err.number
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Call Err.Raise(Error)
End Sub


Private Sub HandleProyectil()
'**************************
'Autor: Lorwik
'Last Modification: 12/07/2020
'**************************

    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
        
    Dim CharSending      As Integer
    Dim CharRecieved     As Integer
    Dim GrhIndex         As Integer
        
    CharSending = incomingData.ReadInteger()
    CharRecieved = incomingData.ReadInteger()
    GrhIndex = incomingData.ReadInteger()
    
    Engine_Projectile_Create CharSending, CharRecieved, GrhIndex, 0
End Sub
