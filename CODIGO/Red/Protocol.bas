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
    PlayIsInChatMode = 123
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
    SendIfCharIsInChatMode = 156
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
            
        Case ServerPacketID.PlayIsInChatMode
            Call HandleSetTypingFlagToCharIndex
            
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
        If Inventario.ObjIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
            End With
        End If
    Next i
    
    ' Fill Npc inventory
    For i = 1 To MAX_NPC_INVENTORY_SLOTS
        If NPCInventory(i).ObjIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .ObjIndex, _
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
            Call InvBanco(1).SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
        End With
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .ObjIndex, _
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
        If Inventario.ObjIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), _
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
    TiempoDopas = IntervaloDopas * 0.04
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
    TiempoDopas = IntervaloDopas * 0.04
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
    TiempoDopas = IntervaloDopas * 0.04
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
    mapInfo.Zona = incomingData.ReadASCIIString

    'Convertimos esto a mayuscula para evitar problemas por que entran mal los datos en los .dat (Recox)
    mapInfo.Zona = UCase(mapInfo.Zona)
    
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
    Dim ObjIndex As Integer
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
    ObjIndex = Buffer.ReadInteger()
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
    
    Call Inventario.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, name, Incompatible)

    If frmComerciar.Visible Then
        Call InvComUsu.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, name, Incompatible)
    End If

    If frmBancoObj.Visible Then
        Call InvBanco(1).SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, name, Incompatible)
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
            Call frmComerciarUsu.UpdateInvCom(.ObjIndex(slot), Amount)
            
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
        .ObjIndex = Buffer.ReadInteger()
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
            Call InvBanco(0).SetItem(slot, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .name, .Incompatible)
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
            .ObjIndex = Buffer.ReadInteger()
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
                    If .Upgrade = ArmasHerrero(k).ObjIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).name = .name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).ObjIndex = .ObjIndex
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
            .ObjIndex = Buffer.ReadInteger()
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
                    If .Upgrade = ArmadurasHerrero(k).ObjIndex Then
                        J = J + 1
                
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        
                        HerreroMejorar(J).name = .name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).ObjIndex = .ObjIndex
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
            .ObjIndex = Buffer.ReadInteger()
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
                    If .Upgrade = ObjCarpintero(k).ObjIndex Then
                        J = J + 1
                
                        ReDim Preserve CarpinteroMejorar(J) As tItemsConstruibles
                        
                        CarpinteroMejorar(J).name = .name
                        CarpinteroMejorar(J).GrhIndex = .GrhIndex
                        CarpinteroMejorar(J).ObjIndex = .ObjIndex
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
            .ObjIndex = Buffer.ReadInteger()
            
            CountCrafteo = Buffer.ReadByte()
            ReDim .ItemsCrafteo(CountCrafteo) As tItemCrafteo
            
            For J = 1 To CountCrafteo
                .ItemsCrafteo(J).name = Buffer.ReadASCIIString()
                .ItemsCrafteo(J).GrhIndex = Buffer.ReadLong()
                .ItemsCrafteo(J).ObjIndex = Buffer.ReadInteger()
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
        .ObjIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger()
        .Incompatible = Buffer.ReadBoolean()
    
        If frmComerciar.Visible Then
            Call InvComNpc.SetItem(slot, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .name, .Incompatible)
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

        ' aobench: auto-select character by name when launched with /bot flag
        If BotAutoLogin And Len(BotCharName) > 0 Then
            Dim autoIdx As Long
            For autoIdx = 1 To NumberOfCharacters
                If UCase$(cPJ(autoIdx).Nombre) = UCase$(BotCharName) Then
                    UserName = cPJ(autoIdx).Nombre
                    Call WriteLoginExistingChar
                    Exit For
                End If
            Next autoIdx
        End If
