Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
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
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit
#If False Then
    Dim Nombre, PicInv, status, length As Variant
#End If
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

    logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildMemberInfo
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails
    
    ShowGuildAlign
    ShowPartyForm
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    DecirPalabrasMagicas
    PlayAttackAnim
    FXtoMap
    AccountLogged

End Enum

Private Enum ClientPacketID

    LoginExistingChar = 1     'OLOGIN
    ThrowDices = 2            'TIRDAD
    LoginNewChar = 3          'NLOGIN
    Talk = 4                  ';
    Yell = 5                  '-
    Whisper = 6                 '\
    Walk = 7                     'M
    RequestPositionUpdate = 8    'RPU
    Attack = 9                  'AT
    PickUp = 10                   'AG
    SafeToggle = 11              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ResuscitationSafeToggle = 12
    RequestGuildLeaderInfo = 13   'GLINFO
    RequestAtributes = 14         'ATR
    RequestFame = 15              'FAMA
    RequestSkills = 16            'ESKI
    RequestMiniStats = 17         'FEST
    CommerceEnd = 18             'FINCOM
    UserCommerceEnd = 19         'FINCOMUSU
    UserCommerceConfirm = 20
    CommerceChat = 21
    BankEnd = 22                'FINBAN
    UserCommerceOk = 23           'COMUSUOK
    UserCommerceReject = 24       'COMUSUNO
    Drop = 25                   'TI
    CastSpell = 26                'LH
    LeftClick = 27               'LC
    DoubleClick = 28             'RC
    Work = 29                     'UK
    UseSpellMacro = 30           'UMH
    UseItem = 31              'USA
    CraftBlacksmith = 32          'CNS
    CraftCarpenter = 33           'CNC
    WorkLeftClick = 34           'WLC
    CreateNewGuild = 35           'CIG
    sadasdA = 36
    EquipItem = 37               'EQUI
    ChangeHeading = 38           'CHEA
    ModifySkills = 39             'SKSE
    Train = 40                   'ENTR
    CommerceBuy = 41              'COMP
    BankExtractItem = 42          'RETI
    CommerceSell = 43            'VEND
    BankDeposit = 44              'DEPO
    ForumPost = 45                'DEMSG
    MoveSpell = 46               'DESPHE
    MoveBank = 47
    ClanCodexUpdate = 48         'DESCOD
    UserCommerceOffer = 49        'OFRECER
    GuildAcceptPeace = 50         'ACEPPEAT
    GuildRejectAlliance = 51      'RECPALIA
    GuildRejectPeace = 52        'RECPPEAT
    GuildAcceptAlliance = 53      'ACEPALIA
    GuildOfferPeace = 54          'PEACEOFF
    GuildOfferAlliance = 55       'ALLIEOFF
    GuildAllianceDetails = 56     'ALLIEDET
    GuildPeaceDetails = 57        'PEACEDET
    GuildRequestJoinerInfo = 58   'ENVCOMEN
    GuildAlliancePropList = 59    'ENVALPRO
    GuildPeacePropList = 60       'ENVPROPP
    GuildDeclareWar = 61          'DECGUERR
    GuildNewWebsite = 62          'NEWWEBSI
    GuildAcceptNewMember = 63     'ACEPTARI
    GuildRejectNewMember = 64     'RECHAZAR
    GuildKickMember = 65         'ECHARCLA
    GuildUpdateNews = 66          'ACTGNEWS
    GuildMemberInfo = 67          '1HRINFO<
    GuildOpenElections = 68       'ABREELEC
    GuildRequestMembership = 69   'SOLICITUD
    GuildRequestDetails = 70      'CLANDETAILS
    Online = 71                  '/ONLINE
    Quit = 72                     '/SALIR
    GuildLeave = 73               '/SALIRCLAN
    RequestAccountState = 74      '/BALANCE
    PetStand = 75                 '/QUIETO
    PetFollow = 76                '/ACOMPAÑAR
    ReleasePet = 77              '/LIBERAR
    TrainList = 78                '/ENTRENAR
    Rest = 79                     '/DESCANSAR
    Meditate = 80                '/MEDITAR
    Resucitate = 81               '/RESUCITAR
    Heal = 82                     '/CURAR
    Help = 83                    '/AYUDA
    RequestStats = 84             '/EST
    CommerceStart = 85           '/COMERCIAR
    BankStart = 86               '/BOVEDA
    Enlist = 87                   '/ENLISTAR
    Information = 88            '/INFORMACION
    Reward = 89                   '/RECOMPENSA
    RequestMOTD = 90              '/MOTD
    UpTime = 91                   '/UPTIME
    PartyLeave = 92               '/SALIRPARTY
    PartyCreate = 93              '/CREARPARTY
    PartyJoin = 94                '/PARTY
    Inquiry = 95                  '/ENCUESTA ( with no params )
    GuildMessage = 96             '/CMSG
    PartyMessage = 97             '/PMSG
    GuildOnline = 98              '/ONLINECLAN
    PartyOnline = 99             '/ONLINEPARTY
    CouncilMessage = 100           '/BMSG
    RoleMasterRequest = 101     '/ROL
    GMRequest = 102              '/GM
    bugReport = 103              '/_BUG
    ChangeDescription = 104      '/DESC
    GuildVote = 105              '/VOTO
    Punishments = 106           '/PENAS
    ChangePassword = 107         '/CONTRASEÑA
    Gamble = 108                '/APOSTAR
    InquiryVote = 109            '/ENCUESTA ( with parameters )
    LeaveFaction = 110          '/RETIRAR ( with no arguments )
    BankExtractGold = 111        '/RETIRAR ( with arguments )
    BankDepositGold = 112        '/DEPOSITAR
    Denounce = 113               '/DENUNCIAR
    GuildFundate = 114          '/FUNDARCLAN
    GuildFundation = 115
    PartyKick = 116              '/ECHARPARTY
    PartySetLeader = 117         '/PARTYLIDER
    PartyAcceptMember = 118      '/ACCEPTPARTY
    Ping = 119                  '/PING
    RequestPartyForm = 120
    ItemUpgrade = 121
    GMCommands = 122
    InitCrafting = 123
    Home = 124
    ShowGuildNews = 125
    ShareNpc = 126               '/COMPARTIR
    StopSharingNpc = 127
    Consultation = 128
    moveItem = 129
    LoginExistingAccount = 130
    LoginNewAccount = 131

End Enum

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS

End Enum

Public FontTypes(20) As tFont

''
' Initializes the fonts array

Public Sub InitFonts()
    
    On Error GoTo InitFonts_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 255
        .Green = 255
        .Blue = 255

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 32
        .Green = 51
        .Blue = 223
        .bold = 1
        .italic = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 65
        .Green = 190
        .Blue = 156

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 65
        .Green = 190
        .Blue = 156
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 130
        .Green = 130
        .Blue = 130
        .bold = 1

    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 255
        .Green = 180
        .Blue = 250

    End With
    
    FontTypes(FontTypeNames.FONTTYPE_VENENO).Green = 255
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 255
        .Green = 255
        .Blue = 255
        .bold = 1

    End With
    
    FontTypes(FontTypeNames.FONTTYPE_SERVER).Green = 185
    
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
        .Green = 255
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
        .Blue = 200
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

    
    Exit Sub

InitFonts_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "InitFonts"
    End If
Resume Next
    
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error Resume Next

    Dim Packet As Byte

    Packet = incomingData.PeekByte()
    Debug.Print Packet
    
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
        
        Case ServerPacketID.ShowBlacksmithForm      ' SFH
            Call HandleShowBlacksmithForm
        
        Case ServerPacketID.ShowCarpenterForm       ' SFC
            Call HandleShowCarpenterForm
        
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
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
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
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
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
        
        Case ServerPacketID.CarpenterObjects        ' OBR
            Call HandleCarpenterObjects
        
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
        
        Case ServerPacketID.TradeOK                 ' TRANSOK
            Call HandleTradeOK
        
        Case ServerPacketID.BankOK                  ' BANCOOK
            Call HandleBankOK
        
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
            
        Case ServerPacketID.DecirPalabrasMagicas
            Call HandleDecirPalabrasMagicas
            
        Case ServerPacketID.PlayAttackAnim
            Call HandleAttackAnim
            
        Case ServerPacketID.FXtoMap
            Call HandleFXtoMap
        
            'CHOTS | Accounts
        Case ServerPacketID.AccountLogged
            Call HandleAccountLogged

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
            
        Case Else
            'ERROR : Abort!
            Exit Sub

    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
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
    '***************************************************
    
    On Error GoTo HandleMultiMessage_Err
    
    Dim BodyPart As Byte
    Dim Daño As Integer
    Dim SpellIndex As Integer
    Dim Nombre     As String
    
    With incomingData
        Call .ReadByte
    
        Select Case .ReadByte

            Case eMessages.NPCSwing
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)
        
            Case eMessages.NPCKillUser
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)
        
            Case eMessages.BlockedWithShieldUser
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
            Case eMessages.BlockedWithShieldOther
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)
        
            Case eMessages.UserSwing
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)
        
            Case eMessages.SafeModeOn
                Call frmMain.ControlSM(eSMType.sSafemode, True)
        
            Case eMessages.SafeModeOff
                Call frmMain.ControlSM(eSMType.sSafemode, False)
        
            Case eMessages.ResuscitationSafeOff
                Call frmMain.ControlSM(eSMType.sResucitation, False)
         
            Case eMessages.ResuscitationSafeOn
                Call frmMain.ControlSM(eSMType.sResucitation, True)
        
            Case eMessages.NobilityLost
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)
        
            Case eMessages.CantUseWhileMeditating
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)
        
            Case eMessages.NPCHitUser

                Select Case incomingData.ReadByte()

                    Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)
                
                    Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)

                End Select
        
            Case eMessages.UserHitNPC
                Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Le has quitado " & CStr(incomingData.ReadLong()) & " puntos de vida a la criatura!!", 255, 0, 0, True, False, True)
        
            Case eMessages.UserAttackedSwing
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).Nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, True)
        
            Case eMessages.UserHittedByUser
                Dim AttackerName As String
            
                AttackerName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Daño = incomingData.ReadInteger()
            
                Select Case BodyPart

                    Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & AttackerName & MENSAJE_RECIVE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)

                End Select
        
            Case eMessages.UserHittedUser

                Dim VictimName As String
            
                VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Daño = incomingData.ReadInteger()
            
                Select Case BodyPart

                    Case bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_CABEZA & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Daño & MENSAJE_2, 255, 0, 0, True, False, True)
                
                    Case bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & VictimName & MENSAJE_PRODUCE_IMPACTO_TORSO & Daño & MENSAJE_2, 255, 0, 0, True, False, True)

                End Select
        
            Case eMessages.WorkRequestTarget
                UsingSkill = incomingData.ReadByte()
            
                frmMain.MousePointer = 2
            
                Select Case UsingSkill

                    Case Magia
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                
                    Case Pesca
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                
                    Case Robar
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                
                    Case Talar
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                
                    Case Mineria
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                
                    Case FundirMetal
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                
                    Case Proyectiles
                        Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)

                End Select

            Case eMessages.HaveKilledUser
                Dim KilledUser As Integer
                Dim Exp        As Long
            
                KilledUser = .ReadInteger
                Exp = .ReadLong
            
                Call ShowConsoleMsg(MENSAJE_HAS_MATADO_A & charlist(KilledUser).Nombre & MENSAJE_22, 255, 0, 0, True, False)
                Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & Exp & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
            
                'Sacamos un screenshot si está activado el FragShooter:
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
            
                Call ShowConsoleMsg(charlist(KillerUser).Nombre & MENSAJE_TE_HA_MATADO, 255, 0, 0, True, False)
            
                'Sacamos un screenshot si está activado el FragShooter:
                If ClientSetup.bDie And ClientSetup.bActive Then
                    FragShooterNickname = charlist(KillerUser).Nombre
                    FragShooterKilledSomeone = False
                
                    FragShooterCapturePending = True

                End If
                
            Case eMessages.EarnExp
                'Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXPE_1 & .ReadLong & MENSAJE_HAS_GANADO_EXPE_2, 255, 0, 0, True, False)
        
            Case eMessages.GoHome
                Dim Distance As Byte
                Dim Hogar    As String
                Dim tiempo   As Integer
                Dim msg      As String
            
                Distance = .ReadByte
                tiempo = .ReadInteger
                Hogar = .ReadASCIIString
            
                If tiempo >= 60 Then
                    If tiempo Mod 60 = 0 Then
                        msg = tiempo / 60 & " minutos."
                    Else
                        msg = CInt(tiempo \ 60) & " minutos y " & tiempo Mod 60 & " segundos."  'Agregado el CInt() asi el número no es con , [C4b3z0n - 09/28/2010]

                    End If

                Else
                    msg = tiempo & " segundos."

                End If
            
                Call ShowConsoleMsg("Te encuentras a " & Distance & " mapas de la " & Hogar & ", este viaje durará " & msg, 255, 0, 0, True)
                Traveling = True

            Case eMessages.CancelGoHome
                Call ShowConsoleMsg(MENSAJE_HOGAR_CANCEL, 255, 0, 0, True)
                Traveling = False
                   
            Case eMessages.FinishHome
                Call ShowConsoleMsg(MENSAJE_HOGAR, 255, 255, 255)
                Traveling = False
            
            Case eMessages.UserMuerto
                Call ShowConsoleMsg(MENSAJE_USER_MUERTO, 255, 255, 255)
        
            Case eMessages.NpcInmune
                Call ShowConsoleMsg(NPC_INMUNE, 210, 220, 220)
            
            Case eMessages.Hechizo_HechiceroMSG_NOMBRE
                SpellIndex = .ReadByte
                Nombre = .ReadASCIIString
         
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " " & Nombre & ".", 210, 220, 220)
         
            Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
                SpellIndex = .ReadByte
         
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " alguien.", 210, 220, 220)
         
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

    
    Exit Sub

HandleMultiMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleMultiMessage"
    End If
Resume Next
    
End Sub

''
' Handles the Logged message.

Private Sub HandleLogged()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleLogged_Err
    
    Call incomingData.ReadByte
    
    ' Variable initialization
    UserClase = incomingData.ReadByte
    EngineRun = True
    Nombres = True
    bRain = False
    
    'Set connected state
    Call SetConnected
    
    If bShowTutorial Then frmTutorial.Show vbModeless
    
    'Show tip
    If tipf = "1" And PrimeraVez Then
        Call CargarTip
        frmtip.Visible = True
        PrimeraVez = False

    End If

    
    Exit Sub

HandleLogged_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleLogged"
    End If
Resume Next
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRemoveDialogs_Err
    
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs

    
    Exit Sub

HandleRemoveDialogs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleRemoveDialogs"
    End If
Resume Next
    
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
    
    On Error GoTo HandleRemoveCharDialog_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check if the packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())

    
    Exit Sub

HandleRemoveCharDialog_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleRemoveCharDialog"
    End If
Resume Next
    
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNavigateToggle_Err
    
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando

    
    Exit Sub

HandleNavigateToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleNavigateToggle"
    End If
Resume Next
    
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleDisconnect_Err
    
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Close connection
    #If UsarWrench = 1 Then
        frmMain.Socket1.Disconnect
    #Else

        If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
    #End If
    ResetAllInfo

    
    Exit Sub

HandleDisconnect_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleDisconnect"
    End If
Resume Next
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCommerceEnd_Err
    
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    
    'Hide form
    Unload frmComerciar
    
    'Reset vars
    Comerciando = False

    
    Exit Sub

HandleCommerceEnd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCommerceEnd"
    End If
Resume Next
    
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBankEnd_Err
    
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False

    
    Exit Sub

HandleBankEnd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBankEnd"
    End If
Resume Next
    
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleCommerceInit_Err
    
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = New clsGrapchicalInventory
    Set InvComNpc = New clsGrapchicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, Inventario.MaxObjs)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS

        If Inventario.ObjIndex(i) <> 0 Then

            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))

            End With

        End If

    Next i
    
    ' Fill Npc inventory
    For i = 1 To 50

        If NPCInventory(i).ObjIndex <> 0 Then

            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name)

            End With

        End If

    Next i
    
    'Set state and show form
    Comerciando = True
    frmComerciar.Show , frmMain

    
    Exit Sub

HandleCommerceInit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCommerceInit"
    End If
Resume Next
    
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleBankInit_Err
    
    Dim i        As Long
    Dim BankGold As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = New clsGrapchicalInventory
    Set InvBanco(1) = New clsGrapchicalInventory
    
    BankGold = incomingData.ReadLong
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, Inventario.MaxObjs)
    
    For i = 1 To Inventario.MaxObjs

        With Inventario
            Call InvBanco(1).SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))

        End With

    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS

        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .ObjIndex, .Amount, .Equipped, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name)

        End With

    Next i
    
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.lblUserGld.Caption = BankGold
    
    frmBancoObj.Show , frmMain

    
    Exit Sub

HandleBankInit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBankInit"
    End If
Resume Next
    
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleUserCommerceInit_Err
    
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    TradingUserName = incomingData.ReadASCIIString
    
    Set InvComUsu = New clsGrapchicalInventory
    Set InvOfferComUsu(0) = New clsGrapchicalInventory
    Set InvOfferComUsu(1) = New clsGrapchicalInventory
    Set InvOroComUsu(0) = New clsGrapchicalInventory
    Set InvOroComUsu(1) = New clsGrapchicalInventory
    Set InvOroComUsu(2) = New clsGrapchicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciarUsu.picInvComercio, Inventario.MaxObjs)
    Call InvOfferComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    Call InvOfferComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
    Call InvOroComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS

        If Inventario.ObjIndex(i) <> 0 Then

            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))

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

    
    Exit Sub

HandleUserCommerceInit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUserCommerceInit"
    End If
Resume Next
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleUserCommerceEnd_Err
    
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

    
    Exit Sub

HandleUserCommerceEnd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUserCommerceEnd"
    End If
Resume Next
    
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
    
    On Error GoTo HandleUserOfferConfirm_Err
    
    Call incomingData.ReadByte
    
    With frmComerciarUsu
        ' Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg TradingUserName & " ha confirmado su oferta!", FontTypeNames.FONTTYPE_CONSE

    End With
    
    
    Exit Sub

HandleUserOfferConfirm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUserOfferConfirm"
    End If
Resume Next
    
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowBlacksmithForm_Err
    
    Call incomingData.ReadByte
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
    Else
        frmHerrero.Show , frmMain
        MirandoHerreria = True

    End If

    
    Exit Sub

HandleShowBlacksmithForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleShowBlacksmithForm"
    End If
Resume Next
    
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowCarpenterForm_Err
    
    Call incomingData.ReadByte
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
    Else
        frmCarp.Show , frmMain
        MirandoCarpinteria = True

    End If

    
    Exit Sub

HandleShowCarpenterForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleShowCarpenterForm"
    End If
Resume Next
    
End Sub

''
' Handles the NPCSwing message.

Private Sub HandleNPCSwing()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNPCSwing_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, True)

    
    Exit Sub

HandleNPCSwing_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleNPCSwing"
    End If
Resume Next
    
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNPCKillUser_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, True)

    
    Exit Sub

HandleNPCKillUser_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleNPCKillUser"
    End If
Resume Next
    
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlockedWithShieldUser_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)

    
    Exit Sub

HandleBlockedWithShieldUser_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBlockedWithShieldUser"
    End If
Resume Next
    
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlockedWithShieldOther_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, True)

    
    Exit Sub

HandleBlockedWithShieldOther_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBlockedWithShieldOther"
    End If
Resume Next
    
End Sub

''
' Handles the UserSwing message.

Private Sub HandleUserSwing()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleUserSwing_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, True)

    
    Exit Sub

HandleUserSwing_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUserSwing"
    End If
Resume Next
    
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleSafeModeOn_Err
    
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, True)

    
    Exit Sub

HandleSafeModeOn_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleSafeModeOn"
    End If
Resume Next
    
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleSafeModeOff_Err
    
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sSafemode, False)

    
    Exit Sub

HandleSafeModeOff_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleSafeModeOff"
    End If
Resume Next
    
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandleResuscitationSafeOff()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleResuscitationSafeOff_Err
    
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, False)

    
    Exit Sub

HandleResuscitationSafeOff_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleResuscitationSafeOff"
    End If
Resume Next
    
End Sub

''
' Handles the ResuscitationSafeOn message.

Private Sub HandleResuscitationSafeOn()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleResuscitationSafeOn_Err
    
    Call incomingData.ReadByte
    
    Call frmMain.ControlSM(eSMType.sResucitation, True)

    
    Exit Sub

HandleResuscitationSafeOn_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleResuscitationSafeOn"
    End If
Resume Next
    
End Sub

''
' Handles the NobilityLost message.

Private Sub HandleNobilityLost()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNobilityLost_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, True)

    
    Exit Sub

HandleNobilityLost_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleNobilityLost"
    End If
Resume Next
    
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCantUseWhileMeditating_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, True)

    
    Exit Sub

HandleCantUseWhileMeditating_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCantUseWhileMeditating"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
    
    On Error GoTo HandleUpdateSta_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    Dim bWidth As Byte
    
    bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 75)
    
    frmMain.shpEnergia.Width = 75 - bWidth
    frmMain.shpEnergia.Left = 584 + (75 - frmMain.shpEnergia.Width)
    
    frmMain.shpEnergia.Visible = (bWidth <> 75)
    
    
    Exit Sub

HandleUpdateSta_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateSta"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
    
    On Error GoTo HandleUpdateMana_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    
    Dim bWidth As Byte
    
    If UserMaxMAN > 0 Then bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 75)
        
    frmMain.shpMana.Width = 75 - bWidth
    frmMain.shpMana.Left = 584 + (75 - frmMain.shpMana.Width)
    
    frmMain.shpMana.Visible = (bWidth <> 75)

    
    Exit Sub

HandleUpdateMana_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateMana"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
    
    On Error GoTo HandleUpdateHP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserMinHP = incomingData.ReadInteger()
    
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    
    Dim bWidth As Byte
    
    bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 75)
    
    frmMain.shpVida.Width = 75 - bWidth
    frmMain.shpVida.Left = 584 + (75 - frmMain.shpVida.Width)
    
    frmMain.shpVida.Visible = (bWidth <> 75)
    
    'Is the user alive??
    If UserMinHP = 0 Then
        UserEstado = 1

        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0

    End If

    
    Exit Sub

HandleUpdateHP_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateHP"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
    
    On Error GoTo HandleUpdateGold_Err
    

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 09/21/10
    'Last Modified By: C4b3z0n
    '- 08/14/07: Tavo - Added GldLbl color variation depending on User Gold and Level
    '- 09/21/10: C4b3z0n - Modified color change of gold ONLY if the player's level is greater than 12 (NOT newbie).
    '***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    If UserGLD >= CLng(UserLvl) * 10000 And UserLvl > 12 Then 'Si el nivel es mayor de 12, es decir, no es newbie.
        'Changes color
        frmMain.GldLbl.ForeColor = &HFF& 'Red
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow

    End If
    
    frmMain.GldLbl.Caption = UserGLD

    
    Exit Sub

HandleUpdateGold_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateGold"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
    
    On Error GoTo HandleUpdateBankGold_Err
    

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
    
    
    Exit Sub

HandleUpdateBankGold_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateBankGold"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
    
    On Error GoTo HandleUpdateExp_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserExp = incomingData.ReadLong()
    frmMain.lblExp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
    frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"

    
    Exit Sub

HandleUpdateExp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateExp"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
    
    On Error GoTo HandleUpdateStrenghtAndDexterity_Err
    

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
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

    
    Exit Sub

HandleUpdateStrenghtAndDexterity_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateStrenghtAndDexterity"
    End If
Resume Next
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
    
    On Error GoTo HandleUpdateStrenght_Err
    

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '***************************************************
    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblStrg.ForeColor = getStrenghtColor()

    
    Exit Sub

HandleUpdateStrenght_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateStrenght"
    End If
Resume Next
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
    
    On Error GoTo HandleUpdateDexterity_Err
    

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '***************************************************
    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblDext.ForeColor = getDexterityColor()

    
    Exit Sub

HandleUpdateDexterity_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateDexterity"
    End If
Resume Next
    
End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
    
    On Error GoTo HandleChangeMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    
    'TODO: Once on-the-fly editor is implemented check for map version before loading....
    'For now we just drop it
    Call incomingData.ReadInteger
    
    If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
        Call SwitchMap(UserMap)

        If bRain And bLluvia(UserMap) = 0 Then
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone

        End If

    Else
        'no encontramos el mapa en el hd
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        
        Call CloseClient

    End If

    
    Exit Sub

HandleChangeMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleChangeMap"
    End If
Resume Next
    
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
    
    On Error GoTo HandlePosUpdate_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
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

    
    Exit Sub

HandlePosUpdate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandlePosUpdate"
    End If
Resume Next
    
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
    
    On Error GoTo HandleNPCHitUser_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Select Case incomingData.ReadByte()

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()) & "!!", 255, 0, 0, True, False, True)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger() & "!!"), 255, 0, 0, True, False, True)

    End Select

    
    Exit Sub

HandleNPCHitUser_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleNPCHitUser"
    End If
Resume Next
    
End Sub

''

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat      As String
    Dim CharIndex As Integer
    Dim r         As Byte
    Dim g         As Byte
    Dim b         As Byte
    
    chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    
    r = Buffer.ReadByte()
    g = Buffer.ReadByte()
    b = Buffer.ReadByte()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If Char_Check(CharIndex) Then Call Dialogos.CreateDialog(Trim$(chat), CharIndex, RGB(r, g, b))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)

ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat      As String
    Dim FontIndex As Integer
    Dim str       As String
    Dim r         As Byte
    Dim g         As Byte
    Dim b         As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            g = 255
        Else
            g = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
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
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08 (NicoNZ)
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat As String
    Dim str  As String
    Dim r    As Byte
    Dim g    As Byte
    Dim b    As Byte
    Dim tmp  As Integer
    Dim Cont As Integer
    
    chat = Buffer.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)

            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)

            End If
            
            str = ReadField(3, chat, 126)

            If Val(str) > 255 Then
                g = 255
            Else
                g = Val(str)

            End If
            
            str = ReadField(4, chat, 126)

            If Val(str) > 255 Then
                b = 255
            Else
                b = Val(str)

            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
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
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim chat      As String
    Dim FontIndex As Integer
    Dim str       As String
    Dim r         As Byte
    Dim g         As Byte
    Dim b         As Byte
    
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            g = 255
        Else
            g = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .Red, .Green, .Blue, .bold, .italic)

        End With

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
    
    On Error GoTo HandleUserIndexInServer_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()

    
    Exit Sub

HandleUserIndexInServer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUserIndexInServer"
    End If
Resume Next
    
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
    
    On Error GoTo HandleUserCharIndexInServer_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Char_UserIndexSet(incomingData.ReadInteger())
                     
    'Update pos label
    Call Char_UserPos

    
    Exit Sub

HandleUserCharIndexInServer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUserCharIndexInServer"
    End If
Resume Next
    
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body      As Integer
    Dim Head      As Integer
    Dim Heading   As E_Heading
    Dim X         As Byte
    Dim Y         As Byte
    Dim weapon    As Integer
    Dim shield    As Integer
    Dim helmet    As Integer
    Dim privs     As Integer
    Dim NickColor As Byte
    
    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    weapon = Buffer.ReadInteger()
    shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()
    
    With charlist(CharIndex)
        Call Char_SetFx(CharIndex, Buffer.ReadInteger(), Buffer.ReadInteger())
        
        .Nombre = Buffer.ReadASCIIString()
        NickColor = Buffer.ReadByte()
        
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0

        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        privs = Buffer.ReadByte()
        
        If privs <> 0 Then

            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil

            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil

            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster

            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0

        End If

    End With
    
    Call Char_Make(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)
    
    Call Char_RefreshAll
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleCharacterChangeNick()
    
    On Error GoTo HandleCharacterChangeNick_Err
    

    '***************************************************
    'Author: Budi
    'Last Modification: 07/23/09
    '
    '***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    
    Call Char_SetName(CharIndex, incomingData.ReadASCIIString)
    
    
    Exit Sub

HandleCharacterChangeNick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCharacterChangeNick"
    End If
Resume Next
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
    
    On Error GoTo HandleCharacterRemove_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    Call Char_Erase(CharIndex)
    Call RefreshAllChars

    
    Exit Sub

HandleCharacterRemove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCharacterRemove"
    End If
Resume Next
    
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
    
    On Error GoTo HandleCharacterMove_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim X         As Byte
    Dim Y         As Byte
    
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
    
    Call Char_RefreshAll

    
    Exit Sub

HandleCharacterMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCharacterMove"
    End If
Resume Next
    
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    On Error GoTo HandleForceCharMove_Err
    
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call Char_MovebyHead(UserCharIndex, Direccion)
    Call Char_MoveScreen(Direccion)
    
    Call Char_RefreshAll

    
    Exit Sub

HandleForceCharMove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleForceCharMove"
    End If
Resume Next
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 21/09/2010 - C4b3z0n
    '25/08/2009: ZaMa - Changed a variable used incorrectly.
    '21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
    '***************************************************
    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim tempint   As Integer
    Dim HeadIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    
    '// Char Body
    Call Char_SetBody(CharIndex, incomingData.ReadInteger())

    '// Char Head
    Call Char_SetHead(CharIndex, incomingData.ReadInteger)
        
    '// Char Heading
    Call Char_SetHeading(CharIndex, incomingData.ReadByte())
        
    '// Char Weapon
    Call Char_SetWeapon(CharIndex, incomingData.ReadInteger())
        
    '// Char Shield
    Call Char_SetShield(CharIndex, incomingData.ReadInteger())
        
    '// Char Casco
    Call Char_SetCasco(CharIndex, incomingData.ReadInteger())
        
    '// Char Fx
    Call Char_SetFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
        
    Call Char_RefreshAll

    
    Exit Sub

HandleCharacterChange_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCharacterChange"
    End If
Resume Next
    
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
    
    On Error GoTo HandleObjectCreate_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X        As Byte
    Dim Y        As Byte
    Dim GrhIndex As Integer
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    GrhIndex = incomingData.ReadInteger()
        
    Call Map_CreateObject(X, Y, GrhIndex)

    
    Exit Sub

HandleObjectCreate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleObjectCreate"
    End If
Resume Next
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
    
    On Error GoTo HandleObjectDelete_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
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

    
    Exit Sub

HandleObjectDelete_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleObjectDelete"
    End If
Resume Next
    
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
    
    On Error GoTo HandleBlockPosition_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim X     As Byte
    Dim Y     As Byte
    Dim block As Boolean
    
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    block = incomingData.ReadBoolean()
    
    If block Then
        Map_SetBlocked X, Y, 1
    Else
        Map_SetBlocked X, Y, 0

    End If

    
    Exit Sub

HandleBlockPosition_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBlockPosition"
    End If
Resume Next
    
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
    
    On Error GoTo HandlePlayMIDI_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim currentMidi As Integer
    Dim Loops       As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMidi = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    If currentMidi Then
        If currentMidi > MP3_INITIAL_INDEX Then
            'Call Audio.MusicMP3Play(App.path & "\MP3\" & currentMidi & ".mp3")
        Else
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Loops)

        End If

    End If
    
    
    Exit Sub

HandlePlayMIDI_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandlePlayMIDI"
    End If
Resume Next
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
    
    On Error GoTo HandlePlayWave_Err
    

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
    If incomingData.length < 3 Then
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

    
    Exit Sub

HandlePlayWave_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandlePlayWave"
    End If
Resume Next
    
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

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
        Dim Upper_guildNames As Long
        
        Upper_guildNames = UBound(GuildNames())
    
        For i = 0 To Upper_guildNames
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
    
    On Error GoTo HandleAreaChanged_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
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

    
    Exit Sub

HandleAreaChanged_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleAreaChanged"
    End If
Resume Next
    
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandlePauseToggle_Err
    
    Call incomingData.ReadByte
    
    pausa = Not pausa

    
    Exit Sub

HandlePauseToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandlePauseToggle"
    End If
Resume Next
    
End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRainToggle_Err
    
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = 1 Or MapData(UserPos.X, UserPos.Y).Trigger = 2 Or MapData(UserPos.X, UserPos.Y).Trigger = 4)
            
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

    
    Exit Sub

HandleRainToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleRainToggle"
    End If
Resume Next
    
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
    
    On Error GoTo HandleCreateFX_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX        As Integer
    Dim Loops     As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call Char_SetFx(CharIndex, fX, Loops)

    
    Exit Sub

HandleCreateFX_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCreateFX"
    End If
Resume Next
    
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
    
    On Error GoTo HandleUpdateUserStats_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 26 Then
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
    
    frmMain.lblExp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
    
    If UserPasarNivel > 0 Then
        frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
    Else
        frmMain.lblPorcLvl.Caption = "[N/A]"

    End If
    
    frmMain.GldLbl.Caption = UserGLD
    frmMain.lblLvl.Caption = UserLvl
    
    'Stats
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    
    Dim bWidth As Byte
    
    '***************************
    If UserMaxMAN > 0 Then bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 75)
        
    frmMain.shpMana.Width = 75 - bWidth
    frmMain.shpMana.Left = 584 + (75 - frmMain.shpMana.Width)
    
    frmMain.shpMana.Visible = (bWidth <> 75)
    '***************************
    
    bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 75)
    
    frmMain.shpVida.Width = 75 - bWidth
    frmMain.shpVida.Left = 584 + (75 - frmMain.shpVida.Width)
    
    frmMain.shpVida.Visible = (bWidth <> 75)
    '***************************
    
    bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 75)
    
    frmMain.shpEnergia.Width = 75 - bWidth
    frmMain.shpEnergia.Left = 584 + (75 - frmMain.shpEnergia.Width)
    
    frmMain.shpEnergia.Visible = (bWidth <> 75)
    '***************************
    
    If UserMinHP = 0 Then
        UserEstado = 1

        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0

    End If
    
    If UserGLD >= CLng(UserLvl) * 10000 Then
        'Changes color
        frmMain.GldLbl.ForeColor = &HFF& 'Red
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = &HFFFF& 'Yellow

    End If

    
    Exit Sub

HandleUpdateUserStats_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateUserStats"
    End If
Resume Next
    
End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
    
    On Error GoTo HandleWorkRequestTarget_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UsingSkill = incomingData.ReadByte()

    frmMain.MousePointer = 2
    
    Select Case UsingSkill

        Case Magia
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)

        Case Pesca
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)

        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)

        Case Talar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)

        Case Mineria
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)

        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)

        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)

    End Select

    
    Exit Sub

HandleWorkRequestTarget_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleWorkRequestTarget"
    End If
Resume Next
    
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot     As Byte
    Dim ObjIndex As Integer
    Dim Name     As String
    Dim Amount   As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Integer
    Dim OBJType  As Byte
    Dim MaxHit   As Integer
    Dim MinHit   As Integer
    Dim MaxDef   As Integer
    Dim MinDef   As Integer
    Dim value    As Single
    
    slot = Buffer.ReadByte()
    ObjIndex = Buffer.ReadInteger()
    Name = Buffer.ReadASCIIString()
    Amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadInteger()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    MaxDef = Buffer.ReadInteger()
    MinDef = Buffer.ReadInteger
    value = Buffer.ReadSingle()
    
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
    
    Call Inventario.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, Name)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()
    '***************************************************
    'Author: Budi
    'Last Modification: 12/01/09
    '
    '***************************************************
    
    On Error GoTo HandleAddSlots_Err
    

    Call incomingData.ReadByte
    
    MaxInventorySlots = incomingData.ReadByte

    
    Exit Sub

HandleAddSlots_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleAddSlots"
    End If
Resume Next
    
End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
    '***************************************************
    'Author: Budi
    'Last Modification: 12/01/09
    '
    '***************************************************
    
    On Error GoTo HandleStopWorking_Err
    

    Call incomingData.ReadByte
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡Has terminado de trabajar!", .Red, .Green, .Blue, .bold, .italic)

    End With
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo

    
    Exit Sub

HandleStopWorking_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleStopWorking"
    End If
Resume Next
    
End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/03/10
    '
    '***************************************************
    
    On Error GoTo HandleCancelOfferItem_Err
    
    Dim slot   As Byte
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
    
    ' Si era el único ítem de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then Call frmComerciarUsu.HabilitarConfirmar(False)
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg("¡No puedes comerciar ese objeto!", FontTypeNames.FONTTYPE_INFO)

    End With

    
    Exit Sub

HandleCancelOfferItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleCancelOfferItem"
    End If
Resume Next
    
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .ObjIndex = Buffer.ReadInteger()
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger
        .Valor = Buffer.ReadLong()
        
        If Comerciando Then
            Call InvBanco(0).SetItem(slot, .ObjIndex, .Amount, .Equipped, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name)

        End If

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
 
    On Error GoTo ErrHandler

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

        If str <> vbNullString Then frmMain.hlst.List(slot - 1) = str
    Else
        str = DevolverNombreHechizo(UserHechizos(slot))

        If str <> vbNullString Then Call frmMain.hlst.AddItem(str)
     
    End If
 
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
 
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
 
    'Destroy auxiliar buffer
    Set Buffer = Nothing
 
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
    
    On Error GoTo HandleAtributes_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 1 + NUMATRIBUTES Then
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

    
    Exit Sub

HandleAtributes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleAtributes"
    End If
Resume Next
    
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count

        With ArmasHerrero(i)
            .Name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .ObjIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()

        End With

    Next i
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvLingosHerreria(i) = New clsGrapchicalInventory
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
                        j = j + 1
                
                        ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                        HerreroMejorar(j).Name = .Name
                        HerreroMejorar(j).GrhIndex = .GrhIndex
                        HerreroMejorar(j).ObjIndex = .ObjIndex
                        HerreroMejorar(j).UpgradeName = ArmasHerrero(k).Name
                        HerreroMejorar(j).UpgradeGrhIndex = ArmasHerrero(k).GrhIndex
                        HerreroMejorar(j).LinH = ArmasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(j).LinP = ArmasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(j).LinO = ArmasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For

                    End If

                Next k

            End If

        End With

    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    
    For i = 1 To Count

        With ArmadurasHerrero(i)
            .Name = Buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .LinH = Buffer.ReadInteger()        'The iron needed
            .LinP = Buffer.ReadInteger()        'The silver needed
            .LinO = Buffer.ReadInteger()        'The gold needed
            .ObjIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()

        End With

    Next i
    
    j = UBound(HerreroMejorar)
    
    For i = 1 To Count

        With ArmadurasHerrero(i)

            If .Upgrade Then

                For k = 1 To Count

                    If .Upgrade = ArmadurasHerrero(k).ObjIndex Then
                        j = j + 1
                
                        ReDim Preserve HerreroMejorar(j) As tItemsConstruibles
                        
                        HerreroMejorar(j).Name = .Name
                        HerreroMejorar(j).GrhIndex = .GrhIndex
                        HerreroMejorar(j).ObjIndex = .ObjIndex
                        HerreroMejorar(j).UpgradeName = ArmadurasHerrero(k).Name
                        HerreroMejorar(j).UpgradeGrhIndex = ArmadurasHerrero(k).GrhIndex
                        HerreroMejorar(j).LinH = ArmadurasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(j).LinP = ArmadurasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(j).LinO = ArmadurasHerrero(k).LinO - .LinO * 0.85
                        
                        Exit For

                    End If

                Next k

            End If

        End With

    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim Count As Integer
    Dim i     As Long
    Dim j     As Long
    Dim k     As Long
    
    Count = Buffer.ReadInteger()
    
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    
    For i = 1 To Count

        With ObjCarpintero(i)
            .Name = Buffer.ReadASCIIString()        'Get the object's name
            .GrhIndex = Buffer.ReadInteger()
            .Madera = Buffer.ReadInteger()          'The wood needed
            .MaderaElfica = Buffer.ReadInteger()    'The elfic wood needed
            .ObjIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()

        End With

    Next i
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvMaderasCarpinteria(i) = New clsGrapchicalInventory
    Next i
    
    With frmCarp
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
                        j = j + 1
                
                        ReDim Preserve CarpinteroMejorar(j) As tItemsConstruibles
                        
                        CarpinteroMejorar(j).Name = .Name
                        CarpinteroMejorar(j).GrhIndex = .GrhIndex
                        CarpinteroMejorar(j).ObjIndex = .ObjIndex
                        CarpinteroMejorar(j).UpgradeName = ObjCarpintero(k).Name
                        CarpinteroMejorar(j).UpgradeGrhIndex = ObjCarpintero(k).GrhIndex
                        CarpinteroMejorar(j).Madera = ObjCarpintero(k).Madera - .Madera * 0.85
                        CarpinteroMejorar(j).MaderaElfica = ObjCarpintero(k).MaderaElfica - .MaderaElfica * 0.85
                        
                        Exit For

                    End If

                Next k

            End If

        End With

    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRestOK_Err
    
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar

    
    Exit Sub

HandleRestOK_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleRestOK"
    End If
Resume Next
    
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call MsgBox(Buffer.ReadASCIIString())
    
    If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
        #If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
        #Else

            If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
        #End If

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlind_Err
    
    Call incomingData.ReadByte
    
    UserCiego = True

    
    Exit Sub

HandleBlind_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBlind"
    End If
Resume Next
    
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleDumb_Err
    
    Call incomingData.ReadByte
    
    UserEstupido = True

    
    Exit Sub

HandleDumb_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleDumb"
    End If
Resume Next
    
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    
    Call InitCartel(tmp, Buffer.ReadInteger())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 21 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim slot As Byte
    slot = Buffer.ReadByte()
    
    With NPCInventory(slot)
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadInteger()
        .ObjIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger

    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
    
    On Error GoTo HandleUpdateHungerAndThirst_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
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
    
    bWidth = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 75)
    
    frmMain.shpHambre.Width = 75 - bWidth
    frmMain.shpHambre.Left = 584 + (75 - frmMain.shpHambre.Width)
    
    frmMain.shpHambre.Visible = (bWidth <> 75)
    '*********************************
    
    bWidth = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 75)
    
    frmMain.shpSed.Width = 75 - bWidth
    frmMain.shpSed.Left = 584 + (75 - frmMain.shpSed.Width)
    
    frmMain.shpSed.Visible = (bWidth <> 75)
    
    
    Exit Sub

HandleUpdateHungerAndThirst_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleUpdateHungerAndThirst"
    End If
Resume Next
    
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
    
    On Error GoTo HandleFame_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 29 Then
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

    
    Exit Sub

HandleFame_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleFame"
    End If
Resume Next
    
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
    
    On Error GoTo HandleMiniStats_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 20 Then
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

    
    Exit Sub

HandleMiniStats_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleMiniStats"
    End If
Resume Next
    
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
    
    On Error GoTo HandleLevelUp_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    
    Call frmMain.LightSkillStar(True)

    
    Exit Sub

HandleLevelUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleLevelUp"
    End If
Resume Next
    
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim ForumType As eForumMsgType
    Dim Title     As String
    Dim Message   As String
    Dim Author    As String
    Dim bAnuncio  As Boolean
    Dim bSticky   As Boolean
    
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
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowForumForm_Err
    
    Call incomingData.ReadByte
    
    frmForo.Privilegios = incomingData.ReadByte
    frmForo.CanPostSticky = incomingData.ReadByte
    
    If Not MirandoForo Then
        frmForo.Show , frmMain

    End If

    
    Exit Sub

HandleShowForumForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleShowForumForm"
    End If
Resume Next
    
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
    
    On Error GoTo HandleSetInvisible_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    
    CharIndex = incomingData.ReadInteger()
    Call Char_SetInvisible(CharIndex, incomingData.ReadBoolean())

    
    Exit Sub

HandleSetInvisible_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleSetInvisible"
    End If
Resume Next
    
End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleDiceRoll_Err
    

    If incomingData.length < 6 Then
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

    
    Exit Sub

HandleDiceRoll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleDiceRoll"
    End If
Resume Next
    
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleMeditateToggle_Err
    
    Call incomingData.ReadByte
    
    UserMeditar = Not UserMeditar

    
    Exit Sub

HandleMeditateToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleMeditateToggle"
    End If
Resume Next
    
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlindNoMore_Err
    
    Call incomingData.ReadByte
    
    UserCiego = False

    
    Exit Sub

HandleBlindNoMore_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBlindNoMore"
    End If
Resume Next
    
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleDumbNoMore_Err
    
    Call incomingData.ReadByte
    
    UserEstupido = False

    
    Exit Sub

HandleDumbNoMore_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleDumbNoMore"
    End If
Resume Next
    
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
    
    On Error GoTo HandleSendSkills_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Now the server send the percentage of progress of the skills.
    '***************************************************
    If incomingData.length < 2 + NUMSKILLS * 2 Then
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

    
    Exit Sub

HandleSendSkills_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleSendSkills"
    End If
Resume Next
    
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatures() As String
    Dim Upper_creatures As Long
    Dim i           As Long
    
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatures = UBound(creatures())
    
    For i = 0 To Upper_creatures
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i

    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    '11/19/09: Pato - Is optional show the frmGuildNews form
    '***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList()     As String
    Dim Upper_guildList As Long
    Dim i               As Long
    Dim sTemp           As String
    
    'Get news' string
    frmGuildNews.news = Buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    'pre-calculate it's upper-bound beforehand to increase performance
    Upper_guildList = UBound(guildList)
    
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    'pre-calculate it's upper-bound beforehand to increase performance
    Upper_guildList = UBound(guildList)
    
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    If ClientSetup.bGuildNews Or bShowGuildNews Then frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim vsGuildList()     As String
    Dim i                 As Long
    Dim Upper_vsGuildList As Long
    
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
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim guildList()     As String
    Dim Upper_guildList As Long
    Dim i               As Long
    
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    Call frmPeaceProp.lista.Clear
    
    Upper_guildList = UBound(guildList())
    
    For i = 0 To Upper_guildList
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
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
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

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
        Dim caos   As Boolean
        
        armada = Buffer.ReadBoolean()
        caos = Buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Legión Oscura"

        End If
        
        .Ciudadanos.Caption = CStr(Buffer.ReadLong())
        .criminales.Caption = CStr(Buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " Ciudadano"
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " Criminal"
            .status.ForeColor = vbRed

        End If
        
        Call .Show(vbModeless, frmMain)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim i                  As Long
    Dim List()             As String
    
    Dim Upper_guildNames   As Long
    Dim Upper_guildMembers As Long
    Dim Upper_list         As Long
    
    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        Upper_guildNames = UBound(GuildNames())
        
        For i = 0 To Upper_guildNames
            Call .guildslist.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        Upper_guildMembers = UBound(GuildMembers())
        
        For i = 0 To Upper_guildMembers
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtguildnews = Buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        Upper_list = UBound(List())
        
        For i = 0 To Upper_list
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain

    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

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
            .eleccion.Caption = "ABIERTA"
        Else
            .eleccion.Caption = "CERRADA"

        End If
        
        .lblAlineacion.Caption = Buffer.ReadASCIIString()
        .Enemigos.Caption = Buffer.ReadInteger()
        .Aliados.Caption = Buffer.ReadInteger()
        .antifaccion.Caption = Buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i          As Long
        
        codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = Buffer.ReadASCIIString()

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    On Error GoTo HandleShowGuildAlign_Err
    
    Call incomingData.ReadByte
    
    frmEligeAlineacion.Show vbModeless, frmMain

    
    Exit Sub

HandleShowGuildAlign_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleShowGuildAlign"
    End If
Resume Next
    
End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowGuildFundationForm_Err
    
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain

    
    Exit Sub

HandleShowGuildFundationForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleShowGuildFundationForm"
    End If
Resume Next
    
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleParalizeOK_Err
    
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado

    
    Exit Sub

HandleParalizeOK_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleParalizeOK"
    End If
Resume Next
    
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the TradeOK message.

Private Sub HandleTradeOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleTradeOK_Err
    
    Call incomingData.ReadByte
    
    If frmComerciar.Visible Then
        Dim i As Long
        
        'Update user inventory
        For i = 1 To MAX_INVENTORY_SLOTS

            ' Agrego o quito un item en su totalidad
            If Inventario.ObjIndex(i) <> InvComUsu.ObjIndex(i) Then

                With Inventario
                    Call InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))

                End With

                ' Vendio o compro cierta cantidad de un item que ya tenia
            ElseIf Inventario.Amount(i) <> InvComUsu.Amount(i) Then
                Call InvComUsu.ChangeSlotItemAmount(i, Inventario.Amount(i))

            End If

        Next i
        
        ' Fill Npc inventory
        For i = 1 To 20

            ' Compraron la totalidad de un item, o vendieron un item que el npc no tenia
            If NPCInventory(i).ObjIndex <> InvComNpc.ObjIndex(i) Then

                With NPCInventory(i)
                    Call InvComNpc.SetItem(i, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name)

                End With

                ' Compraron o vendieron cierta cantidad (no su totalidad)
            ElseIf NPCInventory(i).Amount <> InvComNpc.Amount(i) Then
                Call InvComNpc.ChangeSlotItemAmount(i, NPCInventory(i).Amount)

            End If

        Next i
    
    End If

    
    Exit Sub

HandleTradeOK_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleTradeOK"
    End If
Resume Next
    
End Sub

''
' Handles the BankOK message.

Private Sub HandleBankOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBankOK_Err
    
    Call incomingData.ReadByte
    
    Dim i As Long
    
    If frmBancoObj.Visible Then
        
        For i = 1 To Inventario.MaxObjs

            With Inventario
                Call InvBanco(1).SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), .Valor(i), .ItemName(i))

            End With

        Next i
        
        'Alter order according to if we bought or sold so the labels and grh remain the same
        If frmBancoObj.LasActionBuy Then
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
        Else

            'frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
            'frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
        End If
        
        frmBancoObj.NoPuedeMover = False

    End If
       
    
    Exit Sub

HandleBankOK_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleBankOK"
    End If
Resume Next
    
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    Dim OfferSlot As Byte
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    OfferSlot = Buffer.ReadByte
    
    With Buffer

        If OfferSlot = GOLD_OFFER_SLOT Then
            Call InvOroComUsu(2).SetItem(1, .ReadInteger(), .ReadLong(), 0, .ReadInteger(), .ReadByte(), .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())
        Else
            Call InvOfferComUsu(1).SetItem(OfferSlot, .ReadInteger(), .ReadLong(), 0, .ReadInteger(), .ReadByte(), .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadInteger(), .ReadLong(), .ReadASCIIString())

        End If

    End With
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & " ha modificado su oferta.", FontTypeNames.FONTTYPE_VENENO)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the SendNight message.

Private Sub HandleSendNight()
    
    On Error GoTo HandleSendNight_Err
    

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    '
    '***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tBool As Boolean 'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?
    tBool = incomingData.ReadBoolean()

    
    Exit Sub

HandleSendNight_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleSendNight"
    End If
Resume Next
    
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim creatureList() As String
    Dim Upper_creatureList As Long
    Dim i              As Long
    
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatureList = UBound(creatureList())
    
    For i = 0 To Upper_creatureList
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i

    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim sosList() As String
    Dim Upper_sosList As Long
    Dim i         As Long
    
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_sosList = UBound(sosList())
    
    For i = 0 To Upper_sosList
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim DenounceList() As String
    Dim Upper_denounceList As Long
    Dim DenounceIndex  As Long
    
    DenounceList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_denounceList = UBound(DenounceList())
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)

        For DenounceIndex = 0 To Upper_denounceList
            Call AddtoRichTextBox(frmMain.RecTxt, DenounceList(DenounceIndex), .Red, .Green, .Blue, .bold, .italic)
        Next DenounceIndex

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim members() As String
    Dim Upper_members As Long
    Dim i         As Long
    
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
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '*************************************Su**************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowGMPanelForm_Err
    
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmMain

    
    Exit Sub

HandleShowGMPanelForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleShowGMPanelForm"
    End If
Resume Next
    
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim userList() As String
    Dim Upper_userlist As Long
    Dim i          As Long
    
    userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        
        Upper_userlist = UBound(userList())
        
        For i = 0 To Upper_userlist
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i

        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandlePong_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, True)
    
    pingTime = 0

    
    Exit Sub

HandlePong_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandlePong"
    End If
Resume Next
    
End Sub

''
' Handles the Pong message.

Private Sub HandleGuildMemberInfo()

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i                  As Long
        Dim Upper_guildNames   As Long
        Dim Upper_guildMembers As Long
        
        Upper_guildNames = UBound(GuildNames())

        For i = 0 To Upper_guildNames
            Call .lstClanes.AddItem(GuildNames(i))
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear
        
        'pre-calculate its upper-bound beforehand to increase performance
        Upper_guildMembers = UBound(GuildMembers())
        
        For i = 0 To Upper_guildMembers
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(Buffer)
        
        .Show vbModeless, frmMain

    End With
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag   As String
    
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

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    On Error GoTo WriteLoginExistingAccount_Err
    
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingAccount)
        
        Call .WriteASCIIString(AccountName)
        
        Call .WriteASCIIString(AccountPassword)
        
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)

    End With

    
    Exit Sub

WriteLoginExistingAccount_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLoginExistingAccount"
    End If
Resume Next
    
End Sub

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
    
    On Error GoTo WriteLoginExistingChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

    
    Exit Sub

WriteLoginExistingChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLoginExistingChar"
    End If
Resume Next
    
End Sub

Public Sub WriteLoginNewAccount()
    
    On Error GoTo WriteLoginNewAccount_Err
    

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

    
    Exit Sub

WriteLoginNewAccount_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLoginNewAccount"
    End If
Resume Next
    
End Sub

''
' Writes the "ThrowDices" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteThrowDices()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ThrowDices" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteThrowDices_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ThrowDices)

    
    Exit Sub

WriteThrowDices_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteThrowDices"
    End If
Resume Next
    
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginNewChar" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteLoginNewChar_Err
    
    Dim i As Long
    
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

    
    Exit Sub

WriteLoginNewChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLoginNewChar"
    End If
Resume Next
    
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
    
    On Error GoTo WriteTalk_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Talk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteTalk_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTalk"
    End If
Resume Next
    
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
    
    On Error GoTo WriteYell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Yell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteYell_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteYell"
    End If
Resume Next
    
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal CharName As String, ByVal chat As String)
    
    On Error GoTo WriteWhisper_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/12/10
    'Writes the "Whisper" message to the outgoing data buffer
    '03/12/10: Enanoh - Ahora se envía el nick y no el charindex.
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteASCIIString(CharName)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteWhisper_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWhisper"
    End If
Resume Next
    
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
    
    On Error GoTo WriteWalk_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Walk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)

    End With

    
    Exit Sub

WriteWalk_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWalk"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestPositionUpdate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestPositionUpdate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)

    
    Exit Sub

WriteRequestPositionUpdate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestPositionUpdate"
    End If
Resume Next
    
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Attack" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteAttack_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Attack)
    'Iniciamos la animacion de ataque
    charlist(UserCharIndex).attacking = True

    
    Exit Sub

WriteAttack_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAttack"
    End If
Resume Next
    
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PickUp" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePickUp_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PickUp)

    
    Exit Sub

WritePickUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePickUp"
    End If
Resume Next
    
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSafeToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)

    
    Exit Sub

WriteSafeToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSafeToggle"
    End If
Resume Next
    
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
    
    On Error GoTo WriteResuscitationToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)

    
    Exit Sub

WriteResuscitationToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteResuscitationToggle"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestGuildLeaderInfo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)

    
    Exit Sub

WriteRequestGuildLeaderInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestGuildLeaderInfo"
    End If
Resume Next
    
End Sub

Public Sub WriteRequestPartyForm()
    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "RequestPartyForm" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestPartyForm_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)

    
    Exit Sub

WriteRequestPartyForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestPartyForm"
    End If
Resume Next
    
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
    
    On Error GoTo WriteItemUpgrade_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)

    
    Exit Sub

WriteItemUpgrade_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteItemUpgrade"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAtributes" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestAtributes_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)

    
    Exit Sub

WriteRequestAtributes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestAtributes"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestFame" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestFame()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFame" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestFame_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)

    
    Exit Sub

WriteRequestFame_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestFame"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestSkills" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestSkills_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)

    
    Exit Sub

WriteRequestSkills_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestSkills"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMiniStats" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestMiniStats_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)

    
    Exit Sub

WriteRequestMiniStats_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestMiniStats"
    End If
Resume Next
    
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCommerceEnd_Err
    
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)

    
    Exit Sub

WriteCommerceEnd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCommerceEnd"
    End If
Resume Next
    
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUserCommerceEnd_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)

    
    Exit Sub

WriteUserCommerceEnd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUserCommerceEnd"
    End If
Resume Next
    
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
    
    On Error GoTo WriteUserCommerceConfirm_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)

    
    Exit Sub

WriteUserCommerceConfirm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUserCommerceConfirm"
    End If
Resume Next
    
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBankEnd_Err
    
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)

    
    Exit Sub

WriteBankEnd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBankEnd"
    End If
Resume Next
    
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
    
    On Error GoTo WriteUserCommerceOk_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)

    
    Exit Sub

WriteUserCommerceOk_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUserCommerceOk"
    End If
Resume Next
    
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceReject" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUserCommerceReject_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)

    
    Exit Sub

WriteUserCommerceReject_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUserCommerceReject"
    End If
Resume Next
    
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteDrop_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Drop" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteDrop_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDrop"
    End If
Resume Next
    
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal slot As Byte)
    
    On Error GoTo WriteCastSpell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(slot)

    End With

    
    Exit Sub

WriteCastSpell_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCastSpell"
    End If
Resume Next
    
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
    
    On Error GoTo WriteLeftClick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)

    End With

    
    Exit Sub

WriteLeftClick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLeftClick"
    End If
Resume Next
    
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
    
    On Error GoTo WriteDoubleClick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoubleClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)

    End With

    
    Exit Sub

WriteDoubleClick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDoubleClick"
    End If
Resume Next
    
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
    
    On Error GoTo WriteWork_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Work" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)

    End With

    
    Exit Sub

WriteWork_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWork"
    End If
Resume Next
    
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseSpellMacro" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUseSpellMacro_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)

    
    Exit Sub

WriteUseSpellMacro_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUseSpellMacro"
    End If
Resume Next
    
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal slot As Byte)
    
    On Error GoTo WriteUseItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        
        Call .WriteByte(slot)

    End With

    
    Exit Sub

WriteUseItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUseItem"
    End If
Resume Next
    
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
    
    On Error GoTo WriteCraftBlacksmith_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftBlacksmith" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)

    End With

    
    Exit Sub

WriteCraftBlacksmith_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCraftBlacksmith"
    End If
Resume Next
    
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)
    
    On Error GoTo WriteCraftCarpenter_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)

    End With

    
    Exit Sub

WriteCraftCarpenter_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCraftCarpenter"
    End If
Resume Next
    
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
    
    On Error GoTo WriteShowGuildNews_Err
    
 
    outgoingData.WriteByte (ClientPacketID.ShowGuildNews)

    
    Exit Sub

WriteShowGuildNews_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteShowGuildNews"
    End If
Resume Next
    
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
    
    On Error GoTo WriteWorkLeftClick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        Call .WriteByte(Skill)

    End With

    
    Exit Sub

WriteWorkLeftClick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWorkLeftClick"
    End If
Resume Next
    
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal Desc As String, _
                               ByVal Name As String, _
                               ByVal Site As String, _
                               ByRef Codex() As String)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNewGuild" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCreateNewGuild_Err
    
    Dim Temp As String
    Dim i    As Long
    Dim Lower_codex As Long
    Dim Upper_codex As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Site)
        
        Lower_codex = LBound(Codex())
        Upper_codex = UBound(Codex())
        
        For i = Lower_codex To Upper_codex
            Temp = Temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(Temp) Then Temp = Left$(Temp, Len(Temp) - 1)
        
        Call .WriteASCIIString(Temp)

    End With

    
    Exit Sub

WriteCreateNewGuild_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCreateNewGuild"
    End If
Resume Next
    
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal slot As Byte)
    
    On Error GoTo WriteEquipItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EquipItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(slot)

    End With

    
    Exit Sub

WriteEquipItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteEquipItem"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
    
    On Error GoTo WriteChangeHeading_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeHeading" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)

    End With

    
    Exit Sub

WriteChangeHeading_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeHeading"
    End If
Resume Next
    
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ModifySkills" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteModifySkills_Err
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i

    End With

    
    Exit Sub

WriteModifySkills_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteModifySkills"
    End If
Resume Next
    
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
    
    On Error GoTo WriteTrain_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Train" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)

    End With

    
    Exit Sub

WriteTrain_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTrain"
    End If
Resume Next
    
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteCommerceBuy_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceBuy" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteCommerceBuy_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCommerceBuy"
    End If
Resume Next
    
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteBankExtractItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteBankExtractItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBankExtractItem"
    End If
Resume Next
    
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteCommerceSell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceSell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteCommerceSell_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCommerceSell"
    End If
Resume Next
    
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteBankDeposit_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDeposit" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteBankDeposit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBankDeposit"
    End If
Resume Next
    
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal Title As String, _
                          ByVal Message As String, _
                          ByVal ForumMsgType As Byte)
    
    On Error GoTo WriteForumPost_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForumPost" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteByte(ForumMsgType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteForumPost_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteForumPost"
    End If
Resume Next
    
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
    
    On Error GoTo WriteMoveSpell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MoveSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)

    End With

    
    Exit Sub

WriteMoveSpell_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMoveSpell"
    End If
Resume Next
    
End Sub

''
' Writes the "MoveBank" message to the outgoing data buffer.
'
' @param    upwards True if the item will be moved up in the list, False if it will be moved downwards.
' @param    slot Bank List slot where the item which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal slot As Byte)
    
    On Error GoTo WriteMoveBank_Err
    

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

    
    Exit Sub

WriteMoveBank_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMoveBank"
    End If
Resume Next
    
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ClanCodexUpdate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteClanCodexUpdate_Err
    
    Dim Temp As String
    Dim i    As Long
    Dim Lower_codex As Long
    Dim Upper_codex As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        
        Call .WriteASCIIString(Desc)
        
        Lower_codex = LBound(Codex())
        Upper_codex = UBound(Codex())
        
        For i = Lower_codex To Upper_codex
            Temp = Temp & Codex(i) & SEPARATOR
        Next i
        
        If Len(Temp) Then Temp = Left$(Temp, Len(Temp) - 1)
        
        Call .WriteASCIIString(Temp)

    End With

    
    Exit Sub

WriteClanCodexUpdate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteClanCodexUpdate"
    End If
Resume Next
    
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, _
                                  ByVal Amount As Long, _
                                  ByVal OfferSlot As Byte)
    
    On Error GoTo WriteUserCommerceOffer_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceOffer" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(slot)
        Call .WriteLong(Amount)
        Call .WriteByte(OfferSlot)

    End With

    
    Exit Sub

WriteUserCommerceOffer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUserCommerceOffer"
    End If
Resume Next
    
End Sub

Public Sub WriteCommerceChat(ByVal chat As String)
    
    On Error GoTo WriteCommerceChat_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    'Writes the "CommerceChat" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteCommerceChat_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCommerceChat"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
    
    On Error GoTo WriteGuildAcceptPeace_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildAcceptPeace_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildAcceptPeace"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
    
    On Error GoTo WriteGuildRejectAlliance_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildRejectAlliance_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildRejectAlliance"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
    
    On Error GoTo WriteGuildRejectPeace_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildRejectPeace_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildRejectPeace"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
    
    On Error GoTo WriteGuildAcceptAlliance_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildAcceptAlliance_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildAcceptAlliance"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
    
    On Error GoTo WriteGuildOfferPeace_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

    
    Exit Sub

WriteGuildOfferPeace_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildOfferPeace"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
    
    On Error GoTo WriteGuildOfferAlliance_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

    
    Exit Sub

WriteGuildOfferAlliance_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildOfferAlliance"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
    
    On Error GoTo WriteGuildAllianceDetails_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAllianceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildAllianceDetails_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildAllianceDetails"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
    
    On Error GoTo WriteGuildPeaceDetails_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeaceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildPeaceDetails_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildPeaceDetails"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
    
    On Error GoTo WriteGuildRequestJoinerInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildRequestJoinerInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildRequestJoinerInfo"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAlliancePropList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildAlliancePropList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)

    
    Exit Sub

WriteGuildAlliancePropList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildAlliancePropList"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeacePropList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildPeacePropList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)

    
    Exit Sub

WriteGuildPeacePropList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildPeacePropList"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
    
    On Error GoTo WriteGuildDeclareWar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDeclareWar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildDeclareWar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildDeclareWar"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal URL As String)
    
    On Error GoTo WriteGuildNewWebsite_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNewWebsite" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(URL)

    End With

    
    Exit Sub

WriteGuildNewWebsite_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildNewWebsite"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
    
    On Error GoTo WriteGuildAcceptNewMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildAcceptNewMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildAcceptNewMember"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal Reason As String)
    
    On Error GoTo WriteGuildRejectNewMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)

    End With

    
    Exit Sub

WriteGuildRejectNewMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildRejectNewMember"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
    
    On Error GoTo WriteGuildKickMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildKickMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildKickMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildKickMember"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
    
    On Error GoTo WriteGuildUpdateNews_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildUpdateNews" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)

    End With

    
    Exit Sub

WriteGuildUpdateNews_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildUpdateNews"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
    
    On Error GoTo WriteGuildMemberInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildMemberInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildMemberInfo"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOpenElections" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildOpenElections_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)

    
    Exit Sub

WriteGuildOpenElections_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildOpenElections"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
    
    On Error GoTo WriteGuildRequestMembership_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestMembership" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)

    End With

    
    Exit Sub

WriteGuildRequestMembership_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildRequestMembership"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
    
    On Error GoTo WriteGuildRequestDetails_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildRequestDetails_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildRequestDetails"
    End If
Resume Next
    
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Online" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnline_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Online)

    
    Exit Sub

WriteOnline_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteOnline"
    End If
Resume Next
    
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/16/08
    'Writes the "Quit" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteQuit_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Quit)

    
    Exit Sub

WriteQuit_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteQuit"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeave" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildLeave_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)

    
    Exit Sub

WriteGuildLeave_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildLeave"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAccountState" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestAccountState_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)

    
    Exit Sub

WriteRequestAccountState_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestAccountState"
    End If
Resume Next
    
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PetStand" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePetStand_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PetStand)

    
    Exit Sub

WritePetStand_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePetStand"
    End If
Resume Next
    
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PetFollow" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePetFollow_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)

    
    Exit Sub

WritePetFollow_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePetFollow"
    End If
Resume Next
    
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
    
    On Error GoTo WriteReleasePet_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ReleasePet)

    
    Exit Sub

WriteReleasePet_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReleasePet"
    End If
Resume Next
    
End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTrainList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.TrainList)

    
    Exit Sub

WriteTrainList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTrainList"
    End If
Resume Next
    
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Rest" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Rest)

    
    Exit Sub

WriteRest_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRest"
    End If
Resume Next
    
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Meditate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteMeditate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Meditate)

    
    Exit Sub

WriteMeditate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMeditate"
    End If
Resume Next
    
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Resucitate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteResucitate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)

    
    Exit Sub

WriteResucitate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteResucitate"
    End If
Resume Next
    
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
    
    On Error GoTo WriteConsultation_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Consultation)

    
    Exit Sub

WriteConsultation_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteConsultation"
    End If
Resume Next
    
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Heal" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteHeal_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Heal)

    
    Exit Sub

WriteHeal_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteHeal"
    End If
Resume Next
    
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Help" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteHelp_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Help)

    
    Exit Sub

WriteHelp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteHelp"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestStats" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestStats_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)

    
    Exit Sub

WriteRequestStats_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestStats"
    End If
Resume Next
    
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceStart" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCommerceStart_Err
    
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)

    
    Exit Sub

WriteCommerceStart_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCommerceStart"
    End If
Resume Next
    
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankStart" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBankStart_Err
    
    Call outgoingData.WriteByte(ClientPacketID.BankStart)

    
    Exit Sub

WriteBankStart_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBankStart"
    End If
Resume Next
    
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Enlist" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteEnlist_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Enlist)

    
    Exit Sub

WriteEnlist_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteEnlist"
    End If
Resume Next
    
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Information" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteInformation_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Information)

    
    Exit Sub

WriteInformation_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteInformation"
    End If
Resume Next
    
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Reward" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReward_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Reward)

    
    Exit Sub

WriteReward_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReward"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMOTD" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestMOTD_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)

    
    Exit Sub

WriteRequestMOTD_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestMOTD"
    End If
Resume Next
    
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpTime" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUpTime_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UpTime)

    
    Exit Sub

WriteUpTime_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUpTime"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyLeave()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyLeave" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePartyLeave_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)

    
    Exit Sub

WritePartyLeave_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyLeave"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyCreate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyCreate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyCreate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePartyCreate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)

    
    Exit Sub

WritePartyCreate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyCreate"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyJoin" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyJoin()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyJoin" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePartyJoin_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)

    
    Exit Sub

WritePartyJoin_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyJoin"
    End If
Resume Next
    
End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Inquiry" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteInquiry_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)

    
    Exit Sub

WriteInquiry_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteInquiry"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
    
    On Error GoTo WriteGuildMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteGuildMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyMessage(ByVal Message As String)
    
    On Error GoTo WritePartyMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WritePartyMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnline" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildOnline_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)

    
    Exit Sub

WriteGuildOnline_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildOnline"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyOnline" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePartyOnline_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)

    
    Exit Sub

WritePartyOnline_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyOnline"
    End If
Resume Next
    
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
    
    On Error GoTo WriteCouncilMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteCouncilMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCouncilMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
    
    On Error GoTo WriteRoleMasterRequest_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoleMasterRequest" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteRoleMasterRequest_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRoleMasterRequest"
    End If
Resume Next
    
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMRequest" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGMRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)

    
    Exit Sub

WriteGMRequest_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGMRequest"
    End If
Resume Next
    
End Sub

''
' Writes the "BugReport" message to the outgoing data buffer.
'
' @param    message The message explaining the reported bug.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBugReport(ByVal Message As String)
    
    On Error GoTo WriteBugReport_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BugReport" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteBugReport_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBugReport"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal Desc As String)
    
    On Error GoTo WriteChangeDescription_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(Desc)

    End With

    
    Exit Sub

WriteChangeDescription_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeDescription"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
    
    On Error GoTo WriteGuildVote_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildVote_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildVote"
    End If
Resume Next
    
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
    
    On Error GoTo WritePunishments_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Punishments" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WritePunishments_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePunishments"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
    
    On Error GoTo WriteChangePassword_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/10/07
    'Last Modified By: Rapsodius
    'Writes the "ChangePassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        Call .WriteASCIIString(oldPass)
        Call .WriteASCIIString(newPass)

    End With

    
    Exit Sub

WriteChangePassword_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangePassword"
    End If
Resume Next
    
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
    
    On Error GoTo WriteGamble_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Gamble" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteGamble_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGamble"
    End If
Resume Next
    
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
    
    On Error GoTo WriteInquiryVote_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "InquiryVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)

    End With

    
    Exit Sub

WriteInquiryVote_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteInquiryVote"
    End If
Resume Next
    
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeaveFaction" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteLeaveFaction_Err
    
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)

    
    Exit Sub

WriteLeaveFaction_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLeaveFaction"
    End If
Resume Next
    
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
    
    On Error GoTo WriteBankExtractGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)

    End With

    
    Exit Sub

WriteBankExtractGold_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBankExtractGold"
    End If
Resume Next
    
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
    
    On Error GoTo WriteBankDepositGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)

    End With

    
    Exit Sub

WriteBankDepositGold_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBankDepositGold"
    End If
Resume Next
    
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(ByVal Message As String)
    
    On Error GoTo WriteDenounce_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteDenounce_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDenounce"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildFundate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 03/21/2001
    'Writes the "GuildFundate" message to the outgoing data buffer
    '14/12/2009: ZaMa - Now first checks if the user can foundate a guild.
    '03/21/2001: Pato - Deleted de clanType param.
    '***************************************************
    
    On Error GoTo WriteGuildFundate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)

    
    Exit Sub

WriteGuildFundate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildFundate"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildFundation" message to the outgoing data buffer.
'
' @param    clanType The alignment of the clan to be founded.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildFundation(ByVal clanType As eClanType)
    
    On Error GoTo WriteGuildFundation_Err
    

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "GuildFundation" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundation)
        
        Call .WriteByte(clanType)

    End With

    
    Exit Sub

WriteGuildFundation_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildFundation"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyKick" message to the outgoing data buffer.
'
' @param    username The user to kick fro mthe party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyKick(ByVal UserName As String)
    
    On Error GoTo WritePartyKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
            
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WritePartyKick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyKick"
    End If
Resume Next
    
End Sub

''
' Writes the "PartySetLeader" message to the outgoing data buffer.
'
' @param    username The user to set as the party's leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySetLeader(ByVal UserName As String)
    
    On Error GoTo WritePartySetLeader_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartySetLeader" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WritePartySetLeader_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartySetLeader"
    End If
Resume Next
    
End Sub

''
' Writes the "PartyAcceptMember" message to the outgoing data buffer.
'
' @param    username The user to accept into the party.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartyAcceptMember(ByVal UserName As String)
    
    On Error GoTo WritePartyAcceptMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PartyAcceptMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WritePartyAcceptMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePartyAcceptMember"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
    
    On Error GoTo WriteGuildMemberList_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberList" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildMemberList)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildMemberList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildMemberList"
    End If
Resume Next
    
End Sub

''
' Writes the "InitCrafting" message to the outgoing data buffer.
'
' @param    Cantidad The final aumont of item to craft.
' @param    NroPorCiclo The amount of items to craft per cicle.

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)
    
    On Error GoTo WriteInitCrafting_Err
    

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

    
    Exit Sub

WriteInitCrafting_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteInitCrafting"
    End If
Resume Next
    
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
    
    On Error GoTo WriteHome_Err
    

    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/10
    'Writes the "Home" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Home)

    End With

    
    Exit Sub

WriteHome_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteHome"
    End If
Resume Next
    
End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMMessage(ByVal Message As String)
    
    On Error GoTo WriteGMMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GMMessage)
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteGMMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGMMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowName" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteShowName_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.showName)

    
    Exit Sub

WriteShowName_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteShowName"
    End If
Resume Next
    
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineRoyalArmy_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)

    
    Exit Sub

WriteOnlineRoyalArmy_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteOnlineRoyalArmy"
    End If
Resume Next
    
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineChaosLegion" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineChaosLegion_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)

    
    Exit Sub

WriteOnlineChaosLegion_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteOnlineChaosLegion"
    End If
Resume Next
    
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
    
    On Error GoTo WriteGoNearby_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoNearby" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoNearby)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGoNearby_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGoNearby"
    End If
Resume Next
    
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
    
    On Error GoTo WriteComment_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Comment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Comment)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteComment_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteComment"
    End If
Resume Next
    
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerTime" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteServerTime_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.serverTime)

    
    Exit Sub

WriteServerTime_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteServerTime"
    End If
Resume Next
    
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
    
    On Error GoTo WriteWhere_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Where" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Where)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteWhere_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWhere"
    End If
Resume Next
    
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
    
    On Error GoTo WriteCreaturesInMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreaturesInMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreaturesInMap)
        
        Call .WriteInteger(Map)

    End With

    
    Exit Sub

WriteCreaturesInMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCreaturesInMap"
    End If
Resume Next
    
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpMeToTarget" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteWarpMeToTarget_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)

    
    Exit Sub

WriteWarpMeToTarget_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWarpMeToTarget"
    End If
Resume Next
    
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, _
                         ByVal Map As Integer, _
                         ByVal X As Byte, _
                         ByVal Y As Byte)
    
    On Error GoTo WriteWarpChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(Map)
        
        Call .WriteByte(X)
        Call .WriteByte(Y)

    End With

    
    Exit Sub

WriteWarpChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWarpChar"
    End If
Resume Next
    
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
    
    On Error GoTo WriteSilence_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Silence" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Silence)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteSilence_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSilence"
    End If
Resume Next
    
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSShowList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSOSShowList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)

    
    Exit Sub

WriteSOSShowList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSOSShowList"
    End If
Resume Next
    
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
    
    On Error GoTo WriteSOSRemove_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSRemove" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteSOSRemove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSOSRemove"
    End If
Resume Next
    
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
    
    On Error GoTo WriteGoToChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoToChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGoToChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGoToChar"
    End If
Resume Next
    
End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "invisible" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteInvisible_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.invisible)

    
    Exit Sub

WriteInvisible_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteInvisible"
    End If
Resume Next
    
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMPanel" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGMPanel_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)

    
    Exit Sub

WriteGMPanel_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGMPanel"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestUserList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestUserList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)

    
    Exit Sub

WriteRequestUserList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestUserList"
    End If
Resume Next
    
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Working" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteWorking_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)

    
    Exit Sub

WriteWorking_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWorking"
    End If
Resume Next
    
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Hiding" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteHiding_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)

    
    Exit Sub

WriteHiding_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteHiding"
    End If
Resume Next
    
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal Reason As String, ByVal Time As Byte)
    
    On Error GoTo WriteJail_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

    
    Exit Sub

WriteJail_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteJail"
    End If
Resume Next
    
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPC" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKillNPC_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)

    
    Exit Sub

WriteKillNPC_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteKillNPC"
    End If
Resume Next
    
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal Reason As String)
    
    On Error GoTo WriteWarnUser_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarnUser" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)

    End With

    
    Exit Sub

WriteWarnUser_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteWarnUser"
    End If
Resume Next
    
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, _
                         ByVal EditOption As eEditOptions, _
                         ByVal arg1 As String, _
                         ByVal arg2 As String)
    
    On Error GoTo WriteEditChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

    
    Exit Sub

WriteEditChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteEditChar"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharInfo"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharStats_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharStats" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharStats)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharStats_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharStats"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharGold)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharGold_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharGold"
    End If
Resume Next
    
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharInventory_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInventory" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharInventory_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharInventory"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharBank_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharBank_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharBank"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharSkills_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharSkills" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharSkills_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharSkills"
    End If
Resume Next
    
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
    
    On Error GoTo WriteReviveChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReviveChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteReviveChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReviveChar"
    End If
Resume Next
    
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineGM" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineGM_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)

    
    Exit Sub

WriteOnlineGM_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteOnlineGM"
    End If
Resume Next
    
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap(ByVal Map As Integer)
    
    On Error GoTo WriteOnlineMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    'Writes the "OnlineMap" message to the outgoing data buffer
    '26/03/2009: Now you don't need to be in the map to use the comand, so you send the map to server
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        
        Call .WriteInteger(Map)

    End With

    
    Exit Sub

WriteOnlineMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteOnlineMap"
    End If
Resume Next
    
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive(ByVal UserName As String)
    
    On Error GoTo WriteForgive_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Forgive" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Forgive)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteForgive_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteForgive"
    End If
Resume Next
    
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
    
    On Error GoTo WriteKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Kick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteKick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteKick"
    End If
Resume Next
    
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
    
    On Error GoTo WriteExecute_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Execute" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteExecute_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteExecute"
    End If
Resume Next
    
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal Reason As String)
    
    On Error GoTo WriteBanChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.banChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(Reason)

    End With

    
    Exit Sub

WriteBanChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBanChar"
    End If
Resume Next
    
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
    
    On Error GoTo WriteUnbanChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteUnbanChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUnbanChar"
    End If
Resume Next
    
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCFollow" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteNPCFollow_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)

    
    Exit Sub

WriteNPCFollow_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteNPCFollow"
    End If
Resume Next
    
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
    
    On Error GoTo WriteSummonChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SummonChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteSummonChar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSummonChar"
    End If
Resume Next
    
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnListRequest" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSpawnListRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)

    
    Exit Sub

WriteSpawnListRequest_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSpawnListRequest"
    End If
Resume Next
    
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
    
    On Error GoTo WriteSpawnCreature_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnCreature" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)

    End With

    
    Exit Sub

WriteSpawnCreature_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSpawnCreature"
    End If
Resume Next
    
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetNPCInventory" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteResetNPCInventory_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)

    
    Exit Sub

WriteResetNPCInventory_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteResetNPCInventory"
    End If
Resume Next
    
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
    
    On Error GoTo WriteServerMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteServerMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteServerMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMapMessage(ByVal Message As String)
    
    On Error GoTo WriteMapMessage_Err
    

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

    
    Exit Sub

WriteMapMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMapMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
    
    On Error GoTo WriteNickToIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NickToIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.nickToIP)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteNickToIP_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteNickToIP"
    End If
Resume Next
    
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef Ip() As Byte)
    
    On Error GoTo WriteIPToNick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "IPToNick" message to the outgoing data buffer
    '***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    Dim Upper_ip As Long
    Dim Lower_ip As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        
        Lower_ip = LBound(Ip())
        Upper_ip = UBound(Ip())
        
        For i = Lower_ip To Upper_ip
            Call .WriteByte(Ip(i))
        Next i

    End With

    
    Exit Sub

WriteIPToNick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteIPToNick"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
    
    On Error GoTo WriteGuildOnlineMembers_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnlineMembers" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildOnlineMembers_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildOnlineMembers"
    End If
Resume Next
    
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal Map As Integer, _
                               ByVal X As Byte, _
                               ByVal Y As Byte, _
                               Optional ByVal Radio As Byte = 0)
    
    On Error GoTo WriteTeleportCreate_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

    
    Exit Sub

WriteTeleportCreate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTeleportCreate"
    End If
Resume Next
    
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportDestroy" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTeleportDestroy_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)

    
    Exit Sub

WriteTeleportDestroy_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTeleportDestroy"
    End If
Resume Next
    
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
    
    On Error GoTo WriteExitDestroy_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ExitDestroy)

    
    Exit Sub

WriteExitDestroy_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteExitDestroy"
    End If
Resume Next
    
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRainToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)

    
    Exit Sub

WriteRainToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRainToggle"
    End If
Resume Next
    
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal Desc As String)
    
    On Error GoTo WriteSetCharDescription_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetCharDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetCharDescription)
        
        Call .WriteASCIIString(Desc)

    End With

    
    Exit Sub

WriteSetCharDescription_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSetCharDescription"
    End If
Resume Next
    
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
    
    On Error GoTo WriteForceMIDIToMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIToMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(Map)

    End With

    
    Exit Sub

WriteForceMIDIToMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteForceMIDIToMap"
    End If
Resume Next
    
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, _
                               ByVal Map As Integer, _
                               ByVal X As Byte, _
                               ByVal Y As Byte)
    
    On Error GoTo WriteForceWAVEToMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

    
    Exit Sub

WriteForceWAVEToMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteForceWAVEToMap"
    End If
Resume Next
    
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
    
    On Error GoTo WriteRoyalArmyMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteRoyalArmyMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRoyalArmyMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
    
    On Error GoTo WriteChaosLegionMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteChaosLegionMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChaosLegionMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
    
    On Error GoTo WriteCitizenMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CitizenMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CitizenMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteCitizenMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCitizenMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
    
    On Error GoTo WriteCriminalMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CriminalMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CriminalMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteCriminalMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCriminalMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
    
    On Error GoTo WriteTalkAsNPC_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TalkAsNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteTalkAsNPC_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTalkAsNPC"
    End If
Resume Next
    
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDestroyAllItemsInArea_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)

    
    Exit Sub

WriteDestroyAllItemsInArea_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDestroyAllItemsInArea"
    End If
Resume Next
    
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
    
    On Error GoTo WriteAcceptRoyalCouncilMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteAcceptRoyalCouncilMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAcceptRoyalCouncilMember"
    End If
Resume Next
    
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
    
    On Error GoTo WriteAcceptChaosCouncilMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteAcceptChaosCouncilMember_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAcceptChaosCouncilMember"
    End If
Resume Next
    
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ItemsInTheFloor" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteItemsInTheFloor_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)

    
    Exit Sub

WriteItemsInTheFloor_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteItemsInTheFloor"
    End If
Resume Next
    
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
    
    On Error GoTo WriteMakeDumb_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumb" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumb)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteMakeDumb_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMakeDumb"
    End If
Resume Next
    
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
    
    On Error GoTo WriteMakeDumbNoMore_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumbNoMore" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteMakeDumbNoMore_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMakeDumbNoMore"
    End If
Resume Next
    
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumpIPTables" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDumpIPTables_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.dumpIPTables)

    
    Exit Sub

WriteDumpIPTables_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDumpIPTables"
    End If
Resume Next
    
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
    
    On Error GoTo WriteCouncilKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CouncilKick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteCouncilKick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCouncilKick"
    End If
Resume Next
    
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
    
    On Error GoTo WriteSetTrigger_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetTrigger" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        
        Call .WriteByte(Trigger)

    End With

    
    Exit Sub

WriteSetTrigger_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSetTrigger"
    End If
Resume Next
    
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
    
    On Error GoTo WriteAskTrigger_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)

    
    Exit Sub

WriteAskTrigger_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAskTrigger"
    End If
Resume Next
    
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBannedIPList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)

    
    Exit Sub

WriteBannedIPList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBannedIPList"
    End If
Resume Next
    
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPReload" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBannedIPReload_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)

    
    Exit Sub

WriteBannedIPReload_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBannedIPReload"
    End If
Resume Next
    
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
    
    On Error GoTo WriteGuildBan_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildBan" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildBan_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteGuildBan"
    End If
Resume Next
    
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

Public Sub WriteBanIP(ByVal byIp As Boolean, _
                      ByRef Ip() As Byte, _
                      ByVal Nick As String, _
                      ByVal Reason As String)
    
    On Error GoTo WriteBanIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanIP" message to the outgoing data buffer
    '***************************************************
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    Dim Lower_ip As Long
    Dim Upper_ip As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then
            
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

    
    Exit Sub

WriteBanIP_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteBanIP"
    End If
Resume Next
    
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
    
    On Error GoTo WriteUnbanIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanIP" message to the outgoing data buffer
    '***************************************************
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    Dim Upper_ip As Long
    Dim Lower_ip As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        
        Lower_ip = LBound(Ip())
        Upper_ip = UBound(Ip())
        
        For i = Lower_ip To Upper_ip
            Call .WriteByte(Ip(i))
        Next i

    End With

    
    Exit Sub

WriteUnbanIP_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteUnbanIP"
    End If
Resume Next
    
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long)
    
    On Error GoTo WriteCreateItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)

    End With

    
    Exit Sub

WriteCreateItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCreateItem"
    End If
Resume Next
    
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyItems" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDestroyItems_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)

    
    Exit Sub

WriteDestroyItems_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDestroyItems"
    End If
Resume Next
    
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
    
    On Error GoTo WriteChaosLegionKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteChaosLegionKick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChaosLegionKick"
    End If
Resume Next
    
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
    
    On Error GoTo WriteRoyalArmyKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRoyalArmyKick_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRoyalArmyKick"
    End If
Resume Next
    
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
    
    On Error GoTo WriteForceMIDIAll_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIAll)
        
        Call .WriteByte(midiID)

    End With

    
    Exit Sub

WriteForceMIDIAll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteForceMIDIAll"
    End If
Resume Next
    
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
    
    On Error GoTo WriteForceWAVEAll_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEAll)
        
        Call .WriteByte(waveID)

    End With

    
    Exit Sub

WriteForceWAVEAll_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteForceWAVEAll"
    End If
Resume Next
    
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, _
                                 ByVal punishment As Byte, _
                                 ByVal NewText As String)
    
    On Error GoTo WriteRemovePunishment_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
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

    
    Exit Sub

WriteRemovePunishment_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRemovePunishment"
    End If
Resume Next
    
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TileBlockedToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTileBlockedToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)

    
    Exit Sub

WriteTileBlockedToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTileBlockedToggle"
    End If
Resume Next
    
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKillNPCNoRespawn_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)

    
    Exit Sub

WriteKillNPCNoRespawn_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteKillNPCNoRespawn"
    End If
Resume Next
    
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKillAllNearbyNPCs_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)

    
    Exit Sub

WriteKillAllNearbyNPCs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteKillAllNearbyNPCs"
    End If
Resume Next
    
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
    
    On Error GoTo WriteLastIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LastIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteLastIP_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteLastIP"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMOTD" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteChangeMOTD_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)

    
    Exit Sub

WriteChangeMOTD_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMOTD"
    End If
Resume Next
    
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
    
    On Error GoTo WriteSetMOTD_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetMOTD" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetMOTD)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteSetMOTD_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSetMOTD"
    End If
Resume Next
    
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
    
    On Error GoTo WriteSystemMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SystemMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteSystemMessage_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSystemMessage"
    End If
Resume Next
    
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer)
    
    On Error GoTo WriteCreateNPC_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        
        Call .WriteInteger(NPCIndex)

    End With

    
    Exit Sub

WriteCreateNPC_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCreateNPC"
    End If
Resume Next
    
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NPCIndex As Integer)
    
    On Error GoTo WriteCreateNPCWithRespawn_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPCWithRespawn)
        
        Call .WriteInteger(NPCIndex)

    End With

    
    Exit Sub

WriteCreateNPCWithRespawn_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCreateNPCWithRespawn"
    End If
Resume Next
    
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    
    On Error GoTo WriteImperialArmour_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ImperialArmour" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)

    End With

    
    Exit Sub

WriteImperialArmour_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteImperialArmour"
    End If
Resume Next
    
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    
    On Error GoTo WriteChaosArmour_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosArmour" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)

    End With

    
    Exit Sub

WriteChaosArmour_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChaosArmour"
    End If
Resume Next
    
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteNavigateToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)

    
    Exit Sub

WriteNavigateToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteNavigateToggle"
    End If
Resume Next
    
End Sub

''
' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteServerOpenToUsersToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)

    
    Exit Sub

WriteServerOpenToUsersToggle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteServerOpenToUsersToggle"
    End If
Resume Next
    
End Sub

''
' Writes the "TurnOffServer" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnOffServer()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnOffServer" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTurnOffServer_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)

    
    Exit Sub

WriteTurnOffServer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTurnOffServer"
    End If
Resume Next
    
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
    
    On Error GoTo WriteTurnCriminal_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnCriminal" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TurnCriminal)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteTurnCriminal_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteTurnCriminal"
    End If
Resume Next
    
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
    
    On Error GoTo WriteResetFactions_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetFactions" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResetFactions)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteResetFactions_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteResetFactions"
    End If
Resume Next
    
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
    
    On Error GoTo WriteRemoveCharFromGuild_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRemoveCharFromGuild_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRemoveCharFromGuild"
    End If
Resume Next
    
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharMail_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharMail)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharMail_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRequestCharMail"
    End If
Resume Next
    
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
    
    On Error GoTo WriteAlterPassword_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterPassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)

    End With

    
    Exit Sub

WriteAlterPassword_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAlterPassword"
    End If
Resume Next
    
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
    
    On Error GoTo WriteAlterMail_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)

    End With

    
    Exit Sub

WriteAlterMail_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAlterMail"
    End If
Resume Next
    
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
    
    On Error GoTo WriteAlterName_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterName" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)

    End With

    
    Exit Sub

WriteAlterName_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteAlterName"
    End If
Resume Next
    
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoBackup" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDoBackup_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DoBackUp)

    
    Exit Sub

WriteDoBackup_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDoBackup"
    End If
Resume Next
    
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
    
    On Error GoTo WriteShowGuildMessages_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildMessages" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteShowGuildMessages_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteShowGuildMessages"
    End If
Resume Next
    
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveMap" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSaveMap_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)

    
    Exit Sub

WriteSaveMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSaveMap"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
    
    On Error GoTo WriteChangeMapInfoPK_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)

    End With

    
    Exit Sub

WriteChangeMapInfoPK_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoPK"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoNoOcultar" message to the outgoing data buffer.
'
' @param    PermitirOcultar True if the map permits to hide, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoOcultar_Err
    

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

    
    Exit Sub

WriteChangeMapInfoNoOcultar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoNoOcultar"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoNoInvocar" message to the outgoing data buffer.
'
' @param    PermitirInvocar True if the map permits to invoke, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoInvocar_Err
    

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

    
    Exit Sub

WriteChangeMapInfoNoInvocar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoNoInvocar"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
    
    On Error GoTo WriteChangeMapInfoBackup_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)

    End With

    
    Exit Sub

WriteChangeMapInfoBackup_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoBackup"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
    
    On Error GoTo WriteChangeMapInfoRestricted_Err
    

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

    
    Exit Sub

WriteChangeMapInfoRestricted_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoRestricted"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoMagic_Err
    

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

    
    Exit Sub

WriteChangeMapInfoNoMagic_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoNoMagic"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoInvi_Err
    

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

    
    Exit Sub

WriteChangeMapInfoNoInvi_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoNoInvi"
    End If
Resume Next
    
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoResu_Err
    

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

    
    Exit Sub

WriteChangeMapInfoNoResu_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoNoResu"
    End If
Resume Next
    
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)
    
    On Error GoTo WriteChangeMapInfoLand_Err
    

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

    
    Exit Sub

WriteChangeMapInfoLand_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoLand"
    End If
Resume Next
    
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
    
    On Error GoTo WriteChangeMapInfoZone_Err
    

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

    
    Exit Sub

WriteChangeMapInfoZone_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoZone"
    End If
Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoStealNpc" message to the outgoing data buffer.
'
' @param    forbid TRUE if stealNpc forbiden.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)
    
    On Error GoTo WriteChangeMapInfoStealNpc_Err
    

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

    
    Exit Sub

WriteChangeMapInfoStealNpc_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChangeMapInfoStealNpc"
    End If
Resume Next
    
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveChars" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSaveChars_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)

    
    Exit Sub

WriteSaveChars_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSaveChars"
    End If
Resume Next
    
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanSOS" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCleanSOS_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)

    
    Exit Sub

WriteCleanSOS_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCleanSOS"
    End If
Resume Next
    
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowServerForm" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteShowServerForm_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowServerForm)

    
    Exit Sub

WriteShowServerForm_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteShowServerForm"
    End If
Resume Next
    
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
    
    On Error GoTo WriteShowDenouncesList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowDenouncesList)

    
    Exit Sub

WriteShowDenouncesList_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteShowDenouncesList"
    End If
Resume Next
    
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
    
    On Error GoTo WriteEnableDenounces_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.EnableDenounces)

    
    Exit Sub

WriteEnableDenounces_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteEnableDenounces"
    End If
Resume Next
    
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Night" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteNight_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.night)

    
    Exit Sub

WriteNight_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteNight"
    End If
Resume Next
    
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KickAllChars" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKickAllChars_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)

    
    Exit Sub

WriteKickAllChars_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteKickAllChars"
    End If
Resume Next
    
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadNPCs" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadNPCs_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)

    
    Exit Sub

WriteReloadNPCs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReloadNPCs"
    End If
Resume Next
    
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadServerIni" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadServerIni_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)

    
    Exit Sub

WriteReloadServerIni_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReloadServerIni"
    End If
Resume Next
    
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadSpells" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadSpells_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)

    
    Exit Sub

WriteReloadSpells_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReloadSpells"
    End If
Resume Next
    
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadObjects" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadObjects_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)

    
    Exit Sub

WriteReloadObjects_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteReloadObjects"
    End If
Resume Next
    
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Restart" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRestart_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Restart)

    
    Exit Sub

WriteRestart_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRestart"
    End If
Resume Next
    
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetAutoUpdate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteResetAutoUpdate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)

    
    Exit Sub

WriteResetAutoUpdate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteResetAutoUpdate"
    End If
Resume Next
    
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
    
    On Error GoTo WriteChatColor_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatColor" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(b)

    End With

    
    Exit Sub

WriteChatColor_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteChatColor"
    End If
Resume Next
    
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Ignored" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteIgnored_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Ignored)

    
    Exit Sub

WriteIgnored_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteIgnored"
    End If
Resume Next
    
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
    
    On Error GoTo WriteCheckSlot_Err
    

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

    
    Exit Sub

WriteCheckSlot_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCheckSlot"
    End If
Resume Next
    
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
    
    On Error GoTo WritePing_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    'Prevent the timer from being cut
    If pingTime <> 0 Then Exit Sub
    
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    DoEvents
    
    pingTime = GetTickCount

    
    Exit Sub

WritePing_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WritePing"
    End If
Resume Next
    
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
    
    On Error GoTo WriteShareNpc_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ShareNpc)

    
    Exit Sub

WriteShareNpc_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteShareNpc"
    End If
Resume Next
    
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
    
    On Error GoTo WriteStopSharingNpc_Err
    
    Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)

    
    Exit Sub

WriteStopSharingNpc_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteStopSharingNpc"
    End If
Resume Next
    
End Sub

''
' Writes the "SetIniVar" message to the outgoing data buffer.
'
' @param    sLlave the name of the key which contains the value to edit
' @param    sClave the name of the value to edit
' @param    sValor the new value to set to sClave
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetIniVar(ByRef sLlave As String, _
                          ByRef sClave As String, _
                          ByRef sValor As String)
    
    On Error GoTo WriteSetIniVar_Err
    

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

    
    Exit Sub

WriteSetIniVar_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSetIniVar"
    End If
Resume Next
    
End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, _
                                    ByVal X As Byte, _
                                    ByVal Y As Byte)
    
    On Error GoTo WriteCreatePretorianClan_Err
    

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

    
    Exit Sub

WriteCreatePretorianClan_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteCreatePretorianClan"
    End If
Resume Next
    
End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal Map As Integer)
    
    On Error GoTo WriteDeletePretorianClan_Err
    

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

    
    Exit Sub

WriteDeletePretorianClan_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteDeletePretorianClan"
    End If
Resume Next
    
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************
    
    On Error GoTo FlushBuffer_Err
    
    Dim sndData As String
    
    With outgoingData

        If .length = 0 Then Exit Sub
        
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)

    End With

    
    Exit Sub

FlushBuffer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "FlushBuffer"
    End If
Resume Next
    
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    On Error GoTo SendData_Err
    
    
    'No enviamos nada si no estamos conectados
    #If UsarWrench = 1 Then

        If Not frmMain.Socket1.IsWritable Then
            'Put data back in the bytequeue
            Call outgoingData.WriteASCIIStringFixed(sdData)
        
            Exit Sub

        End If
    
        If Not frmMain.Socket1.Connected Then Exit Sub
    #Else

        If frmMain.Winsock1.State <> sckConnected Then Exit Sub
    #End If
    
    'Send data!
    #If UsarWrench = 1 Then
        Call frmMain.Socket1.Write(sdData, Len(sdData))
    #Else
        Call frmMain.Winsock1.SendData(sdData)
    #End If

    
    Exit Sub

SendData_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "SendData"
    End If
Resume Next
    
End Sub

''
' Writes the "MapMessage" message to the outgoing data buffer.
'
' @param    Dialog The new dialog of the NPC.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetDialog(ByVal dialog As String)
    
    On Error GoTo WriteSetDialog_Err
    

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

    
    Exit Sub

WriteSetDialog_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteSetDialog"
    End If
Resume Next
    
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
    
    On Error GoTo WriteImpersonate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Impersonate)

    
    Exit Sub

WriteImpersonate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteImpersonate"
    End If
Resume Next
    
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
    
    On Error GoTo WriteImitate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Imitate)

    
    Exit Sub

WriteImitate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteImitate"
    End If
Resume Next
    
End Sub

''
' Writes the "RecordAddObs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String)
    
    On Error GoTo WriteRecordAddObs_Err
    

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

    
    Exit Sub

WriteRecordAddObs_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRecordAddObs"
    End If
Resume Next
    
End Sub

''
' Writes the "RecordAdd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal Reason As String)
    
    On Error GoTo WriteRecordAdd_Err
    

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

    
    Exit Sub

WriteRecordAdd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRecordAdd"
    End If
Resume Next
    
End Sub

''
' Writes the "RecordRemove" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte)
    
    On Error GoTo WriteRecordRemove_Err
    

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

    
    Exit Sub

WriteRecordRemove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRecordRemove"
    End If
Resume Next
    
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
    
    On Error GoTo WriteRecordListRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RecordListRequest)

    
    Exit Sub

WriteRecordListRequest_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRecordListRequest"
    End If
Resume Next
    
End Sub

''
' Writes the "RecordDetailsRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte)
    
    On Error GoTo WriteRecordDetailsRequest_Err
    

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

    
    Exit Sub

WriteRecordDetailsRequest_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteRecordDetailsRequest"
    End If
Resume Next
    
End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList()

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte
    
    Dim NumRecords As Byte
    Dim i          As Long
    
    NumRecords = Buffer.ReadByte
    
    'Se limpia el ListBox y se agregan los usuarios
    frmPanelGm.lstUsers.Clear

    For i = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem Buffer.ReadASCIIString
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails()

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

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
            .lblEstado.Caption = "ONLINE"
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = "OFFLINE"

        End If
        
        'IP del personaje
        tmpStr = Buffer.ReadASCIIString

        If LenB(tmpStr) Then
            .txtIP.Text = tmpStr
        Else
            .txtIP.Text = "Usuario offline"

        End If
        
        'Tiempo online
        tmpStr = Buffer.ReadASCIIString

        If LenB(tmpStr) Then
            .txtTimeOn.Text = tmpStr
        Else
            .txtTimeOn.Text = "Usuario offline"

        End If
        
        'Observaciones
        tmpStr = Buffer.ReadASCIIString

        If LenB(tmpStr) Then
            .txtObs.Text = tmpStr
        Else
            .txtObs.Text = "Sin observaciones"

        End If

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(Buffer)
    
ErrHandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Writes the "Moveitem" message to the outgoing data buffer.
'
Public Sub WriteMoveItem(ByVal originalSlot As Integer, _
                         ByVal newSlot As Integer, _
                         ByVal moveType As eMoveType)
    
    On Error GoTo WriteMoveItem_Err
    

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

    
    Exit Sub

WriteMoveItem_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "WriteMoveItem"
    End If
Resume Next
    
End Sub

Private Sub HandleDecirPalabrasMagicas()
    
    On Error GoTo HandleDecirPalabrasMagicas_Err
    

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Spell     As Integer
    Dim CharIndex As Integer
    
    Spell = incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If Char_Check(CharIndex) Then Call Dialogos.CreateDialog(Hechizos(Spell).PalabrasMagicas, CharIndex, RGB(200, 250, 150))

    
    Exit Sub

HandleDecirPalabrasMagicas_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleDecirPalabrasMagicas"
    End If
Resume Next
    
End Sub

Private Sub HandleAttackAnim()
    
    On Error GoTo HandleAttackAnim_Err
    

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim CharIndex As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    'Set the animation trigger on true
    charlist(CharIndex).attacking = True 'should be done in separated sub?

    
    Exit Sub

HandleAttackAnim_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleAttackAnim"
    End If
Resume Next
    
End Sub

Private Sub HandleFXtoMap()
    
    On Error GoTo HandleFXtoMap_Err
    

    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim X, Y, FxIndex, Loops As Integer
    
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

    
    Exit Sub

HandleFXtoMap_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleFXtoMap"
    End If
Resume Next
    
End Sub

Private Sub HandleAccountLogged()
    
    On Error GoTo HandleAccountLogged_Err
    

    If incomingData.length < 32 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim i As Byte

    'Remove packet ID
    Call incomingData.ReadByte
    AccountName = incomingData.ReadASCIIString
    AccountHash = incomingData.ReadASCIIString
    NumberOfCharacters = incomingData.ReadByte

    ReDim cPJ(1 To 10) As PjCuenta
    
    If NumberOfCharacters > 0 Then

        For i = 1 To NumberOfCharacters
            cPJ(i).Nombre = incomingData.ReadASCIIString
            cPJ(i).Body = incomingData.ReadInteger
            cPJ(i).Head = incomingData.ReadInteger
            cPJ(i).weapon = incomingData.ReadInteger
            cPJ(i).shield = incomingData.ReadInteger
            cPJ(i).helmet = incomingData.ReadInteger
            cPJ(i).Class = incomingData.ReadByte
            cPJ(i).Race = incomingData.ReadByte
            cPJ(i).Map = incomingData.ReadInteger
            cPJ(i).Level = incomingData.ReadByte
            cPJ(i).Gold = incomingData.ReadLong
            cPJ(i).Criminal = incomingData.ReadBoolean
            cPJ(i).Dead = incomingData.ReadBoolean
            cPJ(i).GameMaster = incomingData.ReadBoolean
        Next i

    End If
    
    frmPanelAccount.Show

    
    Exit Sub

HandleAccountLogged_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "Protocol" & "->" & "HandleAccountLogged"
    End If
Resume Next
    
End Sub
