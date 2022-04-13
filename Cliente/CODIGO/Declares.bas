Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public particle_group_index_render(0 To 1) As Long

Public INTRO As Boolean
Public EffectIntro As Byte
Public TypeIntroEffect As Byte
Public BastaINTRO As Byte

Public View_FPS_OR_MS As Boolean

Public FormParser As clsCursor
Public LongBlack(3) As Long
Public LongWhite(3) As Long
Public LongYellow(3) As Long
Public LongBlue(3) As Long
Public LongGreen(3) As Long

Public Es_Real As Boolean

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
 
'RGB Type
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type
 
Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
   
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Public PixelOffsetX As Integer
Public PixelOffsetY As Integer

'Objetos públicos
Public DialogosClanes As New clsGuildDlg
Public Audio As New clsAudio
Public Inventario As New clsGrapchicalInventory
Public InvBanco(1) As New clsGrapchicalInventory

'Inventarios de comercio con usuario
Public InvComUsu As New clsGrapchicalInventory ' Inventario del usuario visible en el comercio
Public InvOroComUsu(2) As New clsGrapchicalInventory ' Inventarios de oro (ambos usuarios)
Public InvOfferComUsu(1) As New clsGrapchicalInventory ' Inventarios de ofertas (ambos usuarios)
Public InvComNpc As New clsGrapchicalInventory ' Inventario con los items que ofrece el npc

'Inventarios de herreria
Public Const MAX_LIST_ITEMS As Byte = 4
Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventory
Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As New clsGrapchicalInventory
                
Public CustomKeys As New clsCustomKeys
Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue

''
'The main timer of the game.
Public MainTimer As New clsTimer

'Sonidos
Public Const SND_CLICK As String = "187.wav"
Public Const SND_PASOS1 As String = "23.wav"
Public Const SND_PASOS2 As String = "24.wav"
Public Const SND_PASOS3 As String = "201.wav"
Public Const SND_PASOS4 As String = "202.wav"
Public Const SND_PASOS5 As String = "197.wav"
Public Const SND_PASOS6 As String = "198.wav"
Public Const SND_PASOS7 As String = "199.wav"
Public Const SND_PASOS8 As String = "200.wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.wav"
Public Const SND_DICE As String = "cupdice.wav"

'Gold Sounds
Public Const SND_LOWGOLD As String = "171.wav"
Public Const SND_MIDGOLD As String = "172.wav"
Public Const SND_BIGGOLD As String = "173.wav"

'Ambient Sounds
Public Const SND_MORNING As String = "137.wav"
Public Const SND_AFTERNOON As String = "21.wav"
Public Const SND_EVENING As String = "124.wav"
Public Const SND_EVENINGCITY As String = "141.wav"
Public Const SND_DAYCITY As String = "76.wav"
Public Const SND_TRUENO As String = "61.wav"
Public Const SND_TRUENO2 As String = "60.wav"
Public Const SND_LLUVIA As String = "191.wav"
Public Const SND_LLUVIAEND As String = "192.wav"

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 2788
Public Const INT_MACRO_TRABAJO As Integer = 900
Public Const INT_ATTACK As Integer = 1500
Public Const INT_ARROWS As Integer = 1400
Public Const INT_CAST_SPELL As Integer = 1400
Public Const INT_CAST_ATTACK As Integer = 1000
Public Const INT_WORK As Integer = 700
Public Const INT_USEITEMU As Integer = 250
Public Const INT_USEITEMDCK As Integer = 125
Public Const INT_HUD As Integer = 5000
Public Const INT_SENTRPU As Integer = 2000

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 315
Public Const FRAGATA_FANTASMAL As Integer = 28

Public Const NUMATRIBUTES As Byte = 5

Public Const HUMANO_H_PRIMER_CABEZA As Integer = 1
Public Const HUMANO_H_ULTIMA_CABEZA As Integer = 34
Public Const HUMANO_H_CUERPO_DESNUDO As Integer = 1

Public Const ELFO_H_PRIMER_CABEZA As Integer = 65
Public Const ELFO_H_ULTIMA_CABEZA As Integer = 96
Public Const ELFO_H_CUERPO_DESNUDO As Integer = 1

Public Const DROW_H_PRIMER_CABEZA As Integer = 127
Public Const DROW_H_ULTIMA_CABEZA As Integer = 152
Public Const DROW_H_CUERPO_DESNUDO As Integer = 2

Public Const ENANO_H_PRIMER_CABEZA As Integer = 162
Public Const ENANO_H_ULTIMA_CABEZA As Integer = 191
Public Const ENANO_H_CUERPO_DESNUDO As Integer = 5

Public Const GNOMO_H_PRIMER_CABEZA As Integer = 221
Public Const GNOMO_H_ULTIMA_CABEZA As Integer = 238
Public Const GNOMO_H_CUERPO_DESNUDO As Integer = 5
'**************************************************
Public Const HUMANO_M_PRIMER_CABEZA As Integer = 35
Public Const HUMANO_M_ULTIMA_CABEZA As Integer = 64
Public Const HUMANO_M_CUERPO_DESNUDO As Integer = 3

Public Const ELFO_M_PRIMER_CABEZA As Integer = 97
Public Const ELFO_M_ULTIMA_CABEZA As Integer = 126
Public Const ELFO_M_CUERPO_DESNUDO As Integer = 3

Public Const DROW_M_PRIMER_CABEZA As Integer = 153
Public Const DROW_M_ULTIMA_CABEZA As Integer = 161
Public Const DROW_M_CUERPO_DESNUDO As Integer = 4

Public Const ENANO_M_PRIMER_CABEZA As Integer = 192
Public Const ENANO_M_ULTIMA_CABEZA As Integer = 220
Public Const ENANO_M_CUERPO_DESNUDO As Integer = 6

Public Const GNOMO_M_PRIMER_CABEZA As Integer = 239
Public Const GNOMO_M_ULTIMA_CABEZA As Integer = 255
Public Const GNOMO_M_CUERPO_DESNUDO As Integer = 6

'Musica
Public Const MP3_Inicio As Byte = 101

Public RawServersList As String

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public FogataBufferIndex As Long
Public SoundRainIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public NumEscudosAnims As Integer

Public ArmasHerrero() As tItemsConstruibles
Public ArmadurasHerrero() As tItemsConstruibles
Public ObjCarpintero() As tItemsConstruibles
Public CarpinteroMejorar() As tItemsConstruibles
Public HerreroMejorar() As tItemsConstruibles


Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

Public TradingUserName As String

'Direcciones
Public Enum E_Heading
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
    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Thief       'Ladrón
    Bard        'Bardo
    Druid       'Druida
    Bandit      'Bandido
    Paladin     'Paladín
    Hunter      'Cazador
    Worker      'Trabajador
    Pirat       'Pirata
End Enum

Public Enum eCiudad
    cOnirem = 1
    cOsurac
End Enum

Enum eRaza
    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
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
    Mujer
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
    otLeña = 14
    otFogata = 15
    otescudo = 16
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
    otManchas = 35          'No se usa
    otArbolElfico = 36
    otCualquiera = 1000
End Enum

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
    Comment                 '/REM
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
    banChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    nickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
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
    dumpIPTables            '/DUMPSECURITY
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
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
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
    ToggleCentinelActivated '/CENTINELAACTIVADO
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
End Enum

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "¡¡¡La criatura falló el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "¡¡¡La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "¡¡¡Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "¡¡¡El usuario rechazó el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡¡¡Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_11 As String = "¡"
Public Const MENSAJE_22 As String = "!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te atacó y falló!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la víctima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

Public Const MENSAJE_FRAGSHOOTER_TE_HA_MATADO As String = "te ha matado!"
Public Const MENSAJE_FRAGSHOOTER_HAS_MATADO As String = "Has matado a"
Public Const MENSAJE_FRAGSHOOTER_HAS_GANADO As String = "Has ganado "
Public Const MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA As String = "puntos de experiencia."

Public Const MENSAJE_HAS_MATADO_A As String = "Has matado a "
Public Const MENSAJE_HAS_GANADO_EXPE_1 As String = "Has ganado "
Public Const MENSAJE_HAS_GANADO_EXPE_2 As String = " puntos de experiencia."
Public Const MENSAJE_TE_HA_MATADO As String = " te ha matado!"

Public Const MENSAJE_HOGAR As String = "Has llegado a tu hogar. El viaje ha finalizado."
Public Const MENSAJE_HOGAR_CANCEL As String = "Tu viaje ha sido cancelado."

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
End Enum

'Inventario
Type Inventory
    ObjIndex As Integer
    name As String
    GrhIndex As Integer
    Amount As Long
    Equipped As Byte
    Valor As Single
    ObjType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    ObjIndex As Integer
    name As String
    GrhIndex As Integer
    Amount As Integer
    Valor As Single
    ObjType As Integer
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

Type tReputacion
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
    name As String
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

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
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
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserEnvenenado As Byte
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
'<-------------------------NUEVO-------------------------->

Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 2
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
Public Oscuridad As Integer

Public UsingSkill As Integer

Public pingTime As Long

Public EsPartyLeader As Boolean

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
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
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
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
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends

'CONEXION =============================================================
Public CurServerIp As String
Public CurServerPort As Integer
Public Const MAX_SERVERS As Byte = 2
Type tServerInfo
    port As Integer
    Ip As String
    name As String
End Type
Public lServer(1 To MAX_SERVERS) As tServerInfo
'=======================================================================

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceArmas
    Dir(1 To 4) As Integer
End Type

Public Type tIndiceEscudos
    Dir(1 To 4) As Integer
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
Public clsForos As New clsForum

Public isCapturePending As Boolean
Public Traveling As Boolean

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

Public Const GRH_HALF_STAR As Integer = 5357
Public Const GRH_FULL_STAR As Integer = 5358
Public Const GRH_GLOW_STAR As Integer = 5359

Public Const LH_GRH As Integer = 724
Public Const LP_GRH As Integer = 725
Public Const LO_GRH As Integer = 723

Public Const MADERA_GRH As Integer = 550
Public Const MADERA_ELFICA_GRH As Integer = 1999

Public picMouseIcon As Picture

