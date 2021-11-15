Attribute VB_Name = "modDeclaraciones"
'*****************************************************************
'modDeclaraciones - ImperiumAO - v1.3.0
'
'Main client declares.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public RawServersList As String

Public Type tServerInfo
    IP As String
    Puerto As Integer
    Desc As String
End Type

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer
Public RunWindowed As Byte

Public Const bCabeza As Byte = 1
Public Const bPiernaIzquierda As Byte = 2
Public Const bPiernaDerecha As Byte = 3
Public Const bBrazoDerecho As Byte = 4
Public Const bBrazoIzquierdo As Byte = 5
Public Const bTorso As Byte = 6

Public ArmasHerrero() As Integer
Public ArmadurasHerrero() As Integer
Public ItemsHerrero() As Integer
Public ObjCarpintero() As Integer
Public ObjDruida() As Integer
Public ObjSastre() As Integer

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS As Integer = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]

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
    Name As String
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
    
    speed As Single
    life_counter As Long
End Type

'Direcciones
Public Const NORTH As Byte = 1
Public Const EAST As Byte = 3
Public Const SOUTH As Byte = 5
Public Const WEST As Byte = 7

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Integer = 25
Public Const MAX_COFREINVENTORY_SLOTS As Integer = 10
Public Const MAXHECHI As Integer = 50

Public Const NUMSKILLS As Integer = 27
Public Const NUMATRIBUTOS As Integer = 5
Public Const NUMCLASES As Integer = 18
Public Const NUMRAZAS As Integer = 6

Public Const Fuerza As Integer = 1
Public Const Agilidad As Integer = 2
Public Const Inteligencia As Integer = 3
Public Const Carisma As Integer = 4
Public Const Constitucion As Integer = 5

Public Const MAXSKILLPOINTS As Integer = 100

Public Const FLAGORO As Integer = -777

Public Const Magia As Integer = 2
Public Const Robar As Integer = 3
Public Const Ocultarse As Integer = 8
Public Const Talar As Integer = 10
Public Const Pesca As Integer = 13
Public Const Mineria As Integer = 14
Public Const Carpinteria As Integer = 15
Public Const Herreria As Integer = 16
Public Const Domar As Integer = 18
Public Const Proyectiles As Integer = 19
Public Const Arrojadizas As Integer = 23
Public Const Pociones As Integer = 24
Public Const Sastreria As Integer = 25
Public Const Jardineria As Integer = 26

Public Const FundirMetal As Integer = 88
Public Const Esposas As Integer = 89
Public Const Grupo As Integer = 90

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Long
    Equipped As Byte
    Valor As Long
    ObjType As Integer
    Def As Integer
    MaxHIT As Integer
    MinHIT As Integer
End Type

Type NpcInv
    OBJIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Integer
    Valor As Long
    ObjType As Integer
    Def As Integer
    MaxHIT As Integer
    MinHIT As Integer
    C1 As String
    C2 As String
    c3 As String
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
    Culpabilidad As Long
    Promedio As Long
End Type

Type tListaFamiliares
    Name As String
    Desc As String
    Imagen As String
End Type

Public ListaRazas() As String
Public ListaClases() As String
Public ListaFamiliares() As tListaFamiliares

Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
Public OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public NPCInventory(1 To MAX_INVENTORY_SLOTS) As NpcInv
Public NPCInvDim As Integer

Public CurrentUser As tCurrentUser

Public MD5HushYo As String * 16
Public fMD5HushYo As String * 32

'[Barrin]
Public PuedeTorneo As Byte

Public Const Legal As Byte = 2
Public Const Caotico As Byte = 3
Public Const Neutral As Byte = 1

'[/Barrin]

Public Enum E_MODO
    NORMAL = 1
    CrearNuevoPj = 2
End Enum

Public EstadoLogin As E_MODO

Public Enum E_SISTEMA_MUSICA
    CONST_DESHABILITADA = 0
    CONST_MIDI = 1
    CONST_MP3 = 2
End Enum

Public Musica As E_SISTEMA_MUSICA

'Public Ambiente As Byte
Public fx As Byte
Public FxNavega As Byte
Public DefMidi As Byte

'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtBuffer As String 'Holds temp raw data from server

'String contants
Public ENDC As String 'Endline character for talking with server
Public ENDL As String 'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends
Public FinPres As Boolean 'When presentation is done

Public IPdelServidor As String
Public PuertoDelServidor As String

Public VerLugar As Byte
Public GuardarEXP As Byte
Public CopiarDialogos As Byte
Public MensajesGlobales As Byte

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Type tFamiliar
    TieneFamiliar As Integer
    nombre As String
    ELV As Integer
    MinHP As Integer
    MaxHP As Integer
    ELU As Long
    EXP As Long
    MinHIT As Integer
    MaxHIT As Integer
    Abilidad As String
    Tipo As String
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    offsetX As Integer
    OffsetY As Integer
End Type

Public Type tUserStats
    Clase As Integer
    Raza As Integer
    Genero As Byte
    CiudasMatados As Integer
    CrimisMatados As Integer
    NPCsMatados As Integer
    TimesKilled As Long
End Type

Public Type tIntervalos
    Ataque As Long
    Uso As Long
    Trabajo As Long
    Hechizo As Long
End Type

'[Barrin]

Public Type tCurrentUser
    MapNum As Integer
    MapExt As Byte
    Navegando As Boolean
    Montando As Boolean
    Volando As Boolean
    Paralizado As Boolean
    Transformado As Boolean
    Meditando As Boolean
    Comerciando As Boolean
    Muerto As Boolean
    Reviviendo As Boolean
    Ciego As Boolean
    Estupido As Boolean
    Descansando As Boolean
    UserMaxHP As Integer
    UserMinHP As Integer
    UserMaxMAN As Integer
    UserMinMAN As Integer
    UserMaxSTA As Integer
    UserMinSTA As Integer
    UserMaxAGU As Integer
    UserMinAGU As Integer
    UserMaxHAM As Integer
    UserMinHAM As Integer
    UserGLD As Long
    UserLVL As Integer
    UserPuedeRefrescar As Boolean
    UserPasarNivel As Long
    UserExp As Long
    PetPercExp As Long
    UserPercExp As Long
    ExpCount As Long
    Seguro As Boolean
    Combate As Boolean
    Rol As Boolean
    UserPet As tFamiliar
    UserStats As tUserStats
    UserReputacion As tReputacion
    Intervalos As tIntervalos
    UserHechizos(1 To MAXHECHI) As Integer
    UserSkills(1 To NUMSKILLS) As Integer
    UserAtributos(1 To NUMATRIBUTOS) As Integer
    SendingType As Byte
    sndPrivateTo As String
    CurrentSpeed As Single
    SkillPoints As Integer
    UserClase As Byte
    UserSexo As Byte
    UserRaza As Byte
    UserEmail As String
    UserHogar As Byte
    UserName As String
    UserPassword As String
    CreandoClan As Boolean
    ClanName As String
    Site As String
    Logged As Boolean
    Pausa As Boolean
    Ping As Long
    CurrentChar As Long
    Saliendo As Boolean
    Moved As Boolean
    UsingSkill As Integer
End Type

'Public bAmbiente As Boolean

Public Const VelNormal As Single = 4.3
Public Const VelLenta As Single = 3
Public Const VelRapida As Single = 6.3
Public Const VelUltra As Single = 15

Public MusicVolume As Long
Public FXVolume As Long
Public FadeMod As Integer
Public gldf As Byte
Public DEV_INDEX As Long
Public VSYNC As Boolean
Public InvertirSonido As Boolean

Public Const Campo As String = "CAMPO"
Public Const Ciudad As String = "CIUDAD"

Public Const Bosque As String = "BOSQUE"
Public Const NIEVE As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"

Public Const SND_AVE As Integer = 21
Public Const SND_AVE2 As Integer = 28
Public Const SND_AVE3 As Integer = 34
Public Const SND_AVE4 As Integer = 29
Public Const SND_GRILLO As Integer = 22
Public Const SND_CUERVO As Integer = 126
Public Const SND_TRUENO1 As Integer = 60
Public Const SND_TRUENO2 As Integer = 61
Public Const SND_TRUENO3 As Integer = 62
Public Const SND_TRUENO4 As Integer = 63
Public Const SND_TRUENO5 As Integer = 64

Public Const MAX_HABILIDADES As Integer = 4
Public Const COLOR_ATAQUE As Long = -65536

Public Enum HabilidadesFamiliar
    HABILIDAD_INMO = 1
    HABILIDAD_PARA = 2
    HABILIDAD_DESCARGA = 3
    HABILIDAD_TORMENTA = 4
    HABILIDAD_DESENCANTAR = 5
    HABILIDAD_CURAR = 6
    HABILIDAD_MISIL = 7
    HABILIDAD_DETECTAR = 8
    HABILIDAD_GOLPE_PARALIZA = 9
    HABILIDAD_GOLPE_ENTORPECE = 10
    HABILIDAD_GOLPE_DESARMA = 11
    HABILIDAD_GOLPE_ENCEGA = 12
    HABILIDAD_GOLPE_ENVENENA = 13
End Enum

Public NUMFONTS As Integer

Public FontTypes() As tFontType

Public Type tFontType
    bold As Boolean
    italic As Boolean
    red As Integer
    green As Integer
    blue As Integer
End Type

Public Enum TipoPaso
    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7
End Enum

Public Type tPaso
    CantPasos As Byte
    Wav() As Integer
End Type

Public Const NUM_PASOS As Byte = 6
Public Pasos() As tPaso

Public Const MUS_Inicio As Integer = 54
Public Const MUS_CrearPersonaje As Integer = 48
Public Const MUS_VolverInicio As Integer = 53
Public Const MUS_Noche As Integer = 52
Public Const MUS_Loading As Integer = 72

Public Const SND_CLICK As Integer = 190
Public Const SND_NAVEGANDO As Integer = 50
Public Const SND_OVER As Integer = 0
Public Const SND_DICE As Integer = 188
Public Const SND_FUEGO As Integer = 79

Public Const SND_LLUVIAIN As Integer = 191
Public Const SND_LLUVIAOUT As Integer = 194

Public Const SND_NIEVEIN As Integer = 191
Public Const SND_NIEVEOUT As Integer = 194

Public Const SND_RESUCITAR As Integer = 104
Public Const SND_CURAR As Integer = 101

Public Const CANT_GRH_INDEX As Long = 40000

Public Const CentroInventario As Byte = 1
Public Const CentroHechizos As Byte = 2
Public Const CentroMenu As Byte = 3
Public Const Solapas As Byte = 4

Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

'Mouse location
Public MouseX As Integer
Public MouseY As Integer
Public MouseHitX As Integer
Public MouseHitY As Integer
Public MouseHit As Integer
Public MouseHitText As Integer
Public MouseHitButton As Integer
Public MouseHitComboBox As Integer
Public MouseHitComboBoxY As Integer    'X coord is not necessary
Public MouseHitLabel As Integer

'Used for Non-Click Movement
Public NonClickMovement As Boolean

Public bK As Long
Public bRK As Long

Public Engine As New clsTileEngineX
Public Sound As New clsSoundEngine
Public Meteo_Engine As clsMeteorologic

Public Windows_Temp_Dir As String
Public Win2kXP As Boolean

Public PreloadLevel As Integer
Public modProgress As Single

Public Const CLERIGO As Integer = 1
Public Const MAGO As Integer = 2
Public Const GUERRERO As Integer = 3
Public Const ASESINO As Integer = 4
Public Const LADRON As Integer = 5
Public Const BARDO As Integer = 6
Public Const DRUIDA As Integer = 7
Public Const CAZARECOMPENSAS As Integer = 8
Public Const PALADIN As Integer = 9
Public Const CAZADOR As Integer = 10
Public Const PESCADOR As Integer = 11
Public Const HERRERO As Integer = 12
Public Const LEÑADOR As Integer = 13
Public Const MINERO As Integer = 14
Public Const CARPINTERO As Integer = 15
Public Const SASTRE As Integer = 16
Public Const PIRATA As Integer = 17
Public Const NIGROMANTE As Integer = 18
Public Const gm As Integer = 50

Public Const HUMANO As Byte = 1
Public Const ENANO As Byte = 2
Public Const ELFO As Byte = 3
Public Const DROW As Byte = 4
Public Const GNOMO As Byte = 5
Public Const ORCO As Byte = 6

Public Const Masculino As Byte = 1
Public Const Femenino As Byte = 2

Public ListaIgnorados As String

'[/Barrin]
