Attribute VB_Name = "modClient"
'*****************************************************************
'modClient - ImperiumAO - v1.3.0
'
'Main client functions.
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

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1

Private Type PointAPI
   x As Long
   y As Long
End Type

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Public Const GRH_ORO As Integer = 511
Public Const GRH_FOGATA As Integer = 1521

Public IniPath As String
Public MapPath As String

Public EngineRun As Boolean

Sub DoPasosFx(ByVal CharIndex As Integer)

Static Pie As Integer
Static FileNum As Integer
Static TerrenoDePaso As TipoPaso

Static pos_x As Integer
Static pos_y As Integer

If ((CharIndex <> CurrentUser.CurrentChar) Or (Not CurrentUser.Navegando And Not CurrentUser.Volando And Not CurrentUser.Montando)) Then
        If (Not Engine.Char_Dead_Get(CharIndex)) And (Engine.Char_In_Current_Area(CharIndex)) And Not (Engine.Char_Type_Get(CharIndex) = 4 And Engine.Char_Body_Get(CharIndex) = 0) Then

            If Engine.Char_Pos_Get(CharIndex, pos_x, pos_y) Then
                Pie = Engine.Char_Feet_Switch(CharIndex)
                
                If Pie <> -1 Then
                    FileNum = Engine.Map_FileNum_Get(pos_x, pos_y, 1)
                    TerrenoDePaso = GetTerrenoDePaso(FileNum)
                    
                    If Pie = 0 Then
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(1), , Sound.Calculate_Volume(pos_x, pos_y), Sound.Calculate_Pan(pos_x, pos_y))
                    Else
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(2), , Sound.Calculate_Volume(pos_x, pos_y), Sound.Calculate_Pan(pos_x, pos_y))
                    End If
                End If
            End If
            
        End If
ElseIf CurrentUser.Navegando And FxNavega = 1 Then
    Call Sound.Sound_Play(SND_NAVEGANDO)
ElseIf CurrentUser.Montando Then
    Call Sound.Sound_Play(Pasos(CONST_CABALLO).Wav(1))
End If

End Sub

Private Function GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As TipoPaso

If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Then
    GetTerrenoDePaso = CONST_BOSQUE
    Exit Function
ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
    GetTerrenoDePaso = CONST_DUNGEON
    Exit Function
ElseIf (TerrainFileNum >= 5000 And TerrainFileNum <= 5004) Then
    GetTerrenoDePaso = CONST_NIEVE
    Exit Function
Else
    GetTerrenoDePaso = CONST_PISO
End If

End Function

Public Sub Client_Initialize_DirectX_Objects()

On Error GoTo Error_Handler

Dim ViewHeight As Integer
Dim ViewWidth As Integer
Dim engine_initialized As Boolean

'Initialize the TileEngine
ViewHeight = frmMain.MainViewPic.Height
ViewWidth = frmMain.MainViewPic.Width

Dim MidevM As typDevMODE
Call EnumDisplaySettings(0, 0, MidevM)

If RunWindowed = 0 Then
    If (MidevM.dmBitsPerPel <> 16) Or (MidevM.dmPelsHeight <> 600) Or (MidevM.dmPelsWidth <> 800) Then
        With MidevM
              .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
              .dmPelsWidth = 800
              .dmPelsHeight = 600
              .dmBitsPerPel = 16
              .dmDisplayFrequency = 75
        End With
        Call ChangeDisplaySettings(MidevM, CDS_TEST)
    End If
End If

'Siempre en "ventana" (términos D3D)
engine_initialized = Engine.Engine_Initialize(frmMain.hwnd, frmMain.MainViewPic.hwnd, True, "", , , , , 17, 13, 32, True, True, VSYNC, DEV_INDEX)
'engine_initialized = Engine.Engine_Initialize(frmMain.hwnd, frmMain.hwnd, False, "", 1024, 768, 0, 0, 17, 13, 32, True, True, VSYNC, DEV_INDEX)

If engine_initialized Then
    Engine.Layer_4_Show_Toggle
    Engine.Engine_Label_Render_Set
Else
    MsgBox "¡No se ha logrado iniciar el engine de Direct3D! Reinstale los últimos controladores de DirectX desde www.imperiumao.com.ar", vbCritical, "Saliendo"
    Call EndGame
End If

'Set some data in the tile engine.
Engine.Engine_Base_Speed_Set 0.029

'Font used for almost everything in our game
Engine.Font_Create "Tahoma", 8, False, False

If Sound.Initialize_Engine(frmMain.hwnd, App.Path & "\Recursos", App.Path & "\MP3\", App.Path & "\Recursos", False, True, True, FXVolume, MusicVolume, InvertirSonido) Then
    'frmCargando.picLoad.Width = 300
Else
    MsgBox "¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX desde www.imperiumao.com.ar", vbCritical, "Saliendo"
    Call EndGame
End If

Exit Sub

Error_Handler:
    Call MsgBox("¡Error al iniciar el motor de DirectX! Reinstale los últimos controladores de DirectX desde www.imperiumao.com.ar", vbCritical, "Saliendo")
    Call EndGame
    
End Sub

Private Function HabilidadName(ByVal Habilidad As Integer) As String
    
Select Case Habilidad
    Case HABILIDAD_INMO
        HabilidadName = "Inmoviliza"
    Case HABILIDAD_PARA
        HabilidadName = "Paraliza"
    Case HABILIDAD_DESCARGA
        HabilidadName = "Lanza descargas"
    Case HABILIDAD_TORMENTA
        HabilidadName = "Lanza fuego"
    Case HABILIDAD_DESENCANTAR
        HabilidadName = "Desencanta al amo"
    Case HABILIDAD_CURAR
        HabilidadName = "Cura al amo"
    Case HABILIDAD_MISIL
        HabilidadName = "Lanza misiles mágicos"
    Case HABILIDAD_DETECTAR
        HabilidadName = "Detecta invisibles"
    Case HABILIDAD_GOLPE_PARALIZA
        HabilidadName = "Paraliza con los golpes"
    Case HABILIDAD_GOLPE_ENTORPECE
        HabilidadName "Entorpece con los golpes"
    Case HABILIDAD_GOLPE_DESARMA
        HabilidadName = "Desarma con los golpes"
    Case HABILIDAD_GOLPE_ENCEGA
        HabilidadName = "Encega con los golpes"
    Case HABILIDAD_GOLPE_ENVENENA
        HabilidadName = "Envenena con los golpes"
    Case Else
        HabilidadName = "Desconocida (" & Habilidad & ")"
End Select

End Function

Private Sub CalcularMD5HushYo()
fMD5HushYo = MD5File(App.Path & "\ImperiumAo.exe")
MD5HushYo = txtOffset(hexMd52Asc(fMD5HushYo), 53)
End Sub

Public Function HabilidadToString(ByVal Habilidades As String) As String

On Error GoTo ErrorHandler

Dim t() As String
Dim i As Integer

t = Split(Habilidades, "-")

For i = LBound(t) To UBound(t)
    HabilidadToString = HabilidadToString & HabilidadName(Val(t(i))) & " - "
Next i

If HabilidadToString <> "" Then _
    HabilidadToString = left$(HabilidadToString, Len(HabilidadToString) - 2)

Exit Function

ErrorHandler:
    HabilidadToString = ""

End Function

Public Function IsIp(ByVal IP As String) As Boolean

Dim i As Integer
For i = 1 To UBound(ServersLst)
    If ServersLst(i).IP = IP Then
        IsIp = True
        Exit Function
    End If
Next i

End Function

Public Sub InitServersList(ByVal Lst As String)

On Error Resume Next

Dim NumServers As Integer
Dim i As Integer, Cont As Integer

Cont = General_Field_Count(RawServersList, Asc(";"))

ReDim ServersLst(1 To Cont) As tServerInfo

For i = 1 To Cont
    Dim cur$
    cur$ = General_Field_Read(i, RawServersList, ";")
    ServersLst(i).IP = General_Field_Read(1, cur$, ":")
    ServersLst(i).Puerto = Val(General_Field_Read(2, cur$, ":"))
    ServersLst(i).Desc = General_Field_Read(3, cur$, ":")
Next i

CurServer = 1

End Sub

Public Function CurServerIp() As String

If CurServer <> 0 Then
    CurServerIp = ServersLst(CurServer).IP
Else
    CurServerIp = frmConnect.IPTxt
End If

End Function

Public Function CurServerPort() As Integer

If CurServer <> 0 Then
    CurServerPort = ServersLst(CurServer).Puerto
Else
    CurServerPort = CInt(frmConnect.PortTxt)
End If

End Function

Sub Main()

Dim loopc As Long

On Error Resume Next

If App.PrevInstance Or (FindWindow(vbNullString, "ImperiumAO 1.3") > 0) Then
    Call MsgBox("¡ImperiumAO ya está corriendo o bien hay una ventana con un nombre similar abierta! No es posible correr otra instancia del juego. Relea el reglamento. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If

Call LoadImpAoInit
Call LoadFontTypes

DoEvents

Set Engine = New clsTileEngineX
Set Sound = New clsSoundEngine
Set Meteo_Engine = New clsMeteorologic

Client_Initialize_DirectX_Objects

frmCargando.Show

#If PreLoad = 1 Then
Call PreloadGraphics
Call PreloadSounds
#End If

frmMain.Inet1.URL = "serverlist.imperiumao.com.ar"
RawServersList = frmMain.Inet1.OpenURL

If RawServersList = "" Or RawServersList = "<h1>Service Unavailable</h1>" Then
    ServersRecibidos = True
    RawServersList = "server.imperiumao.com.ar:7666:Argentina (Desconocido);secundario.imperiumao.com.ar:7666:Secundario (Desconocido);españa.imperiumao.com.ar:7666:España (Desconocido);localhost:7666:Local (Desconocido);home.barrin.com.ar:7666:Barrin's Home (Desconocido)"
Else
    ServersRecibidos = True
End If

Call InitServersList(RawServersList)

frmCargando.picLoad.Width = 400
frmCargando.picLoad.Refresh

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Drow"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
ListaRazas(6) = "Orco"

ReDim ListaClases(1 To NUMCLASES) As String
ListaClases(1) = "Clerigo"
ListaClases(2) = "Mago"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Cazarecompensas"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"
ListaClases(18) = "Nigromante"

Call CargarPasos

Load frmMain
Load frmConnect
Load frmPres

ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)

Call CargarParticulas

If Musica <> CONST_DESHABILITADA Then
    Sound.NextMusic = MUS_Inicio
    Sound.Fading = 350
End If

frmPres.Picture = General_Load_Picture_From_Resource("presentacion.bmp")

frmPres.top = 0
frmPres.left = 0
frmPres.Width = 800 * Screen.TwipsPerPixelX
frmPres.Height = 600 * Screen.TwipsPerPixelY

frmCargando.picLoad.Width = 500
frmCargando.picLoad.Refresh

frmPres.Show
Unload frmCargando

Do While Not FinPres
    If Musica <> CONST_DESHABILITADA Then Sound.Sound_Render
    DoEvents
Loop

frmConnect.Visible = True
Unload frmPres

prgRun = True
CurrentUser.Pausa = False

'Obtener el HushMD5
Call CalcularMD5HushYo

Call BuffersBorraTimer(True)
Call FXTimer(True)
Call HoraTimer(True)

Do While prgRun

    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            'Pedimos que nos envie la posicion
            Call SendData("RPU")
        End If
    End If
    
    If EngineRun Then
        If frmMain.WindowState <> vbMinimized Then
            Check_Keys
            If CurrentUser.MapExt Then Meteo_Engine.Meteo_Logic
            Engine.Engine_Render_Start
            Engine.Engine_Render_End
            If (fx = 1 Or Musica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
            If Engine.Engine_Inventory_Render_Get Then Inventory_Render
        End If
    Else
        If (Musica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
    End If
    
    DoEvents
    
Loop

EngineRun = False
Call EndGame(True)

Exit Sub

ManejadorErrores:
    Call MsgBox("Se ha producido el siguiente error grave durante el MainLoop: " & Err.Description, vbCritical, "Saliendo")
    Call EndGame
    
End Sub

Public Sub LoadFontTypes()

Dim lc As Integer, Arch As String, TempStr As String

Arch = App.Path & "\init\" & "FontTypes.ind"

NUMFONTS = Val(General_Var_Get(Arch, "INIT", "NumFonts"))
ReDim Preserve FontTypes(1 To NUMFONTS) As tFontType

For lc = 1 To NUMFONTS
    TempStr = General_Var_Get(Arch, "INIT", Str(lc))
    FontTypes(lc).red = Val(General_Field_Read(2, TempStr, "~"))
    FontTypes(lc).green = Val(General_Field_Read(3, TempStr, "~"))
    FontTypes(lc).blue = Val(General_Field_Read(4, TempStr, "~"))
    FontTypes(lc).bold = Val(General_Field_Read(5, TempStr, "~"))
    FontTypes(lc).italic = Val(General_Field_Read(6, TempStr, "~"))
Next lc

End Sub

Public Sub LoadImpAoInit()

Dim lc As Integer, Sys_Ram As Double, Leer As New clsLeerInis, TmpStr As String

Leer.Abrir (App.Path & "\init\" & "ImpAoInit.bnd")

NUMBOTONES = Val(Leer.DarValor("INIT", "NumBotones"))
NUMBINDS = Val(Leer.DarValor("INIT", "NumBinds"))

ReDim Preserve MacroKeys(1 To NUMBOTONES) As tBoton
ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

For lc = 1 To NUMBOTONES
    MacroKeys(lc).TipoAccion = Val(Leer.DarValor("Bind" & lc, "Accion"))
    MacroKeys(lc).hlist = Val(Leer.DarValor("Bind" & lc, "hlist"))
    MacroKeys(lc).invslot = Val(Leer.DarValor("Bind" & lc, "invslot"))
    MacroKeys(lc).SendString = Leer.DarValor("Bind" & lc, "SndString")
Next lc

lc = 0

For lc = 1 To NUMBINDS
    TmpStr = General_Var_Get(App.Path & "\init\" & "ImpAoInit.bnd", "USER", Str(lc))
    BindKeys(lc).KeyCode = Val(General_Field_Read(1, TmpStr, ","))
    BindKeys(lc).Name = General_Field_Read(2, TmpStr, ",")
Next lc

VerLugar = Val(Leer.DarValor("INIT", "VerLugar"))
FxNavega = Val(Leer.DarValor("INIT", "FxNavega"))
'EfectosM = Val(Leer.DarValor("INIT", "EfectosM"))

GuardarEXP = Val(Leer.DarValor("INIT", "GuardarExp"))
CopiarDialogos = Val(Leer.DarValor("INIT", "CopiarDialogos"))
MensajesGlobales = Val(Leer.DarValor("INIT", "MensajesGlobales"))

DefMidi = Val(Leer.DarValor("INIT", "DefaultMidi"))
frmOpciones.chkMidi.Value = DefMidi

gldf = Val(Leer.DarValor("INIT", "gldf"))

MusicVolume = Val(Leer.DarValor("INIT", "MusicVolume"))
FXVolume = Val(Leer.DarValor("INIT", "FxVolume"))

DEV_INDEX = Val(Leer.DarValor("INIT", "DeviceIndex"))
VSYNC = Val(Leer.DarValor("INIT", "VSYNC"))
RunWindowed = Val(Leer.DarValor("INIT", "RunWindowed"))

fx = Val(Leer.DarValor("INIT", "SonidoHabilitado"))
Musica = Val(Leer.DarValor("INIT", "Musica"))
ListaIgnorados = Leer.DarValor("INIT", "ListaIgnorados")

PreloadLevel = Val(Leer.DarValor("INIT", "BufferTiles"))
InvertirSonido = (Val(Leer.DarValor("INIT", "InvertirSonido")) = 1)

'Primera vez que ejecuta el cliente
If PreloadLevel = -1 Then
    Sys_Ram = General_Get_Total_Ram
    
    If Sys_Ram >= 512 Then
        PreloadLevel = 4
    ElseIf Sys_Ram >= 256 Then
        PreloadLevel = 3
    ElseIf Sys_Ram >= 128 Then
        PreloadLevel = 2
    Else
        PreloadLevel = 1
    End If
End If

Windows_Temp_Dir = General_Get_Temp_Dir
Win2kXP = General_Windows_Is_2000XP

End Sub

Public Sub SaveImpAoInit()

Dim lc As Integer, Arch As String

Arch = App.Path & "\init\" & "ImpAoInit.bnd"

Call General_Var_Write(Arch, "INIT", "NUMBINDS", Str(NUMBINDS))
Call General_Var_Write(Arch, "INIT", "NUMBOTONES", Str(NUMBOTONES))
Call General_Var_Write(Arch, "INIT", "VerLugar", Str(VerLugar))
Call General_Var_Write(Arch, "INIT", "FxNavega", Str(FxNavega))
'Call General_Var_Write(Arch, "INIT", "EfectosM", Str(EfectosM))
Call General_Var_Write(Arch, "INIT", "DefaultMidi", Str(DefMidi))
Call General_Var_Write(Arch, "INIT", "gldf", Str(gldf))
Call General_Var_Write(Arch, "INIT", "GuardarExp", Str(GuardarEXP))
Call General_Var_Write(Arch, "INIT", "CopiarDialogos", Str(CopiarDialogos))
Call General_Var_Write(Arch, "INIT", "MensajesGlobales", Str(MensajesGlobales))
Call General_Var_Write(Arch, "INIT", "CopiarDialogos", Str(CopiarDialogos))
Call General_Var_Write(Arch, "INIT", "MensajesGlobales", Str(MensajesGlobales))
Call General_Var_Write(Arch, "INIT", "MusicVolume", Str(MusicVolume))
Call General_Var_Write(Arch, "INIT", "FXVolume", Str(FXVolume))
Call General_Var_Write(Arch, "INIT", "InvertirSonido", IIf(InvertirSonido = True, "1", "0"))

For lc = 1 To NUMBINDS
    Call General_Var_Write(Arch, "User", Str(lc), Str(BindKeys(lc).KeyCode) & "," & BindKeys(lc).Name)
Next lc

lc = 0

For lc = 1 To NUMBOTONES
    Call General_Var_Write(Arch, "Bind" & lc, "Accion", Str(MacroKeys(lc).TipoAccion))
    Call General_Var_Write(Arch, "Bind" & lc, "hlist", Str(MacroKeys(lc).hlist))
    Call General_Var_Write(Arch, "Bind" & lc, "invslot", Str(MacroKeys(lc).invslot))
    Call General_Var_Write(Arch, "Bind" & lc, "SndString", MacroKeys(lc).SendString)
Next lc

ListaIgnorados = ""

For lc = 0 To frmOpciones.lstIgnore.ListCount
    If frmOpciones.lstIgnore.List(lc) <> "" Then
        ListaIgnorados = ListaIgnorados & frmOpciones.lstIgnore.List(lc) & "¬"
    End If
Next lc

If ListaIgnorados <> "" Then _
    ListaIgnorados = left$(ListaIgnorados, Len(ListaIgnorados) - 1)

Call General_Var_Write(Arch, "INIT", "ListaIgnorados", ListaIgnorados)

End Sub

Public Sub EndGame(Optional ByVal Closed_ByUser As Boolean = False)

prgRun = False

'1. Guardamos datos si se cerró correctamente
If Closed_ByUser Then Call SaveImpAoInit

'2. Cerramos el engine de sonido y borramos buffers
Sound.Engine_DeInitialize
Set Sound = Nothing

'3. Cerramos el engine gráfico y borramos textures
Engine.Engine_DeInitialize
Set Engine = Nothing

'4. Cerramos el engine meteorológico
Set Meteo_Engine = Nothing

'5. Deshabilitamos los timers
Call BuffersBorraTimer(False)
Call FXTimer(False)
Call HoraTimer(False)

'6. Cerramos los forms y nos vamos
Call UnloadAllForms
End

End Sub

Function LegalCharacter(KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************

'if backspace allow
If KeyAscii = 8 Then
    LegalCharacter = True
    Exit Function
End If

'Only allow space,numbers,letters and special characters
If KeyAscii < 32 Or KeyAscii = 44 Then
    LegalCharacter = False
    Exit Function
End If

If KeyAscii > 126 Then
    LegalCharacter = False
    Exit Function
End If

'Check for bad special characters in between
If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
    LegalCharacter = False
    Exit Function
End If

'else everything is cool
LegalCharacter = True

End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************

'Set the nickname
frmMain.lblNick = CurrentUser.UserName

'Show main form
frmMain.Visible = True

CurrentUser.Logged = True
EngineRun = True

'Unload forms (don't waste RAM!)
If frmCrearPersonaje.Visible Then
    Unload frmPasswd
    Unload frmCrearPersonaje
    Unload frmConnect
Else
    Unload frmIniciando
    Unload frmConnect
End If

End Sub

Private Sub MoveNorth(ByVal CurrentUserIndex As Integer)

Dim map_x As Integer
Dim map_y As Integer

Call Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (Engine.Map_Legal_Current_Char_Pos(map_x, map_y - 1) And CurrentUser.Paralizado = False) Then
    If Engine.Engine_View_Move(NORTH) Then
        Engine.Char_Move CurrentUserIndex, NORTH
        Call SendData("M" & NORTH)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
Else
    If Engine.Char_Heading_Get(CurrentUserIndex) <> NORTH Then
        Call SendData("CHEA" & NORTH)
        Call Engine.Char_Heading_Set(CurrentUserIndex, NORTH)
    ElseIf Engine.Char_Dead_Get(Engine.Map_Char_Get(map_x, map_y - 1)) Then
        Call SendData("CHEA" & NORTH)
    End If
End If

End Sub

Sub MoveEast(ByVal CurrentUserIndex As Integer)

Dim map_x As Integer
Dim map_y As Integer

Call Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (Engine.Map_Legal_Current_Char_Pos(map_x + 1, map_y) And CurrentUser.Paralizado = False) Then
    If Engine.Engine_View_Move(EAST) Then
        Engine.Char_Move CurrentUserIndex, EAST
        Call SendData("M" & EAST)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
Else
    If Engine.Char_Heading_Get(CurrentUserIndex) <> EAST Then
        Call SendData("CHEA" & EAST)
        Call Engine.Char_Heading_Set(CurrentUserIndex, EAST)
    ElseIf Engine.Char_Dead_Get(Engine.Map_Char_Get(map_x + 1, map_y)) Then
        Call SendData("CHEA" & EAST)
    End If
End If

End Sub

Private Sub MoveSouth(ByVal CurrentUserIndex As Integer)

Dim map_x As Integer
Dim map_y As Integer

Call Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (Engine.Map_Legal_Current_Char_Pos(map_x, map_y + 1) And CurrentUser.Paralizado = False) Then
    If Engine.Engine_View_Move(SOUTH) Then
        Engine.Char_Move CurrentUserIndex, SOUTH
        Call SendData("M" & SOUTH)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
Else
    If Engine.Char_Heading_Get(CurrentUserIndex) <> SOUTH Then
        Call SendData("CHEA" & SOUTH)
        Call Engine.Char_Heading_Set(CurrentUserIndex, SOUTH)
    ElseIf Engine.Char_Dead_Get(Engine.Map_Char_Get(map_x, map_y + 1)) Then
        Call SendData("CHEA" & SOUTH)
    End If
End If
End Sub

Sub MoveWest(ByVal CurrentUserIndex As Integer)

Dim map_x As Integer
Dim map_y As Integer

Call Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (Engine.Map_Legal_Current_Char_Pos(map_x - 1, map_y) And CurrentUser.Paralizado = False) Then
    If Engine.Engine_View_Move(WEST) Then
        Engine.Char_Move CurrentUserIndex, WEST
        Call SendData("M" & WEST)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
Else
    If Engine.Char_Heading_Get(CurrentUserIndex) <> WEST Then
        Call SendData("CHEA" & WEST)
        Call Engine.Char_Heading_Set(CurrentUserIndex, WEST)
    ElseIf Engine.Char_Dead_Get(Engine.Map_Char_Get(map_x - 1, map_y)) Then
        Call SendData("CHEA" & WEST)
    End If
End If

End Sub

Sub MoveUserChar(ByVal Heading As Byte)

Dim map_x As Integer
Dim map_y As Integer

Dim ran As Integer

If (CurrentUser.CurrentChar <> 0) And (Not CurrentUser.Comerciando) And (Not CurrentUser.Estupido) Then
    
    Select Case Heading
        Case NORTH
            Call MoveNorth(CurrentUser.CurrentChar)
        Case EAST
            Call MoveEast(CurrentUser.CurrentChar)
        Case WEST
            Call MoveWest(CurrentUser.CurrentChar)
        Case SOUTH
            Call MoveSouth(CurrentUser.CurrentChar)
    End Select
ElseIf (CurrentUser.Estupido) Then
    ran = CInt(General_Random_Number(1, 4))
    Select Case ran
        Case 1
            Call MoveNorth(CurrentUser.CurrentChar)
        Case 2
            Call MoveEast(CurrentUser.CurrentChar)
        Case 3
            Call MoveWest(CurrentUser.CurrentChar)
        Case Else
            Call MoveSouth(CurrentUser.CurrentChar)
    End Select
End If

If CurrentUser.Reviviendo Then
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Te has alejado de tu cuerpo!", 0, 0, 0, 0, 0, 0, 2)
    CurrentUser.Reviviendo = False
End If

If frmMain.UltPos = 0 Then
    Call Engine.Char_Pos_Get(CurrentUser.CurrentChar, map_x, map_y)
    frmMain.Label2(0).Caption = "Posición: " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
End If

End Sub

Sub RandomMove()

Dim j As Integer

j = General_Random_Number(1, 4)

Select Case j
    Case 1
        Call MoveEast(CurrentUser.CurrentChar)
    Case 2
        Call MoveNorth(CurrentUser.CurrentChar)
    Case 3
        Call MoveWest(CurrentUser.CurrentChar)
    Case 4
        Call MoveSouth(CurrentUser.CurrentChar)
End Select

End Sub

Public Function SD(ByVal n As Integer) As Integer

On Error Resume Next

'Suma digitos
Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = n

Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10

Loop While (auxint <> 0)

SD = suma

End Function

Public Function SDM(ByVal n As Integer) As Integer
'Suma digitos cada digito menos dos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = n

Do
    digit = (auxint Mod 10)
    
    digit = digit - 1
    
    suma = suma + digit
    
    auxint = auxint \ 10

Loop While (auxint <> 0)

SDM = suma

End Function

Public Function Complex(ByVal n As Integer) As Integer

If n Mod 2 <> 0 Then
    Complex = n * SD(n)
Else
    Complex = n * SDM(n)
End If

End Function

Public Function ValidarLoginMSG(ByVal n As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(n)
AuxInteger2 = SDM(n)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Sub AddtoRichTextBox(RichTextBox As RichTextBox, Text As String, Optional red As Integer = -1, Optional green As Integer, Optional blue As Integer, Optional bold As Boolean, Optional italic As Boolean, Optional bCrLf As Boolean, Optional FontTypeIndex As Integer = 0)
    With RichTextBox
        If FontTypeIndex <= 0 Then
            If (Len(.Text)) > 20000 Then .Text = ""
            .SelStart = Len(RichTextBox.Text)
            .SelLength = 0
        
            .SelBold = IIf(bold, True, False)
            .SelItalic = IIf(italic, True, False)
            
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        Else
            If (Len(.Text)) > 20000 Then .Text = ""
            .SelStart = Len(RichTextBox.Text)
            .SelLength = 0
        
            .SelBold = FontTypes(FontTypeIndex).bold
            .SelItalic = FontTypes(FontTypeIndex).italic
            
            If Not red = -1 Then .SelColor = RGB(FontTypes(FontTypeIndex).red, FontTypes(FontTypeIndex).green, FontTypes(FontTypeIndex).blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        End If
    End With
End Sub
'[END]'


Sub AddtoTextBox(TextBox As TextBox, Text As String)
'******************************************
'Adds text to a text box at the bottom.
'Automatically scrolls to new text.
'******************************************

TextBox.SelStart = Len(TextBox.Text)
TextBox.SelLength = 0


TextBox.SelText = Chr(13) & Chr(10) & Text

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) And (car <> 209) And (car <> 241) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function CheckUserData() As Boolean

Dim loopc As Integer
Dim CharAscii As Integer

If CurrentUser.UserPassword = "" Then
    Call MensajeAdvertencia("¡Ingrese una contraseña!")
    Exit Function
End If

For loopc = 1 To Len(CurrentUser.UserPassword)
    CharAscii = Asc(mid$(CurrentUser.UserPassword, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        Call MensajeAdvertencia("¡Ingrese una contraseña válida!")
        Exit Function
    End If
Next loopc

If CurrentUser.UserName = "" Then
    Call MensajeAdvertencia("¡Ingrese un nombre válido!")
    Exit Function
End If

If Len(CurrentUser.UserName) > 25 Then
    Call MensajeAdvertencia("El nombre debe tener menos de 25 letras")
    Exit Function
End If

For loopc = 1 To Len(CurrentUser.UserName)

    CharAscii = Asc(mid$(CurrentUser.UserName, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        Call MensajeAdvertencia("Nombre inválido")
        Exit Function
    End If
    
Next loopc

CheckUserData = True

End Function

Sub UnloadAllForms()

On Error Resume Next
    
Dim miFrm As Form

For Each miFrm In Forms
    Unload miFrm
Next

End Sub

Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long

    Dim StreamFile As String
    StreamFile = App.Path & "\init\" & "Particulas.ini"

    TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = General_Var_Get(StreamFile, Val(loopc), "X1")
        StreamData(loopc).y1 = General_Var_Get(StreamFile, Val(loopc), "Y1")
        StreamData(loopc).x2 = General_Var_Get(StreamFile, Val(loopc), "X2")
        StreamData(loopc).y2 = General_Var_Get(StreamFile, Val(loopc), "Y2")
        StreamData(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
        StreamData(loopc).vecx1 = General_Var_Get(StreamFile, Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = General_Var_Get(StreamFile, Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = General_Var_Get(StreamFile, Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = General_Var_Get(StreamFile, Val(loopc), "VecY2")
        StreamData(loopc).life1 = General_Var_Get(StreamFile, Val(loopc), "Life1")
        StreamData(loopc).life2 = General_Var_Get(StreamFile, Val(loopc), "Life2")
        StreamData(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
        StreamData(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = General_Var_Get(StreamFile, Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = General_Var_Get(StreamFile, Val(loopc), "XMove")
        StreamData(loopc).YMove = General_Var_Get(StreamFile, Val(loopc), "YMove")
        StreamData(loopc).move_x1 = General_Var_Get(StreamFile, Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = General_Var_Get(StreamFile, Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = General_Var_Get(StreamFile, Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = General_Var_Get(StreamFile, Val(loopc), "move_y2")
        StreamData(loopc).life_counter = General_Var_Get(StreamFile, Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(General_Var_Get(StreamFile, Val(loopc), "Speed"))
        
        StreamData(loopc).NumGrhs = General_Var_Get(StreamFile, Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(Str(i), GrhListing, ",")
        Next i
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
    Next loopc
End Sub

Public Sub PreloadGraphics()

    Dim PreloadFile As String
    Dim strPreload As String
    Dim NumPreload As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    Dim MinVal As Integer
    Dim MaxVal As Integer
    Dim Priority As Byte
    
    Dim TotalPreloads As Integer
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "preload.ind", Windows_Temp_Dir, False) Then
        Err.Description = "No se ha logrado extraer el archivo de recurso."
        GoTo ErrorHandler
    End If
    
    PreloadFile = Windows_Temp_Dir & "Preload.ind"
    
    TotalPreloads = Val(General_Var_Get(PreloadFile, "GRAPHICS", "TotalPreloads"))
    If TotalPreloads = 0 Then TotalPreloads = 1
    
    modProgress = ((200 / TotalPreloads))
    
    NumPreload = Val(General_Var_Get(PreloadFile, "GRAPHICS", "NumGraphics"))
    
    For i = 1 To NumPreload
        strPreload = General_Var_Get(PreloadFile, "GRAPHICS", Str(i))
        MinVal = Val(General_Field_Read(1, strPreload, "-"))
        MaxVal = Val(General_Field_Read(2, strPreload, "-"))
        Priority = Val(General_Field_Read(3, strPreload, "-"))
        
        If Priority <= PreloadLevel Then
            For j = MinVal To MaxVal
                Call Engine.Grh_Load(j)
                frmCargando.picLoad.Width = frmCargando.picLoad.Width + modProgress
                DoEvents
            Next j
        End If
    Next i
    
    Delete_File Windows_Temp_Dir & "Preload.ind"
    
    Exit Sub
    
ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "Preload.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Preload.ind"

End Sub

Public Sub PreloadSounds()

    On Error GoTo ErrorHandler

    Dim PreloadFile As String
    Dim strPreload As String
    
    Dim NumPreload As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    Dim MinVal As Integer
    Dim MaxVal As Integer
    Dim Priority As Byte
    
    Dim TotalPreloads As Integer
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "preload.ind", Windows_Temp_Dir, False) Then
        Err.Description = "No se ha logrado extraer el archivo de recurso."
        GoTo ErrorHandler
    End If
    
    PreloadFile = Windows_Temp_Dir & "Preload.ind"
    
    TotalPreloads = Val(General_Var_Get(PreloadFile, "SOUNDS", "TotalPreloads"))
    If TotalPreloads = 0 Then TotalPreloads = 1

    modProgress = ((200 / TotalPreloads))

    NumPreload = Val(General_Var_Get(PreloadFile, "SOUNDS", "NumSounds"))
    
    For i = 1 To NumPreload
        strPreload = General_Var_Get(PreloadFile, "SOUNDS", Str(i))
        MinVal = Val(General_Field_Read(1, strPreload, "-"))
        MaxVal = Val(General_Field_Read(2, strPreload, "-"))
        Priority = Val(General_Field_Read(3, strPreload, "-"))
        
        If Priority <= PreloadLevel Then
            For j = MinVal To MaxVal
                Call Sound.Sound_Load(j)
                frmCargando.picLoad.Width = frmCargando.picLoad.Width + modProgress
                DoEvents
            Next j
        End If
    Next i
    
    Delete_File Windows_Temp_Dir & "Preload.ind"
    
    Exit Sub
    
ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "Preload.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Preload.ind"
    
End Sub

Public Sub UserExpPerc()

    If CurrentUser.UserExp > 0 And CurrentUser.UserPasarNivel > 0 Then
        CurrentUser.UserPercExp = CLng((CurrentUser.UserExp * 100) / CurrentUser.UserPasarNivel)
        If CurrentUser.UserPercExp = 100 Then CurrentUser.UserPercExp = 99
    Else
        CurrentUser.UserPercExp = 0
    End If

End Sub

Public Sub PetExpPerc()

    If CurrentUser.UserPet.EXP > 0 And CurrentUser.UserPet.ELU > 0 Then
        CurrentUser.PetPercExp = CLng((CurrentUser.UserPet.EXP * 100) / CurrentUser.UserPet.ELU)
        If CurrentUser.PetPercExp = 100 Then CurrentUser.PetPercExp = 99
    Else
        CurrentUser.PetPercExp = 0
    End If

End Sub

Private Sub CargarPasos()

ReDim Pasos(1 To NUM_PASOS) As tPaso

Pasos(CONST_BOSQUE).CantPasos = 2
ReDim Pasos(CONST_BOSQUE).Wav(1 To Pasos(CONST_BOSQUE).CantPasos) As Integer
Pasos(CONST_BOSQUE).Wav(1) = 201
Pasos(CONST_BOSQUE).Wav(2) = 202

Pasos(CONST_NIEVE).CantPasos = 2
ReDim Pasos(CONST_NIEVE).Wav(1 To Pasos(CONST_NIEVE).CantPasos) As Integer
Pasos(CONST_NIEVE).Wav(1) = 199
Pasos(CONST_NIEVE).Wav(2) = 200

Pasos(CONST_CABALLO).CantPasos = 2
ReDim Pasos(CONST_CABALLO).Wav(1 To Pasos(CONST_CABALLO).CantPasos) As Integer
Pasos(CONST_CABALLO).Wav(1) = 23
Pasos(CONST_CABALLO).Wav(2) = 24

Pasos(CONST_DUNGEON).CantPasos = 2
ReDim Pasos(CONST_DUNGEON).Wav(1 To Pasos(CONST_DUNGEON).CantPasos) As Integer
Pasos(CONST_DUNGEON).Wav(1) = 23
Pasos(CONST_DUNGEON).Wav(2) = 24

Pasos(CONST_DESIERTO).CantPasos = 2
ReDim Pasos(CONST_DESIERTO).Wav(1 To Pasos(CONST_DESIERTO).CantPasos) As Integer
Pasos(CONST_DESIERTO).Wav(1) = 197
Pasos(CONST_DESIERTO).Wav(2) = 198

Pasos(CONST_PISO).CantPasos = 2
ReDim Pasos(CONST_PISO).Wav(1 To Pasos(CONST_PISO).CantPasos) As Integer
Pasos(CONST_PISO).Wav(1) = 23
Pasos(CONST_PISO).Wav(2) = 24

End Sub

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = Engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = Engine.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Function UserLocation(UserMap As Integer) As String

Select Case UserMap
    Case 1
        UserLocation = "Ullathorpe"
    Case 2, 3, 4
        UserLocation = "Sendero del Sur"
    Case 8, 9, 10
        UserLocation = "Bosques Sagrados"
    Case 35
        UserLocation = "Colinas de Nix"
    Case 34
        UserLocation = "Nix"
    Case 80, 78, 79, 87, 88
        UserLocation = "Costas de Nix"
    Case 99
        UserLocation = "Costas de Rinkel"
    Case 89 To 103
        UserLocation = "Río Otoren"
    Case 63, 62, 64
        UserLocation = "Lindos"
    Case 104, 105, 106, 152
        UserLocation = "Costas de Lindos"
    Case 153, 154
        UserLocation = "Río Nueva Esperanza"
    Case 155
        UserLocation = "Costas de Nueva Esperanza"
    Case 113, 114
        UserLocation = "Arrecifes Nueva Esperanza"
    Case 111, 112
        UserLocation = "Nueva Esperanza"
    Case 107, 108, 109, 120, 121, 122, 123, 125, 126, 127, 128 To 136
        UserLocation = "Océano Abierto"
    Case 124
        UserLocation = "Los Cuatro Vientos"
    Case 147, 148
        UserLocation = "Río Arghâl"
    Case 149
        UserLocation = "Costas de Arghâl"
    Case 150
        UserLocation = "Puerto de Arghâl"
    Case 151
        UserLocation = "Arghâl Central"
    Case 156
        UserLocation = "Arghâl Oeste"
    Case 138
        UserLocation = "Aguas Abandonadas"
    Case 139
        UserLocation = "Isla Rakj"
    Case 47
        UserLocation = "Isla Beleta"
    Case 49
        UserLocation = "Intermundia"
    Case 137
        UserLocation = "Costas de Banderbille"
    Case 61
        UserLocation = "Puerto de Banderbille"
    Case 59
        UserLocation = "Banderbille Central"
    Case 60
        UserLocation = "Distrito Real"
    Case 58
        UserLocation = "Barrios Bajos"
    Case 66
        UserLocation = "Cárcel del Imperio"
    Case 140 To 145
        UserLocation = "Dungeon Veriil"
    Case 146
        UserLocation = "Fuerte Veriil"
    Case 48
        UserLocation = "Dungeon Dragon"
    Case 37
        UserLocation = "Newbie Dungeon (Nivel 1)"
    Case 115, 116
        UserLocation = "Dungeon Marabel"
    Case 39, 38
        UserLocation = "Bosque Dorck"
    Case 40 To 45
        UserLocation = "Catacumbas"
    Case 15, 16, 17, 21
        UserLocation = "Desierto Rinkel"
    Case 33
        UserLocation = "Minas Thyr"
    Case 50 To 52
        UserLocation = "Minas Dorck"
    Case 28 To 32, 46
        UserLocation = "Bosques de Nix"
    Case 20
        UserLocation = "Rinkel"
    Case 11 To 23, 25, 26, 27
        UserLocation = "Campos Abiertos"
    Case 24
        UserLocation = "Bosque Gran Aullido"
    Case 157, 158, 160, 161, 159, 162
        UserLocation = "Bosques de Banderbille"
    Case 36
        UserLocation = "Campamento Orco"
    Case 5, 6, 7, 53 To 57, 65, 67 To 75
        UserLocation = "Sendero del Norte"
    Case 76
        UserLocation = "Tundra Marabel"
    Case 209, 210, 211
        UserLocation = "Dungeon Zero"
    Case 208
        UserLocation = "Dungeon Newbie (Nivel 2)"
    Case 207
        UserLocation = "Dungeon Gaugin"
    Case 205
        UserLocation = "Minas de Oro"
    Case 183, 184, 185
        UserLocation = "Wonder"
    Case 194
        UserLocation = "Illiandor"
    Case 179
        UserLocation = "Puerto de Illiandor"
    Case 195, 196, 202, 203
        UserLocation = "Bosques Illiandor"
    Case 186 To 193
        UserLocation = "Sendero del Este"
    Case 197 To 200
        UserLocation = "Aguas Sagradas"
    Case 201
        UserLocation = "Tundra Zero"
    Case 163 To 178
        UserLocation = "Río del este"
    Case 182
        UserLocation = "Puerto de Orac"
    Case 181
        UserLocation = "Orac"
    Case 180
        UserLocation = "Costas de Orac"
    Case 204
        UserLocation = "Segundo Piso"
    Case 110
        UserLocation = "Isla Morgolock"
    Case 218
        UserLocation = "Tiama"
    Case 217
        UserLocation = "Costas de Tiama"
    Case 219 To 229
        UserLocation = "Tundra Tiama"
    Case 230 To 232
        UserLocation = "Dungeon Cristal"
    Case 233
        UserLocation = "Mina Kirle"
    Case 234, 235, 236, 211 To 216
        UserLocation = "Río Glacial"
    Case 237, 238
        UserLocation = "Arena"
    Case Else
        UserLocation = "Mapa Desconocido"
End Select

End Function

Public Sub Map_Load(ByVal TempInt As Integer, ByVal TempStr As String)
 
If General_File_Exists(App.Path & "\Mapas\" & "Mapa" & TempInt & ".map", vbNormal) Then
                                                                  
    'Si es la vers correcta cambiamos el mapa
    If Engine.Map_Load_From_File(App.Path & "\Mapas\" & "Mapa" & TempInt & ".map", False, True) Then
        If (InStr(TempStr, "+") <> 0) Then
            Meteo_Engine.ForzarEstado Val(General_Field_Read(2, TempStr, "+"))
            CurrentUser.MapExt = Val(General_Field_Read(3, TempStr, "+"))
        Else
            Call Engine.Engine_Meteo_Particle_Set(-1)
            CurrentUser.MapExt = 0
            If Val(TempStr) <> 0 Then Engine.Map_Base_Light_Set Val(TempStr)
        End If
                            
        Engine.Map_Name_Set (UserLocation(TempInt))
        Engine.Engine_Render_Mini_Map_To_hDC (frmMain.MiniMap.hDC)
        frmMain.MiniMap.Refresh
        CurrentUser.MapNum = TempInt
        frmMain.Label2(0).Caption = Engine.Map_Name_Get
    End If
    
Else
    'no encontramos el mapa en el hd
    Call MsgBox("No se ha encontrado el mapa " & TempInt & " por favor baje nuevamente el juego desde www.imperiumao.com.ar.", vbCritical, "Saliendo")
    Call EndGame
End If

End Sub

Public Sub Make_Transparent_Richtext(ByVal hwnd As Long)

If Win2kXP Then _
    Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

End Sub

Public Sub Make_Transparent_Form(ByVal hwnd As Long, Optional ByVal bytOpacity As Byte = 128)

If Win2kXP Then
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwnd, 0, bytOpacity, LWA_ALPHA)
End If

End Sub

Public Sub UnMake_Transparent_Form(ByVal hwnd As Long)

If Win2kXP Then _
    Call SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) And (Not WS_EX_TRANSPARENT))

End Sub

Public Sub Auto_Drag(ByVal hwnd As Long)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub

Public Sub MensajeAdvertencia(ByVal Mensaje As String)
Call MsgBox(Mensaje, vbInformation + vbOKOnly, "Advertencia")
End Sub

Public Function NickIgnorado(ByVal Nick As String) As Boolean

Dim i As Long

If Nick <> "" Then
    Nick = UCase$(Nick)
    For i = 0 To frmOpciones.lstIgnore.ListCount
        If Nick = UCase$(frmOpciones.lstIgnore.List(i)) Then
            NickIgnorado = True
            Exit Function
        End If
    Next i
End If

End Function

Private Sub Check_Keys()

If Not CurrentUser.Pausa And frmMain.Visible And Not frmForo.Visible And _
    Not frmComerciar.Visible And Not frmComerciarUsu.Visible And Not frmGoliath.Visible And CurrentUser.Logged And Not frmIniciando.Visible Then
    
    'Move Up
    If Engine.Input_Key_Get(BindKeys(14).KeyCode) Then
        Call MoveUserChar(NORTH)
    'Move Right
    ElseIf Engine.Input_Key_Get(BindKeys(17).KeyCode) And Not Engine.Input_Key_Get(vbKeyShift) Then
        Call MoveUserChar(EAST)
    'Move down
    ElseIf Engine.Input_Key_Get(BindKeys(15).KeyCode) Then
        Call MoveUserChar(SOUTH)
    'Move left
    ElseIf Engine.Input_Key_Get(BindKeys(16).KeyCode) And Not Engine.Input_Key_Get(vbKeyShift) Then
          Call MoveUserChar(WEST)
    End If

End If

End Sub

Public Function RealSkillToIndex(ByVal Skill As Integer) As Integer

Select Case Skill
    Case 4
        RealSkillToIndex = 1
    Case 5
        RealSkillToIndex = 2
    Case 20
        RealSkillToIndex = 3
    Case 7
        RealSkillToIndex = 4
    Case 23
        RealSkillToIndex = 5
    Case 19
        RealSkillToIndex = 6
    Case 12
        RealSkillToIndex = 7
    Case 2
        RealSkillToIndex = 8
    Case 22
        RealSkillToIndex = 9
    Case 6
        RealSkillToIndex = 10
    Case 8
        RealSkillToIndex = 11
    Case 18
        RealSkillToIndex = 12
    Case 1
        RealSkillToIndex = 13
    Case 3
        RealSkillToIndex = 14
    Case 11
        RealSkillToIndex = 15
    Case 9
        RealSkillToIndex = 16
    Case 17
        RealSkillToIndex = 17
    Case 13
        RealSkillToIndex = 18
    Case 14
        RealSkillToIndex = 19
    Case 10
        RealSkillToIndex = 20
    Case 26
        RealSkillToIndex = 21
    Case 16
        RealSkillToIndex = 22
    Case 15
        RealSkillToIndex = 23
    Case 24
        RealSkillToIndex = 24
    Case 25
        RealSkillToIndex = 25
    Case 21
        RealSkillToIndex = 26
    Case 27
        RealSkillToIndex = 27
End Select

End Function

Public Function SkillRealToIndex(ByVal SkillIndex As Integer) As Integer

Select Case SkillIndex
    Case 1
        SkillRealToIndex = 4
    Case 2
        SkillRealToIndex = 5
    Case 3
        SkillRealToIndex = 20
    Case 4
        SkillRealToIndex = 7
    Case 5
        SkillRealToIndex = 23
    Case 6
        SkillRealToIndex = 19
    Case 7
        SkillRealToIndex = 12
    Case 8
        SkillRealToIndex = 2
    Case 9
        SkillRealToIndex = 22
    Case 10
        SkillRealToIndex = 6
    Case 11
        SkillRealToIndex = 8
    Case 12
        SkillRealToIndex = 18
    Case 13
        SkillRealToIndex = 1
    Case 14
        SkillRealToIndex = 3
    Case 15
        SkillRealToIndex = 11
    Case 16
        SkillRealToIndex = 9
    Case 17
        SkillRealToIndex = 17
    Case 18
        SkillRealToIndex = 13
    Case 19
        SkillRealToIndex = 14
    Case 20
        SkillRealToIndex = 10
    Case 21
        SkillRealToIndex = 26
    Case 22
        SkillRealToIndex = 16
    Case 23
        SkillRealToIndex = 15
    Case 24
        SkillRealToIndex = 24
    Case 25
        SkillRealToIndex = 25
    Case 26
        SkillRealToIndex = 21
    Case 27
        SkillRealToIndex = 27
End Select

End Function
