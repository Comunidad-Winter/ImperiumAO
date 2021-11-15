Attribute VB_Name = "modClient"
'*****************************************************************
'modClient - ImperiumAO - v1.4.5 R5
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

Private Declare Function EnumDisplaySettings Lib "user32" _
    Alias "EnumDisplaySettingsA" _
    (ByVal lpszDeviceName As Long, ByVal lModeNum As Long, _
    lpudtScreenSettingMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" _
    Alias "ChangeDisplaySettingsA" _
    (lpudtScreenSettingMode As Any, ByVal dwFlags As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, _
   ByVal msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Private Const GWL_WNDPROC       As Long = (-4)
Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_ACTIVATEAPP    As Long = &H1C
Private Const WA_INACTIVE       As Long = 0
Private Const WA_ACTIVE         As Long = 1
Private Const WA_CLICKACTIVE    As Long = 2

Private Const WM_PARENTNOTIFY   As Long = &H210
Private Const WM_CREATE         As Long = &H1
Private Const WM_DESTROY        As Long = &H2

Private Const WM_SHOWWINDOW     As Long = &H18

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

'Minimap UPGrades
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_DISPLAYFREQUENCY = &H400000
Private Const ENUM_CURRENT_SETTINGS = -1

Private Const DISP_CHANGE_SUCCESSFUL = 0

Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H2

Private Type typDEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private curDevMode As typDEVMODE

Public Const GRH_ORO As Integer = 511
Public Const GRH_FOGATA As Integer = 1521

Public IniPath As String
Public MapPath As String

Public EngineRun As Boolean
Public lngClientMutex As Long

Public oldMouseS As Long
Public MouseS As Long

Public FormParser As clsFormParser

Sub DoPasosFx(ByVal CharIndex As Integer)

Static Pie As Integer
Static FileNum As Integer
Static TerrenoDePaso As TipoPaso

Static pos_x As Integer
Static pos_y As Integer

If ((CharIndex <> CurrentUser.CurrentChar) Or (Not CurrentUser.Navegando And Not CurrentUser.Volando And Not CurrentUser.Montando)) Then
        If (Not frmMain.Engine.Char_Dead_Get(CharIndex)) And (frmMain.Engine.Char_In_Current_Area(CharIndex)) And Not (frmMain.Engine.Char_Type_Get(CharIndex) = eGM And frmMain.Engine.Char_Body_Get(CharIndex) = 0) Then

            If frmMain.Engine.Char_Pos_Get(CharIndex, pos_x, pos_y) Then
                Pie = frmMain.Engine.Char_Feet_Switch(CharIndex)
                
                If Pie <> -1 Then
                    FileNum = frmMain.Engine.Map_FileNum_Get(pos_x, pos_y, 1)
                    
                    If Not frmMain.Engine.Char_Big_Get(CharIndex) Then
                        TerrenoDePaso = GetTerrenoDePaso(FileNum)
                    Else
                        TerrenoDePaso = CONST_PESADO
                    End If
                    
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
Dim Engine_Initialized As Boolean
Dim midevM As typDEVMODE

'Initialize the TileEngine
ViewHeight = frmMain.MainViewPic.Height
ViewWidth = frmMain.MainViewPic.Width

Call EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, curDevMode)

If RunWindowed = 0 Then
    If (curDevMode.dmBitsPerPel <> 16) Or (curDevMode.dmPelsHeight <> 600) Or (curDevMode.dmPelsWidth <> 800) Then
        
        midevM = curDevMode
        
        With midevM
              .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
              .dmPelsWidth = 800
              .dmPelsHeight = 600
              .dmBitsPerPel = 16
        End With
        
        Call ChangeDisplaySettings(midevM, 0)
    
    End If
End If

'Siempre en "ventana" (términos D3D)
Engine_Initialized = frmMain.Engine.Engine_Initialize(frmMain, frmMain.MainViewPic.hwnd, True, vbNullString, , , , , 17, 13, 32, True, True, VSYNC, DEV_INDEX, False, False)

If Engine_Initialized Then
    frmMain.Engine.Layer_4_Show_Toggle
    frmMain.Engine.Engine_Label_Render_Set
Else
    MsgBox "¡No se ha logrado iniciar el engine gráfico! Reinstale los últimos controladores de DirectX desde www.imperiumao.com.ar y actualize sus controladores de video. Si el problema persiste por favor consulte los foros de soporte.", vbCritical, "Saliendo"
    Call EndGame
End If

'Set some data in the tile frmMain.Engine.
frmMain.Engine.Engine_Base_Speed_Set 0.029

'Font used for almost everything in our game
frmMain.Engine.Font_Create "Tahoma", 8, False, False

If Sound.Initialize_Engine(frmMain.hwnd, App.Path & "\Recursos", App.Path & "\Recursos", App.Path & "\Recursos", False, (Audio > 0), (sMusica <> CONST_DESHABILITADA), FXVolume, False, InvertirSonido) Then
    'frmCargando.picLoad.Width = 300
Else
    MsgBox "¡No se ha logrado iniciar el engine de DirectSound! Reinstale los últimos controladores de DirectX desde www.imperiumao.com.ar. No habrá soporte de audio en el juego.", vbCritical, "Advertencia"
    frmOpciones.Frame1(0).Enabled = False
End If

Exit Sub

Error_Handler:
    Call MsgBox(Locale_GUI_Frase(348) & " (" & Err.Description & " - " & Err.Number & ")", vbCritical, Locale_GUI_Frase(331))
    Call EndGame
    
End Sub

Public Sub Client_UnInitialize_DirectX_Objects()

On Error Resume Next

'1. Cerramos el engine de sonido y borramos buffers
Sound.Engine_DeInitialize
Set Sound = Nothing

'2. Cerramos el engine gráfico y borramos textures
frmMain.Engine.Engine_DeInitialize

If RunWindowed = 0 Then
    If (curDevMode.dmBitsPerPel <> 16) Or (curDevMode.dmPelsHeight <> 600) Or (curDevMode.dmPelsWidth <> 800) Then
        curDevMode.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        Call ChangeDisplaySettings(curDevMode, 0)
    End If
End If

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

Public Function CalcularMD5HushYo()

If General_File_Exists(App.Path & "\" & App.EXEName & ".exe", vbNormal) Then
    MD5HushYo = MD5File(App.Path & "\" & App.EXEName & ".exe")
Else
    MD5HushYo = MD5File(App.Path & "\ImperiumAO.exe")
End If

End Function

Public Function HabilidadToString(ByVal Habilidades As String) As String

On Error GoTo ErrorHandler

Dim t() As String
Dim i As Integer

t = Split(Habilidades, "-")

For i = LBound(t) To UBound(t)
    If Val(t(i)) > 0 Then
        HabilidadToString = HabilidadToString & HabilidadName(Val(t(i))) & " - "
    End If
Next i

If HabilidadToString <> vbNullString Then _
    HabilidadToString = left$(HabilidadToString, Len(HabilidadToString) - 2)

Exit Function

ErrorHandler:
    HabilidadToString = vbNullString

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
    If LenB(cur$) > 0 Then
        ServersLst(i).IP = General_Field_Read(1, cur$, ":")
        ServersLst(i).Puerto = Val(General_Field_Read(2, cur$, ":"))
        ServersLst(i).Desc = General_Field_Read(3, cur$, ":")
    End If
Next i

CurServer = 1

ServersLstLoaded = True

Call frmConnect.ServerList_Load

End Sub

Public Function CurServerIp() As String

If CurServer <> 0 Then
    CurServerIp = ServersLst(CurServer).IP
End If

End Function

Public Function CurServerPort() As Integer

If CurServer <> 0 Then
    CurServerPort = ServersLst(CurServer).Puerto
End If

End Function

Sub Main()

On Error Resume Next

Dim loopc As Long

If General_File_Exists(App.Path & "\ImperiumAOLauncher.ex_", vbNormal) Then
    Call General_Sleep(2)
    Delete_File App.Path & "\ImperiumAOLauncher.exe"
    Name App.Path & "\ImperiumAOLauncher.ex_" As App.Path & "\ImperiumAOLauncher.exe"
End If

Form_Caption = "ImperiumAO " & App.Major & "." & App.Minor & "." & App.Revision

lngClientMutex = General_CreateMutex("grlcmcl3")

If lngClientMutex = -1 Then
    Call MsgBox("¡ImperiumAO ya está corriendo! No es posible correr otra instancia del juego. Relea el reglamento. Haga click en Aceptar para salir." & vbCrLf & vbCrLf & "ImperiumAO is already running. Game cannot be run. Click OK to quit.", vbApplicationModal + vbInformation + vbOKOnly, "Already running!")
    End
End If
            
Call General_Enable_XPStyle

Set FormParser = New clsFormParser
Call FormParser.Init

Call LoadImpAoInit

If Not Load_Definitions() Then
    MsgBox Locale_GUI_Frase(330), vbCritical, Locale_GUI_Frase(331)
    Call EndGame
End If

Call LoadFontTypes

Call General_SetIcon(frmMain.hwnd, "AAA", True)

DoEvents

Set frmMain.Engine = New clsTileEngineX
Set Sound = New clsSoundEngine
Set Meteo_Engine = New clsMeteorologic
Set ClientTCP = New clsClientTCP

Client_Initialize_DirectX_Objects

frmCargando.Show

'Don't show cursor anymore
If RunWindowed = 0 Then Call General_Cursor_Render(False)

Call PreloadGraphics
'Call PreloadSounds

'Obtener el HushMD5
Call CalcularMD5HushYo

'## SEGURIDAD
Call Main_Logic

If Not GetKeyState(vbKeyShift) < 0 Then
    
    RawServersList = frmCargando.mainInet.OpenURL("http://www.imperiumao.com.ar/serverlist.php")
    
    Do While frmCargando.mainInet.StillExecuting
        DoEvents
    Loop
    
    Call Text_Init

    If LenB(RawServersList) = 0 Or RawServersList = "<h1>Service Unavailable</h1>" Then
        MensajeAdvertencia "No se ha podido cargar la lista de servidores. Le recomendamos verificar el estado de su conexión de internet, en caso de seguir teniendo problemas contacterse con su proveedor de internet. Couldn't load server list. Verify your internet connection, in case of trouble please contact your ISP."
        RawServersList = "server.imperiumao.com.ar:1055:Primario;secundario.imperiumao.com.ar:7891:Secundario;spain.imperiumao.com.ar:7666:España;battle1.imperiumao.com.ar:4655:BattleServer #1;battle2.imperiumao.com.ar:4654:BattleServer #2;battle3.imperiumao.com.ar:4656:BattleServer #3;battle4.imperiumao.com.ar:4657:BattleServer #4;battle5.imperiumao.com.ar:4654:BattleServer #5;"
    End If

Else
    
    RawServersList = "server.imperiumao.com.ar:1055:Primario;secundario.imperiumao.com.ar:7891:Secundario;spain.imperiumao.com.ar:7666:España;battle1.imperiumao.com.ar:4655:BattleServer #1;battle2.imperiumao.com.ar:4654:BattleServer #2;battle3.imperiumao.com.ar:4656:BattleServer #3;battle4.imperiumao.com.ar:4657:BattleServer #4;battle5.imperiumao.com.ar:4654:BattleServer #5;localhost:7666:Local;"

End If

Call InitServersList(RawServersList)

frmCargando.picLoad.Width = 400
frmCargando.picLoad.Refresh

ReDim ListaRazas(1 To NUMRAZAS) As String

For loopc = 1 To NUMRAZAS
    ListaRazas(loopc) = Locale_GUI_Frase(130 + loopc)
Next loopc

ReDim ListaClases(1 To NUMCLASES) As String

For loopc = 1 To NUMCLASES
    ListaClases(loopc) = Locale_GUI_Frase(112 + loopc)
Next loopc

ReDim Head_Range(1 To NUMRAZAS) As tHeadRange

'Male heads
Head_Range(HUMANO).mStart = 1
Head_Range(HUMANO).mEnd = 30
Head_Range(ENANO).mStart = 301
Head_Range(ENANO).mEnd = 315
Head_Range(ELFO).mStart = 101
Head_Range(ELFO).mEnd = 121
Head_Range(DROW).mStart = 202
Head_Range(DROW).mEnd = 212
Head_Range(GNOMO).mStart = 401
Head_Range(GNOMO).mEnd = 409
Head_Range(ORCO).mStart = 501
Head_Range(ORCO).mEnd = 514

'Female heads
Head_Range(HUMANO).fStart = 70
Head_Range(HUMANO).fEnd = 80
Head_Range(ENANO).fStart = 370
Head_Range(ENANO).fEnd = 373
Head_Range(ELFO).fStart = 170
Head_Range(ELFO).fEnd = 189
Head_Range(DROW).fStart = 270
Head_Range(DROW).fEnd = 278
Head_Range(GNOMO).fStart = 470
Head_Range(GNOMO).fEnd = 481
Head_Range(ORCO).fStart = 570
Head_Range(ORCO).fEnd = 573

ReDim SkillsNames(1 To NUMSKILLS) As String

For loopc = 1 To NUMSKILLS
    SkillsNames(loopc) = Locale_GUI_Frase(302 + loopc)
Next loopc

Call CargarPasos
Call CargarParticulas

If sMusica <> CONST_DESHABILITADA Then
    Sound.NextMusic = MUS_Inicio
    Sound.Fading = 350
End If

frmPres.Picture = General_Load_Picture_From_Resource_Ex("presentacion")

frmPres.top = 0
frmPres.left = 0
frmPres.Width = 800 * Screen.TwipsPerPixelX
frmPres.Height = 600 * Screen.TwipsPerPixelY

frmCargando.picLoad.Width = 500
frmCargando.picLoad.Refresh

frmPres.Show
Unload frmCargando

Do While Not FinPres
    If sMusica <> CONST_DESHABILITADA Then Sound.Sound_Render
    DoEvents
Loop

frmConnect.Visible = True
Unload frmPres

'Well let's leave this until GUI is done...
If RunWindowed = 0 Then Call General_Cursor_Render(True)

prgRun = True
CurrentUser.Pausa = False

Call BuffersBorraTimer(True)
Call FXTimer(True)

'On Error Resume Next

Do While prgRun
    
    If EngineRun Then
        If frmMain.WindowState <> vbMinimized Then
            If CurrentUser.MapExt Then Meteo_Engine.Meteo_Logic
            frmMain.Engine.Engine_Render_Start
            frmMain.Engine.Engine_Render_End
            If (Audio = 1 Or sMusica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
            If frmMain.Engine.Engine_Inventory_Render_Get Then Inventory_Render
        End If
    Else
        If (sMusica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
    End If
    
    DoEvents

Loop

EngineRun = False
Call EndGame(True)

Exit Sub

Error_Handler:
    Call MsgBox("Unexpected error: " & Err.Description & " - " & Err.Number, vbCritical, "Quitting")
    Call EndGame
    
End Sub

Public Sub LoadFontTypes()

On Error GoTo ErrorHandler

Dim lc As Integer, Arch As String, tempStr As String

If Not Extract_File(Scripts, App.Path & "\Recursos", "fonttypes.ind", Windows_Temp_Dir, False) Then
    Err.Description = "No se ha logrado extraer el archivo de recurso."
    GoTo ErrorHandler
End If

Arch = Windows_Temp_Dir & "fonttypes.ind"

NUMFONTS = Val(General_Var_Get(Arch, "INIT", "NumFonts"))
ReDim Preserve FontTypes(1 To NUMFONTS) As tFontType

For lc = 1 To NUMFONTS
    tempStr = General_Var_Get(Arch, "INIT", str(lc))
    FontTypes(lc).red = Val(General_Field_Read(2, tempStr, "~"))
    FontTypes(lc).green = Val(General_Field_Read(3, tempStr, "~"))
    FontTypes(lc).blue = Val(General_Field_Read(4, tempStr, "~"))
    FontTypes(lc).bold = Val(General_Field_Read(5, tempStr, "~"))
    FontTypes(lc).italic = Val(General_Field_Read(6, tempStr, "~"))
Next lc

Delete_File Windows_Temp_Dir & "fonttypes.ind"

Exit Sub
    
ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "fonttypes.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "fonttypes.ind"

End Sub

Public Sub LoadImpAoInit()

On Error Resume Next

Dim lc As Integer, Sys_Ram As Double, Leer As New clsIniReader, tmpStr As String

Call Leer.Initialize(App.Path & "\init\" & "ImpAoInit.bnd")

Win2kXP = General_Windows_Is_2000XP
Windows_Temp_Dir = General_Get_Temp_Dir
Publicidad_Contenido = Val(Leer.GetValue("INIT", "Publicidad_Contenido"))
CursoresStandar = Val(Leer.GetValue("INIT", "CursoresStandar"))
LastRunDate = CDate(Leer.GetValue("INIT", "LastRunDate"))

GameLocale = LCase$(Leer.GetValue("INIT", "GameLocale"))
If LenB(GameLocale) = 0 Then GameLocale = General_Language_Default

If Not Load_Locales(GameLocale) Then
    GameLocale = "en"
        
    If Not Load_Locales(GameLocale) Then
        MsgBox "¡No se ha logrado realizar la carga del archivo de idioma! Verifique la integridad del sistema. Si el problema persiste por favor consulte los foros de soporte." & vbCrLf & vbCrLf & "Locale data file could not be loaded. Please refer to tech support if the problem persists.", vbCritical, "Saliendo / Quitting"
        Call EndGame
    End If
    
End If

NombreSkin = Leer.GetValue("INIT", "NombreSkin")

If LenB(General_Field_Read(1, NombreSkin, "_")) = 0 Then
    NombreSkin = NombreSkin & "_" & GameLocale
End If

If General_File_Exists(App.Path & "\Skins\" & NombreSkin & ".ias", vbNormal) = False _
    Or General_Field_Read(2, NombreSkin, "_") <> GameLocale Then
    
    If GameLocale = "es" Then
        NombreSkin = "Principal_es"
    Else
        NombreSkin = "Main_en"
    End If
    
End If

If Not General_File_Exists(App.Path & "\Skins\" & NombreSkin & ".ias", vbNormal) Then
    Call MsgBox(Locale_GUI_Frase(332) & " " & NombreSkin & ".ias " & Locale_GUI_Frase(333), vbCritical, Locale_GUI_Frase(333))
    Call EndGame
End If

Call Set_Skin_Name(NombreSkin)

URL_NEWS = "http://www.imperiumao.com.ar/" & GameLocale & "/noticias.php"

Load frmConnect
Load frmPres
Load frmMensaje
Load frmMain
Load frmCharList
Load frmConnect
Load frmOpciones

NUMBOTONES = 11 'Val(Leer.GetValue("INIT", "NumBotones"))
NUMBINDS = Val(Leer.GetValue("INIT", "NumBinds"))

ReDim Preserve MacroKeys(1 To NUMBOTONES) As tBoton
ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

For lc = 1 To NUMBOTONES
    MacroKeys(lc).TipoAccion = Val(Leer.GetValue("Bind" & lc, "Accion"))
    MacroKeys(lc).hlist = Val(Leer.GetValue("Bind" & lc, "hlist"))
    MacroKeys(lc).invslot = Val(Leer.GetValue("Bind" & lc, "invslot"))
    MacroKeys(lc).SendString = Leer.GetValue("Bind" & lc, "SndString")
Next lc

lc = 0

For lc = 1 To NUMBINDS
    tmpStr = General_Var_Get(App.Path & "\init\" & "ImpAoInit.bnd", "USER", str(lc))
    BindKeys(lc).KeyCode = Val(General_Field_Read(1, tmpStr, ","))
    BindKeys(lc).name = General_Field_Read(2, tmpStr, ",")
    BindKeys(lc).VirtualKey = MapVirtualKey(BindKeys(lc).KeyCode, 0)
    
    If BindKeys(lc).VirtualKey = DIK_NUMPAD4 Then
        BindKeys(lc).VirtualKey = DIK_LEFTARROW
    ElseIf BindKeys(lc).VirtualKey = DIK_NUMPAD6 Then
        BindKeys(lc).VirtualKey = DIK_RIGHTARROW
    ElseIf BindKeys(lc).VirtualKey = DIK_NUMPAD8 Then
        BindKeys(lc).VirtualKey = DIK_UPARROW
    ElseIf BindKeys(lc).VirtualKey = DIK_NUMPAD2 Then
        BindKeys(lc).VirtualKey = DIK_DOWNARROW
    End If
    
Next lc

VerLugar = Val(Leer.GetValue("INIT", "VerLugar"))
FxNavega = Val(Leer.GetValue("INIT", "FxNavega"))

CopiarDialogos = Val(Leer.GetValue("INIT", "CopiarDialogos"))
MensajesGlobales = Val(Leer.GetValue("INIT", "MensajesGlobales"))
MensajesFaccionarios = Val(Leer.GetValue("INIT", "MensajesFaccionarios"))

DefMidi = Val(Leer.GetValue("INIT", "DefaultMidi"))
frmOpciones.chkMidi.value = DefMidi

NombresSimples = Val(Leer.GetValue("INIT", "NombresSimples"))
If NombresSimples = 1 Then frmMain.Engine.Engine_Label_Simple_Set

gldf = Val(Leer.GetValue("INIT", "gldf"))

MusicVolume = Val(Leer.GetValue("INIT", "MusicVolume"))
FXVolume = Val(Leer.GetValue("INIT", "FxVolume"))

DEV_INDEX = Val(Leer.GetValue("INIT", "DeviceIndex"))
VSYNC = Val(Leer.GetValue("INIT", "VSYNC"))
RunWindowed = Val(Leer.GetValue("INIT", "RunWindowed"))

Audio = Val(Leer.GetValue("INIT", "SonidoHabilitado"))
sMusica = Val(Leer.GetValue("INIT", "Musica"))
ListaIgnorados = Leer.GetValue("INIT", "ListaIgnorados")

PreloadLevel = Val(Leer.GetValue("INIT", "BufferTiles"))
InvertirSonido = (Val(Leer.GetValue("INIT", "InvertirSonido")) = 1)

oldMouseS = General_Get_Mouse_Speed
MouseS = Val(Leer.GetValue("INIT", "MouseSpeed"))

If MouseS <= 0 Then
    MouseS = oldMouseS
ElseIf MouseS <> oldMouseS Then
    Call General_Set_Mouse_Speed(MouseS)
End If

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

'MAC_Address = General_Get_MAC_Address

End Sub

Public Sub SaveImpAoInit()

Dim lc As Integer, Arch As String

Arch = App.Path & "\init\" & "ImpAoInit.bnd"

Call General_Var_Write(Arch, "INIT", "NUMBINDS", str(NUMBINDS))
Call General_Var_Write(Arch, "INIT", "NUMBOTONES", str(NUMBOTONES))
Call General_Var_Write(Arch, "INIT", "VerLugar", str(VerLugar))
Call General_Var_Write(Arch, "INIT", "FxNavega", str(FxNavega))
Call General_Var_Write(Arch, "INIT", "DefaultMidi", str(DefMidi))
Call General_Var_Write(Arch, "INIT", "gldf", str(gldf))
Call General_Var_Write(Arch, "INIT", "CopiarDialogos", str(CopiarDialogos))
Call General_Var_Write(Arch, "INIT", "MensajesGlobales", str(MensajesGlobales))
Call General_Var_Write(Arch, "INIT", "MensajesFaccionarios", str(MensajesFaccionarios))
Call General_Var_Write(Arch, "INIT", "CopiarDialogos", str(CopiarDialogos))
Call General_Var_Write(Arch, "INIT", "MusicVolume", str(MusicVolume))
Call General_Var_Write(Arch, "INIT", "FXVolume", str(FXVolume))
Call General_Var_Write(Arch, "INIT", "InvertirSonido", IIf(InvertirSonido = True, "1", "0"))
Call General_Var_Write(Arch, "INIT", "Musica", str(sMusica))
Call General_Var_Write(Arch, "INIT", "SonidoHabilitado", str(Audio))
Call General_Var_Write(Arch, "INIT", "NombreSkin", NombreSkin)
Call General_Var_Write(Arch, "INIT", "NombresSimples", str(NombresSimples))
Call General_Var_Write(Arch, "INIT", "MouseSpeed", str(MouseS))
Call General_Var_Write(Arch, "INIT", "Publicidad_Contenido", str(Publicidad_Contenido))
Call General_Var_Write(Arch, "INIT", "CursoresStandar", str(CursoresStandar))
Call General_Var_Write(Arch, "INIT", "GameLocale", GameLocale)
Call General_Var_Write(Arch, "INIT", "LastRunDate", str(LastRunDate))

For lc = 1 To NUMBINDS
    Call General_Var_Write(Arch, "User", str(lc), str(BindKeys(lc).KeyCode) & "," & BindKeys(lc).name)
Next lc

lc = 0

For lc = 1 To NUMBOTONES
    Call General_Var_Write(Arch, "Bind" & lc, "Accion", str(MacroKeys(lc).TipoAccion))
    Call General_Var_Write(Arch, "Bind" & lc, "hlist", str(MacroKeys(lc).hlist))
    Call General_Var_Write(Arch, "Bind" & lc, "invslot", str(MacroKeys(lc).invslot))
    Call General_Var_Write(Arch, "Bind" & lc, "SndString", MacroKeys(lc).SendString)
Next lc

ListaIgnorados = vbNullString

For lc = 0 To frmOpciones.lstIgnore.ListCount
    If frmOpciones.lstIgnore.List(lc) <> vbNullString Then
        ListaIgnorados = ListaIgnorados & frmOpciones.lstIgnore.List(lc) & "¬"
    End If
Next lc

If ListaIgnorados <> vbNullString Then _
    ListaIgnorados = left$(ListaIgnorados, Len(ListaIgnorados) - 1)

Call General_Var_Write(Arch, "INIT", "ListaIgnorados", ListaIgnorados)

End Sub

Public Sub EndGame(Optional ByVal Closed_ByUser As Boolean = False, Optional ByVal Init_Launcher As Boolean = False)

On Error Resume Next

prgRun = False

'0. Cerramos el socket
If frmMain.MainWinsock.State Then frmMain.MainWinsock.Close

'1. Guardamos datos si se cerró correctamente
If Closed_ByUser Then Call SaveImpAoInit

'2. Eliminamos objetos DX
Call Client_UnInitialize_DirectX_Objects

'3. Cerramos el engine meteorológico
Set Meteo_Engine = Nothing

'4. Borramos otros objetos
Set ClientTCP = Nothing

'5. Deshabilitamos los timers
Call BuffersBorraTimer(False)
Call FXTimer(False)
Call HoraTimer(False)
Call RecSTTimer(False, 0)

'6. Cerramos los forms y nos vamos
Call UnloadAllForms

'7. Adiós MuteX - Restauramos MouseSpeed
Call General_DeleteMutex(lngClientMutex)
Call General_Set_Mouse_Speed(oldMouseS)

'8. ¿Había que prender el launcher?
If Init_Launcher Then ShellExecute GetDesktopWindow, "open", App.Path & "\ImperiumAOLauncher.exe", vbNullString, vbNullString, 1

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
frmMain.lblNick.Caption = CurrentUser.UserName

If Len(CurrentUser.UserName) > 15 Then
    frmMain.lblNick.FontSize = 9
Else
    frmMain.lblNick.FontSize = 14
End If

'Show main form
frmMain.Visible = True

CurrentUser.Logged = True

Call Banner_Init
Call Banner_Logic
If Not Pubilicidad_Deshabilitada Then PubliTimer (True)

EngineRun = True

'Unload forms (don't waste RAM!)
frmIniciando.Visible = False
Unload frmCrearPersonaje

End Sub

Private Function MoveNorth(ByVal CurrentUserIndex As Integer) As Integer

Dim map_x As Integer
Dim map_y As Integer

Call frmMain.Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (frmMain.Engine.Map_Legal_Current_Char_Pos(map_x, map_y - 1) And CurrentUser.Paralizado = False And CurrentUser.Saliendo = False) Then
    If frmMain.Engine.Engine_View_Move(NORTH) Then
        frmMain.Engine.Char_Move CurrentUserIndex, NORTH
        Call ClientTCP.Send_Data(Move_Char_Cl_North)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
    
    MoveNorth = 1
    
Else
    If frmMain.Engine.Char_Heading_Get(CurrentUserIndex) <> NORTH Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_North)
        Call frmMain.Engine.Char_Heading_Set(CurrentUserIndex, NORTH)
    ElseIf frmMain.Engine.Char_Dead_Get(frmMain.Engine.Map_Char_Get(map_x, map_y - 1)) Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_North)
    End If
End If

End Function

Private Function MoveEast(ByVal CurrentUserIndex As Integer) As Integer

Dim map_x As Integer
Dim map_y As Integer

If frmMain.Engine.Char_User_Ladder_Get Then Exit Function

Call frmMain.Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (frmMain.Engine.Map_Legal_Current_Char_Pos(map_x + 1, map_y) And CurrentUser.Paralizado = False And CurrentUser.Saliendo = False) Then
    If frmMain.Engine.Engine_View_Move(EAST) Then
        frmMain.Engine.Char_Move CurrentUserIndex, EAST
        Call ClientTCP.Send_Data(Move_Char_Cl_East)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
    
    MoveEast = 1
    
Else
    If frmMain.Engine.Char_Heading_Get(CurrentUserIndex) <> EAST Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_East)
        Call frmMain.Engine.Char_Heading_Set(CurrentUserIndex, EAST)
    ElseIf frmMain.Engine.Char_Dead_Get(frmMain.Engine.Map_Char_Get(map_x + 1, map_y)) Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_East)
    End If
End If

End Function

Private Function MoveSouth(ByVal CurrentUserIndex As Integer) As Integer

Dim map_x As Integer
Dim map_y As Integer

Call frmMain.Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (frmMain.Engine.Map_Legal_Current_Char_Pos(map_x, map_y + 1) And CurrentUser.Paralizado = False And CurrentUser.Saliendo = False) Then
    If frmMain.Engine.Engine_View_Move(SOUTH) Then
        frmMain.Engine.Char_Move CurrentUserIndex, SOUTH
        Call ClientTCP.Send_Data(Move_Char_Cl_South)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
    
    MoveSouth = 1
    
Else
    If frmMain.Engine.Char_Heading_Get(CurrentUserIndex) <> SOUTH Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_South)
        Call frmMain.Engine.Char_Heading_Set(CurrentUserIndex, SOUTH)
    ElseIf frmMain.Engine.Char_Dead_Get(frmMain.Engine.Map_Char_Get(map_x, map_y + 1)) Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_South)
    End If
End If

End Function

Private Function MoveWest(ByVal CurrentUserIndex As Integer) As Integer

Dim map_x As Integer
Dim map_y As Integer

If frmMain.Engine.Char_User_Ladder_Get Then Exit Function

Call frmMain.Engine.Char_Pos_Get(CurrentUserIndex, map_x, map_y)

If (frmMain.Engine.Map_Legal_Current_Char_Pos(map_x - 1, map_y) And CurrentUser.Paralizado = False And CurrentUser.Saliendo = False) Then
    If frmMain.Engine.Engine_View_Move(WEST) Then
        frmMain.Engine.Char_Move CurrentUserIndex, WEST
        Call ClientTCP.Send_Data(Move_Char_Cl_West)
        Call DoPasosFx(CurrentUserIndex)
        CurrentUser.Moved = True
    End If
    
    MoveWest = 1
    
Else
    If frmMain.Engine.Char_Heading_Get(CurrentUserIndex) <> WEST Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_West)
        Call frmMain.Engine.Char_Heading_Set(CurrentUserIndex, WEST)
    ElseIf frmMain.Engine.Char_Dead_Get(frmMain.Engine.Map_Char_Get(map_x - 1, map_y)) Then
        Call ClientTCP.Send_Data(Change_Heading_Cl_West)
    End If
End If

End Function

Public Function MoveUserChar(ByVal heading As Byte) As Integer

Dim map_x As Integer
Dim map_y As Integer
Dim ran_n As Integer

If (GetActiveWindow <> frmMain.hwnd) And CurrentUser.AutoNavigation = False Then
    MoveUserChar = -1
    Exit Function
End If

If (CurrentUser.CurrentChar <> 0) And (Not CurrentUser.Comerciando) And (Not CurrentUser.Estupido) And (Not CurrentUser.Spectate) Then
    
    Select Case heading
        Case NORTH
            MoveUserChar = MoveNorth(CurrentUser.CurrentChar)
        Case EAST
            MoveUserChar = MoveEast(CurrentUser.CurrentChar)
        Case WEST
            MoveUserChar = MoveWest(CurrentUser.CurrentChar)
        Case SOUTH
            MoveUserChar = MoveSouth(CurrentUser.CurrentChar)
    End Select
    
ElseIf (CurrentUser.Estupido) Then
    ran_n = CInt(General_Random_Number(1, 4))
    
    Select Case ran_n
        Case 1
            Call MoveNorth(CurrentUser.CurrentChar)
        Case 2
            Call MoveEast(CurrentUser.CurrentChar)
        Case 3
            Call MoveWest(CurrentUser.CurrentChar)
        Case Else
            Call MoveSouth(CurrentUser.CurrentChar)
    End Select
    
    MoveUserChar = 1

ElseIf CurrentUser.Spectate Then
    If frmMain.Engine.Engine_View_Move(heading) Then
        
        Select Case heading
            Case NORTH
                Call ClientTCP.Send_Data(Move_Char_Cl_North)
            Case SOUTH
                Call ClientTCP.Send_Data(Move_Char_Cl_South)
            Case WEST
                Call ClientTCP.Send_Data(Move_Char_Cl_West)
            Case EAST
                Call ClientTCP.Send_Data(Move_Char_Cl_East)
        End Select
        
        CurrentUser.Moved = True
        MoveUserChar = 1
        
    End If
End If

Call frmMain.Engine.Engine_View_Pos_Get(map_x, map_y)

If CurrentUser.Reviviendo Then
    Call PrintToConsole(Locale_GUI_Frase(357), 0, 0, 0, 0, 0, 0, 2)
    CurrentUser.Reviviendo = False
End If

If frmMain.UltPos = 0 Then
    If VerLugar = 1 Then
        frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
    End If
ElseIf VerLugar = 0 Then
    frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
End If

General_Update_Minimap map_x, map_y

End Function

Public Function SD(ByVal N As Integer) As Integer

On Error Resume Next

'Suma digitos
Dim auxint As Integer
Dim digit As Byte
Dim suma As Integer
auxint = N

Do
    digit = (auxint Mod 10)
    suma = suma + digit
    auxint = auxint \ 10

Loop While (auxint <> 0)

SD = suma

End Function

Public Function SDM(ByVal N As Integer) As Integer
'Suma digitos cada digito menos dos
Dim auxint As Integer
Dim digit As Integer
Dim suma As Integer
auxint = N

Do
    digit = (auxint Mod 10)
    
    digit = digit - 1
    
    suma = suma + digit
    
    auxint = auxint \ 10

Loop While (auxint <> 0)

SDM = suma

End Function

Public Function Complex(ByVal N As Integer) As Integer

If N Mod 2 <> 0 Then
    Complex = N * SD(N)
Else
    Complex = N * SDM(N)
End If

End Function

Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Sub PrintToConsole(Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean, Optional ByVal italic As Boolean, Optional ByVal bCrLf As Boolean, Optional ByVal FontTypeIndex As Byte = 0)
    
    Dim bUrl As Boolean
    
    With frmMain.RecTxt
        
        .SelFontName = "Tahoma"
        .SelFontSize = 8
        
        If FontTypeIndex <= 0 Then
            
            bUrl = True
            EnableUrlDetect
            
            If (Len(.Text)) > 20000 Then .Text = vbNullString
            .SelStart = Len(frmMain.RecTxt.Text)
            .SelLength = 0
        
            .SelBold = IIf(bold, True, False)
            .SelItalic = IIf(italic, True, False)
            
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        Else
            If (Len(.Text)) > 20000 Then .Text = vbNullString
            
            If FontTypeIndex = FONTTYPE_SERVER Then Text = "Servidor> " & Text
            
            bUrl = (FontTypeIndex = FONTTYPE_SERVER Or FontTypeIndex = FONTTYPE_TALK Or _
                FontTypeIndex = FONTTYPE_GUILDMSG Or FontTypeIndex = FONTTYPE_PIEL Or _
                FontTypeIndex = FONTTYPE_PIEL2)
                        
            If bUrl Then EnableUrlDetect
            
            .SelStart = Len(frmMain.RecTxt.Text)
            .SelLength = 0

            .SelBold = FontTypes(FontTypeIndex).bold
            .SelItalic = FontTypes(FontTypeIndex).italic
            
            If Not red = -1 Then .SelColor = RGB(FontTypes(FontTypeIndex).red, FontTypes(FontTypeIndex).green, FontTypes(FontTypeIndex).blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        End If
    End With
    
    If bUrl Then DisableUrlDetect
    
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

Function AsciiValidos(ByVal Cad As String) As Boolean
Dim car As Byte
Dim i As Integer

Cad = LCase$(Cad)

For i = 1 To Len(Cad)
    car = Asc(mid$(Cad, i, 1))
    
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

If Len(CurrentUser.UserPassword) = 0 Then
    MensajeAdvertencia Locale_GUI_Frase(256)
    Exit Function
End If

For loopc = 1 To Len(CurrentUser.UserPassword)
    CharAscii = Asc(mid$(CurrentUser.UserPassword, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MensajeAdvertencia Locale_GUI_Frase(257)
        Exit Function
    End If
Next loopc

If Len(CurrentUser.AccountName) = 0 Then
    MensajeAdvertencia Locale_GUI_Frase(258)
    Exit Function
End If

If Len(CurrentUser.AccountName) > 20 Then
    MensajeAdvertencia Locale_GUI_Frase(259)
    Exit Function
End If

If Len(CurrentUser.AccountName) > 30 Then
    MensajeAdvertencia Locale_GUI_Frase(260)
    Exit Function
End If

For loopc = 1 To Len(CurrentUser.AccountName)
    CharAscii = Asc(mid$(CurrentUser.AccountName, loopc, 1))
    If LegalCharacter(CharAscii) = False Then
        MensajeAdvertencia Locale_GUI_Frase(251)
        Exit Function
    End If
Next loopc

CurrentUser.AccountName = Trim$(CurrentUser.AccountName)
CheckUserData = True

End Function

Sub UnloadAllForms()

On Error Resume Next
    
Dim miFrm As Form

For Each miFrm In Forms
    Unload miFrm
    Set miFrm = Nothing
Next

Reset

End Sub

Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    
On Error GoTo ErrorHandler
    
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim StreamFile As String
    Dim Leer As New clsIniReader
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "particulas.ini", Windows_Temp_Dir, False) Then
        Err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo ErrorHandler
    End If
    
    StreamFile = Windows_Temp_Dir & "Particulas.ini"
    Leer.Initialize StreamFile
    
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = Leer.GetValue(Val(loopc), "X1")
        StreamData(loopc).y1 = Leer.GetValue(Val(loopc), "Y1")
        StreamData(loopc).x2 = Leer.GetValue(Val(loopc), "X2")
        StreamData(loopc).y2 = Leer.GetValue(Val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
        StreamData(loopc).Speed = Val(Leer.GetValue(Val(loopc), "Speed"))
        
        StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = General_Field_Read(str(i), GrhListing, ",")
        Next i
        
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).b = General_Field_Read(3, TempSet, ",")
        Next ColorSet
        
    Next loopc
    
    Delete_File Windows_Temp_Dir & "particulas.ini"
    Set Leer = Nothing
    
Exit Sub
    
ErrorHandler:
    If General_File_Exists(Windows_Temp_Dir & "particulas.ini", vbNormal) Then Delete_File Windows_Temp_Dir & "particulas.ini"
    
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
        strPreload = General_Var_Get(PreloadFile, "GRAPHICS", str(i))
        MinVal = Val(General_Field_Read(1, strPreload, "-"))
        MaxVal = Val(General_Field_Read(2, strPreload, "-"))
        Priority = Val(General_Field_Read(3, strPreload, "-"))
        
        If Priority <= PreloadLevel Then
            For j = MinVal To MaxVal
                Call frmMain.Engine.Grh_Load(j)
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
        strPreload = General_Var_Get(PreloadFile, "SOUNDS", str(i))
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

On Error GoTo ErrorHandler

    If CurrentUser.UserExp > 0 And CurrentUser.UserPasarNivel > 0 Then
        CurrentUser.UserPercExp = CLng(CurrentUser.UserExp / (CurrentUser.UserPasarNivel / 100))
        If CurrentUser.UserPercExp = 100 Then CurrentUser.UserPercExp = 99
    Else
        CurrentUser.UserPercExp = 0
    End If

Exit Sub

ErrorHandler:

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

Pasos(CONST_PESADO).CantPasos = 3
ReDim Pasos(CONST_PESADO).Wav(1 To Pasos(CONST_PESADO).CantPasos) As Integer
Pasos(CONST_PESADO).Wav(1) = 220
Pasos(CONST_PESADO).Wav(2) = 221
Pasos(CONST_PESADO).Wav(3) = 222

End Sub

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, ByVal PartPos As Byte, Optional ByVal particle_life As Long = 0) As Long

On Error Resume Next

If ParticulaInd <= 0 Then Exit Function

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = frmMain.Engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), PartPos, StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0, Optional ByVal OffsetX As Integer, Optional ByVal OffsetY As Integer) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = frmMain.Engine.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1 + OffsetX, StreamData(ParticulaInd).y1 + OffsetY, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Sub Map_Load(ByVal map_num As Integer, ByVal Ambient_Type As Byte, Optional ByVal base_light As Long = 0, Optional ByVal Day_Night_State As Long = 0)
 
If Extract_File(Maps, App.Path & "\Recursos", "mapa" & map_num & ".csm", Windows_Temp_Dir) Then

    If frmMain.Engine.Map_Load_From_File(Windows_Temp_Dir & "mapa" & map_num & ".csm", True) Then
        CurrentUser.MapExt = Ambient_Type
        
        If CurrentUser.MapExt Then
            Meteo_Engine.ForzarEstado Day_Night_State
        Else
            Call Meteo_Engine.Meteo_Logic
            If base_light <> 0 Then frmMain.Engine.Map_Base_Light_Set General_GetRGB(base_light, 1), General_GetRGB(base_light, 2), General_GetRGB(base_light, 3)
        End If
        
        frmMain.MiniMap.Cls
        frmMain.Engine.Engine_Render_Mini_Map_To_hDC (frmMain.MiniMap.hDC)
        CurrentUser.bLastMiniMap.ColorA = -1
        CurrentUser.bLastMiniMap.ColorB = -1
        CurrentUser.bLastMiniMap.ColorC = -1
        CurrentUser.bLastMiniMap.ColorD = -1
        frmMain.MiniMap.Refresh
        CurrentUser.MapNum = map_num
        
        If VerLugar = 1 Then frmMain.Label2(0).Caption = frmMain.Engine.Map_Name_Get
        
        Call Banner_Logic
        
    End If
    
    Call Delete_File(Windows_Temp_Dir & "mapa" & map_num & ".csm")
    
Else
    'no encontramos el mapa en el hd
    Call MsgBox(Locale_GUI_Frase(349) & " " & map_num & " " & Locale_GUI_Frase(350), vbCritical, Locale_GUI_Frase(331))
    Call EndGame
End If

End Sub

Public Function Map_NameLoad(ByVal map_num As Integer) As String

On Error GoTo ErrorHandler

If Extract_File(Maps, App.Path & "\Recursos", "mapa" & map_num & ".csm", Windows_Temp_Dir) Then

    Map_NameLoad = frmMain.Engine.Map_Name_Load_From_File(Windows_Temp_Dir & "mapa" & map_num & ".csm")

    If LenB(Map_NameLoad) = 0 Then
        Map_NameLoad = "Mapa Desconocido"
    End If
    
    Call Delete_File(Windows_Temp_Dir & "mapa" & map_num & ".csm")
    
Else
    Map_NameLoad = "Mapa Desconocido"
End If

Exit Function

ErrorHandler:
    Map_NameLoad = "Mapa Desconocido"

End Function

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
Call MsgBox(Mensaje, vbInformation + vbOKOnly, Locale_GUI_Frase(351))
End Sub

Public Function NickIgnorado(ByVal Nick As String) As Boolean

Dim i As Long

If Nick <> vbNullString Then
    Nick = UCase$(Nick)
    For i = 0 To frmOpciones.lstIgnore.ListCount
        If Nick = UCase$(frmOpciones.lstIgnore.List(i)) Then
            NickIgnorado = True
            Exit Function
        End If
    Next i
End If

End Function

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

Public Sub ResetCurrentUserEx()

Dim NewCurrUser As tCurrentUser, strAName As String, intCCount As Integer, strName As String

If CurServer < 4 Then
    If CurrentUser.CurrentCharIndex <= 0 Then
        CurrentUser.AccountCharCount = CurrentUser.AccountCharCount + 1
        CurrentUser.CurrentCharIndex = CurrentUser.AccountCharCount
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_clase = CurrentUser.UserClase
    End If
    
    strName = frmMain.Engine.Char_Name_Get_No_Guild(CurrentUser.CurrentChar)
    
    If LenB(strName) > 0 Then
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_level = CurrentUser.UserLVL
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_map = CurrentUser.MapNum
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_body = frmMain.Engine.Char_Body_Get(CurrentUser.CurrentChar)
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_shield = frmMain.Engine.Char_Shield_Get(CurrentUser.CurrentChar)
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_head = frmMain.Engine.Char_Head_Get(CurrentUser.CurrentChar)
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_helmet = frmMain.Engine.Char_Helmet_Get(CurrentUser.CurrentChar)
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_weapon = frmMain.Engine.Char_Weapon_Get(CurrentUser.CurrentChar)
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_type = frmMain.Engine.Char_Type_Get(CurrentUser.CurrentChar)
        ListaPersonajes(CurrentUser.CurrentCharIndex).char_name = strName
        frmCharList.lblAccData(CurrentUser.CurrentCharIndex).Caption = strName
        frmCharList.lblAccData(CurrentUser.CurrentCharIndex).ForeColor = frmMain.Engine.Char_Color_Simple_Get_Ex(ListaPersonajes(CurrentUser.CurrentCharIndex).char_type)
    End If
End If

strAName = CurrentUser.AccountName
intCCount = CurrentUser.AccountCharCount

CurrentUser = NewCurrUser

CurrentUser.CurrentSpeed = VelNormal
CurrentUser.AccountName = strAName
CurrentUser.AccountCharCount = intCCount

frmMain.Engine.Char_Current_Dead_Set (False)
frmMain.Engine.Map_Letter_UnSet
frmMain.Engine.Letter_UnSet
frmMain.Engine.Char_Current_OverWater_Set (False)
frmMain.Engine.Char_Current_OnHorse_Set (False)
frmMain.Engine.Char_Current_Blind_Set (False)
frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)

Call Banner_DeInit

Sound.Sound_Stop_All
Sound.Ambient_Stop

Meteo_Engine.SecondaryStatus = 0

EngineRun = False

End Sub

Public Sub ResetCurrentUser()

On Error Resume Next

Dim NewCurrUser As tCurrentUser

CurrentUser = NewCurrUser

CurrentUser.CurrentSpeed = VelNormal
CurrentUser.bLastMiniMap.ColorA = -1
CurrentUser.bLastMiniMap.ColorB = -1
CurrentUser.bLastMiniMap.ColorC = -1
CurrentUser.bLastMiniMap.ColorD = -1

frmMain.Engine.Map_Letter_UnSet
frmMain.Engine.Letter_UnSet
frmMain.Engine.Char_Current_OverWater_Set (False)
frmMain.Engine.Char_Current_OnHorse_Set (False)
frmMain.Engine.Char_Current_Dead_Set (False)
frmMain.Engine.Char_Current_Blind_Set (False)
frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)

Sound.Sound_Stop_All
Sound.Ambient_Stop

Call Banner_DeInit

Meteo_Engine.SecondaryStatus = 0

EngineRun = False
bK = vbNullString
bRK = 0

End Sub

Public Function RazaToString(ByVal Raza As Byte) As String

Select Case Raza
    Case HUMANO
        RazaToString = "Humano"
    Case ENANO
        RazaToString = "Enano"
    Case ELFO
        RazaToString = "Elfo"
    Case DROW
        RazaToString = "Elfo Drow"
    Case GNOMO
        RazaToString = "Gnomo"
    Case ORCO
        RazaToString = "Orco"
End Select

End Function

Public Function CharClaseValueToString(ByVal Clase As Byte) As String

Select Case Clase

Case CLERIGO
    CharClaseValueToString = "Clérigo"
Case MAGO
    CharClaseValueToString = "Mago"
Case GUERRERO
    CharClaseValueToString = "Guerrero"
Case ASESINO
    CharClaseValueToString = "Asesino"
Case LADRON
    CharClaseValueToString = "Ladrón"
Case BARDO
    CharClaseValueToString = "Bardo"
Case DRUIDA
    CharClaseValueToString = "Druida"
Case CAZARECOMPENSAS
    CharClaseValueToString = "Cazarecompensas"
Case PALADIN
    CharClaseValueToString = "Paladín"
Case CAZADOR
    CharClaseValueToString = "Cazador"
Case PESCADOR
    CharClaseValueToString = "Pescador"
Case HERRERO
    CharClaseValueToString = "Herrero"
Case LEÑADOR
    CharClaseValueToString = "Leñador"
Case MINERO
    CharClaseValueToString = "Minero"
Case CARPINTERO
    CharClaseValueToString = "Carpintero"
Case SASTRE
    CharClaseValueToString = "Sastre"
Case DRAKKAR
    CharClaseValueToString = "Drakkar"
Case NIGROMANTE
    CharClaseValueToString = "Nigromante"
Case gm
    CharClaseValueToString = "Game Master"
Case Else
    CharClaseValueToString = vbNullString

End Select

End Function

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
Porcentaje = (Total * Porc) / 100
End Function

Public Function General_Update_Minimap(ByVal x As Integer, ByVal y As Integer)

'ColorA X = X , Y = Y
'ColorB X = X + 1 , Y = Y
'ColorC X = X , Y = Y - 1
'ColorD X = X + 1 , Y = Y - 1

If CurrentUser.bLastMiniMap.ColorA <> -1 Then _
    SetPixel frmMain.MiniMap.hDC, CurrentUser.bLastMiniMap.x, CurrentUser.bLastMiniMap.y, CurrentUser.bLastMiniMap.ColorA

If CurrentUser.bLastMiniMap.ColorB <> -1 Then _
    SetPixel frmMain.MiniMap.hDC, CurrentUser.bLastMiniMap.x + 1, CurrentUser.bLastMiniMap.y, CurrentUser.bLastMiniMap.ColorB
    
If CurrentUser.bLastMiniMap.ColorC <> -1 Then _
    SetPixel frmMain.MiniMap.hDC, CurrentUser.bLastMiniMap.x, CurrentUser.bLastMiniMap.y - 1, CurrentUser.bLastMiniMap.ColorC

If CurrentUser.bLastMiniMap.ColorD <> -1 Then _
    SetPixel frmMain.MiniMap.hDC, CurrentUser.bLastMiniMap.x + 1, CurrentUser.bLastMiniMap.y - 1, CurrentUser.bLastMiniMap.ColorD

CurrentUser.bLastMiniMap.x = x
CurrentUser.bLastMiniMap.y = y
CurrentUser.bLastMiniMap.ColorA = GetPixel(frmMain.MiniMap.hDC, x, y)
CurrentUser.bLastMiniMap.ColorB = GetPixel(frmMain.MiniMap.hDC, x + 1, y)
CurrentUser.bLastMiniMap.ColorC = GetPixel(frmMain.MiniMap.hDC, x, y - 1)
CurrentUser.bLastMiniMap.ColorD = GetPixel(frmMain.MiniMap.hDC, x + 1, y - 1)

SetPixel frmMain.MiniMap.hDC, x, y, 255
SetPixel frmMain.MiniMap.hDC, x + 1, y, 255
SetPixel frmMain.MiniMap.hDC, x, y - 1, 255
SetPixel frmMain.MiniMap.hDC, x + 1, y - 1, 255

frmMain.MiniMap.Refresh

End Function
