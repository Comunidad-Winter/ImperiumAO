VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIAO_TE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const M_VERSION As String = "V 1.1 fr"
Private Const MAIN_SPEED As Single = 0.08
Private Const SCREEN_PIXELS_POR_MOVIMIENTO = 8 'cada main_speed mueve tantos pixels la pantalla
Private Const ANIM_SPEED As Single = MAIN_SPEED
Private Const VIDADIALOGOS As Single = 8 'vida en segundos
Private Const BufferTiles = 4 'tiles extra para cada lado
Private Const PJ_X_OFFSET = 13
Private Const PJ_X_OFFSET_HEAD = 8
Private Const LIMIT_FPS_ANIM = 0.05 'para 20 fps

Private dx As DirectX8
Private D3D As Direct3D8
Private D3dx As D3DX8
Private D3DDevice As Direct3DDevice8
Private D3DSprite As D3DXSprite
Private Texturas As CTextureMan

Private Type tFont
    Fuente As Direct3DTexture8
    FontMap() As Byte
    FontSize As Byte
    FontName As String
End Type
Private Type tInfoEfecto
    Inicio As Double
    Duracion As Long
    Estado As Integer
    AuxDatas As Single
    AuxDataL As Long
    Direccion As D3DVECTOR2
    Color As Long
    Alpha As Byte
End Type
Private EfectosDesc() As tInfoEfecto

Private MainColor As Long

Private NumeroFuentes As Long
Private Fuentes() As tFont


Private m_Pic As PictureBox
Private WindowTilesX As Long 'tiles
Private WindowTilesY As Long 'tiles
Private BackBufferRect As RECT
Private ScreenRect As RECT

Private m_AlphaValues(0 To 255) As Long

Private RecInv As RECT
Private SourceRect As RECT
Private Scaling As D3DVECTOR2
Private Center As D3DVECTOR2
Private Rotation As Single
Private Translation As D3DVECTOR2

Private MainTimer As PerformanceTimer
Private AnimacionTimer As PerformanceTimer
Private EngineTime As Double
Private Rendering As Boolean

Private MoviendoScreen As Long
Private ScreenAvance As D3DVECTOR2

Private MinMapRender As D3DVECTOR2
Private MaxMapRender As D3DVECTOR2

'FPSs
Private FPSCounter As Long
Private FPSTimer As PerformanceTimer
Private m_FPS As Long

'textos
Private m_DialogoTiempo As Double

'publicas

Public TileSz As Long 'pixels
Public Mouse_Tile_X As Long
Public Mouse_Tile_Y As Long
Public Mouse_Pixel_X As Long
Public Mouse_Pixel_Y As Long
Public PuedoMover As Boolean


Property Get FPSs() As Long
FPSs = m_FPS
End Property



Private Sub Class_Initialize()
PuedoMover = True
Set MainTimer = New PerformanceTimer
Set AnimacionTimer = New PerformanceTimer
Set FPSTimer = New PerformanceTimer
Set Texturas = New CTextureMan
ReDim Ef2ectosDesc(1 To Efectostotales)
FPSTimer.StartTimer
AnimacionTimer.StartTimer

InitAlphaValues
ShowSplash
End Sub

Private Sub InitAlphaValues()
Dim p As Long
For p = 0 To 127
    m_AlphaValues(255 - p) = &HFF000000 - &H1000000 * p
Next p
For p = 0 To 127
    m_AlphaValues(p) = &H1000000 * p
Next p
End Sub


Public Function IniciarEngine(pic As PictureBox, TilesX As Long, TilesY As Long, TamTile As Long) As Boolean
WindowTilesX = TilesX
WindowTilesY = TilesY
TileSz = TamTile
Set m_Pic = pic
pic.width = WindowTilesX * TileSz
pic.height = WindowTilesY * TileSz
With ScreenRect
    .left = 0
    .top = 0
    .Bottom = WindowTilesY * TileSz
    .Right = WindowTilesX * TileSz
End With

DoEvents
If Not InitDX Then
    Exit Function
End If
Texturas.Path = App.Path & "\graficos\"
Texturas.mMaxEntries = 300
Texturas.Init D3dx, D3DDevice
IniciarOtros
IniciarFuentes
ReDim EfectosDesc(1 To Efectostotales)

IniciarEngine = True
End Function

Private Sub IniciarFuentes()
Dim i As Integer
Dim Fh2 As Integer
Dim FN As String
Dim fh As Integer
Dim sData As String
FN = App.Path & "\fonts\"
fh = FreeFile
ReDim Fuentes(1 To 1)
ReDim Fuentes(1).FontMap(1 To 4, 0 To 255)
Open FN & "fd_arial.dat" For Input As fh
    Get #fh, , Fuentes(1).FontMap()
Close fh
    

'Line Input #fh, sData
'NumeroFuentes = Val(sData)

'For i = 1 To NumeroFuentes
'    Line Input #fh, sData
'    Fuentes(i).FontName = sData
'    Line Input #fh, sData
'    Fuentes(i).FontSize = Val(sData)
'    ReDim Fuentes(i).FuentesWidth(1 To 255)
'    Fh2 = FreeFile
'    Open FN & "FW_" & Fuentes(i).FontName & ".dat" For Binary As Fh2
'        Get #Fh2, , Fuentes(i).FuentesWidth
'    Close Fh2
    Set Fuentes(1).Fuente = D3dx.CreateTextureFromFileEx(D3DDevice, FN & "arial.png", _
      D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
      D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)
'Next i
'Close fh
End Sub
Private Sub IniciarOtros()
MinXBorder = XMinMapSize + (WindowTilesX \ 2)
MaxXBorder = XMaxMapSize - (WindowTilesX \ 2)
MinYBorder = YMinMapSize + (WindowTilesY \ 2)
MaxYBorder = YMaxMapSize - (WindowTilesY \ 2)
'Set intial user position
UserPos.X = MinXBorder
UserPos.Y = MinYBorder
inipath = App.Path & "\Init\"
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

If Musica = 0 Or fx = 0 Then
    DirectSound.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
End If
Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs

With RecInv
    .left = 0
    .top = 0
    .Right = frmMain.picInv.width
    .Bottom = frmMain.picInv.height
End With

AddtoRichTextBox frmCargando.Status, "Cargando Gráficos....", 0, 0, 0, , , True
frmCargando.Refresh
'Call LoadGraphics

End Sub
Private Sub Class_Terminate()
Set MainTimer = Nothing
Set AnimacionTimer = Nothing
Set FPSTimer = Nothing
Set Texturas = Nothing
CleanUpDX
End Sub

Private Sub ShowSplash()
Load frmSplash
frmSplash.lblCaption = frmSplash.lblCaption & M_VERSION
frmSplash.Timer1.Enabled = True
DoEvents
frmSplash.Tag = "1"
'frmSplash.Show vbModal
Unload frmSplash
End Sub

Private Function InitDX() As Boolean
On Local Error GoTo errHIDX
Dim mode As D3DDISPLAYMODE
Dim d3dpp As D3DPRESENT_PARAMETERS

Set dx = New DirectX8
Set D3D = dx.Direct3DCreate
Set D3dx = New D3DX8

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode

With d3dpp
    .windowed = 1
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    .BackBufferFormat = mode.format
    .BackBufferCount = 1
    .BackBufferWidth = m_Pic.width + BufferTiles * TileSz * 2 'a los 2 costados
    .BackBufferHeight = m_Pic.height + BufferTiles * TileSz * 2 'arriba y abajo
End With
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, m_Pic.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

If D3DDevice Is Nothing Then Exit Function

With D3DDevice
    .SetRenderState D3DRS_LIGHTING, 0
    .SetRenderState D3DRS_ALPHABLENDENABLE, 1
    .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    .Clear 0, ByVal 0, D3DCLEAR_TARGET, &HFF000000, 1#, 0
    .BeginScene
    .EndScene
    .Present ByVal 0, ByVal 0, 0, ByVal 0
End With
Set D3DSprite = D3dx.CreateSprite(D3DDevice)

InitDX = True
Exit Function
errHIDX:

MsgBox Err.Description
End Function
Private Sub CleanUpDX()
Dim i As Integer
D3DDevice.SetTexture 0, Nothing
For i = 1 To NumeroFuentes
    Set Fuentes(i).Fuente = Nothing
Next i
Set D3DSprite = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set D3dx = Nothing
Set dx = Nothing
End Sub

Private Sub UpdateFPS()
FPSCounter = FPSCounter + 1
FPSTimer.StopTimer
If FPSTimer.TimeElapsed >= 1 Then
    m_FPS = FPSCounter
    FPSCounter = 0
    FPSTimer.StartTimer
End If
End Sub
Private Sub ChkAnims()
AnimacionTimer.StopTimer
If AnimacionTimer.TimeElapsed >= LIMIT_FPS_ANIM Then 'a 20 fps
    UpdateAnims
    AnimacionTimer.StartTimer
End If
End Sub
Private Sub UpdateAnims()
Dim i As Integer
Dim j As Integer
Dim auxGrh As Integer
Dim TempChar As Char
For j = MinMapRender.Y To MaxMapRender.Y
    For i = MinMapRender.X To MaxMapRender.X
    If j >= YMinMapSize And j <= YMaxMapSize Then
        If i >= XMinMapSize And i <= XMaxMapSize Then
            'layers
            For auxGrh = 1 To 3 'no hay techos animados!
                If MapData(i, j).Graphic(auxGrh).started = 1 Then _
                    animGrh MapData(i, j).Graphic(auxGrh)
            Next auxGrh
            'objetos
            If MapData(i, j).ObjGrh.started = 1 Then _
                animGrh MapData(i, j).ObjGrh
            'fx
            If MapData(i, j).fx <> 0 Then _
                animGrh FxData(MapData(i, j).fx).fx, 0, True, i, j
            'pj
            If MapData(i, j).CharIndex > 0 Then
                TempChar = CharList(MapData(i, j).CharIndex)
                If TempChar.fx <> 0 Then _
                    animGrh FxData(TempChar.fx).fx, MapData(i, j).CharIndex, True, i, j
                If TempChar.seMueve Then
                    If TempChar.Body.Walk(TempChar.Heading).started = 1 Then _
                        animGrh TempChar.Body.Walk(TempChar.Heading)
                    If TempChar.Arma.WeaponWalk(TempChar.Heading).started = 1 Then _
                        animGrh TempChar.Arma.WeaponWalk(TempChar.Heading)
                    If TempChar.Escudo.ShieldWalk(TempChar.Heading).started = 1 Then _
                        animGrh TempChar.Arma.WeaponWalk(TempChar.Heading)
                End If
                CharList(MapData(i, j).CharIndex) = TempChar
            End If 'pj
        End If 'pos válido
    End If 'pos válido
    Next i
Next j
End Sub

Private Sub animGrh(Anim As Grh, Optional KillAnim As Integer = 0, Optional AnimateOnce As Boolean = False, Optional AnimX As Integer = 0, Optional AnimY As Integer = 0)
With Anim
    If .speedcounter > 0 Then
        .speedcounter = .speedcounter - 1
        If .speedcounter = 0 Then
            If .grhindex = 0 Then Exit Sub
            .speedcounter = GrhData(.grhindex).speed
            .framecounter = .framecounter + 1
            If .framecounter > GrhData(.grhindex).NumFrames Then _
                .framecounter = 1
            If KillAnim Then
                If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                    If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                    If CharList(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx ;))
                        CharList(KillAnim).fx = 0
                        'bug!
                        'al matar el fx, se mata en charlist, y después cuando asigno charlist(x)= tempchar, la vuelvo a poner :(
                    End If
                End If
            'Barrin
            ElseIf AnimateOnce Then
                MapData(AnimX, AnimY).fx = 0
            End If 'kill anim
        End If
    End If
End With
End Sub


Public Sub Render()
Dim sPos As Single
UpdateFPS
ChkAnims

Scaling.X = 1
Scaling.Y = 1
Rotation = 0 'Rotation + 0.01
Center.X = 0
Center.Y = 0
'Center.x = m_Pic.Width / 2
'Center.y = m_Pic.Height / 2
Translation.X = 0
Translation.Y = 0
With BackBufferRect
    .left = TileSz * BufferTiles
    .top = TileSz * BufferTiles
    .Right = .left + WindowTilesX * TileSz
    .Bottom = .top + WindowTilesY * TileSz
End With
MainColor = m_AlphaValues(255) + &HFFFFFF   ' + MapBaseColor en lugar de &hffffff

EngineTime = MainTimer.TiempoActual
If MoviendoScreen <> 0 Then
    MainTimer.StopTimer
    sPos = TileSz - (MainTimer.TimeElapsed / MAIN_SPEED) * SCREEN_PIXELS_POR_MOVIMIENTO
        
    If sPos <= 0 Then
        sPos = 0
        MoviendoScreen = 0
        PuedoMover = True
    Else
        With BackBufferRect
            .left = .left - ScreenAvance.X * sPos
            .top = .top - ScreenAvance.Y * sPos
            .Right = .left + WindowTilesX * TileSz
            .Bottom = .top + WindowTilesY * TileSz
        End With
    End If
End If 'moviendo screen

MinMapRender.X = UserPos.X - (WindowTilesX \ 2) - BufferTiles
MinMapRender.Y = UserPos.Y - (WindowTilesY \ 2) - BufferTiles
MaxMapRender.X = UserPos.X + (WindowTilesX \ 2) + BufferTiles
MaxMapRender.Y = UserPos.Y + (WindowTilesY \ 2) + BufferTiles

AplicarEfectos

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HFF000000, 0#, 0
D3DDevice.BeginScene
D3DSprite.Begin

RenderLayers1y2

RenderObj3yPjs

RenderLayerFxy4

RenderText 1, "FPS: " & m_FPS, 0, 0, 255, 255
D3DSprite.End
D3DDevice.EndScene
D3DDevice.Present BackBufferRect, ScreenRect, 0, ByVal 0
End Sub

Private Sub AplicarEfectos()
If EfectosDesc(eTerremoto).Duracion <> 0 Then
    With EfectosDesc(eTerremoto)
        .Duracion = .Duracion - 1
        If .Estado <= 0 Then
            .Estado = RandomNumber(.AuxDataL \ 10, .AuxDataL)
            Select Case CInt(RandomNumber(0, 3))
                Case 0
                    .Direccion.X = 1
                    .Direccion.Y = 0
                Case 1
                    .Direccion.X = -1
                    .Direccion.Y = 0
                Case 2
                    .Direccion.X = 0
                    .Direccion.Y = 1
                Case Else
                    .Direccion.X = 0
                    .Direccion.Y = -1
            End Select
            .Inicio = EngineTime
        End If
        
        BackBufferRect.left = BackBufferRect.left + .Direccion.X * .Estado
        BackBufferRect.Right = BackBufferRect.Right + .Direccion.X * .Estado
        BackBufferRect.top = BackBufferRect.top + .Direccion.Y * .Estado
        BackBufferRect.Bottom = BackBufferRect.Bottom + .Direccion.Y * .Estado
        .Estado = .Estado - MainTimer.PasoEntre(.Inicio, EngineTime) * ANIM_SPEED * 50
    End With
End If 'terremoto

If EfectosDesc(eMareado).Duracion <> 0 Then
'sorpresa :)
Else
'ñaca ñaca ñaca
End If

If EfectosDesc(eAbajoAgua).Estado = 1 Then
    MainColor = RGB(255, 255, 100)
End If
With EfectosDesc(eNoche)
    If .Estado = 1 Then
        If .Inicio <> 0 Then
            'estoy animando, no hago nada
        Else
            'termino la animación, pongo color de una
            .Color = RGB(55, 55, 55)
            MainColor = MediaRGB(.Color, MainColor)
        End If
    Else
        If .Inicio <> 0 Then
            'estoy animando, no hago nada
        Else
            'termino la animación, pongo color de una
            .Color = vbWhite
            MainColor = MediaRGB(.Color, MainColor)
        End If
    End If
    If .Inicio > 0 Then
        If .Estado = 1 Then
            'pasamos a noche
            .AuxDatas = .AuxDatas - 1 * MainTimer.PasoEntre(.Inicio, EngineTime)
            If .AuxDatas < 55 Then
                .AuxDatas = 55
                .Inicio = 0
            End If
        Else
            'pasamos a día
            .AuxDatas = .AuxDatas + (ANIM_SPEED * MainTimer.PasoEntre(.Inicio, EngineTime)) / 3
            If .AuxDatas > 255 Then
                .Inicio = 0
                .AuxDatas = 255
            End If
        End If
        .Color = RGB(.AuxDatas, .AuxDatas, .AuxDatas)
        MainColor = MediaRGB(.Color, MainColor)
    End If
End With

If EfectosDesc(eCiego).Estado = 1 Then
    MainColor = vbBlack
End If


End Sub

Private Sub RenderText(FontIndex As Integer, Texto As String, X As Long, Y As Long, Color As Long, Alpha As Byte)
Dim c As Integer
Dim i As Integer
Dim offsetX As Long
Dim sX As Integer
Dim sY As Integer
Dim Rc As RECT
If FontIndex > NumeroFuentes Then Exit Sub
If Len(Texto) = 0 Then Exit Sub
Translation.X = X + BackBufferRect.left
Translation.Y = Y + BackBufferRect.top
For i = 1 To Len(Texto)
    c = Asc(Mid$(Texto, i, 1))
    sX = c Mod 16
    sY = c \ 16
    Rc.left = sX * 16
    Rc.top = sY * 16
    Rc.Bottom = Rc.top + Fuentes(FontIndex).FontSize
    Rc.Right = Rc.left + Fuentes(FontIndex).FontSize
    Translation.X = Translation.X + offsetX
    D3DSprite.Draw Fuentes(FontIndex).Fuente, Rc, Scaling, Center, Rotation, Translation, Color + m_AlphaValues(Alpha)
    offsetX = Fuentes(FontIndex).FuentesWidth(c)
Next i
End Sub

Private Sub RenderTextEx(FontIndex As Integer, Texto As String, X As Long, Y As Long, ScalingX As Single, ScalingY As Single, Color As Long, Alpha As Byte)
Dim c As Integer
Dim i As Integer
Dim offsetX As Long
Dim sX As Integer
Dim sY As Integer
Dim Rc As RECT
Dim Sca As D3DVECTOR2
If FontIndex > NumeroFuentes Then Exit Sub
If Len(Texto) = 0 Then Exit Sub
Translation.X = X
Translation.Y = Y
Sca.X = ScalingX
Sca.Y = ScalingY
For i = 1 To Len(Texto)
    c = Asc(Mid$(Texto, i, 1))
    sX = c Mod 16
    sY = c \ 16
    Rc.left = sX * 16
    Rc.top = sY * 16
    Rc.Bottom = Rc.top + Fuentes(FontIndex).FontSize
    Rc.Right = Rc.left + Fuentes(FontIndex).FontSize
    Translation.X = Translation.X + offsetX
    D3DSprite.Draw Fuentes(FontIndex).Fuente, Rc, Sca, Center, Rotation, Translation, Color + m_AlphaValues(Alpha)
    offsetX = Fuentes(FontIndex).FuentesWidth(c) * Sca.X
Next i
End Sub

Private Sub RenderObj3yPjs()
Dim i As Integer
Dim j As Integer
Dim Mx As Single
Dim My As Single
Dim TempChar As Char
Dim TmpOffset As D3DVECTOR2
Dim sTmp As Single
Dim Looptexto As Integer
Dim bTmp As Byte

'pjs
For j = MinMapRender.Y To MaxMapRender.Y
    For i = MinMapRender.X To MaxMapRender.X
    If j >= YMinMapSize And j <= YMaxMapSize Then
        If i >= XMinMapSize And i <= XMaxMapSize Then
            'obj
            If MapData(i, j).ObjGrh.grhindex > 0 Then
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                RenderGrh MapData(i, j).ObjGrh, Mx, My, True
            End If
            'layer 3
            If MapData(i, j).Graphic(3).grhindex > 0 Then
                'layer 3
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                RenderGrh MapData(i, j).Graphic(3), Mx, My, True
            End If
        
            'pjs
            If MapData(i, j).CharIndex <> 0 Then
                'If Mouse_Tile_X = i And Mouse_Tile_Y = j Then _
                    'GlowChar MapData(i, j).CharIndex
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                TempChar = CharList(MapData(i, j).CharIndex)
                TmpOffset.X = 0
                TmpOffset.Y = 0
                If TempChar.Heading = 0 Then TempChar.Heading = SOUTH
                If TempChar.seMueve Then
                    sTmp = TileSz - (MainTimer.PasoEntre(TempChar.InicioAnim, EngineTime) * SCREEN_PIXELS_POR_MOVIMIENTO) / MAIN_SPEED
                    If sTmp <= 0 Then
                        sTmp = 0
                        TempChar.seMueve = False
                    Else
                        TmpOffset.X = sTmp * TempChar.MiOffset.X * (-1)
                        TmpOffset.Y = sTmp * TempChar.MiOffset.Y * (-1)
                    End If 'stmp<=0
                    Mx = Mx + TmpOffset.X
                    My = My + TmpOffset.Y
                    'If TempChar.Nombre = "fer" Then _
                    'AddtoRichTextBox frmMain.RecTxt, Mx & " " & My & " " & sTmp, 255, 255, 255, False, False, False, 0
                Else
                    TempChar.MiOffset.X = 0
                    TempChar.MiOffset.Y = 0
                    TempChar.Body.Walk(TempChar.Heading).framecounter = 1
                    TempChar.Body.Walk(TempChar.Heading).started = 0
                    TempChar.Arma.WeaponWalk(TempChar.Heading).framecounter = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).started = 0
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).framecounter = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).started = 0
                End If 'se mueve
                If TmpOffset.X <> 0 Then
                    TempChar.Body.Walk(TempChar.Heading).started = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).started = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).started = 1
                End If
                If TmpOffset.Y <> 0 Then
                    TempChar.Body.Walk(TempChar.Heading).started = 1
                    TempChar.Arma.WeaponWalk(TempChar.Heading).started = 1
                    TempChar.Escudo.ShieldWalk(TempChar.Heading).started = 1
                End If
                If TempChar.Arma.WeaponAttack > 0 Then
                    TempChar.Arma.WeaponAttack = TempChar.Arma.WeaponAttack - 1
                    If TempChar.Arma.WeaponAttack = 0 Then _
                        TempChar.Arma.WeaponWalk(TempChar.Heading).started = 0
                End If
                If TempChar.Head.Head(TempChar.Heading).grhindex <> 0 Then
                    If Not CharList(MapData(i, j).CharIndex).invisible Then
                        RenderGrh TempChar.Body.Walk(TempChar.Heading), Mx - PJ_X_OFFSET, My, True
                        If TempChar.Casco.Head(TempChar.Heading).grhindex <> 0 Then
                            RenderGrh TempChar.Head.Head(TempChar.Heading), Mx + TempChar.Body.HeadOffset.X - PJ_X_OFFSET_HEAD, My + TempChar.Body.HeadOffset.Y, True
                            RenderGrh TempChar.Casco.Head(TempChar.Heading), Mx + TempChar.Body.HeadOffset.X - PJ_X_OFFSET_HEAD, My + TempChar.Body.HeadOffset.Y, True
                        Else
                            RenderGrh TempChar.Head.Head(TempChar.Heading), Mx + TempChar.Body.HeadOffset.X - PJ_X_OFFSET_HEAD, My + TempChar.Body.HeadOffset.Y, True
                        End If
                        If TempChar.Arma.WeaponWalk(TempChar.Heading).grhindex <> 0 Then _
                            RenderGrh TempChar.Arma.WeaponWalk(TempChar.Heading), Mx - PJ_X_OFFSET, My, True
                        If TempChar.Escudo.ShieldWalk(TempChar.Heading).grhindex <> 0 Then _
                            RenderGrh TempChar.Escudo.ShieldWalk(TempChar.Heading), Mx - PJ_X_OFFSET, My, True
                    End If 'invisible
                Else 'head <> 0
                    RenderGrh TempChar.Body.Walk(TempChar.Heading), Mx - PJ_X_OFFSET, My, True
                End If 'head <>0
                With TempChar.Dialogos
                    If Len(.Texto) > 0 Then
                        If .TiempoInicial > m_DialogoTiempo And _
                            MainTimer.PasoEntre(.TiempoInicial, EngineTime) < VIDADIALOGOS Then
                            If Not .Iniciado Then IniciarDialogos TempChar.Dialogos
                            For Looptexto = .Lineas To 1 Step -1
                                RenderTextEx 1, .TextoFormateado(.Lineas - Looptexto + 1), CLng(Mx - 32), CLng(My - 32 * Looptexto), 1, 1, .Color, 255
                            Next Looptexto
                        Else
                            .Color = 0
                            .TiempoInicial = 0
                            .Iniciado = False
                            .Texto = vbNullString
                        End If
                    End If
                End With
                
                CharList(MapData(i, j).CharIndex) = TempChar
                If CharList(MapData(i, j).CharIndex).fx <> 0 Then
                    RenderGrh FxData(TempChar.fx).fx, Mx, My, True
                End If
            End If 'hay char en i,j
        End If 'ok en x
    End If 'ok en y
    Next i
Next j

End Sub

Private Sub RenderLayerFxy4()
Dim i As Integer
Dim j As Integer
Dim Mx As Single
Dim My As Single
For j = MinMapRender.Y To MaxMapRender.Y
    For i = MinMapRender.X To MaxMapRender.X
    If j >= YMinMapSize And j <= YMaxMapSize Then
        If i >= XMinMapSize And i <= XMaxMapSize Then
            If MapData(i, j).fx <> 0 Then
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                RenderGrh FxData(MapData(i, j).fx).fx, Mx, My, True
            End If 'fx
        End If
    End If
    Next i
Next j
'layer 4
If Not bTecho Then
For j = MinMapRender.Y To MaxMapRender.Y
    For i = MinMapRender.X To MaxMapRender.X
    If j >= YMinMapSize And j <= YMaxMapSize Then
        If i >= XMinMapSize And i <= XMaxMapSize Then
            If MapData(i, j).Graphic(4).grhindex > 0 Then
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                RenderGrh MapData(i, j).Graphic(4), Mx, My, True
            End If
        End If
    End If
    Next i
Next j
End If 'techo
End Sub

'Private Sub RenderGrhEx2(Grh As Integer, PixelX As Single, PixelY As Single, RotAngle As Single, RotCenterX As Single, RotCenterY As Single, ScalingX As Single, ScalingY As Single, Color As Long, Alpha As Byte)
'Dim Mat As DxVBLibA.D3DMATRIX
'Dim MatTmp As DxVBLibA.D3DMATRIX
'Dim MatTmp2 As DxVBLibA.D3DMATRIX
'Dim MatSca As DxVBLibA.D3DMATRIX
'Dim MatRot As DxVBLibA.D3DMATRIX
'Dim MatTrans As DxVBLibA.D3DMATRIX
'With SourceRect
'    .Left = GrhData(Grh).sX
'    .Top = GrhData(Grh).sY
'    .Right = .Left + GrhData(Grh).pixelWidth
'    .Bottom = .Top + GrhData(Grh).pixelHeight
'End With
'D3DXMatrixScaling MatSca, ScalingX, ScalingY, 1
'D3DXMatrixRotationZ MatRot, RotAngle
'D3DXMatrixTranslation MatTrans, (SourceRect.Right - SourceRect.Left) \ 2, (SourceRect.Bottom - SourceRect.Top) \ 2, 0
'D3DXMatrixMultiply MatTmp, MatSca, MatRot
'D3DXMatrixMultiply MatTmp2, MatTmp, MatRot
'D3DXMatrixMultiply Mat, MatTmp, MatTmp2''

'D3DSprite.DrawTransform Texturas.GetBMP(GrhData(Grh).FileNum), SourceRect, Mat, Color + m_AlphaValues(Alpha)

'End Sub

Private Sub RenderGrhExRotCentrado(Grh As Integer, PixelX As Single, PixelY As Single, RotAngle As Single, ScalingX As Single, ScalingY As Single, Color As Long, Alpha As Byte)
Dim Scaling As D3DVECTOR2
Dim Trans As D3DVECTOR2
Dim RotationCenter As D3DVECTOR2
Dim Rotation As Single
With SourceRect
    .left = GrhData(Grh).sX
    .top = GrhData(Grh).sY
    .Right = .left + GrhData(Grh).pixelWidth
    .Bottom = .top + GrhData(Grh).pixelHeight
End With

Trans.X = PixelX
Trans.Y = PixelY
Rotation = RotAngle
RotationCenter.X = ((SourceRect.Right - SourceRect.left) / 2)  'RotCenterX
RotationCenter.Y = ((SourceRect.Bottom - SourceRect.top) / 2) 'RotCenterY
Scaling.X = ScalingX
Scaling.Y = ScalingY
D3DSprite.Draw Texturas.GetBMP(GrhData(Grh).FileNum), SourceRect, Scaling, RotationCenter, Rotation, Trans, Color + m_AlphaValues(Alpha)

End Sub
Private Sub RenderGrhEx(Grh As Integer, PixelX As Single, PixelY As Single, RotAngle As Single, RotCenterX As Single, RotCenterY As Single, ScalingX As Single, ScalingY As Single, Color As Long, Alpha As Byte)
Dim Scaling As D3DVECTOR2
Dim Trans As D3DVECTOR2
Dim RotationCenter As D3DVECTOR2
Dim Rotation As Single
With SourceRect
    .left = GrhData(Grh).sX
    .top = GrhData(Grh).sY
    .Right = .left + GrhData(Grh).pixelWidth
    .Bottom = .top + GrhData(Grh).pixelHeight
End With

Trans.X = PixelX
Trans.Y = PixelY
Rotation = RotAngle
RotationCenter.X = RotCenterX
RotationCenter.Y = RotCenterY
Scaling.X = ScalingX
Scaling.Y = ScalingY
D3DSprite.Draw Texturas.GetBMP(GrhData(Grh).FileNum), SourceRect, Scaling, RotationCenter, Rotation, Trans, Color + m_AlphaValues(Alpha)
End Sub


Private Sub RenderGrh(Grh As Grh, X As Single, Y As Single, bCenter As Boolean, Anim As Boolean)
Dim iGrhIndex As Integer

If Grh.grhindex = 0 Then Exit Sub
If Anim Then
    With Grh
        If (.started = 1) Then
            If (.speedcounter > 0) Then
                .speedcounter = .speedcounter - 1
                If (.speedcounter = 0) Then
                    .speedcounter = GrhData(.grhindex).speed
                    .framecounter = .framecounter + 1
                    If (.framecounter > GrhData(.grhindex).NumFrames) Then _
                        .framecounter = 1
                End If
            End If
        End If
    End With
End If
iGrhIndex = GrhData(Grh.grhindex).Frames(Grh.framecounter)
If GrhData(iGrhIndex).FileNum = 0 Then Exit Sub
Translation.X = X
Translation.Y = Y
If bCenter Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        Translation.X = Translation.X - (Int(GrhData(iGrhIndex).TileWidth) - 1) * TileSz \ 2 'menos por menos mas ^^
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Translation.Y = Translation.Y - Int(GrhData(iGrhIndex).TileHeight * TileSz) + TileSz
    End If
End If
With SourceRect
    .left = GrhData(iGrhIndex).sX
    .top = GrhData(iGrhIndex).sY
    .Right = .left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .top + GrhData(iGrhIndex).pixelHeight
End With
Center.X = Translation.X + (SourceRect.Right - SourceRect.left) / 2
Center.Y = Translation.Y + (SourceRect.Bottom - SourceRect.top) / 2

D3DSprite.Draw Texturas.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, Scaling, Center, Rotation, Translation, MainColor + m_AlphaValues(255)
End Sub


Private Function MediaRGB(Color1 As Long, Color2 As Long) As Long
Dim r(1 To 2) As Integer
Dim g(1 To 2) As Integer
Dim b(1 To 2) As Integer
RGBFromLong r(1), g(1), b(1), Color1
RGBFromLong r(2), g(2), b(2), Color2
MediaRGB = RGB((r(1) + r(2)) \ 2, (g(1) + g(2)) \ 2, (b(1) + b(2)) \ 2)
End Function

Private Sub RenderObj()
Dim i As Integer
Dim j As Integer
Dim Mx As Single
Dim My As Single
For j = MinMapRender.Y To MaxMapRender.Y
    For i = MinMapRender.X To MaxMapRender.X
    If j >= YMinMapSize And j <= YMaxMapSize Then
        If i >= XMinMapSize And i <= XMaxMapSize Then
            If MapData(i, j).ObjGrh.grhindex > 0 Then
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                'RenderGrh MapData(i, j).ObjGrh, Mx, My, True, True
            End If
        End If
    End If
    Next i
Next j
End Sub
Private Sub RenderLayers1y2()
Dim i As Integer
Dim j As Integer
Dim Mx As Single
Dim My As Single
'layer 1
For j = MinMapRender.Y To MaxMapRender.Y
    For i = MinMapRender.X To MaxMapRender.X
    If j >= YMinMapSize And j <= YMaxMapSize Then
        If i >= XMinMapSize And i <= XMaxMapSize Then
            If MapData(i, j).Graphic(1).grhindex > 0 Then
                'layer 1
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                RenderGrh MapData(i, j).Graphic(1), Mx, My, True
            End If
            'layer 2
            If MapData(i, j).Graphic(2).grhindex > 0 Then
                Mx = TileSz * (i - MinMapRender.X)
                My = TileSz * (j - MinMapRender.Y)
                RenderGrh MapData(i, j).Graphic(2), Mx, My, True
            End If
        End If
    End If
    Next i
Next j
End Sub

Public Sub Mouse_OnMap(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_Pixel_X = X \ Screen.TwipsPerPixelX
Mouse_Pixel_Y = Y \ Screen.TwipsPerPixelY
Mouse_Tile_X = Mouse_Pixel_X \ TileSz + UserPos.X - WindowTilesX \ 2
Mouse_Tile_Y = Mouse_Pixel_Y \ TileSz + UserPos.Y - WindowTilesY \ 2
End Sub
Public Sub Mouse_OnMapUp(button As Integer, Shift As Integer, X As Single, Y As Single)
'hacemos algo en mouseup?
End Sub

Public Sub Mouse_DblClick()
If Not frmForo.Visible Then
    SendData "RC" & Mouse_Tile_X & "," & Mouse_Tile_Y
End If
End Sub
Public Sub Mouse_Click()
If Cartel Then Cartel = False

If Not Comerciando Then
    If UsingSkill = 0 Then
        SendData "LC" & Mouse_Tile_X & "," & Mouse_Tile_Y
    Else
        frmMain.MousePointer = vbDefault
        If (UsingSkill = Magia Or UsingSkill = Proyectiles Or UsingSkill = Arrojadizas) And UserCanAttack = 0 Then Exit Sub
        SendData "WLC" & Mouse_Tile_X & "," & Mouse_Tile_Y & "," & UsingSkill
        If UsingSkill = Magia Or UsingSkill = Proyectiles Or UsingSkill = Arrojadizas Then UserCanAttack = 0
        UsingSkill = 0
    End If
End If
If Not CBool((GetKeyState(vbKeyShift) Or 1) Mod -127) Then
    Call SendData("/TP" & Mouse_Tile_X & ";" & Mouse_Tile_Y)
End If
End Sub
Public Sub Mouse_OnMapDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Mouse_OnMap button, Shift, X, Y 'actualiza valores del mouse
AddtoRichTextBox frmMain.RecTxt, "MouseDown: B" & button & " S" & Shift & " (" & Mouse_Tile_X & ", " & Mouse_Tile_Y & ")", 255, 0, 0, 0, 0, False
End Sub




Public Sub MoveScreen_iao(eDireccion As eDirecciones)
Dim tX As Integer
Dim tY As Integer
PuedoMover = False
If eDireccion = eArriba Then
    If UserPos.Y < MinYBorder Then Exit Sub
    ScreenAvance.Y = -1
    ScreenAvance.X = 0
ElseIf eDireccion = eAbajo Then
    If UserPos.Y > MaxYBorder Then Exit Sub
    ScreenAvance.Y = 1
    ScreenAvance.X = 0
ElseIf eDireccion = eIzquierda Then
    If UserPos.X < MinXBorder Then Exit Sub
    ScreenAvance.Y = 0
    ScreenAvance.X = -1
Else
    If UserPos.X > MaxXBorder Then Exit Sub
    ScreenAvance.Y = 0
    ScreenAvance.X = 1
End If
UserPos.X = UserPos.X + ScreenAvance.X
UserPos.Y = UserPos.Y + ScreenAvance.Y
MoviendoScreen = 1
MainTimer.StartTimer

frmMain.PosX.Caption = UserPos.X
frmMain.PosY.Caption = UserPos.Y
    
If frmMain.UltPos = 0 Then
    frmMain.Label2(frmMain.UltPos).Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y
End If
    
bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
        MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
        MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
End Sub

Public Sub MoveCharByPos_iao(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(X, Y).CharIndex = 0
MapData(nX, nY).CharIndex = CharIndex

With CharList(CharIndex)
    .Pos.X = nX
    .Pos.Y = nY

    .MiOffset.X = addX
    .MiOffset.Y = addY
    .InicioAnim = EngineTime
    .Heading = nHeading
    .seMueve = True
End With
End Sub

Public Sub MoveCharbyHead_iao(CharIndex As Integer, nHeading As Byte)
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading
    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select
nX = X + addX
nY = Y + addY

'MoveCharByPos CharIndex, nX, nY
'Exit Sub
With CharList(CharIndex)
    MapData(nX, nY).CharIndex = CharIndex
    .Pos.X = nX
    .Pos.Y = nY
    MapData(X, Y).CharIndex = 0

    .MiOffset.X = addX
    .MiOffset.Y = addY
    .InicioAnim = EngineTime
    .Heading = nHeading
    .seMueve = True
End With
If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
End Sub

Private Sub RGBFromLong(outRed As Integer, outGreen As Integer, outBlue As Integer, ByVal inColor As Long)
inColor = inColor And &HFFFFFF 'saco alpha si lo tiene
outRed = inColor And &HFF
inColor = inColor \ 256
outGreen = inColor And &HFF
inColor = inColor \ 256
outBlue = inColor
End Sub


Private Sub IniciarDialogos(Dialogo As tDialogo)
Dim i As Long
Dim ProximaLinea As Boolean
Dim Poner As Boolean
If Len(Dialogo.Texto) = 0 Then Exit Sub
Dialogo.Lineas = 1
ReDim Dialogo.TextoFormateado(1 To 1)
For i = 1 To Len(Dialogo.Texto)
    Poner = True
    If ProximaLinea And Mid$(Dialogo.Texto, i, 1) = " " Then
        Dialogo.Lineas = Dialogo.Lineas + 1
        ReDim Preserve Dialogo.TextoFormateado(1 To Dialogo.Lineas)
        Poner = False
        ProximaLinea = False
    End If
    If Poner Then _
        Dialogo.TextoFormateado(Dialogo.Lineas) = Dialogo.TextoFormateado(Dialogo.Lineas) & Mid$(Dialogo.Texto, i, 1)
    If Len(Dialogo.TextoFormateado(Dialogo.Lineas)) > 13 Then _
        ProximaLinea = True
Next i
End Sub
Public Sub BorrarDialogos()
m_DialogoTiempo = EngineTime
End Sub

Public Sub CrearDialogo(sData As String, UserIndex As Integer, Color As Long)
With CharList(UserIndex).Dialogos
    .Iniciado = False
    .Color = Color
    .TiempoInicial = EngineTime
    .Texto = sData
End With
End Sub

Public Sub QuitarDialogo(UserIndex As Integer)
If UserIndex = 0 Then Exit Sub
With CharList(UserIndex).Dialogos
    .Color = 0
    .Iniciado = False
    .Texto = vbNullString
    .TiempoInicial = 0
End With
End Sub
'D3DColorARGB

Public Sub RenderInv()
Dim i As Integer
Dim X As Single
Dim Y As Single

D3DDevice.Clear 1, RecInv, D3DCLEAR_TARGET, &HFF000000, 1#, 0
D3DDevice.BeginScene
D3DSprite.Begin
For i = 1 To MAX_INVENTORY_SLOTS
    If UserInventory(i).Amount > 0 Then
        RenderInvItem i
        X = ((i - 1) Mod XCantItems) * TileSz
        Y = (((i - 1) \ XCantItems) + 0.75) * TileSz
        RenderTextEx 1, Str(UserInventory(i).Amount), _
        CLng(X), _
        CLng(Y), _
        0.75, 0.75, vbWhite, 255
    End If
Next i
D3DSprite.End
D3DDevice.EndScene
D3DDevice.Present RecInv, ByVal 0, frmMain.picInv.hwnd, ByVal 0

End Sub
Private Sub RenderInvItem(InvIndex As Integer, Optional InitDevice As Boolean = False)
If Rendering Then Exit Sub
If InitDevice Then
    D3DDevice.Clear 1, RecInv, D3DCLEAR_TARGET, &HFF000000, 1#, 0
    D3DDevice.BeginScene
    D3DSprite.Begin
End If
Dim ItemX As Single
Dim ItemY As Single
ItemX = ((InvIndex - 1) Mod XCantItems)
ItemY = ((InvIndex - 1) \ XCantItems)

If ItemElegido = InvIndex Then _
    RenderGrhEx 2, ItemX * TileSz, ItemY * TileSz, 0, 0, 0, 1, 1, &HFFFFFFFF, &HFF
'pongo el cuadradito :)

If UserInventory(InvIndex).Equipped Then
    RenderGrhEx UserInventory(InvIndex).grhindex, ItemX * TileSz, ItemY * TileSz, 0, 0, 0, 1, 1, &HFFFFFF, &HFF
Else
    RenderGrhEx UserInventory(InvIndex).grhindex, ItemX * TileSz, ItemY * TileSz, 0, 0, 0, 1, 1, &HFFFFFF, &HA0
End If

If InitDevice Then
    D3DSprite.End
    D3DDevice.EndScene
    D3DDevice.Present RecInv, ByVal 0, frmMain.picInv.hwnd, ByVal 0
End If

End Sub

Public Sub InvClick(X As Single, Y As Single, button As Integer)
'tira item con boton derecho?
ItemElegido = (X \ TileSz) + 1 + (Y \ TileSz) * XCantItems
If ItemElegido > MAX_INVENTORY_SLOTS Or _
    ItemElegido < 1 Then
    ItemElegido = 0
Else
    If UserInventory(ItemElegido).grhindex > 0 Then _
        RenderInv
End If
End Sub

Public Sub EfectoTerremoto(Duracion As eDuracionEfecto, Potencia As Byte)
With EfectosDesc(eTerremoto)
    .Inicio = EngineTime
    .Duracion = Duracion * 50
    .AuxDataL = Potencia
End With
End Sub

Public Sub EfectoNocheDía()
EfectosDesc(eNoche).Inicio = EngineTime
If EfectosDesc(eNoche).Estado = 1 Then
    'es de noche, pasamos a día
    EfectosDesc(eNoche).Estado = 0
    EfectosDesc(eNoche).AuxDatas = 55
Else
    'es de día, pasamos a noche
    EfectosDesc(eNoche).Estado = 1
    EfectosDesc(eNoche).AuxDatas = 255
End If
End Sub

Public Sub EfectoAbajoAgua()
If EfectosDesc(eAbajoAgua).Estado = 1 Then
    EfectosDesc(eAbajoAgua).Estado = 0
Else
    EfectosDesc(eAbajoAgua).Estado = 1
End If
End Sub

Public Sub EfectoCiego()
If EfectosDesc(eCiego).Estado = 1 Then
    EfectosDesc(eCiego).Estado = 0
Else
    EfectosDesc(eCiego).Estado = 1
End If
End Sub

Public Sub SaveScreenShot(FN As String)
'no usar, muy lento

Dim oSurface As Direct3DSurface8
Dim SrcPalette As DxVBLibA.PALETTEENTRY
Dim SrcRect As RECT
Dim DispMode As D3DDISPLAYMODE
Dim SSRect As DxVBLibA.RECT

D3DDevice.GetDisplayMode DispMode
Set oSurface = D3DDevice.CreateImageSurface(DispMode.width, _
            DispMode.height, _
            D3DFMT_A8R8G8B8)

D3DDevice.GetFrontBuffer oSurface

With SSRect
    .left = 0
    .Right = 800 'DispMode.Width
    .top = 0
    .Bottom = 600 'DispMode.Height
End With
D3dx.SaveSurfaceToFile FN, _
            D3DXIFF_BMP, _
            oSurface, _
            SrcPalette, _
            SSRect
Set oSurface = Nothing
End Sub


Public Sub Render2()
UpdateFPS

Scaling.X = 1
Scaling.Y = 1
Rotation = 0
Center.X = 0
Center.Y = 0
Translation.X = 0
Translation.Y = 0
With BackBufferRect
    .left = TileSz * BufferTiles
    .top = TileSz * BufferTiles
    .Right = .left + WindowTilesX * TileSz
    .Bottom = .top + WindowTilesY * TileSz
End With
MainColor = m_AlphaValues(255) + &HFFFFFF   ' + MapBaseColor en lugar de &hffffff

EngineTime = MainTimer.TiempoActual

MinMapRender.X = UserPos.X - WindowTilesX \ 2 - BufferTiles
MinMapRender.Y = UserPos.Y - WindowTilesY \ 2 - BufferTiles
MaxMapRender.X = UserPos.X + WindowTilesX \ 2 + BufferTiles
MaxMapRender.Y = UserPos.Y + WindowTilesY \ 2 + BufferTiles

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HFF000000, 0#, 0
D3DDevice.BeginScene
D3DSprite.Begin

RenderBottom
RenderMedio
RenderTop

RenderText 1, "FPS: " & m_FPS, 0, 0, 255, 255
D3DSprite.End
D3DDevice.EndScene
D3DDevice.Present BackBufferRect, ScreenRect, 0, ByVal 0

End Sub

Private Sub RenderBottom()
Dim i As Integer
Dim j As Integer
Dim Mx As Single
Dim My As Single
'layer 1

My = 0
For j = MinMapRender.Y To MaxMapRender.Y
    Mx = 0
    For i = MinMapRender.X To MaxMapRender.X
    If (j >= YMinMapSize And j <= YMaxMapSize) And _
        (i >= XMinMapSize And i <= XMaxMapSize) Then
            'layer 1
            If MapData(i, j).Graphic(1).grhindex > 0 Then _
                RenderGrh MapData(i, j).Graphic(1), Mx, My, True, True
            'layer 2
            If Not RenderMod.bNoCostas Then
                If MapData(i, j).Graphic(2).grhindex > 0 Then _
                    RenderGrh MapData(i, j).Graphic(2), Mx, My, True, False
            End If
    End If
    Mx = Mx + TileSz
    Next i
    My = My + TileSz
Next j
End Sub

Private Sub RenderMedio()
Dim i As Integer
Dim j As Integer
Dim Mx As Single
Dim My As Single
Dim TempChar As Char
Dim PixelOffsetXTemp As Integer
Dim PixelOffsetYTemp As Integer


My = 0
For j = MinMapRender.Y To MaxMapRender.Y
    Mx = 0
    For i = MinMapRender.X To MaxMapRender.X
    If (j >= YMinMapSize And j <= YMaxMapSize) And _
        (i >= XMinMapSize And i <= XMaxMapSize) Then
        'objetos!
            If MapData(i, j).ObjGrh.grhindex <> 0 Then _
                RenderGrh MapData(i, j).ObjGrh, Mx, My, True, True
            
            'Char layer ************************************
            If MapData(X, Y).CharIndex <> 0 Then
                TempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY

                Moved = 0
            
            
            
    End If
    Mx = Mx + TileSz
    Next i
    My = My + TileSz
Next j
    

End Sub

Private Sub RenderTop()

End Sub



Sub MoveScreen(Heading As Byte)
'******************************************
'Starts the screen moving in a direction
'******************************************
Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer
'Figure out which way to move
Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select

'Fill temp pos
tX = UserPos.X + X
tY = UserPos.Y + Y

If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1

    frmMain.PosX.Caption = UserPos.X
    frmMain.PosY.Caption = UserPos.Y
    
    If frmMain.UltPos = 0 Then
        frmMain.Label2(frmMain.UltPos).Caption = "Posición: " & UserMap & ", " & UserPos.X & ", " & UserPos.Y
    End If
    
    bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
Exit Sub

Stop

    Call DoFogataFx
End If

End Sub


Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte



X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex


CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TileSz * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TileSz * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading
End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TileSz * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TileSz * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)

End Sub

