Attribute VB_Name = "modGeneral"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Const PI As Single = 3.14159265358979

Public bO As Integer
Public bK As Long
Public bRK As Long

Public iplst As String
Public banners As String

Public bInvMod     As Boolean  'El inventario se modificó?

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public lFrameTimer As Long

Public bFPS As Boolean
Public bInvis As Boolean

'Sinuhe_D3d / Barrin DirectSound
Public IAO_TE As New clsTileEngineX
Public IAO_SE As New clsIAO_SE

Public Function DirGraficos() As String
DirGraficos = App.Path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
DirSound = App.Path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
DirMidi = App.Path & "\" & Config_Inicio.DirMusica & "\"
End Function
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

Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
'*****************************************************************
'Author: Aaron Perkins
'Find a Random number between a range
'*****************************************************************
    Randomize Timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()

On Error Resume Next

Dim LoopC As Integer
Dim arch As String
arch = App.Path & "\init\" & "armas.dat"
DoEvents

NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))

ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData

For LoopC = 1 To NumWeaponAnims
    InitGrh WeaponAnimData(LoopC).WeaponWalk(NORTH), Val(GetVar(arch, "ARMA" & LoopC, "Dir1")), 0
    InitGrh WeaponAnimData(LoopC).WeaponWalk(EAST), Val(GetVar(arch, "ARMA" & LoopC, "Dir2")), 0
    InitGrh WeaponAnimData(LoopC).WeaponWalk(SOUTH), Val(GetVar(arch, "ARMA" & LoopC, "Dir3")), 0
    InitGrh WeaponAnimData(LoopC).WeaponWalk(WEST), Val(GetVar(arch, "ARMA" & LoopC, "Dir4")), 0
Next LoopC

End Sub

Sub CargarAnimEscudos()

On Error Resume Next

Dim LoopC As Integer
Dim arch As String
arch = App.Path & "\init\" & "escudos.dat"
DoEvents

NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))

ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData

For LoopC = 1 To NumEscudosAnims
    InitGrh ShieldAnimData(LoopC).ShieldWalk(NORTH), Val(GetVar(arch, "ESC" & LoopC, "Dir1")), 0
    InitGrh ShieldAnimData(LoopC).ShieldWalk(EAST), Val(GetVar(arch, "ESC" & LoopC, "Dir2")), 0
    InitGrh ShieldAnimData(LoopC).ShieldWalk(SOUTH), Val(GetVar(arch, "ESC" & LoopC, "Dir3")), 0
    InitGrh ShieldAnimData(LoopC).ShieldWalk(WEST), Val(GetVar(arch, "ESC" & LoopC, "Dir4")), 0
Next LoopC

End Sub

Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long

    Dim StreamFile As String
    StreamFile = App.Path & "\init\" & "particulas.dat"

    TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = GetVar(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = GetVar(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = GetVar(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = GetVar(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = GetVar(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = GetVar(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = GetVar(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = GetVar(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = GetVar(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = GetVar(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = GetVar(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = GetVar(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = GetVar(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = GetVar(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = GetVar(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = GetVar(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = GetVar(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = GetVar(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = GetVar(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = GetVar(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = GetVar(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = GetVar(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = GetVar(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = GetVar(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = GetVar(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = GetVar(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = GetVar(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = GetVar(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).speed = Val(GetVar(StreamFile, Val(LoopC), "Speed"))
        
        StreamData(LoopC).NumGrhs = GetVar(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = GetVar(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = ReadField(Str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).r = ReadField(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).g = ReadField(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).b = ReadField(3, TempSet, 44)
        Next ColorSet
        
    Next LoopC
End Sub

Sub Addtostatus(RichTextBox As RichTextBox, Text As String, red As Byte, green As Byte, blue As Byte, bold As Byte, italic As Byte)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************

frmCargando.Status.SelStart = Len(RichTextBox.Text)
frmCargando.Status.SelLength = 0
frmCargando.Status.SelColor = RGB(red, green, blue)

If bold Then
    frmCargando.Status.SelBold = True
Else
    frmCargando.Status.SelBold = False
End If

If italic Then
    frmCargando.Status.SelItalic = True
Else
    frmCargando.Status.SelItalic = False
End If

frmCargando.Status.SelText = Chr(13) & Chr(10) & Text

End Sub

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
Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************

On Error Resume Next

Dim LoopC As Integer

For LoopC = 1 To LastChar
    If CharList(LoopC).active = 1 Then
        MapData(CharList(LoopC).Pos.X, CharList(LoopC).Pos.Y).CharIndex = LoopC
    End If
Next LoopC

End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function



Function CheckUserData(checkemail As Boolean) As Boolean
'Validamos los datos del user
Dim LoopC As Integer
Dim CharAscii As Integer

If checkemail Then
 If UserEmail = "" Then
    MsgBox ("Direccion de email invalida")
    Exit Function
 End If
End If

If UserPassword = "" Then
    MsgBox ("Ingrese un password.")
    Exit Function
End If

For LoopC = 1 To Len(UserPassword)
    CharAscii = Asc(Mid$(UserPassword, LoopC, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Password invalido.")
        Exit Function
    End If
Next LoopC

If UserName = "" Then
    MsgBox ("Nombre invalido.")
    Exit Function
End If

If Len(UserName) > 30 Then
    MsgBox ("El nombre debe tener menos de 30 letras.")
    Exit Function
End If

For LoopC = 1 To Len(UserName)

    CharAscii = Asc(Mid$(UserName, LoopC, 1))
    If LegalCharacter(CharAscii) = False Then
        MsgBox ("Nombre invalido.")
        Exit Function
    End If
    
Next LoopC


CheckUserData = True

End Function
Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    For Each mifrm In Forms
        Unload mifrm
    Next
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

'Set Connected
Connected = True

'Unload the connect form
Unload frmConnect

'frmMain.Label8.Caption = UserName
'frmMain.Caption = "ImperiumAO 1.2 - " & UserName
frmMain.lblNick = UserName
'Load main form
frmMain.Visible = True

End Sub
Sub CargarTip()

Dim n As Integer
n = General_Random_Number(1, UBound(Tips))
If n > UBound(Tips) Then n = UBound(Tips)
frmtip.tip.Caption = Tips(n)

End Sub

Sub MoveNorth()
If Cartel Then Cartel = False

If LegalPos(UserPos.X, UserPos.Y - 1) Then
    If IAO_TE.Engine_View_Move(NORTH) Then
        IAO_TE.Char_Move UserCharIndex, NORTH
        Call SendData("M" & NORTH)
        Call DoPasosFx(UserCharIndex)
        DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> NORTH Then
        Call SendData("CHEA" & NORTH)
    End If
End If
End Sub

Sub MoveEast()

If LegalPos(UserPos.X + 1, UserPos.Y) Then
    If IAO_TE.Engine_View_Move(EAST) Then
        IAO_TE.Char_Move UserCharIndex, EAST
        Call SendData("M" & EAST)
        Call DoPasosFx(UserCharIndex)
        Call DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> EAST Then
            Call SendData("CHEA" & EAST)
    End If
End If

End Sub

Private Sub MoveSouth()

If LegalPos(UserPos.X, UserPos.Y + 1) Then
    If IAO_TE.Engine_View_Move(SOUTH) Then
        IAO_TE.Char_Move UserCharIndex, SOUTH
        Call SendData("M" & SOUTH)
        Call DoPasosFx(UserCharIndex)
        DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> SOUTH Then
        Call SendData("CHEA" & SOUTH)
    End If
End If
End Sub

Sub MoveWest()

If LegalPos(UserPos.X - 1, UserPos.Y) Then
    If IAO_TE.Engine_View_Move(WEST) Then
        IAO_TE.Char_Move UserCharIndex, WEST
        Call SendData("M" & WEST)
        Call DoPasosFx(UserCharIndex)
        DoFogataFx
    End If
Else
    If CharList(UserCharIndex).Heading <> WEST Then
            Call SendData("CHEA" & WEST)
    End If
End If

End Sub

Sub MoveUserChar(ByVal Heading As Byte)

If Cartel Then Cartel = False

If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
    Select Case Heading
        Case NORTH
            Call MoveNorth
        Case EAST
            Call MoveEast
        Case WEST
            Call MoveWest
        Case SOUTH
            Call MoveSouth
    End Select
End If

End Sub

Sub RandomMove()

Dim j As Integer

j = General_Random_Number(1, 4)

Select Case j
    Case 1
        Call MoveEast
    Case 2
        Call MoveNorth
    Case 3
        Call MoveWest
    Case 4
        Call MoveSouth
End Select

End Sub

Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************

Dim LoopC As Integer

LoopC = 1
Do While CharList(LoopC).active
    LoopC = LoopC + 1
Loop

NextOpenChar = LoopC

End Function

Public Function DirMapas() As String
DirMapas = App.Path & "\" & Config_Inicio.DirMapas & "\"
End Function

Sub SwitchMap(Map As Integer)

Call IAO_TE.Map_Load_From_File(Map)
CurMap = Map
frmMain.Label2(0).Caption = UserLocation(Map)

End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = Pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = Pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean
If Dir(file, FileType) = "" Then
    FileExist = False
Else
    FileExist = True
End If
End Function

Sub WriteClientVer()

Dim hfile As Integer
    
hfile = FreeFile()
Open App.Path & "\init\Ver.bin" For Binary Access Write As #hfile
Put #hfile, , CLng(777)
Put #hfile, , CLng(777)
Put #hfile, , CLng(777)

Put #hfile, , CInt(App.Major)
Put #hfile, , CInt(App.Minor)
Put #hfile, , CInt(App.Revision)

Close #hfile

End Sub

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
i = 1

Do While (ReadField(i, RawServersList, Asc(";")) <> "")
    i = i + 1
    Cont = Cont + 1
Loop

ReDim ServersLst(1 To Cont) As tServerInfo

For i = 1 To Cont
    Dim cur$
    cur$ = ReadField(i, RawServersList, Asc(";"))
    ServersLst(i).IP = ReadField(1, cur$, Asc(":"))
    ServersLst(i).Puerto = Val(ReadField(2, cur$, Asc(":")))
    ServersLst(i).desc = ReadField(3, cur$, Asc(":"))
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

On Error Resume Next

Call WriteClientVer

If App.PrevInstance Then
    Call MsgBox("¡ImperiumAO ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
    End
End If

Dim UltTick As Long, EstTick As Long
Dim Timers(1 To 5) As Long

ChDrive App.Path
ChDir App.Path

'Cargamos el archivo de configuracion inicial

If FileExist(App.Path & "\init\Inicio.con", vbNormal) Then
    Config_Inicio = LeerGameIni()
End If

If FileExist(App.Path & "\init\ao.dat", vbNormal) Then
    Open App.Path & "\init\ao.dat" For Binary As #53
        Get #53, , RenderMod
    Close #53

    Musica = IIf(RenderMod.bNoMusic = 1, 1, 0)
    fx = IIf(RenderMod.bNoSound = 1, 1, 0)

    Select Case RenderMod.iImageSize
        Case 4
            RenderMod.iImageSize = 0
        Case 3
            RenderMod.iImageSize = 1
        Case 2
            RenderMod.iImageSize = 2
        Case 1
            RenderMod.iImageSize = 3
        Case 0
            RenderMod.iImageSize = 4
    End Select
End If

tipf = Config_Inicio.tip

Call LoadImpAoInit
Call LoadFontTypes

frmCargando.Show
frmCargando.Refresh

UserParalizado = False

frmConnect.version = "ImperiumAO! 1.2"
AddtoRichTextBox frmCargando.Status, "Buscando servidores...", 255, 255, 255, 1, 0, 1
frmCargando.Refresh

'frmMain.Inet1.URL = "http://serverlist.imperiumao.com.ar"
RawServersList = "server.imperiumao.com.ar:7666:Barrin's Home" 'frmMain.Inet1.OpenURL
'server.imperiumao.com.ar
If RawServersList = "" Then
    ServersRecibidos = True
    RawServersList = "server.imperiumao.com.ar:7666:Argentina (server.imperiumao.com.ar);secundario.imperiumao.com.ar:7666:Secundario (secundario.imperiumao.com.ar);españa.imperiumao.com.ar:7666:España (españa.imperiumao.com.ar);localhost:7666:Local (localhost)"
    AddtoRichTextBox frmCargando.Status, "No se encontró la lista", 255, , , 1
    frmCargando.Refresh
Else
    ServersRecibidos = True
    AddtoRichTextBox frmCargando.Status, "Encontrados", , , , 1
    frmCargando.Refresh
End If

Call InitServersList(RawServersList)

AddtoRichTextBox frmCargando.Status, "Iniciando constantes....", 255, 255, 255, 1, 0, 1
frmCargando.Refresh

ReDim ListaRazas(1 To NUMRAZAS) As String
ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Drow"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"
'[KEVIN]
ListaRazas(6) = "Orco"
'[/KEVIN]

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

ReDim SkillsNames(1 To NUMSKILLS) As String
SkillsNames(1) = "Música"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar árboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Minería"
SkillsNames(15) = "Carpintería"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Artes Marciales"
SkillsNames(21) = "Navegación"
SkillsNames(22) = "Resistencia mágica"
SkillsNames(23) = "Armas arrojadizas"
SkillsNames(24) = "Alquimia"
SkillsNames(25) = "Sastrería"
SkillsNames(26) = "Botánica"
SkillsNames(27) = "Equitación"

ReDim UserSkills(1 To NUMSKILLS) As Integer
ReDim UserAtributos(1 To NUMATRIBUTOS) As Integer
ReDim AtributosNames(1 To NUMATRIBUTOS) As String
AtributosNames(1) = "Fuerza"
AtributosNames(2) = "Agilidad"
AtributosNames(3) = "Inteligencia"
AtributosNames(4) = "Carisma"
AtributosNames(5) = "Constitucion"

AddtoRichTextBox frmCargando.Status, "Hecho", 0, 255, 0, True
frmCargando.Refresh


'cargo el form así tengo los hwnd desde ahora
'sinuhe_d3d
Load frmMain

Set IAO_TE = New clsTileEngineX
Set IAO_SE = New clsIAO_SE

IniciarObjetosDirectX

Dim LoopC As Integer

ENDL = Chr(13) & Chr(10)
ENDC = Chr(1)
                                  
Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extra...", 255, 255, 255, 1)
frmCargando.Refresh

Call CargarAnimsExtra
Call CargarTips
Call CargarArrayLluvia
Call CargarAnimArmas
Call CargarAnimEscudos
Call CargarParticulas

Call AddtoRichTextBox(frmCargando.Status, "Hecho", 0, 255, 0, 1, 0, 1)
frmCargando.Refresh

AddtoRichTextBox frmCargando.Status, "                    ¡Bienvenido a ImperiumAO!", 255, 255, 0, 1, 0
frmCargando.Refresh

Unload frmCargando

If Musica = 0 Then
    Call IAO_SE.PlayMidi(Midi_Inicio)
End If

frmPres.Picture = LoadPicture(App.Path & "\Interface\Presentacion.jpg")

frmPres.top = 0
frmPres.left = 0
frmPres.width = 800 * Screen.TwipsPerPixelX
frmPres.height = 600 * Screen.TwipsPerPixelY 'Screen.Height

frmPres.Show

Do While Not finpres
    DoEvents
Loop

Unload frmPres

frmConnect.Visible = True


CurrentSpeed = VelNormal
'AdvertenciasMacro = 0
'AdvertenciasSH = 0

PrimeraVez = True
prgRun = True
Pausa = False
bInvMod = True
'lFrameLimiter = DirectX.TickCount

'Obtener el HushMD5
Call CalcularMD5HushYo

If SoftICELoaded Then ' check if softice is loaded
    'Feo mensaje:P
    'Sinuhe
   MsgBox "DXError -2001531853", vbMsgBoxSetForeground + vbInformation
   'sinuhe_d3d
   Set IAO_TE = Nothing
   End ' if true finish the app
End If

Do While prgRun

    If RequestPosTimer > 0 Then
        RequestPosTimer = RequestPosTimer - 1
        If RequestPosTimer = 0 Then
            'Pedimos que nos envie la posicion
            Call SendData("RPU")
        End If
    End If

    Call RefreshAllChars

    'sinuhe para que no se mueva el pj mientras comercia
    
    If Not Pausa And frmMain.Visible And Not frmForo.Visible And _
        Not frmComerciar.Visible And Not frmComerciarUsu.Visible Then
        CheckKeys
        CheckMouse
    End If

    If EngineRun Then
        If frmMain.WindowState <> vbMinimized Then
            IAO_TE.Engine_Render_Start
            If bInvMod Then DibujarInv
            IAO_TE.Engine_Render_End
            Call RenderSounds
        End If
    End If
    
    If Musica = 0 Then
        'If Not SegState Is Nothing Then
        '    If Not Perf.IsPlaying(Seg, SegState) Then Play_Midi
        'End If
    End If
    
    If Logged Then
        EstTick = GetTickCount
        For LoopC = 1 To UBound(Timers)
            
            Timers(LoopC) = Timers(LoopC) + (EstTick - UltTick)
            
            'Timer de trabajo
            If Timers(1) >= tUs Then
                Timers(1) = 0
                NoPuedeUsar = False
            End If
            
            'Timer de attaque (77)
            If Timers(2) >= tAt Then
                Timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If
        
        Next LoopC
    End If
    
    UltTick = GetTickCount
    
    DoEvents
Loop

EngineRun = False
frmCargando.Show
AddtoRichTextBox frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 1
LiberarObjetosDX
'sinuhe_d3d
Set IAO_TE = Nothing

Call UnloadAllForms

Config_Inicio.tip = tipf
Call EscribirGameIni(Config_Inicio)
Call SaveImpAoInit

End

ManejadorErrores:
    'LogError "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
    
End Sub

Sub WriteVar(file As String, Main As String, Var As String, value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************

writeprivateprofilestring Main, Var, value, file

End Sub

Function GetVar(file As String, Main As String, Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************

Dim l As Integer
Dim Char As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

szReturn = ""

sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish


getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

GetVar = RTrim(sSpaces)
GetVar = left(GetVar, Len(GetVar) - 1)

End Function

Function HayAgua(X As Integer, Y As Integer) As Boolean

If MapData(X, Y).Graphic(1).grhindex >= 1505 And _
   MapData(X, Y).Graphic(1).grhindex <= 1520 And _
   MapData(X, Y).Graphic(2).grhindex = 0 Then
            HayAgua = True
Else
            HayAgua = False
End If

End Function
    
'[Barrin]
Public Sub UserExpPerc()

    If UserExp > 0 Then
        UserPercentageExp = (UserExp * 100) / UserPasarNivel
    Else
        UserPercentageExp = 0
    End If
End Sub

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

Public Sub LoadFontTypes()

Dim lc As Integer, arch As String, TempStr As String

arch = App.Path & "\init\" & "FontTypes.ind"

NUMFONTS = Val(GetVar(arch, "INIT", "NumFonts"))
ReDim Preserve FontTypes(1 To NUMFONTS) As tFontType

For lc = 1 To NUMFONTS
    TempStr = GetVar(arch, "INIT", Str(lc))
    FontTypes(lc).red = Val(ReadField(2, TempStr, 126))
    FontTypes(lc).green = Val(ReadField(3, TempStr, 126))
    FontTypes(lc).blue = Val(ReadField(4, TempStr, 126))
    FontTypes(lc).bold = Val(ReadField(5, TempStr, 126))
    FontTypes(lc).italic = Val(ReadField(6, TempStr, 126))
Next lc

End Sub

Public Sub LoadImpAoInit()

Dim lc As Integer, arch As String

arch = App.Path & "\init\" & "ImpAoInit.bnd"

NUMBOTONES = Val(GetVar(arch, "INIT", "NumBotones"))
NUMBINDS = Val(GetVar(arch, "INIT", "NumBinds"))

ReDim Preserve MacroKeys(1 To NUMBOTONES) As tBoton
ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

For lc = 1 To NUMBOTONES
    MacroKeys(lc).TipoAccion = Val(GetVar(arch, "Bind" & lc, "Accion"))
    MacroKeys(lc).hlist = Val(GetVar(arch, "Bind" & lc, "hlist"))
    MacroKeys(lc).invslot = Val(GetVar(arch, "Bind" & lc, "invslot"))
    MacroKeys(lc).SendString = GetVar(arch, "Bind" & lc, "SndString")
Next lc

lc = 0

For lc = 1 To NUMBINDS
    BindKeys(lc).KeyCode = Val(ReadField(1, GetVar(arch, "USER", Str(lc)), 44))
    BindKeys(lc).Name = ReadField(2, GetVar(arch, "USER", Str(lc)), 44)
Next lc

VerLugar = Val(GetVar(arch, "INIT", "VerLugar"))
FxNavega = Val(GetVar(arch, "INIT", "FxNavega"))
EfectosM = Val(GetVar(arch, "INIT", "EfectosM"))
Ambiente = Val(GetVar(arch, "INIT", "Ambiente"))

ConAlfaB = Val(GetVar(arch, "INIT", "ConAlfaB"))

ResChange = Val(GetVar(arch, "INIT", "ResChange"))
GuardarEXP = Val(GetVar(arch, "INIT", "GuardarExp"))
CopiarDialogos = Val(GetVar(arch, "INIT", "CopiarDialogos"))
MensajesGlobales = Val(GetVar(arch, "INIT", "MensajesGlobales"))

DefMidi = Val(GetVar(arch, "INIT", "DefaultMidi"))
frmOpciones.chkMidi.value = DefMidi

gldf = Val(GetVar(arch, "INIT", "gldf"))
frmGuildNews.ChkShow.value = gldf

VolumenInicial = 120

End Sub

Public Sub SaveImpAoInit()

Dim lc As Integer, arch As String

arch = App.Path & "\init\" & "ImpAoInit.bnd"

Call WriteVar(arch, "INIT", "NUMBINDS", Str(NUMBINDS))
Call WriteVar(arch, "INIT", "NUMBOTONES", Str(NUMBOTONES))
Call WriteVar(arch, "INIT", "VerLugar", Str(VerLugar))
Call WriteVar(arch, "INIT", "FxNavega", Str(FxNavega))
Call WriteVar(arch, "INIT", "EfectosM", Str(EfectosM))
Call WriteVar(arch, "INIT", "Ambiente", Str(Ambiente))
Call WriteVar(arch, "INIT", "DefaultMidi", Str(DefMidi))
Call WriteVar(arch, "INIT", "ConAlfaB", Str(ConAlfaB))
Call WriteVar(arch, "INIT", "ResChange", Str(ResChange))
Call WriteVar(arch, "INIT", "gldf", Str(gldf))
Call WriteVar(arch, "INIT", "GuardarExp", Str(GuardarEXP))
Call WriteVar(arch, "INIT", "CopiarDialogos", Str(CopiarDialogos))
Call WriteVar(arch, "INIT", "MensajesGlobales", Str(MensajesGlobales))

For lc = 1 To NUMBINDS
    Call WriteVar(arch, "User", Str(lc), Str(BindKeys(lc).KeyCode) & "," & BindKeys(lc).Name)
Next lc

lc = 0

For lc = 1 To NUMBOTONES
    Call WriteVar(arch, "Bind" & lc, "Accion", Str(MacroKeys(lc).TipoAccion))
    Call WriteVar(arch, "Bind" & lc, "hlist", Str(MacroKeys(lc).hlist))
    Call WriteVar(arch, "Bind" & lc, "invslot", Str(MacroKeys(lc).invslot))
    Call WriteVar(arch, "Bind" & lc, "SndString", MacroKeys(lc).SendString)
Next lc

End Sub
    
Public Sub ResetAtributos()

Atributos = ATT_INICIALES
frmCrearPersonaje.lbAtributos.Caption = Atributos

Dim i As Integer

For i = 1 To NUMATRIBUTOS
    frmCrearPersonaje.lbAtt(i - 1).Caption = "6"
    UserAtributos(i) = 6
Next i

End Sub

Public Sub EndGame()

prgRun = False

AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
frmCargando.Refresh
LiberarObjetosDX
AddtoRichTextBox frmCargando.Status, "Hecho", 0, 255, 0, 1, 0, 1
AddtoRichTextBox frmCargando.Status, "¡Gracias por jugar ImperiumAO!", 0, 0, 0, 1, 0, 1
frmCargando.Refresh
Call UnloadAllForms

End Sub

Public Function HabilidadToString(Habilidades As String) As String

Dim HabilidadesActuales(1 To MAX_HABILIDADES) As Integer, i As Integer, PrimeraHabilidad As Boolean

PrimeraHabilidad = True

For i = 1 To MAX_HABILIDADES
    HabilidadesActuales(i) = Val(ReadField(i, Habilidades, Asc("-")))

    If HabilidadesActuales(i) > 0 Then
        If PrimeraHabilidad Then
            HabilidadToString = HabilidadName(HabilidadesActuales(i))
            PrimeraHabilidad = False
        Else
            HabilidadToString = HabilidadToString & " - " & HabilidadName(HabilidadesActuales(i))
        End If
    End If

Next i

End Function

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

Dim MD5Temp As String

fMD5HushYo = MD5File(App.Path & "\ImperiumAo.exe")
MD5Temp = txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\ImperiumAoNoDinamico.exe")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\Cabezas.ind")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\Cascos.ind")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\Escudos.dat")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\Graficos.ind")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\Fxs.ind")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\FK.ind")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\Init\Personajes.ind")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
fMD5HushYo = MD5File(App.Path & "\aamd532.dll")
MD5Temp = MD5Temp & txtOffset(hexMd52Asc(fMD5HushYo), 53)
MD5HushYo = MD5String(MD5Temp)

End Sub

Public Function General_Get_Elapsed_Time() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If

    'Get current time
    QueryPerformanceCounter start_time
    
    'Calculate elapsed time
    General_Get_Elapsed_Time = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time
End Function

Public Function General_RGB_Color_to_Long(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim c As Long
    
    r = r '* 255
    g = g '* 255
    b = b '* 255
    a = a '* 255
    
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    General_RGB_Color_to_Long = c

End Function

'[/Barrin]
