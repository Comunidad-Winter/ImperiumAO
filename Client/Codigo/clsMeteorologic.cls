VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeteorologic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsMeteorologic - ImperiumAO - v1.4.5 R5
'
'Meteorologic frmMain.Engine.
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
'Sinuhe (sinuhe@imperiumao.com.ar)
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Const DURACION_TRANS = 300

'Main Status
Private Const DIA As Byte = 1
Private Const ATARDECER As Byte = 2
Private Const NOCHE As Byte = 3
Private Const AMANECER As Byte = 4

Private Const FIN_ANIMACION As Byte = &HFF

'Secondary Status
Private Const Normal As Byte = 0
Private Const NUBLADO As Byte = 1
Private Const LLUVIA As Byte = 2
Private Const NIEVE As Byte = 3
Private Const TORMENTA As Byte = 4

Private Const Particula_Lluvia As Long = 58 '8
Private Const Particula_Nieve As Long = 56

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type MeteoStat
    EstadoActual As Byte
    EstadoProximo As Byte
End Type

Private Type Efectos
    MainStatus As MeteoStat
    SecondaryStatus As MeteoStat
End Type

Private Type D3DCOLORVALUE
    a As Long
    r As Long
    g As Long
    b As Long
End Type

Private Meteo As Efectos

Private m_Color As D3DCOLORVALUE
Private TimeElapsed As Single

Private m_Color_Dia As D3DCOLORVALUE
Private m_Color_Noche As D3DCOLORVALUE
Private m_Color_Tarde As D3DCOLORVALUE
Private m_Color_Manana As D3DCOLORVALUE

Private m_Hora_Actual As Long
Private m_Last_Hora_Actual As Long
Private Meteo_Visible As Boolean

Private Declare Sub GetSystemTime Lib "Kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private theTime As SYSTEMTIME

'Inicializa colores de los diferentes estados
Private Sub Class_Initialize()

'No animo nada al iniciar
Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
Meteo.MainStatus.EstadoActual = DIA

With m_Color_Dia
    .a = 255
    .b = 255
    .r = 255
    .g = 255
End With

With m_Color_Noche
    .a = 255
    .b = 170
    .r = 170
    .g = 170
End With

With m_Color_Tarde
    .a = 255
    .b = 200
    .r = 230
    .g = 200
End With

With m_Color_Manana
    .a = 255
    .b = 230
    .r = 200
    .g = 200
End With

m_Last_Hora_Actual = -1

End Sub

Public Sub Meteo_Logic()

Dim Diferencia As Integer
Dim ParticleIndex As Long
Dim CambioValor As Single
Dim C1 As D3DCOLORVALUE
Dim C2 As D3DCOLORVALUE

If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
    TimeElapsed = TimeElapsed + General_Get_Elapsed_Time
    CambioValor = TimeElapsed / DURACION_TRANS
    If CambioValor > 1 Then
        Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo
        Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
    End If
End If

Select Case Meteo.MainStatus.EstadoActual
    Case DIA
        If Meteo.MainStatus.EstadoProximo = ATARDECER Then
            C1 = m_Color_Dia
            C2 = m_Color_Tarde
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
            
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
            
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Dia
        End If
    Case NOCHE
        If Meteo.MainStatus.EstadoProximo = AMANECER Then
            C1 = m_Color_Noche
            C2 = m_Color_Manana
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
                
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
                
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Noche
        End If
    Case AMANECER
        If Meteo.MainStatus.EstadoProximo = DIA Then
            C1 = m_Color_Manana
            C2 = m_Color_Dia
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
                
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
                
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Manana
        End If
    Case ATARDECER
        If Meteo.MainStatus.EstadoProximo = NOCHE Then
            C1 = m_Color_Tarde
            C2 = m_Color_Noche
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
                
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
                
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Tarde
        End If
End Select

ParticleIndex = frmMain.Engine.Engine_Meteo_Particle_Get

Select Case Meteo.SecondaryStatus.EstadoActual
    Case LLUVIA
        If (ParticleIndex = 0) And (CurrentUser.MapExt = 1) Then
            Call frmMain.Engine.Engine_Meteo_Particle_Set(Particula_Lluvia)
            Meteo_Visible = True
            Draw_Time_Image
        Else
            If (CurrentUser.MapExt <> 1) And (ParticleIndex <> 0) Then
                Call frmMain.Engine.Engine_Meteo_Particle_Set(-1)
                Meteo_Visible = False
               Draw_Time_Image
            End If
        End If
    Case NIEVE
        If (ParticleIndex = 0) And (CurrentUser.MapExt = 2) Then
            Call frmMain.Engine.Engine_Meteo_Particle_Set(Particula_Nieve)
            Meteo_Visible = True
            Draw_Time_Image
        Else
            If (CurrentUser.MapExt <> 2) And (ParticleIndex <> 0) Then
                Call frmMain.Engine.Engine_Meteo_Particle_Set(-1)
                Meteo_Visible = False
                Draw_Time_Image
            End If
        End If
    Case Normal
        If (ParticleIndex <> 0) Then
            Call frmMain.Engine.Engine_Meteo_Particle_Set(-1)
            Meteo_Visible = False
            Draw_Time_Image
        End If
    Case TORMENTA
        If (ParticleIndex = 0) And (CurrentUser.MapExt = 1) Then
            Call frmMain.Engine.Engine_Meteo_Particle_Set(Particula_Lluvia)
            Meteo_Visible = True
            Draw_Time_Image
        Else
            If (CurrentUser.MapExt <> 1) And (ParticleIndex <> 0) Then
                Call frmMain.Engine.Engine_Meteo_Particle_Set(-1)
                Meteo_Visible = False
                Draw_Time_Image
            End If
        End If
    Case NUBLADO
        'ASD
End Select

Call frmMain.Engine.Map_Base_Light_Set(m_Color.r, m_Color.g, m_Color.b)

End Sub

Public Sub Time_Logic(ByVal hora_ac As Byte)

m_Hora_Actual = hora_ac

If m_Hora_Actual <> m_Last_Hora_Actual Then
    If Not Meteo_Visible Then frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex(format(m_Hora_Actual, "0#"))
    m_Last_Hora_Actual = m_Hora_Actual
End If

End Sub

Private Sub Draw_Time_Image()

If Not Meteo_Visible Then
    frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex(format(m_Last_Hora_Actual, "0#"))
Else
    Select Case Meteo.SecondaryStatus.EstadoActual
        Case LLUVIA
            frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex("lluvia")
        Case NIEVE
            frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex("nieve")
        Case TORMENTA
            frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex("electrica")
        Case NUBLADO
            frmMain.imgHora.Picture = General_Load_Picture_From_Resource_Ex("nublado")
    End Select
End If

End Sub

Public Function Get_Time_String() As String

GetSystemTime theTime
Get_Time_String = CStr(m_Hora_Actual) & ":" & format(theTime.wMinute, "00") & "... "

If Not Meteo_Visible Then
    Select Case m_Hora_Actual
        Case 5, 6, 7
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(237)
        Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(238)
        Case 18, 19
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(239)
        Case Else
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(240)
    End Select
Else
    Select Case Meteo.SecondaryStatus.EstadoActual
        Case LLUVIA
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(241)
        Case NIEVE
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(242)
        Case TORMENTA
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(243)
        Case NUBLADO
            Get_Time_String = Get_Time_String & Locale_GUI_Frase(244)
    End Select
End If

End Function

Public Sub SetNuevoEstado(ByVal Estado As Byte)

TimeElapsed = 0

If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
    'Como no termino la animación anterior, al menos le pongo que la que estaba haciendo sea desde donde sale
    Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo
End If

Meteo.MainStatus.EstadoProximo = Estado

End Sub

Public Property Get SecondaryStatus() As Byte
SecondaryStatus = Meteo.SecondaryStatus.EstadoActual
End Property

Public Property Let SecondaryStatus(ByVal vNewValue As Byte)
Meteo.SecondaryStatus.EstadoActual = vNewValue
End Property

Public Property Get PrimaryStatus() As Byte
PrimaryStatus = Meteo.MainStatus.EstadoActual
End Property

Public Property Let PrimaryStatus(ByVal vNewValue As Byte)
Meteo.MainStatus.EstadoProximo = vNewValue
End Property

Public Sub ForzarEstado(ByVal Estado As Byte)
TimeElapsed = 0
Meteo.MainStatus.EstadoActual = Estado
Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
End Sub

Public Sub StartLighting()
'Call PrintToConsole("Rayo xD", 0, 0, 0, 0, 0, 0, 4)
End Sub

Public Sub NextEstado()

TimeElapsed = 0

If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
    'Como no termino la animación anterior, al menos le pongo que la que estaba haciendo sea desde donde sale
    Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo
End If

Meteo.MainStatus.EstadoProximo = Meteo.MainStatus.EstadoActual + 1
If Meteo.MainStatus.EstadoProximo > AMANECER Then Meteo.MainStatus.EstadoProximo = DIA

End Sub
