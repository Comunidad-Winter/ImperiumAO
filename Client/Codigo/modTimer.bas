Attribute VB_Name = "modTimer"
'*****************************************************************
'modTimer - ImperiumAO - v1.4.5 R5
'
'Windows API timer functions and handles.
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

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private hBuffersTimer As Long
Private hFXTimer As Long
Private hHourTimer As Long
Private hRecSTTimer As Long
Private hRecMTimer As Long
Private hPubliTimer As Long

'Tolerancia por delay
Public Const CONST_INTERVALO_TOLERANCIA As Long = 0

Private Const CONST_INTERVALO_USAR As Long = 250
Private Const CONST_INTERVALO_TRABAJAR As Long = 600
Private Const CONST_INTERVALO_RPU As Long = 600
Private Const CONST_INTERVALO_ENDGAME As Long = 3000

Public Sub BuffersBorraTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 120000)
    If Enabled Then
        If hBuffersTimer <> 0 Then KillTimer 0, hBuffersTimer
        hBuffersTimer = SetTimer(0, 0, Intervalo, AddressOf BuffersBorraTimerProc)
    Else
        If hBuffersTimer = 0 Then Exit Sub
        KillTimer 0, hBuffersTimer
        hBuffersTimer = 0
    End If
End Sub

Public Sub FXTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 4000)
    If Enabled Then
        If hFXTimer <> 0 Then KillTimer 0, hFXTimer
        hFXTimer = SetTimer(0, 0, Intervalo, AddressOf FXTimerProc)
    Else
        If hFXTimer = 0 Then Exit Sub
        KillTimer 0, hFXTimer
        hFXTimer = 0
    End If
End Sub

Public Sub HoraTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 3600000)
    If Enabled Then
        If hHourTimer <> 0 Then KillTimer 0, hHourTimer
        hHourTimer = SetTimer(0, 0, Intervalo, AddressOf HoraLogicProc)
        Call HoraLogicProc
    Else
        If hHourTimer = 0 Then Exit Sub
        KillTimer 0, hHourTimer
        hHourTimer = 0
    End If
End Sub

Public Sub RecSTTimer(ByVal Enabled As Boolean, ByVal Intervalo As Long)
    If Enabled Then
        If hRecSTTimer <> 0 Then KillTimer 0, hRecSTTimer
        hRecSTTimer = SetTimer(0, 0, Intervalo, AddressOf RecSTProc)
    Else
        If hRecSTTimer = 0 Then Exit Sub
        KillTimer 0, hRecSTTimer
        hRecSTTimer = 0
    End If
End Sub

Public Sub RecMANTimer(ByVal Enabled As Boolean, ByVal Intervalo As Long)
    If Enabled Then
        If hRecMTimer <> 0 Then KillTimer 0, hRecMTimer
        hRecMTimer = SetTimer(0, 0, Intervalo, AddressOf RecMProc)
    Else
        If hRecMTimer = 0 Then Exit Sub
        KillTimer 0, hRecMTimer
        hRecMTimer = 0
    End If
End Sub

Public Sub PubliTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 900000)
    If Enabled Then
        If hPubliTimer <> 0 Then KillTimer 0, hPubliTimer
        hPubliTimer = SetTimer(0, 0, Intervalo, AddressOf PubliTimerProc)
    Else
        If hPubliTimer = 0 Then Exit Sub
        KillTimer 0, hPubliTimer
        hPubliTimer = 0
    End If
End Sub

Private Sub RecSTProc()

Dim intSta As Integer
Dim PorcRec As Integer

If CurrentUser.Logged = False Then Exit Sub
If CurrentUser.Trabajando Then Exit Sub
If CurrentUser.UserMinSTA = CurrentUser.UserMaxSTA Or CurrentUser.Muerto Then Exit Sub
If (CurrentUser.UserMinHAM = 0 Or CurrentUser.UserMinAGU = 0 Or frmMain.Engine.Char_Current_Naked_Get = True) And (CurrentUser.Descansando = False) Then Exit Sub
If (CurrentUser.MapExt <> 0) And (frmMain.Engine.Engine_Meteo_Particle_Get > 0 And frmMain.Engine.Char_User_Roof_Get = False) Then Exit Sub

PorcRec = CInt(10 + (CurrentUser.Supervivencia / 10))
intSta = CInt(General_Random_Number(1, Porcentaje(CurrentUser.UserMaxSTA, PorcRec)))

CurrentUser.UserMinSTA = CurrentUser.UserMinSTA + intSta

If CurrentUser.UserMinSTA >= CurrentUser.UserMaxSTA Then
    CurrentUser.UserMinSTA = CurrentUser.UserMaxSTA
    Call ClientTCP.Send_Data(Stats_Sync, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA))
End If

Call ClientTCP.ActualizarEst(, , , , , CurrentUser.UserMinSTA)

End Sub

Private Sub RecMProc()

Dim lngPorc As Long
Dim intMAN As Integer

If CurrentUser.Logged = False Then Exit Sub
If CurrentUser.Meditando = False Then Exit Sub
If CurrentUser.UserMinMAN = CurrentUser.UserMaxMAN Or CurrentUser.Muerto Then Exit Sub

If (GetTickCount - CurrentUser.Intervalos.InicioMeditar > (TIEMPO_INICIOMEDITAR * 1000)) Then
    lngPorc = CLng((CurrentUser.Meditar + 50) / 10)
Else
    Exit Sub
End If

intMAN = Porcentaje(CurrentUser.UserMaxMAN, lngPorc)

CurrentUser.UserMinMAN = CurrentUser.UserMinMAN + intMAN

If CurrentUser.UserMinMAN > CurrentUser.UserMaxMAN Then
    CurrentUser.UserMinMAN = CurrentUser.UserMaxMAN
    Call ClientTCP.Send_Data(Stats_Sync, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA))
End If

Call ClientTCP.ActualizarEst(, , , CurrentUser.UserMinMAN)

End Sub

Private Sub FXTimerProc()

Dim N As Long

On Error Resume Next

If CurrentUser.Logged And CurrentUser.MapExt = 1 And (Meteo_Engine.SecondaryStatus = 2 Or Meteo_Engine.SecondaryStatus = 4) Then
    If General_Random_Number(1, 100) > 25 Then
        N = General_Random_Number(1, 100)
        If Meteo_Engine.SecondaryStatus = 4 Then
             If N < 30 And N >= 15 Then
                 N = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO1, , , N)
                 Call Meteo_Engine.StartLighting
             ElseIf N < 30 And N < 15 Then
                 N = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO2, , , N)
                 Call Meteo_Engine.StartLighting
             ElseIf N >= 30 And N <= 35 Then
                 N = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO3, , , N)
                 Call Meteo_Engine.StartLighting
             ElseIf N >= 35 And N <= 40 Then
                 N = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO4, , , N)
                 Call Meteo_Engine.StartLighting
             ElseIf N >= 40 And N <= 45 Then
                 N = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO5, , , N)
             End If
        ElseIf Meteo_Engine.SecondaryStatus = 2 Then
             If N >= 40 And N <= 45 Then
                 N = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO5, , , N)
             End If
        End If
    End If
End If

End Sub

Private Sub HoraLogicProc()

If Meteo_Engine Is Nothing Then Exit Sub

CurrentUser.HoraActual = CurrentUser.HoraActual + 1
If CurrentUser.HoraActual > 24 Then CurrentUser.HoraActual = 0
Call Meteo_Engine.Time_Logic(CurrentUser.HoraActual)

End Sub

Private Sub BuffersBorraTimerProc()
If Sound Is Nothing Then Exit Sub
Call Sound.BorraTimer
End Sub

Public Function IntervaloPermiteTrabajar() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Trabajo >= CONST_INTERVALO_TRABAJAR Then
    CurrentUser.Intervalos.Trabajo = TActual
    IntervaloPermiteTrabajar = True
Else
    IntervaloPermiteTrabajar = False
End If

End Function

Public Function IntervaloPermiteUsar() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Uso >= CONST_INTERVALO_USAR Then
    CurrentUser.Intervalos.Uso = TActual
    IntervaloPermiteUsar = True
Else
    IntervaloPermiteUsar = False
End If

End Function

Public Function IntervaloPermiteAtacar() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Ataque >= CurrentUser.IAtaque Then
    CurrentUser.Intervalos.Ataque = TActual
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If

End Function

Public Function IntervaloPermiteLanzarSpell() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Hechizo >= CurrentUser.IMagia Then
    CurrentUser.Intervalos.Hechizo = TActual
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function

Public Function IntervaloPermiteRefrescar() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.RequestPos >= CONST_INTERVALO_RPU Then
    CurrentUser.Intervalos.RequestPos = TActual
    IntervaloPermiteRefrescar = True
Else
    IntervaloPermiteRefrescar = False
End If

End Function

Private Function PubliTimerProc()

If CurrentUser.Logged = False Or Pubilicidad_Deshabilitada Then Exit Function

If Not Publicidad_Visible And Not (CurrentUser.Muerto = True Or frmMain.Engine.Map_Combat_Get = 1) Then
    Call Random_Announce
Else
    frmMain.publi.Refresh
End If

End Function
