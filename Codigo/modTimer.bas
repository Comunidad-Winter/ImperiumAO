Attribute VB_Name = "modTimer"
'*****************************************************************
'modTimer - ImperiumAO - v1.3.0
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

'Despreciar siempre un poco esto, hacer el intervalo más corto
Private Const CONST_INTERVALO_CASTEO As Long = 1400
Private Const CONST_INTERVALO_ATAQUE As Long = 1400
Private Const CONST_INTERVALO_USAR As Long = 250
Private Const CONST_INTERVALO_TRABAJAR As Long = 600

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

Public Sub HoraTimer(ByVal Enabled As Boolean, Optional ByVal Intervalo As Long = 60000)
    If Enabled Then
        If hHourTimer <> 0 Then KillTimer 0, hHourTimer
        hHourTimer = SetTimer(0, 0, Intervalo, AddressOf HoraLogicProc)
        'Para cargar la imágen desde ya...
        Call HoraLogicProc
    Else
        If hFXTimer = 0 Then Exit Sub
        KillTimer 0, hHourTimer
        hHourTimer = 0
    End If
End Sub

Private Sub FXTimerProc()

Dim n As Long

If CurrentUser.Logged Then
    If General_Random_Number(1, 100) > 25 Then
        n = General_Random_Number(1, 100)
        If (Meteo_Engine.SecondaryStatus = 2) And (CurrentUser.MapExt = 1) Then
             If n < 30 And n >= 15 Then
                 n = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO1, , , n)
             ElseIf n < 30 And n < 15 Then
                 n = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO2, , , n)
             ElseIf n >= 30 And n <= 35 Then
                 n = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO3, , , n)
             ElseIf n >= 35 And n <= 40 Then
                 n = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO4, , , n)
             ElseIf n >= 40 And n <= 45 Then
                 n = CLng(General_Random_Number(-10000, 10000))
                 Call Sound.Sound_Play(SND_TRUENO5, , , n)
             End If
        End If
    End If
End If

End Sub

Private Sub HoraLogicProc()
If Meteo_Engine Is Nothing Then Exit Sub
Call Meteo_Engine.Time_Logic
End Sub

Private Sub BuffersBorraTimerProc()
If Sound Is Nothing Then Exit Sub
Call Sound.BorraTimer
End Sub

Public Function IntervaloPermiteTrabajar() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Trabajo >= CONST_INTERVALO_TRABAJAR Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Trabajar OK.", 0, 0, 0, 0, 0, 0, 4)
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
    Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 0, 0, 0, 0, 0, 0, 4)
    CurrentUser.Intervalos.Uso = TActual
    IntervaloPermiteUsar = True
Else
    IntervaloPermiteUsar = False
End If

End Function

Public Function IntervaloPermiteAtacar() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Ataque >= CONST_INTERVALO_ATAQUE Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Atacar OK.", 0, 0, 0, 0, 0, 0, 4)
    CurrentUser.Intervalos.Ataque = TActual
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If

End Function

Public Function IntervaloPermiteLanzarSpell() As Boolean

Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - CurrentUser.Intervalos.Hechizo >= CONST_INTERVALO_CASTEO Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Lanzar OK.", 0, 0, 0, 0, 0, 0, 4)
    CurrentUser.Intervalos.Hechizo = TActual
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function
