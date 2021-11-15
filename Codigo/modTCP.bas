Attribute VB_Name = "modTCP"
'*****************************************************************
'modTCP - ImperiumAO - v1.3.0
'
'TCP protocol handle.
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
'Pablo Ignacio Márquez (morgolock@speedy.com.ar)
'   - First Relase
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - Recoding
'*****************************************************************

Option Explicit

Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean
Public LlegoFami As Boolean
Public LlegoEst As Boolean

Public Sub HandleData(ByVal rdata As String)
    
    On Error Resume Next
    
    Dim x As Integer
    Dim y As Integer
    Dim CharIndex As Integer
    Dim TempInt As Integer
    Dim TempStr As String
    Dim i As Integer, k As Integer
    Dim cad$, m As Integer
    Dim t() As String
    
    Dim sData As String
    
    Dim part_life() As Long
    Dim part_type() As Integer
        
    sData = UCase$(rdata)
                
    'Handle específico del protocolo a implementar
                
End Sub

Sub SendData(ByVal sdData As String)

Dim retcode
Dim lsdData As Long
Dim abk As Long
Dim AuxCmd As String

If left$(sdData, 1) = "/" Then
    frmPanelGm.LastStr = sdData
End If

If InStr(1, sdData, ENDC) <> 0 Then Exit Sub

sdData = sdData & ENDC

If frmMain.mainWinsock.State = sckConnected Then
    frmMain.mainWinsock.SendData (sdData)
End If

End Sub

Sub Login(ByVal valcode As Integer)

'Personaje grabado
If EstadoLogin = NORMAL Then
    SendData ("OLOGIN" & CurrentUser.UserName & "," & CurrentUser.UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
'Crear personaje
ElseIf EstadoLogin = CrearNuevoPj Then
    If CurrentUser.UserClase = MAGO Or CurrentUser.UserClase = DRUIDA Or CurrentUser.UserClase = CAZADOR Then
        SendData ("NLOGIN" & CurrentUser.UserName & "," & CurrentUser.UserPassword _
        & "," & 0 & "," & 0 & "," _
        & App.Major & "." & App.Minor & "." & App.Revision & _
        "," & CurrentUser.UserRaza & "," & CurrentUser.UserSexo & "," & CurrentUser.UserClase & "," & _
        CurrentUser.UserAtributos(1) & "," & CurrentUser.UserAtributos(2) & "," & CurrentUser.UserAtributos(3) _
        & "," & CurrentUser.UserAtributos(4) & "," & CurrentUser.UserAtributos(5) _
        & "," & CurrentUser.UserSkills((1)) & "," & CurrentUser.UserSkills((2)) _
        & "," & CurrentUser.UserSkills((3)) & "," & CurrentUser.UserSkills((4)) _
        & "," & CurrentUser.UserSkills((5)) & "," & CurrentUser.UserSkills((6)) _
        & "," & CurrentUser.UserSkills((7)) & "," & CurrentUser.UserSkills((8)) _
        & "," & CurrentUser.UserSkills((9)) & "," & CurrentUser.UserSkills((10)) _
        & "," & CurrentUser.UserSkills((11)) & "," & CurrentUser.UserSkills((12)) _
        & "," & CurrentUser.UserSkills((13)) & "," & CurrentUser.UserSkills((14)) _
        & "," & CurrentUser.UserSkills((15)) & "," & CurrentUser.UserSkills((16)) _
        & "," & CurrentUser.UserSkills((17)) & "," & CurrentUser.UserSkills((18)) _
        & "," & CurrentUser.UserSkills((19)) & "," & CurrentUser.UserSkills((20)) _
        & "," & CurrentUser.UserSkills((21)) & "," & CurrentUser.UserSkills((22)) _
        & "," & CurrentUser.UserSkills((23)) & "," & CurrentUser.UserSkills((24)) _
        & "," & CurrentUser.UserSkills((25)) & "," & CurrentUser.UserSkills((26)) _
        & "," & CurrentUser.UserSkills((27)) & "," & CurrentUser.UserEmail & "," & CurrentUser.UserHogar & "," & "1" _
        & "," & CurrentUser.UserPet.nombre & "," & CurrentUser.UserPet.Tipo & "," & valcode & MD5HushYo)
    Else
        SendData ("NLOGIN" & CurrentUser.UserName & "," & CurrentUser.UserPassword _
        & "," & 0 & "," & 0 & "," _
        & App.Major & "." & App.Minor & "." & App.Revision & _
        "," & CurrentUser.UserRaza & "," & CurrentUser.UserSexo & "," & CurrentUser.UserClase & "," & _
        CurrentUser.UserAtributos(1) & "," & CurrentUser.UserAtributos(2) & "," & CurrentUser.UserAtributos(3) _
        & "," & CurrentUser.UserAtributos(4) & "," & CurrentUser.UserAtributos(5) _
        & "," & CurrentUser.UserSkills(1) & "," & CurrentUser.UserSkills(2) _
        & "," & CurrentUser.UserSkills(3) & "," & CurrentUser.UserSkills(4) _
        & "," & CurrentUser.UserSkills(5) & "," & CurrentUser.UserSkills(6) _
        & "," & CurrentUser.UserSkills(7) & "," & CurrentUser.UserSkills(8) _
        & "," & CurrentUser.UserSkills(9) & "," & CurrentUser.UserSkills(10) _
        & "," & CurrentUser.UserSkills(11) & "," & CurrentUser.UserSkills(12) _
        & "," & CurrentUser.UserSkills(13) & "," & CurrentUser.UserSkills(14) _
        & "," & CurrentUser.UserSkills(15) & "," & CurrentUser.UserSkills(16) _
        & "," & CurrentUser.UserSkills(17) & "," & CurrentUser.UserSkills(18) _
        & "," & CurrentUser.UserSkills(18) & "," & CurrentUser.UserSkills(20) _
        & "," & CurrentUser.UserSkills(21) & "," & CurrentUser.UserSkills(22) _
        & "," & CurrentUser.UserSkills(23) & "," & CurrentUser.UserSkills(24) _
        & "," & CurrentUser.UserSkills(25) & "," & CurrentUser.UserSkills(26) _
        & "," & CurrentUser.UserSkills(27) & "," & CurrentUser.UserEmail & "," & CurrentUser.UserHogar _
        & "," & "0" & "," & valcode & MD5HushYo)
    End If
End If

End Sub

Private Sub CopiarDialogoAConsola(ByVal NickName As String, Dialogo As String, color As Long)

If NickName = "" Then Exit Sub
If Right$(Dialogo, 1) = " " Then Exit Sub

If InStr(NickName, "<") Then
    NickName = left$(NickName, InStr(NickName, "<") - 2)
End If

Select Case color
    Case vbWhite
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 255, 255, 255, False, True, False)
    Case -987136
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 225, 225, 0, False, True, False)
    Case -3670016
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 255, 0, 0, False, True, False)
    Case vbGreen
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 0, 255, 0, False, True, False)
    Case -14117888
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 0, 201, 197, False, True, False)
    Case &HC0C0C0 'Gris
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 164, 164, 164, False, True, False)

End Select

End Sub

Private Sub MostrarEstadisticas()

If LlegaronSkills And LlegaronAtrib And LlegoFama And LlegoFami And LlegoEst Then
    If frmMain.PedimosEst Then
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show vbModeless, frmMain
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoFami = False
        LlegoEst = False
        frmMain.PedimosEst = False
    End If
End If

End Sub

Public Function ActualizarEst(Optional ByVal MaxHP As Integer = -1, Optional ByVal MinHP As Integer = -1, Optional ByVal MaxMAN As Integer = -1, _
    Optional ByVal MinMAN As Integer = -1, Optional ByVal MaxSTA As Integer = -1, Optional ByVal MinSTA As Integer = -1, _
    Optional ByVal GLD As Long = -1, Optional ByVal Nivel As Integer = -1, Optional PasarNivel As Long = -1, Optional EXP As Long = -1, _
    Optional Fuerza As Integer = -1, Optional Agilidad As Integer = -1, _
    Optional ActualizarTodos As Boolean = False)

Dim ActualizarCual As Byte

If MaxHP <> -1 Then
    CurrentUser.UserMaxHP = MaxHP
    ActualizarCual = 1
End If

If MinHP <> -1 Then
    If MinHP < 0 Then MinHP = 0
    CurrentUser.UserMinHP = MinHP
    ActualizarCual = 1
End If

If MaxMAN <> -1 Then
    CurrentUser.UserMaxMAN = MaxMAN
    ActualizarCual = 2
End If

If MinMAN <> -1 Then
    CurrentUser.UserMinMAN = MinMAN
    ActualizarCual = 2
End If

If MaxSTA <> -1 Then
    CurrentUser.UserMaxSTA = MaxSTA
    ActualizarCual = 3
End If

If MinSTA <> -1 Then
    CurrentUser.UserMinSTA = MinSTA
    ActualizarCual = 3
End If

If GLD <> -1 Then
    CurrentUser.UserGLD = GLD
    ActualizarCual = 4
End If

If Nivel <> -1 Then
    CurrentUser.UserLVL = Nivel
    ActualizarCual = 5
End If

If PasarNivel <> -1 Then
    CurrentUser.UserPasarNivel = PasarNivel
    ActualizarCual = 5
End If
    
If EXP <> -1 Then
    CurrentUser.UserExp = EXP
    ActualizarCual = 5
End If

If Fuerza <> -1 Then
    frmMain.lblFU = Fuerza
    frmMain.lblAG = Agilidad
End If

If Not ActualizarTodos Then
    Select Case ActualizarCual
        Case 1
            Call ActualizarHP
        Case 2
            Call ActualizarMAN
        Case 3
            Call ActualizarSTA
        Case 4
            Call ActualizarGLD
        Case 5
            Call ActualizarExp
    End Select
Else
    Call ActualizarHP
    Call ActualizarMAN
    Call ActualizarSTA
    Call ActualizarGLD
    Call ActualizarExp
End If

End Function

Private Sub ActualizarMAN()

If CurrentUser.UserMaxMAN > 0 Then
    frmMain.MANShp.Width = (((CurrentUser.UserMinMAN + 1 / 100) / (CurrentUser.UserMaxMAN + 1 / 100)) * 91)
    frmMain.lblMP.Visible = True
    frmMain.lblMP.Caption = CurrentUser.UserMinMAN & "/" & CurrentUser.UserMaxMAN
Else
    frmMain.MANShp.Width = 0
    frmMain.lblMP.Visible = False
End If

End Sub

Private Sub ActualizarGLD()
frmMain.GldLbl.Caption = CurrentUser.UserGLD
End Sub

Private Sub ActualizarSTA()
frmMain.STAShp.Width = (((CurrentUser.UserMinSTA / 100) / (CurrentUser.UserMaxSTA / 100)) * 91)
frmMain.lblST.Caption = CurrentUser.UserMinSTA & "/" & CurrentUser.UserMaxSTA
End Sub

Private Sub ActualizarHP()

If CurrentUser.UserMinHP = 0 Then
    CurrentUser.Muerto = True
    CurrentUser.CurrentSpeed = VelRapida
    Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
    frmMain.lblHP.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.Hpshp.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.Hpshp.FillColor = &H808080
Else
    CurrentUser.Muerto = False
    If CurrentUser.Logged Then
        If (CurrentUser.Montando = False) And (Engine.Char_Type_Get(CurrentUser.CurrentChar) <> 4) Then
            CurrentUser.CurrentSpeed = VelNormal
            Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
        End If
    End If
    frmMain.lblHP.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.Hpshp.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.Hpshp.FillColor = &HC0&
End If

End Sub

Private Sub ActualizarExp()

frmMain.LvlLbl.Caption = CurrentUser.UserLVL

Call UserExpPerc

If CurrentUser.UserPercExp <> 0 Then
    frmMain.ExpShp.Width = (((CurrentUser.UserExp / 100) / (CurrentUser.UserPasarNivel / 100)) * 121)
Else
    frmMain.ExpShp.Width = 0
End If
            
frmMain.Label2(1).Caption = IIf(frmMain.UltPos = 1, CurrentUser.UserExp & "/" & CurrentUser.UserPasarNivel, CurrentUser.UserPercExp & "%")

If CurrentUser.UserPasarNivel = 0 Then
    frmMain.Label2(1).Caption = "¡Nivel máximo!"
End If

End Sub

Public Sub ResetCurrentUser()

Dim NewCurrUser As tCurrentUser

CurrentUser = NewCurrUser
CurrentUser.CurrentSpeed = VelNormal

Engine.Char_Current_OverWater_Set (False)
Engine.Char_Current_OnHorse_Set (False)
Engine.Char_Current_Blind_Set (False)
Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)

Sound.Sound_Stop_All
Sound.Ambient_Stop

Meteo_Engine.SecondaryStatus = 0

EngineRun = False
bK = 0
bRK = 0

End Sub

'[/Barrin]
