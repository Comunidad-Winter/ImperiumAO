VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientTCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsClientTCP.cls - ImperiumAO - v1.4.0 - www.imperiumao.com.ar
'
'Protocolo de comunicación del cliente
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
'Augusto Rando(barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Enum TIPOEFECTO
    EFECTO_PARTICULA = 0
    EFECTO_GRAFICO = 1
End Enum

Private LlegaronSkills As Boolean
Private LlegaronAtrib As Boolean
Private LlegoFami As Boolean
Private LlegoEst As Boolean

Public Sub Reset_Skill_Data()

LlegaronSkills = False
LlegaronAtrib = False
LlegoFami = False
LlegoEst = False

End Sub

Public Sub Handle_Data(ByVal rData As String)
    
    On Error Resume Next
    
    Dim x As Integer
    Dim y As Integer
    Dim i As Long
    
    Dim CharIndex As Integer
    Dim tempInt As Integer
    Dim tempStr As String
    Dim tempLng As Long
    Dim tempCol As Long
    Dim tmpCad As String
    Dim t() As String
    Dim fxgr_extra() As Long
    Dim fxgr_type() As Long
    Dim fxgr_pos() As Byte
    Dim Paquet_ID As Byte
    Dim tempByte As Byte
    Dim tmpBt As Byte
        
    Paquet_ID = Asc(left$(rData, 1))
    'Paquet_ID = AscB(rData)
    
    rData = Right$(rData, Len(rData) - 1)
    
    Select Case Paquet_ID
        Case Conn_Succ

            frmIniciando.Show
            frmCharList.Visible = False
            frmCrearPersonaje.Visible = False
            
            Call Map_Load(String_To_Integer(rData, 1), String_To_Byte(rData, 3), String_To_Long(rData, 7), String_To_Long(rData, 7))
            
            CurrentUser.Seguro = True
            frmMain.modocombate.Visible = False
            frmMain.nomodocombate.Visible = True
            frmMain.modoseguro.Visible = True
            frmMain.nomodorol.Visible = True
            frmMain.nomodoseguro.Visible = False
            frmMain.modorol.Visible = False
            
            '[Barrin]
            CurrentUser.SendingType = 1 'Normal
            frmMensaje.mnuNormal.Checked = True
            frmMensaje.mnuGritar.Checked = False
            frmMensaje.mnuPrivado.Checked = False
            frmMensaje.mnuClan.Checked = False
            frmMensaje.mnuGMs.Checked = False
            frmMensaje.mnuGrupo.Checked = False
            frmMensaje.mnuGlobal.Checked = False
            
            '¿Es GM? ¿Tiene clan? Si es asi habilitamos los menús
            frmMensaje.mnuClan.Visible = String_To_Byte(rData, 4)
            frmMensaje.mnuGMs.Visible = String_To_Byte(rData, 5)
            
            CurrentUser.HoraActual = CInt(String_To_Byte(rData, 6)) - 1
            CurrentUser.Supervivencia = String_To_Integer(rData, 11)
            CurrentUser.ISNormal = String_To_Long(rData, 13)
            CurrentUser.ISDescansar = String_To_Long(rData, 17)
            CurrentUser.IMana = String_To_Long(rData, 21)
            CurrentUser.IAtaque = String_To_Long(rData, 25) + CONST_INTERVALO_TOLERANCIA
            CurrentUser.IMagia = String_To_Long(rData, 29) + CONST_INTERVALO_TOLERANCIA
            
            Call HoraTimer(True)
            Call RecSTTimer(True, CurrentUser.ISNormal)
            Call RecMANTimer(True, CurrentUser.IMana)
            
            '[/Barrin]
            
            Call DibujarMenuMacros
            Call Inventory_Render
            
            'Barrin
            Call Send_Data(Loading_Finished, Byte_To_String(MensajesGlobales) & Byte_To_String(MensajesFaccionarios))
        
        Case Spectate_Toggle
            
            x = String_To_Integer(rData, 1)
            y = String_To_Integer(rData, 3)
            
            Call frmMain.Engine.Char_Remove(CurrentUser.CurrentChar)
            Call frmMain.Engine.Char_Current_Set(0)
            CurrentUser.CurrentChar = 0
            
            Call frmMain.Engine.Engine_View_Pos_Set(x, y)
            
            If Not EngineRun Then
                Call SetConnected
            Else
                frmMain.Visible = True
                frmIniciando.Visible = False
            End If
            
            If VerLugar = 0 Then frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & x & ", " & y
            
            Call frmMain.Engine.Map_Letter_Fade_Set(frmMain.Engine.Map_Letter_GRH_Get)
            
            General_Update_Minimap x, y
            
            CurrentUser.Spectate = True
            
            CurrentUser.CurrentSpeed = VelRapida
            frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
            
        Case Navigating_Toggle
            CurrentUser.Navegando = Not CurrentUser.Navegando
            frmMain.Engine.Char_Current_OverWater_Set (CurrentUser.Navegando)
        
        Case Horse_Toggle
            CurrentUser.Montando = Not CurrentUser.Montando
            frmMain.Engine.Char_Current_OnHorse_Set (CurrentUser.Montando)
            
            If CurrentUser.Montando Then
                CurrentUser.CurrentSpeed = VelRapida
                frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
                With frmPanelGm
                    .mnuNormal.Checked = False
                    .mnuMuy.Checked = False
                    .mnuRapida.Checked = True
                End With
            Else
                CurrentUser.CurrentSpeed = VelNormal
                frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
                With frmPanelGm
                    .mnuNormal.Checked = True
                    .mnuMuy.Checked = False
                    .mnuRapida.Checked = False
                End With
            End If
        
        'Case Fly_Toggle
        '    CurrentUser.Volando = Not CurrentUser.Volando

        Case End_Succ
            
            If CurrentUser.EndingGame = False Then
                Call ResetCurrentUserEx
                Call FormParser.Parse_Form(frmCharList)
                frmCharList.Visible = True
                frmMain.Visible = False
                
                If sMusica <> CONST_DESHABILITADA Then
                    Sound.NextMusic = MUS_VolverInicio
                    Sound.Fading = 200
                End If
                
            Else
                Call EndGame(True)
            End If
                        
        Case Comm_End_Succ
            Unload frmComerciar
            CurrentUser.Comerciando = False
        
        Case Bank_End_Succ
            Unload frmBancoObj
            CurrentUser.Comerciando = False

        Case Comm_Init
            
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            
            i = 1
            
            Do While i <= UBound(UserInventory)
                If UserInventory(i).ObjIndex <> 0 Then
                    frmComerciar.List1(1).AddItem UserInventory(i).name
                Else
                    frmComerciar.List1(1).AddItem Locale_GUI_Frase(269)
                End If
                i = i + 1
            Loop
            
            i = 1
            Do While i <= UBound(NPCInventory)
                If NPCInventory(i).ObjIndex <> 0 Then
                    frmComerciar.List1(0).AddItem NPCInventory(i).name
                Else
                    frmComerciar.List1(0).AddItem Locale_GUI_Frase(269)
                End If
                i = i + 1
            Loop
            
            CurrentUser.Comerciando = True
            frmComerciar.Show vbModeless, frmMain
        
        Case Bank_Data
            
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            
            i = 1
            
            Do While i <= UBound(UserInventory)
                If UserInventory(i).ObjIndex <> 0 Then
                    frmBancoObj.List1(1).AddItem UserInventory(i).name
                Else
                    frmBancoObj.List1(1).AddItem Locale_GUI_Frase(269)
                End If
                i = i + 1
            Loop
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).ObjIndex <> 0 Then
                    frmBancoObj.List1(0).AddItem UserBancoInventory(i).name
                Else
                    frmBancoObj.List1(0).AddItem Locale_GUI_Frase(269)
                End If
                i = i + 1
            Loop
            
            CurrentUser.Comerciando = True
            frmBancoObj.Show vbModeless, frmMain
        
        Case Herrero_Init
            frmHerrero.Show vbModeless, frmMain
        
        Case Carp_Init
            frmCarp.Show vbModeless, frmMain
        
        Case Druid_Init
            frmDruida.Show vbModeless, frmMain
        
        Case Tailor_Init
            frmSastre.Show vbModeless, frmMain
        
        Case Wedding_Event
            
            tempStr = Locale_Parse_ServerMessage(Wedding_Celebration, Map_NameLoad(String_To_Integer(rData, 1)))
            tmpBt = String_To_Byte(rData, 3)
            
            Call PrintToConsole(tempStr & " " & mid$(rData, 4, tmpBt) & " " & Locale_GUI_Frase(471) & " " & mid$(rData, 4 + tmpBt) & "!", 0, 0, 0, 0, 0, 0, 18)
            Call Sound.Sound_Play(SND_BODA)
        
        Case Char_Desc_Localized
            Call Parse_CharDesc(rData)
            
        Case Item_Desc_Localized
            
            tempInt = String_To_Integer(rData, 1)
            tempStr = Locale_UserItem_Desc(tempInt)
            
            If Len(tempStr) > 0 Then
                tempStr = Locale_UserItem(tempInt) & " - " & tempStr
            Else
                tempStr = Locale_UserItem(tempInt)
            End If
            
            If Len(rData) = 2 Then
                Call PrintToConsole(tempStr, 0, 0, 0, 0, 0, 0, 4)
            Else
                Call PrintToConsole(tempStr & " (" & String_To_Long(rData, 3) & ")", 0, 0, 0, 0, 0, 0, 4)
            End If
            
        Case Duel_Init
            CurrentUser.Duel_Mode = Not CurrentUser.Duel_Mode
            
            If CurrentUser.Duel_Mode Then
                Call PrintToConsole(Locale_GUI_Frase(362), 0, 0, 0, 0, 0, 0, 2)
                
                If (sMusica <> CONST_DESHABILITADA) Then
                    Sound.NextMusic = MUS_Fight
                    Sound.Fading = 300
                End If
                
            Else
            
                If (sMusica <> CONST_DESHABILITADA) Then
                    Sound.NextMusic = Sound.LastMapMusic
                    Sound.Fading = 300
                End If
            
            End If
                        
        Case Ping_Received
            CurrentUser.PingRequested = False
            CurrentUser.Ping = GetTickCount - CurrentUser.Ping
            Call PrintToConsole("Ping: " & CurrentUser.Ping & "ms", 0, 0, 0, 0, 0, 0, 8)
            Call Send_Data(Ping_Request)
        
        Case Map_Load_Request
            frmIniciando.Show
            frmMain.Visible = False
            CurrentUser.CurrentChar = 0
            Call Map_Load(String_To_Integer(rData, 1), String_To_Byte(rData, 3), String_To_Long(rData, 4), String_To_Byte(rData, 4))
            Call Send_Data(Loading_Finished)

        Case Main_Map_Data, Main_Map_Data_Compressed
            If Paquet_ID = Main_Map_Data_Compressed Then rData = UncompressString(rData)
            tempInt = String_To_Integer(rData, 1)
            CurrentUser.Loading = True

            Do While tempInt > 0
                Call ClientTCP.Handle_Data(mid$(rData, 3, tempInt))
                rData = mid$(rData, tempInt + 3)
                If Len(rData) > 1 Then
                    tempInt = String_To_Integer(rData, 1)
                Else
                    tempInt = 0
                End If
            Loop
            
            '## SEGURIDAD
            Call Main_Logic
            
            Call frmMain.Engine.Massive_Load_End
            
            CurrentUser.Loading = False
            
        Case Closing
            Call PrintToConsole(Locale_GUI_Frase(363) & " " & String_To_Byte(rData, 1) & " " & Locale_GUI_Frase(364), 65, 190, 156, False, True, False)
            CurrentUser.Saliendo = True
        
        Case Close_Cancel
            Call PrintToConsole(Locale_GUI_Frase(365), 0, 0, 0, False, False, False, 14)
            CurrentUser.Saliendo = False
        
        Case Dialog_Hit_Spell
            tempLng = String_To_Long(rData, 1)
            
            If tempLng > 200 Then
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & tempLng & "!", COLOR_ATAQUE, 5, True, 2
            Else
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, CStr(tempLng), COLOR_ATAQUE, 5, True, 2
            End If
        
        Case Dialog_Hit_Weap
            tempLng = String_To_Long(rData, 1)
            
            If tempInt > 200 Then
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & tempLng & "!", COLOR_ATAQUE, 5, True, 2
            Else
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, CStr(tempLng), COLOR_ATAQUE, 5, True, 2
            End If
        
        Case Update_CombatTimers
            CurrentUser.IAtaque = String_To_Long(rData, 1)
            CurrentUser.IMagia = String_To_Long(rData, 5)
            
        Case Position_Current
            x = String_To_Integer(rData, 1)
            y = String_To_Integer(rData, 3)
            Call frmMain.Engine.Char_Current_Pos_Refresh(x, y)
            Call frmMain.Engine.Engine_View_Pos_Set(x, y)
            Call frmMain.Engine.Input_Mouse_Start
            General_Update_Minimap x, y
            
        Case Position_Other
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            x = String_To_Integer(rData, 3)
            y = String_To_Integer(rData, 5)
            Call frmMain.Engine.Char_Pos_Set(CharIndex, x, y)
            If CharIndex = CurrentUser.CurrentChar Then General_Update_Minimap x, y
            
        Case Fly_Toggle
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Fly_Set(CharIndex, String_To_Byte(rData, 3))
        
        Case NPC_Hit
            tempByte = String_To_Byte(rData, 1)
            tempInt = String_To_Integer(rData, 2)
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 4))
            
            Select Case tempByte
                Case bCabeza
                    Call PrintToConsole(Locale_GUI_Frase(367) & " " & tempInt, 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call PrintToConsole(Locale_GUI_Frase(368) & " " & tempInt, 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call PrintToConsole(Locale_GUI_Frase(369) & " " & tempInt, 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call PrintToConsole(Locale_GUI_Frase(370) & " " & tempInt, 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call PrintToConsole(Locale_GUI_Frase(371) & " " & tempInt, 255, 0, 0, True, False, False)
                Case bTorso
                    Call PrintToConsole(Locale_GUI_Frase(372) & " " & tempInt, 255, 0, 0, True, False, False)
            End Select
        
            If tempInt > 150 Then
                frmMain.Engine.Char_Dialog_Set CharIndex, "¡" & tempInt & "!", COLOR_ATAQUE, 5, True, 2
            Else
                frmMain.Engine.Char_Dialog_Set CharIndex, tempInt, COLOR_ATAQUE, 5, True, 2
            End If
        
        Case User_NPC_Hit
            tempLng = String_To_Long(rData, 1)
            
            If tempLng > 150 Then
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, Locale_GUI_Frase(373) & tempLng & "!", COLOR_ATAQUE, 5, True, 2
            Else
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, CStr(tempLng), COLOR_ATAQUE, 5, True, 2
            End If
        
        Case User_Failed_User
            tempInt = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            tempStr = frmMain.Engine.Char_Name_Get_No_Guild(tempInt)
            Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(374), 255, 0, 0, True, False, False)
        
        Case User_Hit_Current_User
            tempByte = String_To_Byte(rData, 1)
            tempLng = String_To_Long(rData, 2)
            tempInt = frmMain.Engine.Char_Find(String_To_Integer(rData, 6))
            tempStr = frmMain.Engine.Char_Name_Get_No_Guild(tempInt)
            
            Select Case tempByte
                Case bCabeza
                    Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(375) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(376) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(377) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(378) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(379) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call PrintToConsole(Locale_GUI_Frase(373) & tempStr & " " & Locale_GUI_Frase(380) & " " & tempLng & "!", 255, 0, 0, True, False, False)
            End Select
        
            If tempLng > 150 Then
                frmMain.Engine.Char_Dialog_Set tempInt, Locale_GUI_Frase(373) & tempLng & "!", COLOR_ATAQUE, 5, True, 2
            Else
                frmMain.Engine.Char_Dialog_Set tempInt, tempLng, COLOR_ATAQUE, 5, True, 2
            End If
        
        Case User_Hit_User
            tempByte = String_To_Byte(rData, 1)
            tempLng = String_To_Long(rData, 2)
            tempStr = frmMain.Engine.Char_Name_Get_No_Guild(frmMain.Engine.Char_Find(String_To_Integer(rData, 6)))
            
            Select Case tempByte
                Case bCabeza
                    Call PrintToConsole(Locale_GUI_Frase(381) & " " & tempStr & " " & Locale_GUI_Frase(382) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call PrintToConsole(Locale_GUI_Frase(381) & " " & tempStr & " " & Locale_GUI_Frase(383) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call PrintToConsole(Locale_GUI_Frase(381) & " " & tempStr & " " & Locale_GUI_Frase(384) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call PrintToConsole(Locale_GUI_Frase(381) & " " & tempStr & " " & Locale_GUI_Frase(385) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call PrintToConsole(Locale_GUI_Frase(381) & " " & tempStr & " " & Locale_GUI_Frase(386) & " " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call PrintToConsole(Locale_GUI_Frase(381) & " " & tempStr & " " & Locale_GUI_Frase(387) & " " & tempLng & "!", 255, 0, 0, True, False, False)
            End Select
            
            If tempLng > 150 Then
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, Locale_GUI_Frase(373) & tempLng & "!", COLOR_ATAQUE, 5, True, 2
            Else
                frmMain.Engine.Char_Dialog_Set CurrentUser.CurrentChar, tempLng, COLOR_ATAQUE, 5, True, 2
            End If

        Case Main_Dialog, Main_Dialog_Compressed
        
            If Paquet_ID = Main_Dialog_Compressed Then rData = UncompressString(rData)
        
            tempStr = General_Field_Read(1, rData, ">")
            
            If Not NickIgnorado(tempStr) Then
                If InStr(rData, "~") = 0 Then
                    tempByte = String_To_Byte(mid$(rData, Len(rData)), 1)
                    Call PrintToConsole(left$(rData, Len(rData) - 1), 0, 0, 0, 0, 0, False, tempByte)
                Else
                    tempStr = General_Field_Read(2, rData, "~")
                    Call PrintToConsole(General_Field_Read(1, rData, "~"), String_To_Integer(tempStr, 1), String_To_Integer(tempStr, 3), String_To_Integer(tempStr, 5), String_To_Byte(tempStr, 7), String_To_Byte(tempStr, 8))
                End If
            End If
        
        Case NPC_Name_Status_Localized
            
            tempStr = Locale_NPC_Name(String_To_Integer(rData, 1)) & " [" & Generate_Char_Status(String_To_Long(rData, 3), String_To_Byte(rData, 7), String_To_Byte(rData, 8), String_To_Byte(rData, 9)) & "]"
            Call PrintToConsole(tempStr, 0, 0, 0, 0, 0, 0, 4)
        
        Case NPC_Name_Status_Localized_Invoked
        
            tempStr = Locale_NPC_Name(String_To_Integer(rData, 1)) & " [" & Generate_Char_Status(String_To_Long(rData, 3), String_To_Byte(rData, 7), String_To_Byte(rData, 8), String_To_Byte(rData, 9)) & "] " & Locale_GUI_Frase(469) & " " & mid$(rData, 10)
            Call PrintToConsole(tempStr, 0, 0, 0, 0, 0, 0, 4)

        Case NPC_Name_Status_Localized_Owner
        
            tempStr = Locale_NPC_Name(String_To_Integer(rData, 1)) & " [" & String_To_Integer(rData, 3) & "/" & String_To_Integer(rData, 5) & Generate_Char_Status(-1, String_To_Byte(rData, 7), String_To_Byte(rData, 8), String_To_Byte(rData, 9)) & "] " & Locale_GUI_Frase(469) & " " & mid$(rData, 10)
            Call PrintToConsole(tempStr, 0, 0, 0, 0, 0, 0, 4)

        Case Spell_Caster_Localized
            
            tempInt = String_To_Integer(rData, 1)
            tempStr = mid$(rData, 3)
            If LenB(tempStr) = 0 Then tempStr = Locale_GUI_Frase(470)
            Call PrintToConsole(Locale_Spell_Caster(tempInt) & " " & tempStr, 255, 0, 0, True, False, False)
        
        Case Spell_Caster_Localized_Terrain

            tempInt = String_To_Integer(rData, 1)
            Call PrintToConsole(Locale_Spell_Caster(tempInt), 255, 0, 0, True, False, False)

        Case Spell_Target_Localized
            
            tempInt = String_To_Integer(rData, 1)
            tempStr = mid$(rData, 3)
            Call PrintToConsole(tempStr & " " & Locale_Spell_Target(tempInt), 255, 0, 0, True, False, False)
        
        Case Spell_Owner_Localized
        
            tempInt = String_To_Integer(rData, 1)
            Call PrintToConsole(Locale_Spell_Owner(tempInt), 255, 0, 0, True, False, False)
        
        Case OverChar_Dialog_Localized
        
            tempStr = Locale_NPC_Dialog(String_To_Integer(rData, 3))
            
            If LenB(tempStr) > 0 Then
                CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
                frmMain.Engine.Char_Dialog_Set CharIndex, tempStr, vbWhite, 10, True
            End If
        
        Case Centinel_Info
            CharIndex = String_To_Integer(rData, 1)
            tempInt = String_To_Integer(rData, 3)
            
            If CharIndex > 0 Then
                CharIndex = frmMain.Engine.Char_Find(CharIndex)
                tmpCad = Replace$(Locale_GUI_Frase(472), "#1", CStr(tempInt))
                tmpCad = Replace$(tmpCad, "%N", mid$(rData, 5))
                frmMain.Engine.Char_Dialog_Set CharIndex, tmpCad, -3670016, 10, True
            Else
                tmpCad = Replace$(Locale_GUI_Frase(473), "#1", CStr(tempInt))
                frmMensaje.msg.Caption = tmpCad
                frmMensaje.Show vbModeless, frmMain
                Call Sound.Sound_Play(SND_MENSAJE)
            End If
        
        Case OverChar_Dialog, OverChar_Dialog_Compressed
            
            If Paquet_ID = OverChar_Dialog_Compressed Then rData = UncompressString(rData)
        
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            tempStr = frmMain.Engine.Char_Name_Get_No_Guild(CharIndex)
            tmpCad = mid$(rData, 7)
            
            If Not NickIgnorado(tempStr) Then
                frmMain.Engine.Char_Dialog_Set CharIndex, tmpCad, String_To_Long(rData, 3), 10, True
                
                If CopiarDialogos = 1 And LenB(tempStr) > 0 Then
                    Call CopiarDialogoAConsola(tempStr, tmpCad, String_To_Long(rData, 3))
                End If
            End If

        Case Magic_Words
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            
            If CharIndex > 0 Then _
                frmMain.Engine.Char_Dialog_Set CharIndex, mid$(rData, 7), String_To_Long(rData, 3), 5, True

        Case Global_Chat_Message
            
            rData = UncompressString(rData)
        
            If MensajesGlobales = 1 Then
                tempStr = General_Field_Read(1, rData, ">")
                If Not NickIgnorado(tempStr) Then _
                    Call PrintToConsole(rData, 0, 0, 0, 0, 0, 0, 23)
            End If
        
        Case Stats_HP
            Call ActualizarEst(, String_To_Integer(rData, 1))
            
        Case Stats_STA
            Call ActualizarEst(, , , , , String_To_Integer(rData, 1))
        
        Case Stats_MAN
            Call ActualizarEst(, , , String_To_Integer(rData, 1))
        
        Case Stats_Update_MAN
            Call ActualizarEst(, , , CurrentUser.UserMinMAN + String_To_Integer(rData, 1))
        
        Case War_Mode_Engaged
            
            If String_To_Byte(rData, 1) = 1 Then
                Call PrintToConsole(Locale_GUI_Frase(388), 0, 0, 0, 0, 0, 0, 3)
                Call Sound.Sound_Play(SND_WARHORN)
            Else
                Call PrintToConsole(Locale_GUI_Frase(389), 0, 0, 0, 0, 0, 0, 3)
            End If
            
        Case Stats_GLD
            Call ActualizarEst(, , , , , , String_To_Long(rData, 1))
        
        Case Stats_EXP
            Call ActualizarEst(, , , , , , , String_To_Integer(rData, 1), String_To_Long(rData, 3), String_To_Long(rData, 7))
        
        Case Stats_ATT
            Call ActualizarEst(, , , , , , , , , , String_To_Integer(rData, 1), String_To_Integer(rData, 3))
        
        Case Server_Message
            frmMensaje.msg.Caption = rData
            frmMensaje.Show vbModeless, frmMain
            Call Sound.Sound_Play(SND_MENSAJE)
            
        Case Create_Char
            tempByte = String_To_Byte(rData, 31)
            tempStr = mid$(rData, 32 + (tempByte * 9))
            
            If tempByte > 0 Then
                ReDim fxgr_type(1 To tempByte) As Long
                ReDim fxgr_extra(1 To tempByte) As Long
                ReDim fxgr_pos(1 To tempByte) As Byte
            End If
            
            'Primer byte = 31
            'Último byte = 32 + (tempByte * 4)
            tempLng = -9
            For i = 1 To tempByte
                tempLng = tempLng + 9
                fxgr_type(i) = String_To_Long(rData, 32 + tempLng)
                fxgr_extra(i) = String_To_Long(rData, 36 + tempLng)
                fxgr_pos(i) = String_To_Byte(rData, 40 + tempLng)
            Next i
        
            Call frmMain.Engine.Char_Create(String_To_Integer(rData, 8), String_To_Integer(rData, 10), String_To_Byte(rData, 5), String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 16), String_To_Integer(rData, 12), String_To_Integer(rData, 14), tempStr, _
                String_To_Byte(rData, 18), 0, 0, String_To_Integer(rData, 6), _
                String_To_Byte(rData, 19), String_To_Byte(rData, 20), String_To_Long(rData, 23), String_To_Long(rData, 27), tempByte, fxgr_extra(), fxgr_type(), fxgr_pos(), String_To_Integer(rData, 21))
            
            If Not CurrentUser.Loading Then frmMain.Engine.Massive_Load_End
            
        Case Create_Char_Current
            tempByte = String_To_Byte(rData, 31)
            tempStr = mid$(rData, 32 + (tempByte * 9))
            
            If tempByte > 0 Then
                ReDim fxgr_type(1 To tempByte) As Long
                ReDim fxgr_extra(1 To tempByte) As Long
                ReDim fxgr_pos(1 To tempByte) As Byte
            End If
            
            'Primer byte = 31
            'Último byte = 32 + (tempByte * 4)
            tempLng = -9
            For i = 1 To tempByte
                tempLng = tempLng + 9
                fxgr_type(i) = String_To_Long(rData, 32 + tempLng)
                fxgr_extra(i) = String_To_Long(rData, 36 + tempLng)
                fxgr_pos(i) = String_To_Byte(rData, 40 + tempLng)
            Next i
            
            CurrentUser.CharID = String_To_Integer(rData, 6)
            CurrentUser.CharType = String_To_Byte(rData, 18)
        
            CharIndex = frmMain.Engine.Char_Create(String_To_Integer(rData, 8), String_To_Integer(rData, 10), String_To_Byte(rData, 5), String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 16), String_To_Integer(rData, 12), String_To_Integer(rData, 14), tempStr, _
                CurrentUser.CharType, 0, 0, CurrentUser.CharID, _
                String_To_Byte(rData, 19), String_To_Byte(rData, 20), String_To_Long(rData, 23), String_To_Long(rData, 27), tempByte, fxgr_extra(), fxgr_type(), fxgr_pos(), String_To_Integer(rData, 21))
            
            CurrentUser.CurrentChar = CharIndex
            Call frmMain.Engine.Char_Current_Set(CharIndex)
            Call frmMain.Engine.Char_Pos_Get(CharIndex, x, y)
            Call frmMain.Engine.Engine_View_Pos_Set(x, y)
            
            If Not EngineRun Then
                Call SetConnected
            Else
                frmMain.Visible = True
                frmIniciando.Visible = False
            End If
            
            If VerLugar = 0 Then frmMain.Label2(0).Caption = Locale_GUI_Frase(170) & ": " & CurrentUser.MapNum & ", " & x & ", " & y
            
            If (CurrentUser.CurrentCharIndex > 0 Or CurrentUser.AccountCharCount > 0 Or CurrentUser.bShowLetter = True) Then
                Call frmMain.Engine.Map_Letter_Fade_Set(frmMain.Engine.Map_Letter_GRH_Get)
            Else
                Call frmMain.Engine.Map_Letter_Fade_Set(frmMain.Engine.Map_Letter_GRH_Get, 30093)
                CurrentUser.bShowLetter = True
            End If
            
            General_Update_Minimap x, y
            
            If CurrentUser.Spectate Then
                CurrentUser.Spectate = False
                CurrentUser.CurrentSpeed = VelNormal
                frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
            End If
            
        Case Create_Char_NPC
            Call frmMain.Engine.Char_Create(String_To_Integer(rData, 8), String_To_Integer(rData, 10), String_To_Byte(rData, 5), String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 16), String_To_Integer(rData, 12), String_To_Integer(rData, 14), vbNullString, _
                0, 0, 0, String_To_Integer(rData, 6), _
                0, 0, String_To_Long(rData, 18), String_To_Long(rData, 22), 0, fxgr_extra(), fxgr_type(), fxgr_pos(), 0)
        
        Case Change_Heading_North
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Heading_Set(CharIndex, NORTH)
        
        Case Change_Heading_South
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Heading_Set(CharIndex, SOUTH)
        
        Case Change_Heading_East
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Heading_Set(CharIndex, EAST)
        
        Case Change_Heading_West
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Heading_Set(CharIndex, WEST)
        
        Case Change_Type
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Type_Set(CharIndex, String_To_Byte(rData, 3))
        
            If CharIndex = CurrentUser.CurrentChar Then _
                CurrentUser.CharType = String_To_Byte(rData, 3)
        
        Case Change_Name
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            tempStr = mid$(rData, 3)
            Call frmMain.Engine.Char_Name_Set(CharIndex, tempStr)
            
            If CharIndex = CurrentUser.CurrentChar Then _
                frmMensaje.mnuClan.Enabled = (InStr(tempStr, "<") > 0)
                        
        Case Change_Group
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Group_Set(CharIndex, String_To_Integer(rData, 3))
        
        Case Erase_Char
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Remove(CharIndex)
                        
        Case Guild_Info
        
        Case Move_Char_North
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call frmMain.Engine.Char_Move(CharIndex, NORTH)
            If frmMain.Engine.Char_Current_Pos_Check = False And CurrentUser.Spectate = False Then Call ClientTCP.Send_Data(Request_Pos)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Move_Char_East
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call frmMain.Engine.Char_Move(CharIndex, EAST)
            If frmMain.Engine.Char_Current_Pos_Check = False And CurrentUser.Spectate = False Then Call ClientTCP.Send_Data(Request_Pos)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Move_Char_South
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call frmMain.Engine.Char_Move(CharIndex, SOUTH)
            If frmMain.Engine.Char_Current_Pos_Check = False And CurrentUser.Spectate = False Then Call ClientTCP.Send_Data(Request_Pos)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Move_Char_West
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call frmMain.Engine.Char_Move(CharIndex, WEST)
            If frmMain.Engine.Char_Current_Pos_Check = False And CurrentUser.Spectate = False And CurrentUser.Spectate = False Then Call ClientTCP.Send_Data(Request_Pos)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Change_Char
        
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            
            Call frmMain.Engine.Char_Change(CharIndex, String_To_Byte(rData, 7), _
            String_To_Integer(rData, 3), String_To_Integer(rData, 5), String_To_Integer(rData, 12), _
            String_To_Integer(rData, 8), String_To_Integer(rData, 10), 0, _
            0)
            
        Case Item_Create
            x = String_To_Integer(rData, 5)
            y = String_To_Integer(rData, 7)
            
            If Len(rData) = 8 Then
                Call frmMain.Engine.Map_Item_Grh_Add(x, y, String_To_Long(rData, 1))
            Else
                tempByte = String_To_Byte(rData, 9)
                tempLng = String_To_Long(rData, 10)
                Call frmMain.Engine.Map_Item_Grh_Add(x, y, String_To_Long(rData, 1), tempByte, tempLng)
            End If
        
        Case Item_Erase
            x = String_To_Integer(rData, 1)
            y = String_To_Integer(rData, 3)
            Call frmMain.Engine.Map_Item_Grh_Remove(x, y)
        
        Case Pos_Bloq
            Call frmMain.Engine.Map_Blocked_Set(String_To_Integer(rData, 1), String_To_Integer(rData, 3), String_To_Byte(rData, 5))

        Case Music_Change
            tempInt = String_To_Integer(rData, 1)
            Sound.LastMapMusic = tempInt
            
            If (sMusica <> CONST_DESHABILITADA) And (DefMidi = 1) Then
                If tempInt <> 0 Then
                    Sound.NextMusic = tempInt
                    Sound.Fading = 200
                End If
            End If
                    
        Case Sound_Play
            If Audio = 1 Then
                tempInt = String_To_Integer(rData, 1)
                x = String_To_Integer(rData, 3)
                y = String_To_Integer(rData, 5)
                
                If frmMain.Engine.Map_In_Bounds(x, y) = False Then
                    Call Sound.Sound_Play(tempInt)
                Else
                    Call Sound.Sound_Play(tempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
                End If
            End If
        
        Case Main_Message
            tempByte = String_To_Byte(rData, 1)
            If tempByte = 0 Then Exit Sub
            
            Select Case tempByte
                Case Server_Message_ID.Failed
                    tempStr = Locale_Parse_ServerMessage(tempByte)
                    Call frmMain.Engine.Char_Dialog_Set(CurrentUser.CurrentChar, tempStr, COLOR_ATAQUE, 5, True, 2)
                    Exit Sub
                Case Server_Message_ID.NPC_Failed
                    tempStr = Locale_Parse_ServerMessage(tempByte)
                    Call frmMain.Engine.Char_Dialog_Set(frmMain.Engine.Char_Find(String_To_Integer(rData, 2)), tempStr, COLOR_ATAQUE, 5, True, 2)
                    Exit Sub
                Case Server_Message_ID.Cured
                    Call frmMain.Engine.Char_FX_Set(CurrentUser.CurrentChar, FXCURAR, 1)
                    Call Sound.Sound_Play(SND_CURAR)
                Case Server_Message_ID.Revived
                    Call General_Char_Particle_Create(PARESUCITAR, CurrentUser.CurrentChar, 0, 100)
                    Call Sound.Sound_Play(SND_RESUCITAR)
            End Select
            
            If Len(rData) > 1 Then
                tempStr = Locale_Parse_ServerMessage(tempByte, mid$(rData, 2))
            Else
                tempStr = Locale_Parse_ServerMessage(tempByte)
            End If
            
            Call PrintToConsole(tempStr, 204, 193, 115, 0, 1)
            
        Case Atack_FX
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))  'CharIndex Attacker
            tempInt = String_To_Integer(rData, 3) 'WavOver
            
            'DoFX OverChar (Attacker)
            'TODO_BARRIN
            
            If Len(rData) > 4 Then 'WavOver Receiver?
                CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 5))  'CharIndex Receiver
            End If
            
            If frmMain.Engine.Char_Pos_Get(CharIndex, x, y) Then
                Call Sound.Sound_Play(tempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
            End If
            
        Case Guild_List
            Call frmGuildAdm.ParseGuildList(UncompressString(rData))
            
        Case Commerce_Data
            Call frmComercioSeguro.ParseData(rData)
        
        Case Main_Pause
            CurrentUser.Pausa = Not CurrentUser.Pausa
            
            If CurrentUser.Pausa Then
                frmMain.Engine.Char_Current_Direction_Set (0)
            Else
                Call frmMain.Check_Main_Keys
            End If
            
        Case Client_Validation
            bRK = String_To_Long(rData, 1)
            bK = CStr(String_To_Long(rData, 5) + CLng(Month(Date)) + CLng(Year(Date)) + bRK)
            Call Login(ValidarLoginMSG(CInt(bRK)))
        
        Case Char_List
            ReDim ListaPersonajes(0 To MAX_CHAR_ACCOUNTS)
            frmCharList.ClearChars
            
            CurrentUser.AccountCharCount = 0
            
            If LenB(rData) > 0 Then
                    tempLng = 1
                    i = 1
                    Do While tempLng < Len(rData) And i <= 10
                        ListaPersonajes(i).char_level = String_To_Integer(rData, tempLng)
                        ListaPersonajes(i).char_clase = String_To_Byte(rData, tempLng + 2)
                        ListaPersonajes(i).char_map = String_To_Integer(rData, tempLng + 3)
                        ListaPersonajes(i).char_body = String_To_Integer(rData, tempLng + 5)
                        ListaPersonajes(i).char_head = String_To_Integer(rData, tempLng + 7)
                        ListaPersonajes(i).char_weapon = String_To_Integer(rData, tempLng + 9)
                        ListaPersonajes(i).char_shield = String_To_Integer(rData, tempLng + 11)
                        ListaPersonajes(i).char_helmet = String_To_Integer(rData, tempLng + 13)
                        ListaPersonajes(i).char_familiar = String_To_Integer(rData, tempLng + 15)
                        ListaPersonajes(i).char_type = String_To_Byte(rData, tempLng + 17)
                        
                        tempLng = tempLng + 18
                        tempByte = String_To_Byte(rData, tempLng)
                        ListaPersonajes(i).char_name = mid$(rData, tempLng + 1, tempByte)
                        tempLng = tempLng + tempByte + 1
                        
                        frmCharList.lblAccData(i) = ListaPersonajes(i).char_name
                        
                        frmCharList.lblAccData(i).ForeColor = frmMain.Engine.Char_Color_Simple_Get_Ex(ListaPersonajes(i).char_type)
                        
                        i = i + 1
                    Loop
            End If
            
            CurrentUser.AccountCharCount = i - 1
            If CurrentUser.AccountCharCount < 0 Then CurrentUser.AccountCharCount = 0
            
            Call FormParser.Parse_Form(frmCharList)
                        
            frmCharList.lblAccData(0).Caption = CurrentUser.AccountName
                        
            If Not frmCharList.Visible Then
                frmCharList.Show
            Else
                Call frmCharList.PaintChars
            End If

            If Not frmConnect Is Nothing Then
                frmConnect.NameTxt = vbNullString
                frmConnect.PwdTxt = vbNullString
                frmConnect.Visible = False
            End If
            
        Case Rain_Toggle
            If Meteo_Engine.SecondaryStatus = 2 Or Meteo_Engine.SecondaryStatus = 4 Then
                Meteo_Engine.SecondaryStatus = 0
            Else
                If LenB(rData) = 0 Then
                    Meteo_Engine.SecondaryStatus = 2
                Else
                    Meteo_Engine.SecondaryStatus = 4
                End If
            End If
        
        Case Snow_Toggle
            If Meteo_Engine.SecondaryStatus = 3 Then
                Meteo_Engine.SecondaryStatus = 0
            Else
                Meteo_Engine.SecondaryStatus = 3
            End If
        
        Case Hour_Night_Toggle
            tempByte = String_To_Byte(rData, 1)
            x = String_To_Integer(rData, 2)
            Call Meteo_Engine.SetNuevoEstado(tempByte)
            If x <> 0 Then Call Sound.Sound_Play(x)

        Case Dialog_Erase
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Dialog_Remove(CharIndex)

        Case FX_Create
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            tempInt = String_To_Integer(rData, 7)
            Call frmMain.Engine.Char_FX_Set(CharIndex, String_To_Integer(rData, 3), String_To_Integer(rData, 5))
            
            If tempInt <> 0 Then
                If frmMain.Engine.Char_Pos_Get(CharIndex, x, y) Then
                    Call Sound.Sound_Play(tempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
                End If
            End If
        
        Case Map_FX_Create
            x = String_To_Integer(rData, 3)
            y = String_To_Integer(rData, 5)
            tempInt = String_To_Integer(rData, 7)
            Call frmMain.Engine.Map_Fx_Add(x, y, String_To_Integer(rData, 1))
            If tempInt <> 0 Then Call Sound.Sound_Play(tempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
        
        Case FXGroup_Char_Add
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            tempInt = String_To_Integer(rData, 9)
            tempByte = String_To_Byte(rData, 11) 'Type
            tmpBt = String_To_Byte(rData, 12) 'Pos
            
            If tempByte = EFECTO_PARTICULA Then
                Call General_Char_Particle_Create(String_To_Long(rData, 3), CharIndex, tmpBt, CLng(String_To_Integer(rData, 7)))
            Else
                Call frmMain.Engine.Char_FX_Group_Create(CharIndex, String_To_Long(rData, 3), String_To_Long(rData, 13), tmpBt)
            End If
            
            If tempInt <> 0 Then
                If frmMain.Engine.Char_Pos_Get(CharIndex, x, y) Then
                    Call Sound.Sound_Play(tempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
                End If
            End If

        Case Particle_Map_Add
            x = String_To_Integer(rData, 5)
            y = String_To_Integer(rData, 7)
            Call General_Particle_Create(String_To_Integer(rData, 1), x, y, CLng(String_To_Integer(rData, 3)))
            tempInt = String_To_Integer(rData, 9)
            If tempInt <> 0 Then Call Sound.Sound_Play(tempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))

        Case Particle_Map_Remove
            x = String_To_Integer(rData, 1)
            y = String_To_Integer(rData, 3)
            tempLng = frmMain.Engine.Map_Particle_Group_Get(x, y)
            If tempLng > 0 Then frmMain.Engine.Particle_Group_Remove (tempLng)
            
        Case FXGroup_Char_Remove
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_FX_Group_Remove(CharIndex, String_To_Integer(rData, 3))

        Case FXGroup_All_Char_Remove
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_FX_Group_Remove_All(CharIndex)

        Case GM_Panel
            frmPanelGm.Show vbModeless, frmMain
            
        Case User_Stats
            Call ActualizarEst(String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 5), String_To_Integer(rData, 7), _
                String_To_Integer(rData, 9), String_To_Integer(rData, 11), _
                String_To_Long(rData, 13), String_To_Integer(rData, 17), _
                String_To_Long(rData, 19), String_To_Long(rData, 23), _
                String_To_Integer(rData, 27), String_To_Integer(rData, 29), String_To_Integer(rData, 31), String_To_Integer(rData, 33), String_To_Integer(rData, 35), String_To_Integer(rData, 37), True)
        
        Case Quest_Info
            If rData = vbNullString Then
                Call PrintToConsole(Locale_GUI_Frase(352), 204, 193, 115, 0, 1)
            Else
                frmQuestActual.Show vbModeless, frmMain
                frmQuestActual.ParseQuestInfo (rData)
            End If
        
        Case Working_Main
            CurrentUser.UsingSkill = String_To_Byte(rData, 1)
            
            Select Case CurrentUser.UsingSkill
                Case Magia
                    Call FormParser.Parse_Form(frmMain, E_CAST)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                Case Robar
                    Call FormParser.Parse_Form(frmMain, E_SHOOT)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call FormParser.Parse_Form(frmMain, E_ARROW)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                '[Barrin]
                Case Arrojadizas
                    Call FormParser.Parse_Form(frmMain, E_ATTACK)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                Case Esposas
                    Call FormParser.Parse_Form(frmMain, E_SHOOT)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                Case Musica
                    Call FormParser.Parse_Form(frmMain, E_SHOOT)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                Case Grupo
                    Call FormParser.Parse_Form(frmMain, E_SHOOT)
                    Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)
                '[/Barrin]
            End Select
            
        Case Inventory_Update, Inventory_Update_Localized
            tempByte = String_To_Byte(rData, 1)
            
            If tempByte <> 255 Then
                UserInventory(tempByte).ObjIndex = String_To_Integer(rData, 2)
                UserInventory(tempByte).Amount = String_To_Integer(rData, 4)
                UserInventory(tempByte).Equipped = String_To_Byte(rData, 6)
                UserInventory(tempByte).GrhIndex = String_To_Long(rData, 7)
                UserInventory(tempByte).ObjType = String_To_Integer(rData, 11)
                UserInventory(tempByte).MaxHIT = String_To_Integer(rData, 13)
                UserInventory(tempByte).MinHIT = String_To_Integer(rData, 15)
                UserInventory(tempByte).Def = String_To_Integer(rData, 17)
                UserInventory(tempByte).Valor = String_To_Long(rData, 19)
                UserInventory(tempByte).PuedeUsar = String_To_Byte(rData, 23)
                
                If Paquet_ID = Inventory_Update Then
                    tmpBt = String_To_Byte(rData, 24)
                    UserInventory(tempByte).name = mid$(rData, 25, tmpBt)
                    
                    If 25 + tmpBt <> Len(rData) Then
                        UserInventory(tempByte).ExtraStr = mid$(rData, 25 + tmpBt)
                    Else
                        UserInventory(tempByte).ExtraStr = vbNullString
                    End If
                Else
                
                    tempInt = String_To_Integer(rData, 24)
                
                    UserInventory(tempByte).name = Locale_UserItem(tempInt)
                    UserInventory(tempByte).ExtraStr = Locale_UserItem_Desc(tempInt)
                    
                End If
                
            Else
                tempByte = String_To_Byte(rData, 2)
                UserInventory(tempByte).name = Locale_GUI_Frase(269)
                UserInventory(tempByte).ObjIndex = 0
                UserInventory(tempByte).Amount = 0
                UserInventory(tempByte).Equipped = 0
                UserInventory(tempByte).GrhIndex = 0
                UserInventory(tempByte).ObjType = 0
                UserInventory(tempByte).MaxHIT = 0
                UserInventory(tempByte).MinHIT = 0
                UserInventory(tempByte).Def = 0
                UserInventory(tempByte).PuedeUsar = 0
                UserInventory(tempByte).Valor = 0
                UserInventory(tempByte).ExtraStr = vbNullString
            End If
            
            If CurrentUser.Logged Then
                frmMain.Engine.Engine_Inventory_Render_Set
                
                If frmComerciar.Visible Then
                    frmComerciar.List1(1).List(tempByte - 1) = UserInventory(tempByte).name
                    Call frmComerciar.Refresh_User_Inventory
                ElseIf frmBancoObj.Visible Then
                    frmBancoObj.List1(1).List(tempByte - 1) = UserInventory(tempByte).name
                    Call frmBancoObj.Refresh_User_Inventory
                End If
            End If
            
        Case Bank_Update, Bank_Update_Localized
            tempByte = String_To_Byte(rData, 1)
            
            If tempByte <> 111 Then
                UserBancoInventory(tempByte).ObjIndex = String_To_Integer(rData, 2)
                UserBancoInventory(tempByte).Amount = String_To_Long(rData, 4)
                UserBancoInventory(tempByte).GrhIndex = String_To_Long(rData, 8)
                UserBancoInventory(tempByte).ObjType = String_To_Integer(rData, 12)
                UserBancoInventory(tempByte).MaxHIT = String_To_Integer(rData, 14)
                UserBancoInventory(tempByte).MinHIT = String_To_Integer(rData, 16)
                UserBancoInventory(tempByte).Def = String_To_Integer(rData, 18)
                                
                If Paquet_ID = Bank_Update Then
                    tmpBt = String_To_Byte(rData, 20)
                    UserBancoInventory(tempByte).name = mid$(rData, 21, tmpBt)
                    
                    If Len(rData) > 21 + tmpBt Then
                        UserBancoInventory(tempByte).ExtraStr = mid$(rData, 21 + tmpBt)
                    Else
                        UserBancoInventory(tempByte).ExtraStr = vbNullString
                    End If
                Else
                
                    tempInt = String_To_Integer(rData, 20)
                
                    UserBancoInventory(tempByte).name = Locale_UserItem(tempInt)
                    UserBancoInventory(tempByte).ExtraStr = Locale_UserItem_Desc(tempInt)
                    
                End If
                
            Else
                tempByte = String_To_Byte(rData, 2)
                UserBancoInventory(tempByte).name = Locale_GUI_Frase(269)
                UserBancoInventory(tempByte).ObjIndex = 0
                UserBancoInventory(tempByte).Amount = 0
                UserBancoInventory(tempByte).GrhIndex = 0
                UserBancoInventory(tempByte).ObjType = 0
                UserBancoInventory(tempByte).MaxHIT = 0
                UserBancoInventory(tempByte).MinHIT = 0
                UserBancoInventory(tempByte).Def = 0
                UserBancoInventory(tempByte).ExtraStr = vbNullString
            End If
        
            If frmBancoObj.Visible Then
                frmBancoObj.List1(0).List(tempByte - 1) = UserBancoInventory(tempByte).name
                Call frmBancoObj.Refresh_Bank_Inventory
            End If
        
        Case Spell_Add
            tempByte = String_To_Byte(rData, 1)
            CurrentUser.UserHechizos(tempByte) = String_To_Integer(rData, 2)
            
            If tempByte > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem IIf(CurrentUser.UserHechizos(tempByte) = 0, Locale_GUI_Frase(269), Right$(rData, Len(rData) - 3))
            Else
                frmMain.hlst.List(tempByte - 1) = IIf(CurrentUser.UserHechizos(tempByte) = 0, Locale_GUI_Frase(269), Right$(rData, Len(rData) - 3))
            End If
            
        Case Spell_Add_Localized
            tempByte = String_To_Byte(rData, 1)
            CurrentUser.UserHechizos(tempByte) = String_To_Integer(rData, 2)
            tempInt = String_To_Integer(rData, 4)
            
            If tempByte > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem IIf(CurrentUser.UserHechizos(tempByte) = 0, Locale_GUI_Frase(269), Locale_Spell_Name(tempInt))
            Else
                frmMain.hlst.List(tempByte - 1) = IIf(CurrentUser.UserHechizos(tempByte) = 0, Locale_GUI_Frase(269), Locale_Spell_Name(tempInt))
            End If
        
        Case Att_Receive
            CurrentUser.UserAtributos(1) = String_To_Integer(rData, 1)
            CurrentUser.UserAtributos(2) = String_To_Integer(rData, 3)
            CurrentUser.UserAtributos(3) = String_To_Integer(rData, 5)
            CurrentUser.UserAtributos(4) = String_To_Integer(rData, 7)
            CurrentUser.UserAtributos(5) = String_To_Integer(rData, 9)
            LlegaronAtrib = True
            Call MostrarEstadisticas

        Case Weapon_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim ArmasHerrero(1 To tempInt) As Integer
            
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                ArmasHerrero(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmHerrero.lstArmas.AddItem tmpCad
            Next i
        
        Case Armor_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim ArmadurasHerrero(1 To tempInt) As Integer
                        
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                ArmadurasHerrero(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmHerrero.lstArmaduras.AddItem tmpCad
            Next i
        
        Case Shield_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim EscudosHerrero(1 To tempInt) As Integer
                        
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                EscudosHerrero(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmHerrero.lstEscudos.AddItem tmpCad
            Next i

        Case Helmet_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim CascosHerrero(1 To tempInt) As Integer
                        
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                CascosHerrero(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmHerrero.lstCascos.AddItem tmpCad
            Next i

        Case Potion_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim ObjDruida(1 To tempInt) As Integer
                        
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                ObjDruida(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmDruida.lstPociones.AddItem tmpCad
            Next i
        
        Case Tailor_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim ObjSastre(1 To tempInt) As Integer
                 
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                ObjSastre(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmSastre.lstRopas.AddItem tmpCad
            Next i
        
        Case Carpenter_List
            tempInt = Val(General_Field_Count(rData, 44) / 2)
            ReDim ObjCarpintero(1 To tempInt) As Integer
            
            tempLng = 0
            For i = 1 To tempInt
                tempLng = tempLng + 1
                If i <> 1 Then tempLng = tempLng + 1
                tmpCad = General_Field_Read(tempLng, rData, ",")
                ObjCarpintero(i) = Val(General_Field_Read(tempLng + 1, rData, ","))
                If tmpCad <> vbNullString Then frmCarp.lstArmas.AddItem tmpCad
            Next i
        
        Case Rest_Toggle
            
            If CurrentUser.Descansando Then
                Call RecSTTimer(True, CurrentUser.ISNormal)
                Call ClientTCP.Send_Data(Stats_Sync, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA))
                CurrentUser.Descansando = False
            Else
                Call RecSTTimer(True, CurrentUser.ISDescansar)
                Sound.Sound_Play (SND_DESCANSAR)
                CurrentUser.Descansando = True
            End If
            
        Case GM_Spawn_List
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmSpawnList.lstCriaturas.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmSpawnList.Show vbModeless, frmMain
                
        Case Skill_Improved
            tempByte = String_To_Byte(rData, 1)
            tempInt = String_To_Integer(rData, 2)
            
            If tempByte = Supervivencia Then
                CurrentUser.Supervivencia = tempInt
            ElseIf tempByte = Meditar Then
                CurrentUser.Meditar = tempInt
            End If
            
            Call PrintToConsole(Locale_GUI_Frase(354) & " " & SkillsNames(tempByte) & " " & Locale_GUI_Frase(355) & " " & tempInt & " " & Locale_GUI_Frase(356), 0, 0, 0, 0, 0, 0, 4)
        
        Case Error_Message, Main_Error_Message
            
            If Paquet_ID = Main_Error_Message Then
                rData = Locale_Error(String_To_Byte(rData, 1))
            End If
            
            If frmConnect.Visible Then
                
                Call ResetCurrentUser
                
                Call FormParser.Parse_Form(frmConnect)
                frmMensaje.msg.Caption = rData
                frmMensaje.Show vbModal, frmConnect
            
            ElseIf frmCharList.Visible Then
                
                Call FormParser.Parse_Form(frmCharList)
                frmMensaje.msg.Caption = rData
                frmMensaje.Show vbModal, frmCharList
            
            ElseIf frmIniciando.Visible Then
                
                Call FormParser.Parse_Form(frmCharList)
                frmCharList.Visible = True
                frmIniciando.Visible = False
                frmMensaje.msg.Caption = rData
                frmMensaje.Show vbModal, frmCharList
            
            ElseIf frmCrearPersonaje.Visible Then
                
                Call FormParser.Parse_Form(frmCrearPersonaje)
                frmMensaje.msg.Caption = rData
                frmMensaje.Show vbModal, frmCrearPersonaje
                
            End If
            
        Case Error_Message_Update
            Call FormParser.Parse_Form(frmConnect)
            If frmMain.MainWinsock.State Then frmMain.MainWinsock.Close
            Call ResetCurrentUser
            frmPregunta.msg.Caption = Locale_GUI_Frase(171)
            frmPregunta.Unload_And_Update = True
            frmPregunta.Show vbModal, frmConnect
            
        Case Bounty_List
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmHunter.lstBuscados.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmHunter.Show vbModeless, frmMain
        
        Case Group_List
            frmGrupo.lstGrupo.Clear
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmGrupo.lstGrupo.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            tempByte = General_Field_Read(Val(General_Field_Read(1, rData, ",")) + 2, rData, ",")
            If tempByte = 1 Then
                frmGrupo.cmdExpulsar.Enabled = True
                frmGrupo.cmdInvitar.Enabled = True
            Else
                frmGrupo.cmdExpulsar.Enabled = False
                frmGrupo.cmdInvitar.Enabled = False
            End If
            frmGrupo.Show vbModeless, frmMain
                
        Case Stupid_True
            CurrentUser.Estupido = True
        
        Case Letter_Show
            Call frmMain.Engine.Letter_Set(CLng(String_To_Integer(rData, 1)), Locale_UserItem_Desc(String_To_Integer(rData, 3)))
        
        Case NPC_Inv_Receive, NPC_Inv_Receive_Localized
            tempByte = String_To_Byte(rData, 1)
            
            If tempByte <> 111 Then
                NPCInventory(tempByte).Amount = String_To_Integer(rData, 2)
                NPCInventory(tempByte).Valor = String_To_Long(rData, 4)
                NPCInventory(tempByte).GrhIndex = String_To_Long(rData, 8)
                NPCInventory(tempByte).ObjIndex = String_To_Integer(rData, 12)
                NPCInventory(tempByte).ObjType = String_To_Integer(rData, 14)
                NPCInventory(tempByte).MaxHIT = String_To_Integer(rData, 16)
                NPCInventory(tempByte).MinHIT = String_To_Integer(rData, 18)
                NPCInventory(tempByte).Def = String_To_Integer(rData, 20)
                    
                If Paquet_ID = NPC_Inv_Receive Then
                    tmpBt = String_To_Byte(rData, 22)
                    NPCInventory(tempByte).name = mid$(rData, 23, tmpBt)
                    
                    If Len(rData) > 23 + tmpBt Then
                        NPCInventory(tempByte).ExtraStr = mid$(rData, 23 + tmpBt)
                    Else
                        NPCInventory(tempByte).ExtraStr = vbNullString
                    End If
                Else
                    tempInt = String_To_Integer(rData, 22)
                    NPCInventory(tempByte).name = Locale_UserItem(tempInt)
                    NPCInventory(tempByte).ExtraStr = Locale_UserItem_Desc(tempInt)
                End If
                
            Else
                tempByte = String_To_Byte(rData, 2)
                NPCInventory(tempByte).name = Locale_GUI_Frase(269)
                NPCInventory(tempByte).Amount = 0
                NPCInventory(tempByte).Valor = 0
                NPCInventory(tempByte).GrhIndex = 0
                NPCInventory(tempByte).ObjIndex = 0
                NPCInventory(tempByte).ObjType = 0
                NPCInventory(tempByte).MaxHIT = 0
                NPCInventory(tempByte).MinHIT = 0
                NPCInventory(tempByte).Def = 0
                NPCInventory(tempByte).ExtraStr = vbNullString
            End If
        
            If frmComerciar.Visible Then
                frmComerciar.List1(0).List(tempByte - 1) = NPCInventory(tempByte).name
                Call frmComerciar.Refresh_NPC_Inventory
            End If

        Case HYS_Update
            Call ActualizarEst(, , , , , , , , , , , , String_To_Integer(rData, 1), , String_To_Integer(rData, 3))
        
        Case Skill_Data
            CurrentUser.SkillPoints = String_To_Integer(rData, 1)
        
        Case Number_Enemys
            Call PrintToConsole(String_To_Integer(rData, 1) & " NPCs.", 255, 255, 255, 0, 0)
        
        Case GM_Sos_Message
            rData = UncompressString(rData)
            tempInt = InStr(1, rData, "µ")
                        
            If tempInt > 0 Then
                tempStr = left$(rData, tempInt - 1) & " (" & General_Field_Read(3, rData, "µ") & ")"
                frmPanelGm.List1.AddItem tempStr
                frmPanelGm.List1.ItemData(frmPanelGm.List1.NewIndex) = Val(General_Field_Read(4, rData, "µ"))
                frmPanelGm.MensajePoner tempStr, General_Field_Read(2, rData, "µ")
            Else
                frmPanelGm.List1.AddItem rData
                frmPanelGm.List1.ItemData(frmPanelGm.List1.NewIndex) = Val(General_Field_Read(4, rData, "µ"))
            End If
            
        Case GM_Sos_Show
            frmPanelGm.Show vbModeless, frmMain
                
        Case GameMaster_Message_Show
            If Val(rData) = 1 Then
                frmGMAyuda.Label2.Caption = Locale_GUI_Frase(172)
                frmGMAyuda.optConsulta.Item(6).value = True
            End If
            If Not frmGMAyuda.Visible Then
                frmGMAyuda.Show vbModeless, frmMain
                frmGMAyuda.txtMotivo.SetFocus
            End If
        
        Case Map_Show
            If Not frmMapa.Visible Then
                  frmMapa.Show vbModeless, frmMain
            Else
                  frmMapa.SetFocus
            End If
        
        Case Familiar_Info
            CurrentUser.UserPet.TieneFamiliar = String_To_Byte(rData, 1)
            
            If CurrentUser.UserPet.TieneFamiliar <> 0 Then
                CurrentUser.UserPet.ELU = String_To_Long(rData, 2)
                CurrentUser.UserPet.ELV = String_To_Integer(rData, 6)
                CurrentUser.UserPet.EXP = String_To_Long(rData, 8)
                CurrentUser.UserPet.MaxHP = String_To_Integer(rData, 12)
                CurrentUser.UserPet.MinHP = String_To_Integer(rData, 14)
                CurrentUser.UserPet.MinHIT = String_To_Integer(rData, 16)
                CurrentUser.UserPet.MaxHIT = String_To_Integer(rData, 18)
                tempStr = mid$(rData, 20)
                CurrentUser.UserPet.nombre = General_Field_Read(1, tempStr, ",")
                CurrentUser.UserPet.Abilidad = HabilidadToString(General_Field_Read(2, tempStr, ","))
            End If
            
            LlegoFami = True
            Call MostrarEstadisticas
        
        Case Goliath_Info
            Call frmGoliath.ParseBancoInfo(rData)
        
        Case Blind_True
            Call frmMain.Engine.Char_Current_Blind_Set(True)
        
        Case Stats_Mini
            CurrentUser.UserStats.ImperialesMatados = String_To_Long(rData, 1)
            CurrentUser.UserStats.RenegadosMatados = String_To_Long(rData, 5)
            CurrentUser.UserStats.ArmadasMatados = String_To_Long(rData, 9)
            CurrentUser.UserStats.MiliciasMatados = String_To_Long(rData, 13)
            CurrentUser.UserStats.CaosMatados = String_To_Long(rData, 17)
            CurrentUser.UserStats.RepublicanosMatados = String_To_Long(rData, 21)
            CurrentUser.UserStats.NPCsMatados = String_To_Long(rData, 25)
            CurrentUser.UserStats.Clase = String_To_Byte(rData, 29)
            CurrentUser.UserStats.TimesKilled = String_To_Long(rData, 30)
            CurrentUser.UserStats.Raza = String_To_Byte(rData, 34)
            CurrentUser.UserStats.Genero = String_To_Byte(rData, 35)
            LlegoEst = True
            Call MostrarEstadisticas
        
        Case Meditating_Toggle
            CurrentUser.Meditando = Not CurrentUser.Meditando
            
            If CurrentUser.Meditando Then
                CurrentUser.Intervalos.InicioMeditar = GetTickCount
            Else
                Call ClientTCP.Send_Data(Stats_Sync, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA))
            End If
            
        Case Working_Toggle
            CurrentUser.Trabajando = Not CurrentUser.Trabajando
            
            If LenB(rData) > 0 Then _
                Call PrintToConsole(rData, 0, 0, 0, 0, 0, 0, 4)
            
            If CurrentUser.Trabajando Then
                Call PrintToConsole(Locale_GUI_Frase(390), 0, 0, 0, 0, 0, 0, 4)
            Else
                Call PrintToConsole(Locale_GUI_Frase(391), 0, 0, 0, 0, 0, 0, 4)
            End If
            
        Case char_invisible
            CharIndex = frmMain.Engine.Char_Find(String_To_Integer(rData, 1))
            Call frmMain.Engine.Char_Invisible_Set(CharIndex, String_To_Byte(rData, 3))
                
        Case Familiar_Adopt_Show
            frmSeleccionFamiliar.Show vbModeless, frmMain
        
        Case Quest_Data
            frmQuest.Show vbModeless, frmMain
            frmQuest.ParseQuestInfo (rData)
        
        Case Blind_False
            Call frmMain.Engine.Char_Current_Blind_Set(False)
        
        Case Stupid_False
            CurrentUser.Estupido = False
        
        Case Skills_Rec
            tempLng = -1
            
            For i = 1 To NUMSKILLS
                tempLng = tempLng + 2
                CurrentUser.UserSkills(i) = String_To_Integer(rData, CByte(tempLng))
            Next i
            
            LlegaronSkills = True
            Call MostrarEstadisticas
        
        Case Train_NPC_List
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmEntrenador.lstCriaturas.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmEntrenador.Show vbModeless, frmMain

        Case GM_User_List
            t = Split(UncompressString(rData), ",")
            frmPanelGm.cboListaUsus.Clear
            For i = LBound(t) To UBound(t)
                frmPanelGm.cboListaUsus.AddItem t(i)
            Next i
            If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
        
        Case Freezed_Toggle
            CurrentUser.Paralizado = Not CurrentUser.Paralizado
        
            If CurrentUser.Paralizado Then
                x = String_To_Integer(rData, 1)
                y = String_To_Integer(rData, 3)
                Call frmMain.Engine.Char_Current_Pos_Refresh(x, y)
                Call frmMain.Engine.Engine_View_Pos_Set(x, y)
            End If
        
        Case Guild_News
            Call frmGuildNews.ParseGuildNews(rData)
        
        Case Guild_Petition
            Call frmUserRequest.recievePeticion(rData)
                
        Case Guild_Char_Info
            Call frmCharInfo.parseCharInfo(rData)
        
        Case Bounty_Info
            Call frmBuscadoInfo.parseBuscadoInfo(rData)
        
        Case Guild_Lead_Info
            Call frmGuildLeader.ParseLeaderInfo(rData)
        
        Case Guild_Clan_Details
            Call frmGuildBrief.ParseGuildInfo(rData)
        
        Case Guild_Found_Show
            CurrentUser.CreandoClan = True
            Call frmGuildFoundation.Show(vbModeless, frmMain)
            frmGuildFoundation.txtClanName = rData
            frmGuildFoundation.txtClanName.Enabled = (Len(rData) = 0)
            
        Case Metam_Toggle
            CurrentUser.Transformado = Not CurrentUser.Transformado
            Call frmMain.Engine.Char_Current_OnHorse_Set(CurrentUser.Transformado)
        
        Case Challenge_Main
            Call frmTorneo.Parse_Torneo_Info(rData)
        
        Case Earthshock_Set
            tempInt = String_To_Integer(rData, 1)
            tempLng = String_To_Long(rData, 3)
            Call frmMain.Engine.Engine_EarthShock_Set(CSng(tempInt), tempLng)
            
        Case Compressed_Data
            rData = UncompressString(rData)
            Call Handle_Data(rData)
                
    End Select

End Sub

Public Sub Parse_Command_Str(ByVal strCmd As String)

On Error Resume Next

Dim strArg As String
Dim strCompa As String
Dim bFound As Boolean

If Len(strCmd) > 255 Then strCmd = mid$(strCmd, 1, 255)

bFound = True
strCompa = UCase$(strCmd)

Select Case mid$(strCompa, 2)
    Case Locale_CMD_Get(1)
        If CurrentUser.PingRequested Then Exit Sub
        CurrentUser.PingRequested = True
        CurrentUser.Ping = GetTickCount
        Call Send_Data_Command(cmdPing)
    Case Locale_CMD_Get(2)
        Call Send_Data_Command(cmdSalir, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA) & Byte_To_String(0))
    Case Locale_CMD_Get(3)
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdMeditar)
    Case Locale_CMD_Get(4)
        Call Send_Data_Command(cmdOnline)
    Case Locale_CMD_Get(5)
        Call Send_Data_Command(cmdOnlineClan)
    Case Locale_CMD_Get(6)
        Call Send_Data_Command(cmdGM)
    Case Locale_CMD_Get(7)
        Call Send_Data_Command(cmdEst, Integer_To_String(CurrentUser.UserMinMAN) & Integer_To_String(CurrentUser.UserMinSTA))
    Case Locale_CMD_Get(8)
        Call Send_Data_Command(cmdSubasta)
    Case Locale_CMD_Get(9)
        Call Send_Data_Command(cmdFuerte)
    Case Locale_CMD_Get(10)
        Call Send_Data_Command(cmdFami)
    Case Locale_CMD_Get(11)
        Call Send_Data_Command(cmdAyuda)
    Case Locale_CMD_Get(12)
        Call Send_Data_Command(cmdBoveda)
    Case Locale_CMD_Get(13)
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdComerciar)
    Case Locale_CMD_Get(14)
        Call Send_Data_Command(cmdCuanto)
    Case Locale_CMD_Get(15)
        Call Send_Data_Command(cmdRechazarMatrimonio)
    Case Locale_CMD_Get(16)
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdAcompañar)
    Case Locale_CMD_Get(17)
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdInformacion)
    Case Locale_CMD_Get(18)
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdDuelo)
    Case Locale_CMD_Get(19)
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdBalance)
    Case Locale_CMD_Get(20)
        Call Send_Data_Command(cmdFundarClan)
    Case Locale_CMD_Get(21)
        Call Send_Data_Command(cmdSalirClan)
    Case Locale_CMD_Get(22)
        Call Send_Data_Command(cmdDivorcio)
    Case Locale_CMD_Get(23)
        Call Send_Data_Command(cmdPena)
    Case Else
        bFound = False
End Select

If bFound Then Exit Sub

bFound = True
strCompa = UCase$(General_Field_Read(1, strCmd, " "))

Select Case mid$(strCompa, 2)
    Case Locale_CMD_Get(24)
        CurrentUser.SendingType = 5
        Call Send_Data_Command_GM(cmdGMSG, Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1)))
    
    Case Locale_CMD_Get(25)
        If DeadCheck Then Exit Sub
        CurrentUser.SendingType = 4
        Call Send_Data_Command(cmdCMSG, Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1)))
    
    Case Locale_CMD_Get(26)
        If DeadCheck Then Exit Sub
        CurrentUser.SendingType = 6
        Call Send_Data_Command(cmdGRMG, Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1)))
    
    Case Locale_CMD_Get(27)
        If DeadCheck Then Exit Sub
        CurrentUser.SendingType = 8
        Call Send_Data_Command(cmdFMMG, Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1)))
    
    Case Locale_CMD_Get(28)
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        If strArg <> vbNullString Then Call frmOpciones.AgregarIgnorado(strArg)
    
    Case Locale_CMD_Get(29)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdRetirar, strArg)
    
    Case Locale_CMD_Get(30)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdOfertar, strArg)
    
    Case Locale_CMD_Get(31)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdOfertaInicial, strArg)
    
    Case Locale_CMD_Get(32)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdDepositar, strArg)
    
    Case Locale_CMD_Get(33)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdApostar, Long_To_String(strArg))
    
    Case Locale_CMD_Get(34)
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdBug, strArg)
    
    Case Locale_CMD_Get(35)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdProponer, strArg)
    
    Case Locale_CMD_Get(36)
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        Call Send_Data_Command(cmdCentinela, Integer_To_String(strArg))
    
    Case Locale_CMD_Get(37)
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - (Len(strCompa) + 1))
        
        If Not AsciiValidos(strArg) Then
            Call PrintToConsole(Locale_GUI_Frase(392), 0, 0, 0, 0, 0, 0, 4)
        ElseIf Len(strArg) > 50 Then
            Call PrintToConsole(Locale_GUI_Frase(393), 0, 0, 0, 0, 0, 0, 4)
        Else
            Call Send_Data_Command(cmdDesc, strArg)
        End If
                        
    Case Else
        CurrentUser.SendingType = 1
        bFound = False
End Select

If bFound Then Exit Sub

bFound = True
strCompa = UCase$(strCmd)

Select Case strCompa
    Case "/INVISIBLE"
        Call Send_Data_Command_GM(cmdInvisible)
    
    Case "/TRABAJANDO"
        Call Send_Data_Command_GM(cmdTrabajando)
    
    Case "/ENGRUPO"
        Call Send_Data_Command_GM(cmdEnGrupo)
    
    Case "/SHOWSOS"
        Call Send_Data_Command_GM(cmdPanelGM)
    
    Case "/PANELGM"
        Call Send_Data_Command_GM(cmdPanelGM)
        
    Case "/ONLINEGM"
        Call Send_Data_Command_GM(cmdOnlineGM)
    
    Case "/CS"
        Call Send_Data_Command_GM(cmdCS)
    
    Case "/CTO"
        Call Send_Data_Command_GM(cmdCTO)
        
    Case "/UNBANGEN"
        Call Send_Data_Command_GM(cmdUnBanGen)
    
    Case "/CHATGEN"
        Call Send_Data_Command_GM(cmdChatGen)
    
    Case "/SEGUIR"
        Call Send_Data_Command_GM(cmdSeguir)
    
    Case "/LIMPIAR"
        Call Send_Data_Command_GM(cmdLimpiar)
    
    Case "/BANIPLIST"
        Call Send_Data_Command_GM(cmdBanIPList)
    
    Case "/DT"
        Call Send_Data_Command_GM(cmdDT)
    
    Case "/DEST"
        Call Send_Data_Command_GM(cmdDest)
    
    Case "/BLOQ"
        Call Send_Data_Command_GM(cmdBloq)
    
    Case "/MASSKILL"
        Call Send_Data_Command_GM(cmdMassKill)
    
    Case "/MATA"
        Call Send_Data_Command_GM(cmdMata)
    
    Case "/CC"
         Call Send_Data_Command_GM(cmdCC)
    
    Case "/APAGAR"
        Call Send_Data_Command_GM(cmdApagar)
    
    Case "/GUARDARMAPA"
        Call Send_Data_Command_GM(cmdGuardarMapa)
    
    Case "/REINICIAR"
        Call Send_Data_Command_GM(cmdReiniciar)
    
    Case "/DOBACKUP"
        Call Send_Data_Command_GM(cmdDoBackUP)
    
    Case "/GRABAR"
        Call Send_Data_Command_GM(cmdGrabar)
    
    Case "/BORRARSOS"
        Call Send_Data_Command_GM(cmdBorrarSos)
    
    Case "/SHOWINT"
        Call Send_Data_Command_GM(cmdShowInt)
    
    Case "/LLUVIA"
        Call Send_Data_Command_GM(cmdLluvia)
    
    Case "/NIEVE"
        Call Send_Data_Command_GM(cmdNieve)
    
    Case "/ONLINEMAP"
        Call Send_Data_Command_GM(cmdOnlineMap)
    
    Case "/MODOROL"
        Call Send_Data_Command_GM(cmdModoRol)
    
    Case "/MASSDEST"
        Call Send_Data_Command_GM(cmdMassDest)
    
    Case "/MD5OK"
        Call Send_Data_Command_GM(cmdMD5OK)
    
    Case "/NOCREAR"
        Call Send_Data_Command_GM(cmdNoCrear)
    
    Case Else
        bFound = False
End Select

If bFound Then Exit Sub

bFound = True
strCompa = UCase$(General_Field_Read(1, strCmd, " "))
strArg = mid$(strCmd, Len(strCompa) + 2)

Select Case strCompa
    Case "/REM"
        Call Send_Data_Command_GM(cmdRem, strArg)
    
    Case "/DONDE"
        Call Send_Data_Command_GM(cmdDonde, strArg)
    
    Case "/SETA" '1.6 centro
        frmMain.Engine.SetSetAngle = CSng(strArg)
    
    Case "/NENE"
        Call Send_Data_Command_GM(cmdNene, strArg)
    
    Case "/ATT"
        Call Send_Data_Command_GM(cmdAtt, strArg)
    
    Case "/TELEP"
        Call Send_Data_Command_GM(cmdTelep, strArg)
    
    Case "/IRA"
        Call Send_Data_Command_GM(cmdIra, strArg)
    
    Case "/TP"
        Call Send_Data_Command_GM(cmdTeleploc, strArg)
    
    Case "/SILENCIO"
        Call Send_Data_Command_GM(cmdSilencio, strArg)
    
    Case "/ECHAR"
        Call Send_Data_Command_GM(cmdEchar, strArg)
    
    Case "/PATEAR"
        Call Send_Data_Command_GM(cmdPatear, strArg)
    
    Case "/ONSTAT"
        Call Send_Data_Command_GM(cmdOnStat, Byte_To_String(strArg))
    
    Case "/ONFAC"
        Call Send_Data_Command_GM(cmdOnFac, Byte_To_String(strArg))
    
    Case "/MODOGUERRA"
        Call Send_Data_Command_GM(cmdModoGuerra, Integer_To_String(strArg))
    
    Case "/CS"
        Call Send_Data_Command_GM(cmdCS, strArg)

    Case "/APAGAR"
        Call Send_Data_Command_GM(cmdApagar, strArg)

    Case "/CTO"
        Call Send_Data_Command_GM(cmdCTO, strArg)
    
    Case "/BANINFO"
        Call Send_Data_Command_GM(cmdBanInfo, strArg)
    
    Case "/BANCHARINFO"
        Call Send_Data_Command_GM(cmdBanCharInfo, strArg)
    
    Case "/REVIVIR"
        Call Send_Data_Command_GM(cmdRevivir, strArg)
    
    Case "/LASTIP"
        Call Send_Data_Command_GM(cmdLastIP, strArg)
        
    Case "/INFO"
        Call Send_Data_Command_GM(cmdInfo, strArg)
    
    Case "/INFOACC"
        Call Send_Data_Command_GM(cmdInfoAcc, strArg)
    
    Case "/INFOFAMI"
        Call Send_Data_Command_GM(cmdInfoFami, strArg)
    
    Case "/INV"
        Call Send_Data_Command_GM(cmdInv, strArg)
    
    Case "/BOV"
        Call Send_Data_Command_GM(cmdBov, strArg)
    
    Case "/MAIL"
        Call Send_Data_Command_GM(cmdMail, strArg)
    
    Case "/HMR"
        Call Send_Data_Command_GM(cmdHmr, strArg)
    
    Case "/SKILLS"
        Call Send_Data_Command_GM(cmdSkills, strArg)
    
    Case "/CARCEL"
        Call Send_Data_Command_GM(cmdCarcel, strArg)
    
    Case "/BAN"
        Call Send_Data_Command_GM(cmdBan, strArg)
    
    Case "/BANCHAR"
        Call Send_Data_Command_GM(cmdBanChar, strArg)
    
    Case "/BANACC"
        Call Send_Data_Command_GM(cmdBanAcc, strArg)
    
    Case "/AUTORIZAR"
        Call Send_Data_Command_GM(cmdAutorizar, strArg)
    
    Case "/UNBAN"
        Call Send_Data_Command_GM(cmdUnBan, strArg)
    
    Case "/UNBANCHAR"
        Call Send_Data_Command_GM(cmdUnBanChar, strArg)
    
    Case "/SUM"
        Call Send_Data_Command_GM(cmdSum, strArg)
    
    Case "/RMSG"
        Call Send_Data_Command_GM(cmdRMSG, strArg)
    
    Case "/AMSG"
        Call Send_Data_Command_GM(cmdAMSG, strArg)
    
    Case "/NICK2IP"
        Call Send_Data_Command_GM(cmdNick2IP, strArg)
    
    Case "/SAMEIP"
        Call Send_Data_Command_GM(cmdSameIP, strArg)
    
    Case "/IP2NICK"
        Call Send_Data_Command_GM(cmdIP2Nick, strArg)
    
    Case "/CLEANMAP"
        Call Send_Data_Command_GM(cmdCleanMap, strArg)
    
    Case "/FACCION"
        Call Send_Data_Command_GM(cmdFaccion, strArg)
    
    Case "/PERDON"
        Call Send_Data_Command_GM(cmdPerdon, strArg)
    
    Case "/TRIGGER"
        Call Send_Data_Command_GM(cmdTrigger, strArg)
    
    Case "/BANIP"
        Call Send_Data_Command_GM(cmdBanIP, strArg)
    
    Case "/UNBANIP"
        Call Send_Data_Command_GM(cmdUnBanIP, strArg)
    
    Case "/CT"
        Call Send_Data_Command_GM(cmdCT, strArg)
    
    Case "/NICKPROHIBIDO"
        Call Send_Data_Command_GM(cmdNickProhibido, strArg)
    
    Case "/CI"
        Call Send_Data_Command_GM(cmdCI, strArg)
    
    Case "/SMSG"
        Call Send_Data_Command_GM(cmdSMSG, strArg)
    
    Case "/ACC"
        Call Send_Data_Command_GM(cmdAcc, strArg)
    
    Case "/RACC"
        Call Send_Data_Command_GM(cmdRacc, strArg)
    
    Case "/NAVE"
        Call Send_Data_Command_GM(cmdNave, strArg)
    
    Case "/CONDEN"
        
        Call Send_Data_Command_GM(cmdConden, strArg)
    
    Case "/MODMAPINFO"
        Call Send_Data_Command_GM(cmdModMapInfo, strArg)
    
    Case "/RELOAD"
        Call Send_Data_Command_GM(cmdReload, strArg)
    
    Case "/INICIAR"
        Call Send_Data_Command_GM(cmdIniciar, strArg)
    
    Case "/KILLAPP"
        Call Send_Data_Command_GM(cmdKilLApp, strArg)
    
    Case "/MOD"
        Call Send_Data_Command_GM(cmdMod, strArg)
    
    Case "/PLAYWAV"
        Call Send_Data_Command_GM(cmdPlayWav, strArg)
    
    Case "/MODNICK"
        Call Send_Data_Command_GM(cmdModNick, strArg)
    
    Case "/SQL_QUERY"
        Call Send_Data_Command_GM(cmdQuery, strArg)
    
    Case "/MODPASS"
        Call Send_Data_Command_GM(cmdModPass, strArg)
    
    Case "/DELCHAR"
        Call Send_Data_Command_GM(cmdDelChar, strArg)
    
    Case Else
        Call PrintToConsole(Locale_GUI_Frase(394), 0, 0, 0, 0, 0, 0, 4)
End Select

End Sub

Public Sub Send_Data(ByVal sndPaquet As Byte, Optional ByRef sndData As String = vbNullString)

Dim strToSend As String

strToSend = Integer_To_String(Len(sndData) + 1) & Chr$(sndPaquet) & sndData

If frmMain.MainWinsock.State = sckConnected Then
    'If LenB(bK) > 0 Then
    '    Call RC4_E(Integer_To_String(Len(sndData) + 1) & Chr$(sndPaquet) & sndData, CStr(bK), sndArrData)
    'Else
    '    Call RC4_E(Integer_To_String(Len(sndData) + 1) & Chr$(sndPaquet) & sndData, "gGuard~'!t0", sndArrData)
    'End If
        
    Call frmMain.MainWinsock.SendData(strToSend)
End If

End Sub

Public Sub Login(ByVal valCode As Integer)

Call Send_Data(Account_Login, CurrentUser.AccountName & "," & CurrentUser.UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valCode & MD5HushYo)

End Sub

Public Sub CreateNewChar(ByVal valCode As Integer, ByVal intHeadInd As Integer)

CurrentUser.CurrentCharIndex = 0

If CurrentUser.UserClase = MAGO Or CurrentUser.UserClase = DRUIDA Or CurrentUser.UserClase = CAZADOR Then
    Call Send_Data(New_Char, Integer_To_String(intHeadInd) _
    & Integer_To_String(App.Major) & Integer_To_String(App.Minor) & Integer_To_String(App.Revision) _
    & Byte_To_String(CurrentUser.UserRaza) & Byte_To_String(CurrentUser.UserSexo) & Byte_To_String(CurrentUser.UserClase) _
    & Integer_To_String(CurrentUser.UserAtributos(1)) & Integer_To_String(CurrentUser.UserAtributos(2)) & Integer_To_String(CurrentUser.UserAtributos(3)) _
    & Integer_To_String(CurrentUser.UserAtributos(4)) & Integer_To_String(CurrentUser.UserAtributos(5)) _
    & Integer_To_String(CurrentUser.UserSkills(1)) & Integer_To_String(CurrentUser.UserSkills(2)) _
    & Integer_To_String(CurrentUser.UserSkills(3)) & Integer_To_String(CurrentUser.UserSkills(4)) _
    & Integer_To_String(CurrentUser.UserSkills(5)) & Integer_To_String(CurrentUser.UserSkills(6)) _
    & Integer_To_String(CurrentUser.UserSkills(7)) & Integer_To_String(CurrentUser.UserSkills(8)) _
    & Integer_To_String(CurrentUser.UserSkills(9)) & Integer_To_String(CurrentUser.UserSkills(10)) _
    & Integer_To_String(CurrentUser.UserSkills(11)) & Integer_To_String(CurrentUser.UserSkills(12)) _
    & Integer_To_String(CurrentUser.UserSkills(13)) & Integer_To_String(CurrentUser.UserSkills(14)) _
    & Integer_To_String(CurrentUser.UserSkills(15)) & Integer_To_String(CurrentUser.UserSkills(16)) _
    & Integer_To_String(CurrentUser.UserSkills(17)) & Integer_To_String(CurrentUser.UserSkills(18)) _
    & Integer_To_String(CurrentUser.UserSkills(19)) & Integer_To_String(CurrentUser.UserSkills(20)) _
    & Integer_To_String(CurrentUser.UserSkills(21)) & Integer_To_String(CurrentUser.UserSkills(22)) _
    & Integer_To_String(CurrentUser.UserSkills(23)) & Integer_To_String(CurrentUser.UserSkills(24)) _
    & Integer_To_String(CurrentUser.UserSkills(25)) & Integer_To_String(CurrentUser.UserSkills(26)) _
    & Integer_To_String(CurrentUser.UserSkills(27)) & Byte_To_String(CurrentUser.UserHogar) & Byte_To_String(1) & Integer_To_String(valCode) _
    & MD5HushYo & CurrentUser.UserName & "," & CurrentUser.UserPet.nombre & "," & CurrentUser.UserPet.Tipo)
Else
    Call Send_Data(New_Char, Integer_To_String(intHeadInd) _
    & Integer_To_String(App.Major) & Integer_To_String(App.Minor) & Integer_To_String(App.Revision) _
    & Byte_To_String(CurrentUser.UserRaza) & Byte_To_String(CurrentUser.UserSexo) & Byte_To_String(CurrentUser.UserClase) _
    & Integer_To_String(CurrentUser.UserAtributos(1)) & Integer_To_String(CurrentUser.UserAtributos(2)) & Integer_To_String(CurrentUser.UserAtributos(3)) _
    & Integer_To_String(CurrentUser.UserAtributos(4)) & Integer_To_String(CurrentUser.UserAtributos(5)) _
    & Integer_To_String(CurrentUser.UserSkills(1)) & Integer_To_String(CurrentUser.UserSkills(2)) _
    & Integer_To_String(CurrentUser.UserSkills(3)) & Integer_To_String(CurrentUser.UserSkills(4)) _
    & Integer_To_String(CurrentUser.UserSkills(5)) & Integer_To_String(CurrentUser.UserSkills(6)) _
    & Integer_To_String(CurrentUser.UserSkills(7)) & Integer_To_String(CurrentUser.UserSkills(8)) _
    & Integer_To_String(CurrentUser.UserSkills(9)) & Integer_To_String(CurrentUser.UserSkills(10)) _
    & Integer_To_String(CurrentUser.UserSkills(11)) & Integer_To_String(CurrentUser.UserSkills(12)) _
    & Integer_To_String(CurrentUser.UserSkills(13)) & Integer_To_String(CurrentUser.UserSkills(14)) _
    & Integer_To_String(CurrentUser.UserSkills(15)) & Integer_To_String(CurrentUser.UserSkills(16)) _
    & Integer_To_String(CurrentUser.UserSkills(17)) & Integer_To_String(CurrentUser.UserSkills(18)) _
    & Integer_To_String(CurrentUser.UserSkills(19)) & Integer_To_String(CurrentUser.UserSkills(20)) _
    & Integer_To_String(CurrentUser.UserSkills(21)) & Integer_To_String(CurrentUser.UserSkills(22)) _
    & Integer_To_String(CurrentUser.UserSkills(23)) & Integer_To_String(CurrentUser.UserSkills(24)) _
    & Integer_To_String(CurrentUser.UserSkills(25)) & Integer_To_String(CurrentUser.UserSkills(26)) _
    & Integer_To_String(CurrentUser.UserSkills(27)) & Byte_To_String(CurrentUser.UserHogar) & Byte_To_String(0) & Integer_To_String(valCode) _
    & MD5HushYo & CurrentUser.UserName)
End If

'Ventana de espera en el frmCrearPersonaje

End Sub

Private Sub CopiarDialogoAConsola(ByVal NickName As String, ByVal Dialogo As String, ByVal color As Long)

If NickName = vbNullString Then Exit Sub
If Right$(Dialogo, 1) = " " Then Exit Sub

If InStr(NickName, "<") Then
    NickName = left$(NickName, InStr(NickName, "<") - 2)
End If

Select Case color
    Case 267514112 'Privado
        Call PrintToConsole(NickName & "> " & Dialogo, 225, 225, 0, False, True, False)
    Case 264830976 'Rojo
        Call PrintToConsole(NickName & "> " & Dialogo, 255, 0, 0, False, True, False)
    Case 65280 'Verde
        Call PrintToConsole(NickName & "> " & Dialogo, 0, 255, 0, False, True, False)
    Case 254383360 'Privado
        Call PrintToConsole(NickName & "> " & Dialogo, 0, 201, 197, False, True, False)
    Case &HC0C0C0 'Gris
        Call PrintToConsole(NickName & "> " & Dialogo, 164, 164, 164, False, True, False)
    Case Else 'Blanco
        Call PrintToConsole(NickName & "> " & Dialogo, 255, 255, 255, False, True, False)
End Select

End Sub

Private Sub MostrarEstadisticas()

If LlegaronSkills And LlegaronAtrib And LlegoFami And LlegoEst Then
    If frmMain.PedimosEst Then
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show vbModeless, frmMain
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFami = False
        LlegoEst = False
        frmMain.PedimosEst = False
    End If
End If

End Sub

'[BARRIN]

'Public Function RC4_D(ByteArray() As Byte, ByVal Password As String) As String
'
'On Error Resume Next
'
'Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, temp As Byte
'
'If Len(Password) = 0 Then
'    Exit Function
'End If
'
'If Len(Password) > 256 Then
'    Key() = StrConv(left$(Password, 256), vbFromUnicode)
'Else
'    Key() = StrConv(Password, vbFromUnicode)
'End If
'
'For X = 0 To 255
'    RB(X) = X
'Next X
'
'X = 0
'Y = 0
'Z = 0
'
'For X = 0 To 255
'    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
'    temp = RB(X)
'    RB(X) = RB(Y)
'    RB(Y) = temp
'Next X
'
'X = 0
'Y = 0
'Z = 0
'
''For X = 0 To UBound(ByteArray)
''    Y = (Y + 1) Mod 256
''    Z = (Z + RB(Y)) Mod 256
''    temp = RB(Y)
''    RB(Y) = RB(Z)
''    RB(Z) = temp
''    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
''Next X
'
'RC4_D = StrConv(ByteArray, vbUnicode)
'
'End Function

'Public Sub RC4_E(ByVal Expression As String, ByVal Password As String, ByRef ByteArray() As Byte)
'
'On Error Resume Next
'
'Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, temp As Byte
'
'If Len(Password) = 0 Then
'    Exit Sub
'End If
'
'If Len(Expression) = 0 Then
'    Exit Sub
'End If
'
'If Len(Password) > 256 Then
'    Key() = StrConv(left$(Password, 256), vbFromUnicode)
'Else
'    Key() = StrConv(Password, vbFromUnicode)
'End If
'
'For X = 0 To 255
'    RB(X) = X
'Next X
'
'X = 0
'Y = 0
'Z = 0
'
'For X = 0 To 255
'    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
'    temp = RB(X)
'    RB(X) = RB(Y)
'    RB(Y) = temp
'Next X
'
'X = 0
'Y = 0
'Z = 0
'
'ByteArray() = StrConv(Expression, vbFromUnicode)
'
''For X = 0 To UBound(ByteArray())
''    Y = (Y + 1) Mod 256
''    Z = (Z + RB(Y)) Mod 256
''    temp = RB(Y)
''    RB(Y) = RB(Z)
''    RB(Z) = temp
''    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
''Next X
'
'End Sub

'[/BARRIN]

Public Function ActualizarEst(Optional ByVal MaxHP As Integer = -1, Optional ByVal MinHP As Integer = -1, Optional ByVal MaxMAN As Integer = -1, _
    Optional ByVal MinMAN As Integer = -1, Optional ByVal MaxSTA As Integer = -1, Optional ByVal MinSTA As Integer = -1, _
    Optional ByVal GLD As Long = -1, Optional ByVal Nivel As Integer = -1, Optional ByVal PasarNivel As Long = -1, Optional ByVal EXP As Long = -1, _
    Optional ByVal Fuerza As Integer = -1, Optional ByVal Agilidad As Integer = -1, _
    Optional ByVal MinHAM As Integer = -1, Optional ByVal MaxHAM As Integer = -1, _
    Optional ByVal MinSED As Integer = -1, Optional ByVal MaxSED As Integer = -1, _
    Optional ByVal ActualizarTodos As Boolean = False)

Dim ActualizarCual As Byte

If MaxHP <> -1 Then
    CurrentUser.UserMaxHP = MaxHP
    ActualizarCual = 1
End If

If MinHP <> -1 Then

    If Not ActualizarTodos Then
        If MinHP < CurrentUser.UserMinHP Then CurrentUser.LastHostile = GetTickCount
    End If

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
    
    If CurrentUser.UserMinMAN > CurrentUser.UserMaxMAN Then _
        CurrentUser.UserMinMAN = CurrentUser.UserMaxMAN
        
    ActualizarCual = 2
End If

If MaxSTA <> -1 Then
    CurrentUser.UserMaxSTA = MaxSTA
    ActualizarCual = 3
End If

If MinSTA <> -1 Then

    If CurrentUser.UserMinSTA > CurrentUser.UserMaxSTA Then _
        CurrentUser.UserMinSTA = CurrentUser.UserMaxSTA

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
    frmMain.lblFU.ForeColor = IIf(Fuerza < 35, &HFFFFFF, &H40C0&)
    frmMain.lblAG = Agilidad
    frmMain.lblAG.ForeColor = IIf(Agilidad < 35, &HFFFFFF, &H40C0&)
End If

If MinHAM <> -1 Then
    CurrentUser.UserMinHAM = MinHAM
    ActualizarCual = 6
End If

If MinSED <> -1 Then
    CurrentUser.UserMinAGU = MinSED
    ActualizarCual = 6
End If

If MaxHAM <> -1 Then
    CurrentUser.UserMaxHAM = MaxHAM
    ActualizarCual = 6
End If

If MaxSED <> -1 Then
    CurrentUser.UserMaxAGU = MaxSED
    ActualizarCual = 6
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
        Case 6
            Call ActualizarHAM
    End Select
Else
    Call ActualizarHP
    Call ActualizarMAN
    Call ActualizarSTA
    Call ActualizarGLD
    Call ActualizarExp
    Call ActualizarHAM
End If

End Function

Private Sub ActualizarHAM()

frmMain.AGUAsp.Width = (((CurrentUser.UserMinAGU / 100) / (CurrentUser.UserMaxAGU / 100)) * 91)
frmMain.COMIDAsp.Width = (((CurrentUser.UserMinHAM / 100) / (CurrentUser.UserMaxHAM / 100)) * 91)
frmMain.lblHAM.Caption = CurrentUser.UserMinHAM & "/" & CurrentUser.UserMaxHAM
frmMain.lblSED.Caption = CurrentUser.UserMinAGU & "/" & CurrentUser.UserMaxAGU

End Sub

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

If CurrentUser.UserMinHP <= 0 Then
    CurrentUser.UserMinHP = 0
    CurrentUser.Muerto = True
    Call Banner_Logic
    CurrentUser.CurrentSpeed = VelRapida
    frmMain.Engine.Char_Current_Dead_Set (True)
    frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
    frmMain.lblHP.Caption = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP
    frmMain.Hpshp.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 91)
    frmMain.Hpshp.FillColor = &H808080
Else
    If CurrentUser.Muerto Then
        Call Banner_Logic
        CurrentUser.Muerto = False
        Call frmMain.Engine.Char_Current_Dead_Set(False)
    End If
    
    If CurrentUser.Logged Then
        If (CurrentUser.Montando = False) And (CurrentUser.CharType <> eGM) Then
            CurrentUser.CurrentSpeed = VelNormal
            frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
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
    frmMain.Label2(1).Caption = Locale_GUI_Frase(173)
End If

End Sub

Public Function DeadCheck() As Boolean

If CurrentUser.Muerto Then
    Call PrintToConsole(Locale_GUI_Frase(245), 0, 0, 0, 0, 0, 0, 4)
    DeadCheck = True
End If

End Function

Public Function CombateCheck() As Boolean

If CurrentUser.Combate = False Then
    Call PrintToConsole(Locale_GUI_Frase(248), 0, 0, 0, 0, 0, 0, 4)
    CombateCheck = True
End If

End Function

Public Function MeditarCheck() As Boolean

If CurrentUser.Meditando = True Then
    Call PrintToConsole(Locale_GUI_Frase(247), 0, 0, 0, 0, 0, 0, 4)
    MeditarCheck = True
End If

End Function

Public Function StealCheck() As Boolean

If CurrentUser.Seguro = True Then
    Call PrintToConsole(Locale_GUI_Frase(246), 0, 0, 0, 0, 0, 0, 4)
    StealCheck = True
End If

End Function

Private Function Generate_Char_Status(ByVal PercVida As Long, ByVal Paralizado As Byte, ByVal Inmovilizado As Byte, ByVal Envenenado As Byte, Optional ByVal Trabajando As Byte = 0, Optional ByVal Silenciado As Byte = 0, Optional ByVal Ciego As Byte = 0, Optional ByVal Incinerado As Byte = 0, Optional ByVal Transformado As Byte = 0, Optional ByVal Comerciando As Byte = 0, Optional ByVal Inactivo As Byte = 0) As String

If PercVida <> -1 Then
    If PercVida = 100 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(445) & " "
    ElseIf PercVida >= 80 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(446) & " "
    ElseIf PercVida >= 50 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(447) & " "
    ElseIf PercVida >= 30 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(448) & " "
    ElseIf PercVida <> 0 Then
        Generate_Char_Status = " " & Locale_GUI_Frase(449) & " "
    Else
        Generate_Char_Status = " " & Locale_GUI_Frase(450) & " "
    End If
End If

If Paralizado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(451) & " "
End If

If Inmovilizado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(452) & " "
End If

If Envenenado > 0 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(453) & " "
End If

If Trabajando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(463) & " "
End If

If Silenciado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(461) & " "
End If

If Ciego = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(455) & " "
End If

If Incinerado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(454) & " "
End If

If Transformado = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(457) & " "
End If

If Comerciando = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(462) & " "
End If

If Inactivo = 1 Then
    Generate_Char_Status = Generate_Char_Status & "| " & Locale_GUI_Frase(464) & " "
End If

End Function

Private Sub Parse_CharDesc(ByRef rData As String)

Dim tempStr As String

Dim btStatus As Byte
Dim btArmada As Byte
Dim lngPorcVida As Long
Dim btRango As Byte
Dim btFounder As Byte
Dim btGenero As Byte
Dim btParalizado As Byte
Dim btInmovilizado As Byte
Dim btTrabajando As Byte
Dim btSilencio As Byte
Dim btCiego As Byte
Dim btIncinerado As Byte
Dim btEnvenenado As Byte
Dim btTransformado As Byte
Dim btComerciando As Byte
Dim btInactivo As Byte

Dim btNameLen As Byte
Dim btGuildLen As Byte
Dim btMatrimonioLen As Byte

Dim btOffset As Byte

Dim btRed As Byte
Dim btGreen As Byte
Dim btBlue As Byte

btStatus = String_To_Byte(rData, 1)
btArmada = String_To_Byte(rData, 2)
lngPorcVida = String_To_Long(rData, 3)
btRango = String_To_Byte(rData, 7)
btFounder = String_To_Byte(rData, 8)
btGenero = String_To_Byte(rData, 9)
btParalizado = String_To_Byte(rData, 10)
btInmovilizado = String_To_Byte(rData, 11)
btTrabajando = String_To_Byte(rData, 12)
btSilencio = String_To_Byte(rData, 13)
btCiego = String_To_Byte(rData, 14)
btIncinerado = String_To_Byte(rData, 15)
btEnvenenado = String_To_Byte(rData, 16)
btTransformado = String_To_Byte(rData, 17)
btComerciando = String_To_Byte(rData, 18)
btInactivo = String_To_Byte(rData, 19)

btNameLen = String_To_Byte(rData, 20)
btGuildLen = String_To_Byte(rData, 21)
btMatrimonioLen = String_To_Byte(rData, 22)

btOffset = 23

tempStr = Locale_GUI_Frase(466) & " " & mid$(rData, btOffset, btNameLen) & " [" & Generate_Char_Status(lngPorcVida, btParalizado, btInmovilizado, btEnvenenado, btTrabajando, btSilencio, btCiego, btIncinerado, btTransformado, btComerciando, btInactivo) & "]"

If btArmada > 0 Then
    Select Case btArmada
        Case E_ARMADA.eArCaos
            tempStr = tempStr & " <" & Locale_GUI_Frase(150) & "> <" & Locale_Facc_Frase(btRango) & ">"
            btRed = 196
            btGreen = 0
            btBlue = 15
        Case E_ARMADA.eArImperial
            tempStr = tempStr & " <" & Locale_GUI_Frase(148) & "> <" & Locale_Facc_Frase(btRango) & ">"
            btRed = 32
            btGreen = 81
            btBlue = 251
        Case E_ARMADA.eArRepublicano
            tempStr = tempStr & " <" & Locale_GUI_Frase(149) & "> <" & Locale_Facc_Frase(btRango) & ">"
            btRed = 204
            btGreen = 107
            btBlue = 0
    End Select
Else
    Select Case btStatus
        Case E_USER_STATUS.eImperial
            tempStr = tempStr & " <" & Locale_GUI_Frase(152) & ">"
            btRed = 32
            btGreen = 81
            btBlue = 251
        Case E_USER_STATUS.eRenegado
            tempStr = tempStr & " <" & Locale_GUI_Frase(154) & ">"
            btRed = 114
            btGreen = 115
            btBlue = 108
        Case E_USER_STATUS.eRepublicano
            tempStr = tempStr & " <" & Locale_GUI_Frase(153) & ">"
            btRed = 204
            btGreen = 107
            btBlue = 0
        Case E_USER_STATUS.eGM
            tempStr = tempStr & " <GM>"
            btRed = 2
            btGreen = 162
            btBlue = 38
    End Select
End If

If btGuildLen > 0 Then
    tempStr = tempStr & " <" & mid$(rData, btOffset + btNameLen, btGuildLen) & ">"
End If

If btMatrimonioLen > 0 Then
    If btGenero = Masculino Then
        tempStr = tempStr & " <" & Locale_GUI_Frase(468) & " " & mid$(rData, btOffset + btNameLen + btGuildLen, btMatrimonioLen) & ">"
    Else
        tempStr = tempStr & " <" & Locale_GUI_Frase(467) & " " & mid$(rData, btOffset + btNameLen + btGuildLen, btMatrimonioLen) & ">"
    End If
End If

If Len(rData) > btOffset + btNameLen + btGuildLen + btMatrimonioLen Then
    tempStr = tempStr & " - " & mid$(rData, btOffset + btNameLen + btGuildLen + btMatrimonioLen)
End If

Call PrintToConsole(tempStr, btRed, btGreen, btBlue, 1, 0, 0)

End Sub
