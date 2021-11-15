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

Public Sub HandleData(ByVal rData As String)
    
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
        
    sData = UCase$(rData)
    
    Select Case left(sData, 6)
    
        Case "CONNET"
            
            rData = Right$(rData, Len(rData) - 6)
            
            If EstadoLogin = CrearNuevoPj Then
                Unload frmPasswd
                Unload frmCrearPersonaje
                frmIniciando.Show
            ElseIf EstadoLogin = NORMAL Then
                frmIniciando.Show
                frmConnect.Visible = False
            End If
            
            Call Map_Load(Val(General_Field_Read(1, rData, ",")), General_Field_Read(2, rData, ","))
            
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
            frmMensaje.mnuClan.Enabled = General_Field_Read(3, rData, ",")
            frmMensaje.mnuGMs.Enabled = General_Field_Read(4, rData, ",")
            '[/Barrin]
            
            Call DibujarMenuMacros
            Call Inventory_Render
            
            'Barrin
            Call tcp.Send_Data(Loading_Finished, MensajesGlobales)
            
            Exit Sub
    End Select
    
    Select Case sData
        Case "NAVEG"
            CurrentUser.Navegando = Not CurrentUser.Navegando
            Engine.Char_Current_OverWater_Set (CurrentUser.Navegando)
            Exit Sub
        Case "MONTA"
            CurrentUser.Montando = Not CurrentUser.Montando
            Engine.Char_Current_OnHorse_Set (CurrentUser.Montando)
            
            If CurrentUser.Montando Then
                CurrentUser.CurrentSpeed = VelRapida
                Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
                With frmPanelGm
                    .mnuNormal.Checked = False
                    .mnuMuy.Checked = False
                    .mnuRapida.Checked = True
                End With
            Else
                CurrentUser.CurrentSpeed = VelNormal
                Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
                With frmPanelGm
                    .mnuNormal.Checked = True
                    .mnuMuy.Checked = False
                    .mnuRapida.Checked = False
                End With
            End If
            
            Exit Sub
        Case "VOLAO"
            CurrentUser.Volando = Not CurrentUser.Volando
            Exit Sub
        Case "FINOK"
            frmMain.mainWinsock.Close
            frmConnect.Visible = True
            frmMain.Visible = False
            frmMain.modocombate.Visible = False
            frmMain.nomodocombate.Visible = True
            frmMain.modoseguro.Visible = True
            frmMain.nomodoseguro.Visible = False
            frmMain.modorol.Visible = False
            Call ResetCurrentUser
            
            If Musica <> CONST_DESHABILITADA Then
                Sound.NextMusic = MUS_VolverInicio
                Sound.Fading = 200
            End If
                                                
            Exit Sub
        Case "FINCOMOK"
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            CurrentUser.Comerciando = False
            Exit Sub
        Case "FINBANOK"
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            CurrentUser.Comerciando = False
            Exit Sub
        Case "INITCOM"
            i = 1
            Do While i <= UBound(UserInventory)
                If UserInventory(i).OBJIndex <> 0 Then
                        frmComerciar.List1(1).AddItem UserInventory(i).Name
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            CurrentUser.Comerciando = True
            frmComerciar.Show vbModeless, frmMain
            Exit Sub
        Case "INITBANCO"
            Dim ii As Integer
            ii = 1
            Do While ii <= UBound(UserInventory)
                If UserInventory(ii).OBJIndex <> 0 Then
                        frmBancoObj.List1(1).AddItem UserInventory(ii).Name
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            CurrentUser.Comerciando = True
            frmBancoObj.Show vbModeless, frmMain
            Exit Sub
        Case "SFH"
            frmHerrero.Show vbModeless, frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show vbModeless, frmMain
            Exit Sub
        Case "SFD"
            frmDruida.Show vbModeless, frmMain
            Exit Sub
        Case "SFS"
            frmSastre.Show vbModeless, frmMain
            Exit Sub
        Case "ROLN"
            Call AddtoRichTextBox(frmMain.RecTxt, "No se te ha permitido el uso del modo rol.", 61, 142, 36, True, True, False)
            CurrentUser.Rol = False
            frmMain.modorol.Visible = False
            frmMain.nomodorol.Visible = True
            Exit Sub
        Case "RE" ' <--- Resucitado
            Call AddtoRichTextBox(frmMain.RecTxt, "El cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar... ¡Has sido resucitado!", 0, 0, 0, 0, 0, 0, 4)
            Call Sound.Sound_Play(SND_RESUCITAR)
            Exit Sub
        Case "CU" ' <--- Curado
            Call AddtoRichTextBox(frmMain.RecTxt, "El sacerdote levanta sus manos, recita unas palabras, y comienzas a sentir un fuerte ardor. Luego ves como van cerrando tus heridas... ¡Has sido curado!", 0, 0, 0, 0, 0, 0, 4)
            Call Engine.Char_FX_Set(CurrentUser.CurrentChar, 9, 1)
            Call Sound.Sound_Play(SND_CURAR)
            Exit Sub
        Case "VP" ' <--- Obstruye
            Call AddtoRichTextBox(frmMain.RecTxt, "Estás obstruyendo la vía pública, ¡muévete o serás encarcelado!", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, "¡La criatura falló el golpe!", 0, 0, 0, 0, 0, 0, 2)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, "¡La criatura te ha matado!", 0, 0, 0, 0, 0, 0, 2)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has rechazado el ataque con el escudo!", 0, 0, 0, 0, 0, 0, 2)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, "¡El usuario rechazó el ataque con su escudo!", 0, 0, 0, 0, 0, 0, 2)
            Exit Sub
        Case "9" ' <--- Menos cansado
            Call AddtoRichTextBox(frmMain.RecTxt, "Te sentis menos cansado.", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "10" ' <--- No se esconde
            Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado esconderte!", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "11" ' <--- Offline
            Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> Usuario offline.", 0, 0, 0, 0, 0, 0, 8)
            Exit Sub
        Case "12" ' <--- Nada interesante
            Call AddtoRichTextBox(frmMain.RecTxt, "No ves nada interesante.", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "13" ' <--- Seleccionar pj
            Call AddtoRichTextBox(frmMain.RecTxt, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "14" ' <--- No tiene cantidad
            Call AddtoRichTextBox(frmMain.RecTxt, "No tenés esa cantidad.", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "15" ' <--- Muere criatura
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has matado a la criatura!", 0, 0, 0, 0, 0, 0, 2)
            Exit Sub
        Case "16" ' <--- Muerto
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Estás muerto! Ve al sacerdote más cercano para que puedas ser revivido.", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "17" ' <--- Mov. Especial
            Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado realizar un movimiento especial!", 0, 0, 0, 0, 0, 0, 16)
            Exit Sub
        Case "18" ' <--- Apuñalar
            Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado apuñalar a tu enemigo!", 0, 0, 0, 0, 0, 0, 16)
            Exit Sub
        Case "19" ' <--- Inexistente
            Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> El personaje no existe.", 0, 0, 0, 0, 0, 0, 8)
            Exit Sub
        Case "20" ' <--- Resucitando
            Call AddtoRichTextBox(frmMain.RecTxt, "Tu cuerpo comienza a tomar forma...", 0, 0, 0, 0, 0, 0, 2)
            CurrentUser.Reviviendo = True
            Exit Sub
        Case "22" ' <--- Resucitado (spell)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has vuelto a la vida!", 0, 0, 0, 0, 0, 0, 2)
            CurrentUser.Reviviendo = False
            Exit Sub
        Case "23" ' <--- Invalida reglamento
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Tu consulta ha sido ignorada debido a que invalida el reglamento del juego! Te recomendamos leerlo en el sitio oficial: www.imperiumao.com.ar", 0, 0, 0, 0, 0, 0, 8)
            Exit Sub
        Case "24" ' <--- Invalida manual
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Tu consulta ha sido ignorada debido a que está respondida en el manual básico o bien las preguntas frecuentes (FAQ)! Te recomendamos leer los textos en el sitio oficial: www.imperiumao.com.ar", 0, 0, 0, 0, 0, 0, 8)
            Exit Sub
        Case "25" ' <--- Mensaje a GMs
            Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> ¡Gracias por tu mensaje! Será respondido a la brevedad.", 0, 0, 0, 0, 0, 0, 8)
            Exit Sub
        Case "26" ' <--- Mensaje a GMs (repetido)
            Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola.", 0, 0, 0, 0, 0, 0, 8)
            Exit Sub
        Case "27" ' <--- Oro insuficiente
            Call AddtoRichTextBox(frmMain.RecTxt, "Oro insuficiente.", 0, 0, 0, 0, 0, 0, 4)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Engine.Char_Dialog_Set CurrentUser.CurrentChar, "*Fallas*", COLOR_ATAQUE, 5, 2
            Exit Sub
        Case "SZ"
            Call AddtoRichTextBox(frmMain.RecTxt, "Estás saliendo de una zona segura. Recuerda que aquí corres riesgo de ser atacado por otros.", 0, 0, 0, 0, 0, 0, 3)
            Exit Sub
        Case "PONG"
            CurrentUser.Ping = GetTickCount - CurrentUser.Ping
            Call AddtoRichTextBox(frmMain.RecTxt, "Ping: " & CurrentUser.Ping & "ms", 0, 0, 0, 0, 0, 0, 4)
            Call tcp.Send_Data(Ping_Request)
            Exit Sub
    End Select
    
    Select Case left(sData, 2)
        Case "CM"
            rData = Right$(rData, Len(rData) - 2)
            TempInt = Val(General_Field_Read(1, rData, ","))
            TempStr = General_Field_Read(2, rData, ",")
            frmIniciando.Visible = True
            frmMain.Visible = False
            CurrentUser.CurrentChar = 0
            Call Map_Load(TempInt, TempStr)
            Call tcp.Send_Data(Loading_Finished)
            Exit Sub
        Case "MD"
            rData = (Right$(rData, Len(rData) - 2))
            t = Split(rData, "¬")
            
            For x = 0 To UBound(t)
                Call HandleData(t(x))
            Next x
            
            Exit Sub
        Case "CE"
            rData = Right$(rData, Len(rData) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "En " & rData & " segundos se cerrará el juego...", 65, 190, 156, False, True, False)
            CurrentUser.Saliendo = True
            Exit Sub
        Case "CR"
            Call AddtoRichTextBox(frmMain.RecTxt, "¡El cierre del juego ha sido cancelado!", 0, 0, 0, False, False, False, 14)
            CurrentUser.Saliendo = False
            Exit Sub
        Case "XP"
            rData = Right$(rData, Len(rData) - 2)
            
            If GuardarEXP = 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt, "¡Has ganado " & rData & " puntos de experiencia!", 51, 183, 247, True, False, False)
            Else
                CurrentUser.ExpCount = CurrentUser.ExpCount + Val(rData)
            End If
                
            Exit Sub
        Case "XG"
            rData = Right$(rData, Len(rData) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡El grupo ha ganado " & rData & " puntos de experiencia!", 51, 183, 247, True, False, False)
            Exit Sub
        Case "GH"
            rData = Right$(rData, Len(rData) - 2)
            TempStr = General_Field_Read(1, rData, ",")
            
            If Val(TempStr) > 200 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & TempStr & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, TempStr, COLOR_ATAQUE, 5, 2
            End If
            
            Exit Sub
        Case "GU"
            rData = General_Field_Read(1, Right$(rData, Len(rData) - 2), ",")
            TempStr = General_Field_Read(1, rData, ",")
            
            If Val(rData) > 200 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & TempStr & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, TempStr, COLOR_ATAQUE, 5, 2
            End If
            
            Exit Sub
        Case "PU"
            rData = Right$(rData, Len(rData) - 2)
            x = CInt(General_Field_Read(1, rData, ","))
            y = CInt(General_Field_Read(2, rData, ","))
            Call Engine.Char_Current_Pos_Refresh(x, y)
            Call Engine.Engine_View_Pos_Set(x, y)
            Exit Sub
        Case "PS"
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(CInt(General_Field_Read(1, rData, ",")))
            x = CInt(General_Field_Read(2, rData, ","))
            y = CInt(General_Field_Read(3, rData, ","))
            Call Engine.Char_Pos_Set(CharIndex, x, y)
            Exit Sub
        Case "PP"
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(CInt(General_Field_Read(1, rData, ",")))
            x = Val(General_Field_Read(2, rData, ","))
            Call Engine.Char_Fly_Set(CharIndex, x)
            Exit Sub
        Case "N2" ' <<--- Npc nos impacto
            rData = Right$(rData, Len(rData) - 2)
            i = Val(General_Field_Read(1, rData, ","))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado en la cabeza por " & Val(General_Field_Read(2, rData, ",")), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado el brazo izquierdo por " & Val(General_Field_Read(2, rData, ",")), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado el brazo derecho por " & Val(General_Field_Read(2, rData, ",")), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado la pierna izquierda por " & Val(General_Field_Read(2, rData, ",")), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado la pierna derecha por " & Val(General_Field_Read(2, rData, ",")), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado en el torso por " & Val(General_Field_Read(2, rData, ",")), 255, 0, 0, True, False, False)
            End Select
                        
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            rData = Right$(rData, Len(rData) - 2)
            
            If Val(rData) > 150 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & rData & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, rData, COLOR_ATAQUE, 5, 2
            End If
            
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            rData = Right$(rData, Len(rData) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡" & rData & " te ataco y fallo!", 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            rData = Right$(rData, Len(rData) - 2)
            i = Val(General_Field_Read(1, rData, ","))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & General_Field_Read(3, rData, ",") & " te ha pegado en la cabeza por " & Val(General_Field_Read(2, rData, ",")) & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & General_Field_Read(3, rData, ",") & " te ha pegado el brazo izquierdo por " & Val(General_Field_Read(2, rData, ",")) & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & General_Field_Read(3, rData, ",") & " te ha pegado el brazo derecho por " & Val(General_Field_Read(2, rData, ",")) & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & General_Field_Read(3, rData, ",") & " te ha pegado la pierna izquierda por " & Val(General_Field_Read(2, rData, ",")) & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & General_Field_Read(3, rData, ",") & " te ha pegado la pierna derecha por " & Val(General_Field_Read(2, rData, ",")) & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & General_Field_Read(3, rData, ",") & " te ha pegado en el torso por " & Val(General_Field_Read(2, rData, ",")) & "!", 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            rData = Right$(rData, Len(rData) - 2)
            i = Val(General_Field_Read(1, rData, ","))
            x = Val(General_Field_Read(2, rData, ","))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & General_Field_Read(3, rData, ",") & " en la cabeza por " & x & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & General_Field_Read(3, rData, ",") & " en el brazo izquierdo por " & x & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & General_Field_Read(3, rData, ",") & " en el brazo derecho por " & x & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & General_Field_Read(3, rData, ",") & " en la pierna izquierda por " & x & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & General_Field_Read(3, rData, ",") & " en la pierna derecha por " & x & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & General_Field_Read(3, rData, ",") & " en el torso por " & x & "!", 255, 0, 0, True, False, False)
            End Select
            
            If Val(rData) > 150 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & x & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, x, COLOR_ATAQUE, 5, 2
            End If
            
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Val((General_Field_Read(3, rData, "°")))
            TempInt = Val(General_Field_Read(2, rData, "«"))
            
            If CharIndex <> 0 Then
                CharIndex = Engine.Char_Find(Val(General_Field_Read(3, rData, "°")))
                TempStr = Engine.Char_Name_Get(CharIndex)
                
                If InStr(TempStr, "<") Then
                    cad$ = General_Field_Read(1, TempStr, "<")
                    cad$ = left$(cad$, Len(cad$) - 1)
                Else
                    cad$ = TempStr
                End If
                
                If Not NickIgnorado(cad$) Then
                    Engine.Char_Dialog_Set CharIndex, General_Field_Read(2, rData, "°"), Val(General_Field_Read(1, rData, "°")), 10
                    
                    If CopiarDialogos = 1 And TempStr <> "" Then
                        Call CopiarDialogoAConsola(TempStr, General_Field_Read(2, rData, "°"), Val(General_Field_Read(1, rData, "°")))
                    End If
                End If
                
            Else
                TempStr = General_Field_Read(1, rData, 62)
                                
                If Not NickIgnorado(TempStr) Then
                    If TempInt > 0 Then
                        AddtoRichTextBox frmMain.RecTxt, General_Field_Read(1, rData, "«"), 0, 0, 0, 0, 0, False, TempInt
                    Else
                        AddtoRichTextBox frmMain.RecTxt, General_Field_Read(1, rData, "~"), Val(General_Field_Read(2, rData, "~")), Val(General_Field_Read(3, rData, "~")), Val(General_Field_Read(4, rData, "~")), Val(General_Field_Read(5, rData, "~")), Val(General_Field_Read(6, rData, "~"))
                    End If
                End If
                
            End If
            Exit Sub
        Case "()" 'Palabras mágicas
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Val((General_Field_Read(3, rData, "°")))
            TempInt = Val(General_Field_Read(2, rData, "«"))
            
            If CharIndex <> 0 Then
                CharIndex = Engine.Char_Find(Val(General_Field_Read(3, rData, "°")))
                Engine.Char_Dialog_Set CharIndex, General_Field_Read(2, rData, "°"), Val(General_Field_Read(1, rData, "°")), 5
            End If
            
            Exit Sub
        Case "[]"
            rData = Right$(rData, Len(rData) - 2)
            If MensajesGlobales = 1 Then
                TempStr = General_Field_Read(1, rData, 62)
                If Not NickIgnorado(TempStr) Then _
                    Call AddtoRichTextBox(frmMain.RecTxt, rData, 0, 0, 0, 0, 0, 0, 23)
            End If
            Exit Sub
        Case "S1" 'HP
            rData = Right$(rData, Len(rData) - 2)
            Call ActualizarEst(Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")))
        Case "S2" 'STA
            rData = Right$(rData, Len(rData) - 2)
            Call ActualizarEst(, , , , Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")))
        Case "S3" 'MAN
            rData = Right$(rData, Len(rData) - 2)
            Call ActualizarEst(, , Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")))
        Case "S4" 'GLD
            rData = Right$(rData, Len(rData) - 2)
            Call ActualizarEst(, , , , , , Val(rData))
        Case "S5" 'EXP
            rData = Right$(rData, Len(rData) - 2)
            Call ActualizarEst(, , , , , , , Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")), Val(General_Field_Read(3, rData, ",")))
        Case "S6" 'ATRIBUTOS
            rData = Right$(rData, Len(rData) - 2)
            Call ActualizarEst(, , , , , , , , , , Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")))
        Case "!!"                ' >>>>> Msgbox :: !!
            rData = Right$(rData, Len(rData) - 2)
            frmMensaje.msg.Caption = rData
            frmMensaje.Show vbModeless, frmMain
            Call Sound.Sound_Play(118)
            Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            rData = Right$(rData, Len(rData) - 2)
            k = General_Field_Count(rData, 44)
            
            x = General_Field_Read(5, rData, ",")
            y = General_Field_Read(6, rData, ",")
            
            If k > 11 Then
                TempStr = General_Field_Read(16, rData, ",")
                TempInt = Val(General_Field_Read(1, TempStr, "@")) '64=@
                
                If TempInt > 0 Then
                    TempStr = Right$(TempStr, Len(TempStr) - 2)
                    ReDim part_type(1 To TempInt) As Integer
                    ReDim part_life(1 To TempInt) As Long
                End If
                
                For i = 1 To TempInt
                    part_type(i) = Val(General_Field_Read(IIf(i > 1, i + 1, i), TempStr, "@"))
                    part_life(i) = Val(General_Field_Read(IIf(i > 1, i + 2, i + 1), TempStr, "@"))
                Next i
            
                CharIndex = Engine.Char_Create(x, y, Val(General_Field_Read(3, rData, ",")), Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")), _
                    Val(General_Field_Read(11, rData, ",")), Val(General_Field_Read(7, rData, ",")), Val(General_Field_Read(8, rData, ",")), General_Field_Read(12, rData, ","), _
                    Val(General_Field_Read(13, rData, ",")), Val(General_Field_Read(9, rData, ",")), Val(General_Field_Read(10, rData, ",")), General_Field_Read(4, rData, ","), _
                    Val(General_Field_Read(14, rData, ",")), Val(General_Field_Read(15, rData, ",")), TempInt, part_life(), part_type())
                                
                If k = 17 Then
                    CurrentUser.CurrentChar = CharIndex
                    Call Engine.Char_Current_Set(CharIndex)
                    Call Engine.Char_Pos_Get(CharIndex, x, y)
                    Call Engine.Engine_View_Pos_Set(x, y)
                    
                    If Not EngineRun Then
                        Call SetConnected
                    Else
                        frmMain.Visible = True
                        frmIniciando.Visible = False
                    End If
                End If
            Else
                Call Engine.Char_Create(x, y, Val(General_Field_Read(3, rData, ",")), Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")), _
                    Val(General_Field_Read(11, rData, ",")), Val(General_Field_Read(7, rData, ",")), Val(General_Field_Read(8, rData, ",")), "", _
                    0, Val(General_Field_Read(9, rData, ",")), Val(General_Field_Read(10, rData, ",")), General_Field_Read(4, rData, ","), _
                    0, 0, 0, part_life(), part_type())
            End If
            
            Exit Sub
        Case "CH" ' Barrin: cambiar heading
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call Engine.Char_Heading_Set(CharIndex, Val(General_Field_Read(2, rData, ",")))
            Exit Sub
        Case "CT" ' Barrin: cambiar char type
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call Engine.Char_Type_Set(CharIndex, Val(General_Field_Read(2, rData, ",")))
            Exit Sub
        Case "CN" ' Barrin: cambiar nombre
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call Engine.Char_Name_Set(CharIndex, General_Field_Read(2, rData, ","))
            Exit Sub
        Case "CG" ' Barrin: cambiar grupo
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call Engine.Char_Group_Set(CharIndex, Val(General_Field_Read(2, rData, ",")))
            Exit Sub
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(Val(rData))
            Call Engine.Char_Remove(CharIndex)
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            rData = Right$(rData, Len(rData) - 2)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            If CharIndex <= 0 Then Exit Sub
            x = Val(General_Field_Read(2, rData, ","))
            y = Val(General_Field_Read(3, rData, ","))
            Call Engine.Char_Move_By_Pos(CharIndex, x, y)
            If fx = 1 Then Call DoPasosFx(CharIndex)
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            rData = Right$(rData, Len(rData) - 2)
            
            Call Engine.Char_Change(Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(4, rData, ",")), _
            Val(General_Field_Read(2, rData, ",")), Val(General_Field_Read(3, rData, ",")), Val(General_Field_Read(9, rData, ",")), _
            Val(General_Field_Read(5, rData, ",")), Val(General_Field_Read(6, rData, ",")), Val(General_Field_Read(7, rData, ",")), _
            Val(General_Field_Read(8, rData, ",")))
            
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            rData = Right$(rData, Len(rData) - 2)
            x = Val(General_Field_Read(2, rData, ","))
            y = Val(General_Field_Read(3, rData, ","))
            Call Engine.Map_Item_Grh_Add(x, y, Val(General_Field_Read(1, rData, ",")))
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            rData = Right$(rData, Len(rData) - 2)
            x = Val(General_Field_Read(1, rData, ","))
            y = Val(General_Field_Read(2, rData, ","))
            Call Engine.Map_Item_Grh_Remove(x, y)
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            rData = Right$(rData, Len(rData) - 2)
            Call Engine.Map_Blocked_Set(Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")), Val(General_Field_Read(3, rData, ",")))
            Exit Sub
        Case "TM"
            If (Musica <> CONST_DESHABILITADA) And (DefMidi = 1) Then
                rData = Right$(rData, Len(rData) - 2)
                If Val(General_Field_Read(1, rData, "-")) <> 0 Then
                    Sound.NextMusic = Val(General_Field_Read(1, rData, "-"))
                    Sound.Fading = 200
                End If
            End If
            Exit Sub
        Case "TA"
            rData = Right$(rData, Len(rData) - 2)
            x = Val(rData)
            If x <> 0 Then Sound.AmbienteActual = x
            Exit Sub
        Case "TW"
            If fx = 1 Then
                rData = Right$(rData, Len(rData) - 2)
                x = Val(General_Field_Read(2, rData, ","))
                y = Val(General_Field_Read(3, rData, ","))
                
                If Engine.Map_In_Bounds(x, y) = False Then
                    Call Sound.Sound_Play(rData)
                Else
                    TempInt = Val(General_Field_Read(1, rData, ","))
                    Call Sound.Sound_Play(TempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
                End If
            End If
            Exit Sub
        Case "GL" 'Lista de guilds
            rData = Right$(rData, Len(rData) - 2)
            Call frmGuildAdm.ParseGuildList(rData)
            Exit Sub
        Case "C2"
            rData = Right$(rData, Len(rData) - 2)
            Call frmComerciarUsu.ParseData(rData)
    End Select
    
    Select Case left(sData, 3)
        Case "VAL"
            rData = Right$(rData, Len(rData) - 3)
            bK = CLng(General_Field_Read(1, rData, ","))
            bRK = General_Field_Read(2, rData, ",")
            Call Login(ValidarLoginMSG(CInt(bRK)))
            Exit Sub
        Case "BKW"
            CurrentUser.Pausa = Not CurrentUser.Pausa
            Exit Sub
        Case "LLU"
            If Meteo_Engine.SecondaryStatus = 2 Then
                Meteo_Engine.SecondaryStatus = 0
            Else
                Meteo_Engine.SecondaryStatus = 2
            End If
            Exit Sub
        Case "NEV"
            If Meteo_Engine.SecondaryStatus = 3 Then
                Meteo_Engine.SecondaryStatus = 0
            Else
                Meteo_Engine.SecondaryStatus = 3
            End If
            Exit Sub
        Case "NOC"
            rData = Right$(rData, Len(rData) - 3)
            x = Val(General_Field_Read(1, rData, ","))
            y = Val(General_Field_Read(2, rData, ","))
            Call Meteo_Engine.SetNuevoEstado(CByte(x))
            If y <> 0 Then Call Sound.Sound_Play(y)
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            rData = Right$(rData, Len(rData) - 3)
            CharIndex = Engine.Char_Find(Val(rData))
            Call Engine.Char_Dialog_Remove(CharIndex)
            Exit Sub
        Case "CFX"
            rData = Right$(rData, Len(rData) - 3)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            TempInt = Val(General_Field_Read(4, rData, ","))
            Call Engine.Char_FX_Set(CharIndex, Val(General_Field_Read(2, rData, ",")), Val(General_Field_Read(3, rData, ",")))
            
            If TempInt <> 0 Then
                If Engine.Char_Pos_Get(CharIndex, x, y) Then
                    Call Sound.Sound_Play(TempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
                End If
            End If
            
            Exit Sub
        '[Barrin: FX sobre mapa]
        Case "MFX"
            rData = Right$(rData, Len(rData) - 3)
            x = Val(General_Field_Read(2, rData, ","))
            y = Val(General_Field_Read(3, rData, ","))
            TempInt = Val(General_Field_Read(4, rData, ","))
            Call Engine.Map_Fx_Add(x, y, Val(General_Field_Read(1, rData, ",")))
            If TempInt <> 0 Then Call Sound.Sound_Play(TempInt, , Sound.Calculate_Volume(x, y), Sound.Calculate_Pan(x, y))
            Exit Sub
        '[/Barrin: FX sobre mapa]
        
        '[Barrin: Partículas]
        'Formato de cadena para personajes: (Personaje, Partícula, Vida)
        'Formato de cadena para mapa: (X, Y, Partícula, Vida)
        'Formato de cadena para borrado de partícula: (Personaje, Partícula)
        'Formato de cadena para borrado de todas las partículas: (Personaje)
        
        '1. Partícula sobre personaje
        Case "XAX"
            rData = Right$(rData, Len(rData) - 3)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call General_Char_Particle_Create(Val(General_Field_Read(2, rData, ",")), CharIndex, Val(General_Field_Read(3, rData, ",")))
            Exit Sub
        '2. Partícula en mapa
        Case "XMX"
            rData = Right$(rData, Len(rData) - 3)
            x = Val(General_Field_Read(2, rData, ","))
            y = Val(General_Field_Read(3, rData, ","))
            Call General_Particle_Create(Val(General_Field_Read(1, rData, ",")), x, y, Val(General_Field_Read(2, rData, ",")))
            Exit Sub
        '3. Borrado de partícula en personaje
        Case "XAB"
            rData = Right$(rData, Len(rData) - 3)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call Engine.Char_Particle_Group_Remove(CharIndex, Val(General_Field_Read(2, rData, ",")))
            Exit Sub
        '4. Borrado de todas las partículas en personaje
        Case "XAT"
            rData = Right$(rData, Len(rData) - 3)
            CharIndex = Engine.Char_Find(Val(rData))
            Call Engine.Char_Particle_Group_Remove_All(CharIndex)
            Exit Sub
        '[/Barrin: Partículas]
        
        Case "AYM"
            Dim n As String, n2 As String
            rData = Right$(rData, Len(rData) - 3)
            n = General_Field_Read(2, rData, "°")
            n2 = General_Field_Read(1, rData, "°")
            frmPanelGm.Show vbModeless, frmMain
            Exit Sub
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            rData = Right$(rData, Len(rData) - 3)
            Call ActualizarEst(Val(General_Field_Read(1, rData, ",")), Val(General_Field_Read(2, rData, ",")), _
                Val(General_Field_Read(3, rData, ",")), Val(General_Field_Read(4, rData, ",")), _
                Val(General_Field_Read(5, rData, ",")), Val(General_Field_Read(6, rData, ",")), _
                Val(General_Field_Read(7, rData, ",")), Val(General_Field_Read(8, rData, ",")), _
                Val(General_Field_Read(9, rData, ",")), Val(General_Field_Read(10, rData, ",")), _
                Val(General_Field_Read(11, rData, ",")), Val(General_Field_Read(12, rData, ",")), True)
        '[Barrin]
        Case "QUE"
            rData = Right$(rData, Len(rData) - 3)
            
            If rData = "" Then
                Call AddtoRichTextBox(frmMain.RecTxt, "¡Aún no has aceptado ninguna propuesta!", 204, 193, 115, 0, 1)
            Else
                frmQuestActual.Show vbModeless, frmMain
                frmQuestActual.ParseQuestInfo (rData)
            End If
            
            Exit Sub
        '[/Barrin]
        
        Case "T19"                  ' >>>>> TRABAJANDO :: TRA
            rData = Right$(rData, Len(rData) - 3)
            CurrentUser.UsingSkill = Val(rData)
            frmMain.MousePointer = 2
            Select Case CurrentUser.UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el sitio donde quieres pescar...", 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el árbol...", 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el yacimiento...", 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la fragua...", 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                '[Barrin]
                Case Arrojadizas
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre la victima...", 100, 100, 120, 0, 0)
                Case Jardineria
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el recurso...", 100, 100, 120, 0, 0)
                Case Esposas
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el criminal...", 100, 100, 120, 0, 0)
                Case Musica
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el objetivo...", 100, 100, 120, 0, 0)
                Case Grupo
                    Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre un personaje...", 100, 100, 120, 0, 0)
                '[/Barrin]
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            rData = Right$(rData, Len(rData) - 3)
            TempInt = General_Field_Read(1, rData, ",")
            UserInventory(TempInt).OBJIndex = Val(General_Field_Read(2, rData, ","))
            UserInventory(TempInt).Name = IIf(General_Field_Read(3, rData, ",") = "0", "(Nada)", General_Field_Read(3, rData, ","))
            UserInventory(TempInt).Amount = Val(General_Field_Read(4, rData, ","))
            UserInventory(TempInt).Equipped = Val(General_Field_Read(5, rData, ","))
            UserInventory(TempInt).GrhIndex = Val(General_Field_Read(6, rData, ","))
            UserInventory(TempInt).ObjType = Val(General_Field_Read(7, rData, ","))
            UserInventory(TempInt).MaxHIT = Val(General_Field_Read(8, rData, ","))
            UserInventory(TempInt).MinHIT = Val(General_Field_Read(9, rData, ","))
            UserInventory(TempInt).Def = Val(General_Field_Read(10, rData, ","))
            UserInventory(TempInt).Valor = Val(General_Field_Read(11, rData, ","))
        
            TempStr = ""
            If UserInventory(TempInt).Equipped = 1 Then
                TempStr = TempStr & "(Eqp)"
            End If
            
            If UserInventory(TempInt).Amount > 0 Then
                TempStr = TempStr & "(" & UserInventory(TempInt).Amount & ") " & UserInventory(TempInt).Name
            Else
                TempStr = TempStr & UserInventory(TempInt).Name
            End If
            
            If CurrentUser.Logged Then Engine.Engine_Inventory_Render_Set
            
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            rData = Right$(rData, Len(rData) - 3)
            TempInt = General_Field_Read(1, rData, ",")
            UserBancoInventory(TempInt).OBJIndex = General_Field_Read(2, rData, ",")
            UserBancoInventory(TempInt).Name = IIf(General_Field_Read(3, rData, ",") = "0", "(Nada)", General_Field_Read(3, rData, ","))
            UserBancoInventory(TempInt).Amount = General_Field_Read(4, rData, ",")
            UserBancoInventory(TempInt).GrhIndex = Val(General_Field_Read(5, rData, ","))
            UserBancoInventory(TempInt).ObjType = Val(General_Field_Read(6, rData, ","))
            UserBancoInventory(TempInt).MaxHIT = Val(General_Field_Read(7, rData, ","))
            UserBancoInventory(TempInt).MinHIT = Val(General_Field_Read(8, rData, ","))
            UserBancoInventory(TempInt).Def = Val(General_Field_Read(9, rData, ","))
        
            TempStr = ""
            
            If UserBancoInventory(TempInt).Amount > 0 Then
                TempStr = TempStr & "(" & UserBancoInventory(TempInt).Amount & ") " & UserBancoInventory(TempInt).Name
            Else
                TempStr = TempStr & UserBancoInventory(TempInt).Name
            End If
                        
            Exit Sub
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            rData = Right$(rData, Len(rData) - 3)
            TempInt = General_Field_Read(1, rData, ",")
            CurrentUser.UserHechizos(TempInt) = Val(General_Field_Read(2, rData, ","))
            If TempInt > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem IIf(General_Field_Read(3, rData, ",") = "0", "(Nada)", General_Field_Read(3, rData, ","))
            Else
                frmMain.hlst.List(TempInt - 1) = IIf(General_Field_Read(3, rData, ",") = "0", "(Nada)", General_Field_Read(3, rData, ","))
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            rData = Right$(rData, Len(rData) - 3)
            For i = 1 To NUMATRIBUTOS
                CurrentUser.UserAtributos(i) = Val(General_Field_Read(i, rData, ","))
            Next i
            LlegaronAtrib = True
            Call MostrarEstadisticas
            Exit Sub
        Case "LAH"
            rData = Right$(rData, Len(rData) - 3)
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ArmasHerrero(1 To m) As Integer
                        
            For i = 1 To m
                cad$ = General_Field_Read(i, rData, ",")
                ArmasHerrero(i) = Val(General_Field_Read(i + 1, rData, ","))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
            Next i
            
            Exit Sub
         Case "LAR"
            rData = Right$(rData, Len(rData) - 3)
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ArmadurasHerrero(1 To m) As Integer
                        
            For i = 1 To m
                cad$ = General_Field_Read(i, rData, ",")
                ArmadurasHerrero(i) = Val(General_Field_Read(i + 1, rData, ","))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
            Next i
            
            Exit Sub
        '[KEVIN]
        Case "DRP" 'para la lista de pociones
            rData = Right$(rData, Len(rData) - 3)
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ObjDruida(1 To m) As Integer
                        
            For i = 1 To m
                cad$ = General_Field_Read(i, rData, ",")
                ObjDruida(i) = Val(General_Field_Read(i + 1, rData, ","))
                If cad$ <> "" Then frmDruida.lstPociones.AddItem cad$
            Next i
            
            Exit Sub
        
        Case "SAR"
            rData = Right$(rData, Len(rData) - 3)
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ObjSastre(1 To m) As Integer
                        
            For i = 1 To m
                cad$ = General_Field_Read(i, rData, ",")
                ObjSastre(i) = Val(General_Field_Read(i + 1, rData, ","))
                If cad$ <> "" Then frmSastre.lstRopas.AddItem cad$
            Next i
            
            Exit Sub
            
         Case "OBR"
            rData = Right$(rData, Len(rData) - 3)
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ObjCarpintero(1 To m) As Integer
                        
            For i = 1 To m
                cad$ = General_Field_Read(i, rData, ",")
                ObjCarpintero(i) = Val(General_Field_Read(i + 1, rData, ","))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
            Next i
            
            Exit Sub
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            CurrentUser.Descansando = Not CurrentUser.Descansando
            Exit Sub
        Case "SPL"
            rData = Right(rData, Len(rData) - 3)
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmSpawnList.lstCriaturas.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmSpawnList.Show vbModeless, frmMain
            Exit Sub
        Case "ERR"
            rData = Right$(rData, Len(rData) - 3)
            frmConnect.MousePointer = 1
            frmPasswd.MousePointer = 1
            If frmMain.mainWinsock.State Then frmMain.mainWinsock.Close
            Call ResetCurrentUser
            
            If EstadoLogin = CrearNuevoPj Then
                If frmPasswd.Visible Then frmPasswd.lblStatus.Caption = "Servidor> " & rData
            ElseIf EstadoLogin = NORMAL Then
                If frmIniciando.Visible Then
                    frmConnect.Show
                    Unload frmIniciando
                End If
                
                frmMensaje.msg.Caption = rData
                frmMensaje.Show vbModal, frmConnect
                
            Else
                Call MsgBox(rData, vbExclamation, "Mensaje del servidor")
            End If
            
            Exit Sub
        '[Barrin]
        Case "BCL"
            rData = Right(rData, Len(rData) - 3)
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmHunter.lstBuscados.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmHunter.Show vbModeless, frmMain
            Exit Sub
        Case "GCL"
            Dim EsLider As Byte
            frmGrupo.lstGrupo.Clear
            rData = Right(rData, Len(rData) - 3)
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmGrupo.lstGrupo.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            EsLider = General_Field_Read(Val(General_Field_Read(1, rData, ",")) + 2, rData, ",")
            If EsLider = 1 Then
                frmGrupo.cmdExpulsar.Enabled = True
                frmGrupo.cmdInvitar.Enabled = True
            Else
                frmGrupo.cmdExpulsar.Enabled = False
                frmGrupo.cmdInvitar.Enabled = False
            End If
            frmGrupo.Show vbModeless, frmMain
            Exit Sub
        Case "ERT"
            rData = Right$(rData, Len(rData) - 3)
            If rData = "JOYA" Then
                Unload frmTorneoCrear
                Exit Sub
            End If
            MsgBox rData
            Exit Sub
        '[/Barrin]
    End Select
    
    '[Alejo-21-5]
    Select Case left(sData, 4)
        Case "DUMB"
            CurrentUser.Estupido = True
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            rData = Right$(rData, Len(rData) - 4)
            Call Engine.Letter_Set(CInt(General_Field_Read(2, rData, "°")), General_Field_Read(1, rData, "°"))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            rData = Right(rData, Len(rData) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = General_Field_Read(1, rData, ",")
            NPCInventory(NPCInvDim).Amount = General_Field_Read(2, rData, ",")
            NPCInventory(NPCInvDim).Valor = General_Field_Read(3, rData, ",")
            NPCInventory(NPCInvDim).GrhIndex = General_Field_Read(4, rData, ",")
            NPCInventory(NPCInvDim).OBJIndex = General_Field_Read(5, rData, ",")
            NPCInventory(NPCInvDim).ObjType = General_Field_Read(6, rData, ",")
            NPCInventory(NPCInvDim).MaxHIT = General_Field_Read(7, rData, ",")
            NPCInventory(NPCInvDim).MinHIT = General_Field_Read(8, rData, ",")
            NPCInventory(NPCInvDim).Def = General_Field_Read(9, rData, ",")
            NPCInventory(NPCInvDim).C1 = General_Field_Read(10, rData, ",")
            NPCInventory(NPCInvDim).C2 = General_Field_Read(11, rData, ",")
            NPCInventory(NPCInvDim).c3 = General_Field_Read(12, rData, ",")
            NPCInventory(NPCInvDim).C4 = General_Field_Read(13, rData, ",")
            NPCInventory(NPCInvDim).C5 = General_Field_Read(14, rData, ",")
            NPCInventory(NPCInvDim).C6 = General_Field_Read(15, rData, ",")
            NPCInventory(NPCInvDim).C7 = General_Field_Read(16, rData, ",")
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            rData = Right$(rData, Len(rData) - 4)
            CurrentUser.UserMaxAGU = Val(General_Field_Read(1, rData, ","))
            CurrentUser.UserMinAGU = Val(General_Field_Read(2, rData, ","))
            CurrentUser.UserMaxHAM = Val(General_Field_Read(3, rData, ","))
            CurrentUser.UserMinHAM = Val(General_Field_Read(4, rData, ","))
            frmMain.AGUAsp.Width = (((CurrentUser.UserMinAGU / 100) / (CurrentUser.UserMaxAGU / 100)) * 91)
            frmMain.COMIDAsp.Width = (((CurrentUser.UserMinHAM / 100) / (CurrentUser.UserMaxHAM / 100)) * 91)
            frmMain.lblHAM.Caption = CurrentUser.UserMinHAM & "/" & CurrentUser.UserMaxHAM
            frmMain.lblSED.Caption = CurrentUser.UserMinAGU & "/" & CurrentUser.UserMaxAGU
            Exit Sub
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            rData = Right$(rData, Len(rData) - 4)
            CurrentUser.UserReputacion.AsesinoRep = Val(General_Field_Read(1, rData, ","))
            CurrentUser.UserReputacion.BandidoRep = Val(General_Field_Read(2, rData, ","))
            CurrentUser.UserReputacion.BurguesRep = Val(General_Field_Read(3, rData, ","))
            CurrentUser.UserReputacion.LadronesRep = Val(General_Field_Read(4, rData, ","))
            CurrentUser.UserReputacion.NobleRep = Val(General_Field_Read(5, rData, ","))
            CurrentUser.UserReputacion.PlebeRep = Val(General_Field_Read(6, rData, ","))
            CurrentUser.UserReputacion.Promedio = Val(General_Field_Read(7, rData, ","))
            CurrentUser.UserReputacion.Culpabilidad = Val(General_Field_Read(8, rData, ","))
            LlegoFama = True
            Call MostrarEstadisticas
            Exit Sub
        Case "SUNI"
            rData = Right$(rData, Len(rData) - 4)
            CurrentUser.SkillPoints = Val(rData)
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            rData = Right$(rData, Len(rData) - 4)
            AddtoRichTextBox frmMain.RecTxt, "Hay " & rData & " npcs.", 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            rData = Right$(rData, Len(rData) - 4)
            TempInt = InStr(1, rData, "µ")
            k = Val(General_Field_Read(4, rData, "µ"))
                        
            If TempInt > 0 Then
                If k = 1 Then 'Barrin: Está online!
                    TempStr = left(rData, TempInt - 1) & " (" & General_Field_Read(3, rData, "µ") & ")"
                    frmPanelGm.List1.AddItem TempStr
                    frmPanelGm.MensajePoner TempStr, General_Field_Read(2, rData, "µ")
                Else
                    TempStr = left(rData, TempInt - 1) & " (" & General_Field_Read(3, rData, "µ") & ")"
                    frmPanelGm.List2.AddItem TempStr
                    frmPanelGm.MensajePoner TempStr, General_Field_Read(2, rData, "µ")
                End If
            Else
                frmPanelGm.List1.AddItem rData
            End If
            
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmPanelGm.Show vbModeless, frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            rData = Right$(rData, Len(rData) - 4)
            frmForo.List.AddItem General_Field_Read(1, rData, "°")
            frmForo.Text(frmForo.List.ListCount - 1).Text = General_Field_Read(2, rData, "°")
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show vbModeless, frmMain
            End If
            Exit Sub
        Case "AGMS"
            rData = Right$(rData, Len(rData) - 4)
            If Val(rData) = 1 Then
                frmGMAyuda.Label2.Caption = "Deje los siguientes datos: nombre del clan (sensitivo a mayúsculas y minúsculas), alineamiento y otros datos que considere de importancia."
                frmGMAyuda.optConsulta.Item(6).Value = True
            End If
            If Not frmGMAyuda.Visible Then
                frmGMAyuda.Show vbModeless, frmMain
                frmGMAyuda.txtMotivo.SetFocus
            End If
            Exit Sub
        Case "VMAP"
            If Not frmMapa.Visible Then
                  frmMapa.Show vbModeless, frmMain
            Else
                  frmMapa.SetFocus
            End If
            Exit Sub
        Case "FAMI"         'Información del familiar o mascota
            rData = Right$(rData, Len(rData) - 4)
            CurrentUser.UserPet.TieneFamiliar = Val(General_Field_Read(1, rData, ","))
            
            If CurrentUser.UserPet.TieneFamiliar <> 0 Then
                CurrentUser.UserPet.ELU = Val(General_Field_Read(2, rData, ","))
                CurrentUser.UserPet.ELV = Val(General_Field_Read(3, rData, ","))
                CurrentUser.UserPet.EXP = Val(General_Field_Read(4, rData, ","))
                CurrentUser.UserPet.MaxHP = Val(General_Field_Read(5, rData, ","))
                CurrentUser.UserPet.MinHP = Val(General_Field_Read(6, rData, ","))
                CurrentUser.UserPet.nombre = General_Field_Read(7, rData, ",")
                CurrentUser.UserPet.MinHIT = Val(General_Field_Read(8, rData, ","))
                CurrentUser.UserPet.MaxHIT = Val(General_Field_Read(9, rData, ","))
                CurrentUser.UserPet.Abilidad = HabilidadToString(General_Field_Read(10, rData, ","))
            End If
            
            LlegoFami = True
            Call MostrarEstadisticas
            Exit Sub
        Case "GOLI"
            rData = Right$(rData, Len(rData) - 4)
            Call frmGoliath.ParseBancoInfo(rData)
        Case "YEGU"
            Call Engine.Char_Current_Blind_Set(True)
            Exit Sub
        Case "YEGS"
            rData = Right$(rData, Len(rData) - 4)
            CurrentUser.UserStats.CiudasMatados = Val(General_Field_Read(1, rData, ","))
            CurrentUser.UserStats.CrimisMatados = Val(General_Field_Read(2, rData, ","))
            CurrentUser.UserStats.NPCsMatados = Val(General_Field_Read(3, rData, ","))
            CurrentUser.UserStats.Clase = Val(General_Field_Read(4, rData, ","))
            CurrentUser.UserStats.TimesKilled = Val(General_Field_Read(5, rData, ","))
            CurrentUser.UserStats.Raza = Val(General_Field_Read(6, rData, ","))
            CurrentUser.UserStats.Genero = Val(General_Field_Read(7, rData, ","))
            LlegoEst = True
            Call MostrarEstadisticas
            Exit Sub
        Case "TPRO"
            rData = Right$(rData, Len(rData) - 4)
            PuedeTorneo = rData
            Exit Sub
        '[El Yind]
        Case "INFT"
            rData = Right$(rData, Len(rData) - 4)
            Dim TInscriptos As Integer
            Dim loopc As Integer
            If left$(rData, 2) = "LI" Then
                    rData = Right$(rData, Len(rData) - 2)
                    TInscriptos = Val(General_Field_Read(1, rData, ","))
                    If TInscriptos > 0 Then
                        For loopc = 1 To TInscriptos
                            frmTorneosLider.members.AddItem General_Field_Read(1 + loopc, rData, ",")
                        Next loopc
                    End If
                frmTorneosLider.txtguildnews.Text = General_Field_Read(2 + TInscriptos, rData, ",")
                If Not frmTorneosLider.Visible Then frmTorneosLider.Show vbModeless, frmMain
            Else
                If rData = "NO" Then
                    frmTorneo.TXTNombreT = "No hay ningun torneo."
                    frmTorneo.COT.Enabled = True
                    frmTorneo.txtguildnews = ""
                    frmTorneo.ListIns.Clear
                    frmTorneo.TXTLider = ""
                    frmTorneo.Command2.Enabled = False
                Else
                    frmTorneo.TXTNombreT = General_Field_Read(1, rData, ",")
                    frmTorneo.COT.Enabled = False
                    frmTorneo.TXTPjs = General_Field_Read(2, rData, ",")
                    frmTorneo.Val1.Value = IIf(General_Field_Read(3, rData, ",") = 0, True, False)
                    frmTorneo.Val2.Value = IIf(General_Field_Read(3, rData, ",") = 1, True, False)
                    frmTorneo.Modo1.Value = IIf(General_Field_Read(4, rData, ",") = 0, True, False)
                    frmTorneo.Modo2.Value = IIf(General_Field_Read(4, rData, ",") = 1, True, False)
                    frmTorneo.TXTPR = General_Field_Read(5, rData, ",")
                    frmTorneo.TXTPrecio = IIf(General_Field_Read(6, rData, ",") = 0, "GRATIS", General_Field_Read(6, rData, ","))
                    TInscriptos = Val(General_Field_Read(7, rData, ","))
                    If TInscriptos > 0 Then
                        For loopc = 1 To TInscriptos
                            frmTorneo.ListIns.AddItem General_Field_Read(7 + loopc, rData, ",")
                        Next loopc
                    End If
                    frmTorneo.txtguildnews.Text = General_Field_Read(8 + TInscriptos, rData, ",")
                    frmTorneo.TXTLider = General_Field_Read(9 + TInscriptos, rData, ",")
                End If
                frmTorneo.Show vbModeless, frmMain
            End If
            Exit Sub
        '[/El Yind]
    End Select
    
    Select Case left(sData, 5)
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            CurrentUser.Meditando = Not CurrentUser.Meditando
            Exit Sub
        Case "UNSEK"             ' >>>>> Invisible :: NOVER
            rData = Right$(rData, Len(rData) - 5)
            CharIndex = Engine.Char_Find(Val(General_Field_Read(1, rData, ",")))
            Call Engine.Char_Invisible_Set(CharIndex, Val(General_Field_Read(2, rData, ",")))
            Exit Sub
        Case "ZMOTD"
            rData = Right$(rData, Len(rData) - 5)
            frmCambiaMotd.Show vbModeless, frmMain
            frmCambiaMotd.txtMotd.Text = rData
            Exit Sub
        '[Barrin]
        Case "NFAMI"
            frmSeleccionFamiliar.Show vbModeless, frmMain
            Exit Sub
        Case "QUNFO"
            rData = Right$(rData, Len(rData) - 5)
            frmQuest.Show vbModeless, frmMain
            frmQuest.ParseQuestInfo (rData)
        '[/Barrin]
        
    End Select
    
    Select Case left(sData, 6)
        Case "NSEGUE"
            Call Engine.Char_Current_Blind_Set(False)
            Exit Sub
        Case "NESTUP"
            CurrentUser.Estupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            rData = Right$(rData, Len(rData) - 6)
            For i = 1 To NUMSKILLS
                CurrentUser.UserSkills(i) = Val(General_Field_Read(i, rData, ","))
            Next i
            LlegaronSkills = True
            Call MostrarEstadisticas
            Exit Sub
        Case "LSTCRI"
            rData = Right(rData, Len(rData) - 6)
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmEntrenador.lstCriaturas.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmEntrenador.Show vbModeless, frmMain
            Exit Sub
        Case "LISTUS"
            rData = Right(rData, Len(rData) - 6)
            t = Split(rData, ",")
            frmPanelGm.cboListaUsus.Clear
            For i = LBound(t) To UBound(t)
                frmPanelGm.cboListaUsus.AddItem t(i)
            Next i
            If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            Exit Sub
        Case "FREZOK"         ' >>>>> Paralizar OK :: PARADOK
            CurrentUser.Paralizado = Not CurrentUser.Paralizado
            
            If CurrentUser.Paralizado Then
                rData = Right(rData, Len(rData) - 6)
                x = CInt(General_Field_Read(1, rData, ","))
                y = CInt(General_Field_Read(2, rData, ","))
                Call Engine.Char_Current_Pos_Refresh(x, y)
                Call Engine.Engine_View_Pos_Set(x, y)
            End If
            
            Exit Sub
    End Select
    
    Select Case left(sData, 7)
        Case "KUILDNE"
            rData = Right(rData, Len(rData) - 7)
            Call frmGuildNews.ParseGuildNews(rData)
            Exit Sub
        Case "PEACEDE"
            rData = Right(rData, Len(rData) - 7)
            Call frmUserRequest.recievePeticion(rData)
            Exit Sub
        Case "PETICIO"
            rData = Right(rData, Len(rData) - 7)
            Call frmUserRequest.recievePeticion(rData)
            Exit Sub
        Case "PEACEPR"
            rData = Right(rData, Len(rData) - 7)
            Call frmPeaceProp.ParsePeaceOffers(rData)
            Exit Sub
        Case "GIRINFO"
            rData = Right(rData, Len(rData) - 7)
            Call frmCharInfo.parseCharInfo(rData)
            Exit Sub
        Case "BCDINFO"
            rData = Right(rData, Len(rData) - 7)
            Call frmBuscadoInfo.parseBuscadoInfo(rData)
            Exit Sub
        Case "LEADERI"
            rData = Right(rData, Len(rData) - 7)
            Call frmGuildLeader.ParseLeaderInfo(rData)
            Exit Sub
        Case "CLANDET"
            rData = Right(rData, Len(rData) - 7)
            Call frmGuildBrief.ParseGuildInfo(rData)
            Exit Sub
        Case "SHOWFUN"
            rData = Right(rData, Len(rData) - 7)
            CurrentUser.CreandoClan = True
            frmGuildFoundation.txtClanName = rData
            Call frmGuildFoundation.Show(vbModeless, frmMain)
            Exit Sub
        Case "METAMOK"
            CurrentUser.Transformado = Not CurrentUser.Transformado
            Call Engine.Char_Current_OnHorse_Set(CurrentUser.Transformado)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= UBound(UserInventory)
                    If UserInventory(i).OBJIndex <> 0 Then
                        frmComerciar.List1(1).AddItem UserInventory(i).Name
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                rData = Right(rData, Len(rData) - 7)
                
                If General_Field_Read(2, rData, ",") = "0" Then
                    frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                i = 1
                Do While i <= UBound(UserInventory)
                    If UserInventory(i).OBJIndex <> 0 Then
                            frmBancoObj.List1(1).AddItem UserInventory(i).Name
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                
                ii = 1
                Do While ii <= UBound(UserBancoInventory)
                    If UserBancoInventory(ii).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    ii = ii + 1
                Loop
                
                rData = Right(rData, Len(rData) - 7)
                
                If General_Field_Read(2, rData, ",") = "0" Then
                    frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                    frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
    End Select
        
End Sub

Public Sub Login(ByVal valcode As Integer)

'Personaje grabado
If EstadoLogin = NORMAL Then
    Call tcp.Send_Data(Old_Login, CurrentUser.UserName & "," & CurrentUser.UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
'Crear personaje
ElseIf EstadoLogin = CrearNuevoPj Then
    If CurrentUser.UserClase = MAGO Or CurrentUser.UserClase = DRUIDA Or CurrentUser.UserClase = CAZADOR Then
        Call tcp.Send_Data(New_Login, CurrentUser.UserName & "," & CurrentUser.UserPassword _
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
        Call tcp.Send_Data(New_Login, CurrentUser.UserName & "," & CurrentUser.UserPassword _
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

Function Decrypt_Data(ByVal strText As String, ByVal strPwd As String) As String
    
    Dim i As Long, C As Integer
    Dim strBuff As String

    If Len(strPwd) Then
        For i = 1 To Len(strText)
            C = Asc(mid$(strText, i, 1))
            C = C - Asc(mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(C And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    
    Decrypt_Data = strBuff

End Function

Function Encrypt_Data(ByVal strText As String, ByVal strPwd As String) As String
    
    Dim i As Long, C As Integer
    Dim strBuff As String

    If Len(strPwd) Then
        For i = 1 To Len(strText)
            C = Asc(mid$(strText, i, 1))
            C = C + Asc(mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(C And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    
    Encrypt_Data = strBuff

End Function

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
