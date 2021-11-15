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
'clsClientTCP.cls - ImperiumAO - v1.3 - www.imperiumao.com.ar
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

Public Enum Client_Paquet_ID
    Auth_Start = 1
    Old_Login
    New_Login
    Normal_Chat
    Shout_Chat
    Private_Chat
    Global_Chat
    Move_Char_Cl_North
    Move_Char_Cl_South
    Move_Char_Cl_East
    Move_Char_Cl_West
    Drop_Item
    Cast_Spell
    Left_Click
    Right_Click
    Working_Click  '"UK"
    Get_Item
    User_Commerce
    Cancel_Exit  'CU
    Request_Pos
    Attack
    Combat_Mode
    Global_Option
    Loading_Finished
    Gold_Transfer
    Safe_Mode
    Group_Remove_Self
    Group_Remove_Member
    Group_Toggle  'GRI
    Group_Member_List
    Quest_Data_Cl
    Role_Mode
    Corpse_Take_Gold
    Game_Master_Support
    Use_Item
    Druid_Build_Item
    Tailor_Build_Item
    Guild_Create
    Working_Left_Click  'WLC
    Familiar_Adopt
    Client_MD5
    Spell_Info_Request
    Equip_Item
    Change_Heading_Cl_North
    Change_Heading_Cl_South
    Change_Heading_Cl_East
    Change_Heading_Cl_West
    Skill_Data_Cl  'SKSE
    Training_Spawn
    Main_Command
    Main_GM_Command
    Sell_Item
    Carpenter_Build
    Herrero_Build
    Get_Item_Bank
    Buy_Item
    Insert_Item_Bank
    Bank_End
    Bounty_Info_Cl  'BCDI
    Bounty_Offer
    Bounty_Add
    Challenge_Desc_Set  'TACT
    Challenge_Init  'ATOR
    Challenge_Inscription  'INSC
    Challenge_Start  'TRUN
    Challenge_Create  'CTOR
    Stats_Att_Request
    Stats_Fame_Request
    Stats_Familiar_Request
    Stats_General_Request
    Stats_Skills_Request
    Quest_Start
    Ping_Request
    Coffer_End
    Corpse_End
    Corpse_Take_Item  'ARET
    Spell_Move
    Guild_Desc_Codec_Update
    Guild_Info_Request
    Commerce_End
    Forum_Acc
    Quest_End
    Bank_Init  'INITBOV
    Midi_Request
    Amb_Request
    Guild_Peace_OK  'ACEPPEAT
    Guild_Peace_Send_All
    Guild_Peace_Receive  'PEACEOFF
    Guild_Peace_Request  'PEACEDET
    Guild_Send_Petition  'ENVCOMEN
    Guild_Declare_War  'DECGUERR
    Guild_Declare_Allie  'DECALIAD
    Guild_Set_URL  'NEWWEBSI
    Guild_Accept_Member  'ACEPTARI
    Guild_Deny_Member  'RECHAZAR
    Guild_Remove_Member  'ECHARCLA
    Guild_News_Set  'ACTGNEWS
    Guild_Char_Info_Cl  '1HRINFO<
    Guild_Sol_Get  'SOLICITUD
    Guild_Details_Request  'CLANDETAI
    GM_User_List_Cl
    GM_Sos_Erase
    GM_Sos_Inv
    GM_Sos_Manual
    GM_Motd_Change
    GM_Spawn
End Enum

Public Enum Server_Paquet_ID
    Conn_Succ = 1
    Navigating_Toggle
    Horse_Toggle
    'Fly_Toggle
    End_Succ
    Comm_End_Succ
    Bank_End_Succ
    Comm_Init
    Bank_Data
    Herrero_Init
    Carp_Init
    Druid_Init
    Tailor_Init
    Rol_Toggle
    Message_Revived
    Message_Cured
    Message_Piquete
    Message_NPC_Failed
    Message_NPC_Killed
    Message_Shield_Blocked
    Message_Shield_Blocked_Our
    Message_Less_Tired
    Message_Not_Hide
    Message_Offline_User
    Message_Nothing
    Message_Select_Char
    Message_Not_Amount
    Message_NPC_Die
    Message_Dead
    Message_Not_Special_Move
    Message_Not_Stab
    Message_User_Not_Exist
    Message_Reviving
    Message_Corpse_Formed
    Message_Invalid_Rules
    Message_Invalid_Manual
    Message_GameMasters_Sent
    Message_GameMasters_Already_Sent
    Message_Not_Gold
    Message_Failed
    Message_Safe_Exit
    Ping_Received
    Map_Load_Request
    Main_Map_Data
    Closing
    Close_Cancel
    Exp_Won
    Group_Exp_Won
    Dialog_Hit_Spell  'GH
    Dialog_Hit_Weap  'GU
    Position_Current  'PU
    Position_Other  'PS
    Fly_Toggle
    NPC_Hit
    User_NPC_Hit
    User_Failed_User
    User_Hit_Current_User
    User_Hit_User
    Main_Dialog  '||
    Magic_Words  '()
    Global_Chat_Message
    Stats_HP
    Stats_STA
    Stats_MAN
    Stats_GLD
    Stats_EXP
    Stats_ATT
    Server_Message
    Create_Char
    Create_Char_NPC
    Create_Char_Current
    Change_Heading_North
    Change_Heading_South
    Change_Heading_East
    Change_Heading_West
    Change_Type
    Change_Name
    Change_Group
    Erase_Char
    Guild_Info
    Move_Char_North
    Move_Char_South
    Move_Char_East
    Move_Char_West
    Change_Char
    Item_Create
    Item_Erase
    Pos_Bloq
    Music_Change
    Ambient_Change
    Sound_Play
    Guild_List
    Commerce_Data  'C2
    Main_Pause
    Client_Validation
    Rain_Toggle
    Snow_Toggle
    Hour_Night_Toggle  'NOC
    Dialog_Erase
    FX_Create
    Map_FX_Create
    Particle_Char_Add
    Particle_Map_Add
    Particle_Char_Remove
    Particle_All_Char_Remove
    GM_Panel  '"AYM"
    User_Stats  'EST
    Quest_Info  'QUE
    Working_Main  'T19
    Inventory_Update  'CSI
    Bank_Update  'SBO
    Spell_Add  'SHS
    Att_Receive  'ATR
    Weapon_List  'LAH
    Armor_List  'LAR
    Potion_List  'DRP
    Tailor_List  'SAR
    Carpenter_List  'OBR
    Rest_Toggle  'DOK
    GM_Spawn_List  'SPL
    Error_Message  'ERR
    Bounty_List  'BCL
    Group_List  'GCL
    Challenge_Data  'ERT
    Stupid_True  'DUMB
    Letter_Show
    NPC_Inv_Receive
    HYS_Update  'EHYS
    Fame_Update  'FAMA
    Skill_Data  'SUNI
    Number_Enemys  'NENE
    GM_Sos_Message  'RSOS
    GM_Sos_Show  'MSOS
    Forum_Main  'FMSG
    Forum_Show  'MFOR
    GameMaster_Message_Show  'AGMS
    Map_Show
    Familiar_Info  'FAMI
    Goliath_Info  'GOLI
    Blind_True  'YEGU
    Stats_Mini  'YEGS
    Challenge_Can  'TPRO
    Challenge_Info  'INFT
    Meditating_Toggle  'MEDOK
    char_invisible  'UNSEK
    GM_Motd_Info
    Familiar_Adopt_Show
    Quest_Data
    Blind_False
    Stupid_False
    Skills_Rec
    Train_NPC_List
    GM_User_List
    Freezed_Toggle
    Guild_News  'KUILDNE
    Guild_Petition  'PEACEDE, PETICIO
    Guild_Peace_Prop  'PEACEPR
    Guild_Char_Info  'GIRINFO
    Bounty_Info  'BCDINFO
    Guild_Lead_Info 'LEADERI
    Guild_Clan_Details 'CLANDET
    Guild_Found_Show 'SHOWFUN
    Metam_Toggle 'METAMOK
    Commerce_OK 'TRANSOK
    Bank_OK 'BANCOOK
End Enum

Public Enum Client_User_Command_ID
    cmdSalir = 1
    cmdMeditar
    cmdCurar
    cmdOnline
    cmdOnlineClan
    cmdGM
    cmdEst
    cmdSubasta
    cmdFami
    cmdAyuda
    cmdPing
    cmdBoveda
    cmdComerciar
    cmdCuanto
    cmdRechazarMatrimonio
    cmdAcompañar
    cmdInformacion
    cmdTorneos
    cmdBalance
    cmdFundarClan
    cmdSalirClan
    cmdDivorcio
    cmdPena
    cmdRetirar
    cmdPasswd
    cmdOfertar
    cmdOfertaInicial
    cmdDepositar
    cmdGRMG
    cmdCMSG
    cmdApostar
    cmdDesc
    cmdBug
    cmdProponer
    cmdCentinela
    cmdErcmd666
End Enum

Public Enum Client_GM_Command_ID
    cmdGMSG = 1
    cmdRem
    cmdDonde
    cmdNene
    cmdAtt
    cmdTelep
    cmdIra
    cmdTeleploc
    cmdInvisible
    cmdTrabajando
    cmdEnGrupo
    cmdOnlineGM
    cmdPanelGM
    cmdCaos
    cmdReal
    cmdSilencio
    cmdEchar
    cmdPatear
    cmdCS
    cmdBanInfo
    cmdRevivir
    cmdLastIP
    cmdMotdCambia
    cmdInfo
    cmdInfoFami
    cmdInv
    cmdBov
    cmdMail
    cmdHmr
    cmdSkills
    cmdCarcel
    cmdBan
    cmdAutorizar
    cmdUnBanGen
    cmdChatGen
    cmdUnBan
    cmdSeguir
    cmdSum
    cmdLimpiar
    cmdRMSG
    cmdAMSG
    cmdNick2IP
    cmdSameIP
    cmdIP2Nick
    cmdCleanMap
    cmdFaccion
    cmdPerdon
    cmdTrigger
    cmdBanIPList
    cmdBanIP
    cmdUnBanIP
    cmdCT
    cmdDT
    cmdDest
    cmdNickProhibido
    cmdCI
    cmdBloq
    cmdMassKill
    cmdSMSG
    cmdMata
    cmdCC
    cmdAcc
    cmdRacc
    cmdNave
    cmdApagar
    cmdConden
    cmdGuardarMapa
    cmdModMapInfo
    cmdReload
    cmdReiniciar
    cmdIniciar
    cmdKilLApp
    cmdMod
    cmdDoBackUP
    cmdGrabar
    cmdBorrarSos
    cmdShowInt
    cmdLluvia
    cmdNieve
    cmdOnlineMap
    cmdModoRol
    cmdMassDest
    cmdMD5OK
    cmdPlayWav
    cmdModNick
    cmdQuery
    cmdListProc
    cmdModPass
    cmdDelChar
    cmdNoCrear
End Enum

Private LlegaronSkills As Boolean
Private LlegaronAtrib As Boolean
Private LlegoFama As Boolean
Private LlegoFami As Boolean
Private LlegoEst As Boolean

Public Sub Reset_Skill_Data()

LlegaronSkills = False
LlegaronAtrib = False
LlegoFama = False
LlegoFami = False
LlegoEst = False

End Sub

Public Sub Handle_Data(ByVal rData As String)
    
    On Error Resume Next
    
    Dim X As Integer
    Dim Y As Integer
    Dim CharIndex As Integer
    Dim TempInt As Integer
    Dim tempStr As String
    Dim tempLng As Long
    Dim i As Integer, k As Integer
    Dim cad$, m As Integer
    Dim t() As String
    Dim part_life() As Long
    Dim part_type() As Integer
    Dim Paquet_ID As Byte
    Dim tempByte As Byte
    
    Paquet_ID = Asc(left$(rData, 1))
    'Paquet_ID = AscB(rData)
    
    rData = Right$(rData, Len(rData) - 1)
    
    Select Case Paquet_ID
        Case Conn_Succ
            If EstadoLogin = CrearNuevoPj Then
                Unload frmPasswd
                Unload frmCrearPersonaje
                frmIniciando.Show
            ElseIf EstadoLogin = NORMAL Then
                frmIniciando.Show
                frmConnect.Visible = False
            End If
            
            Call Map_Load(String_To_Integer(rData, 1), String_To_Byte(rData, 3), String_To_Long(rData, 6), String_To_Byte(rData, 6))
            
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
            frmMensaje.mnuClan.Enabled = String_To_Byte(rData, 4)
            frmMensaje.mnuGMs.Enabled = String_To_Byte(rData, 5)
            '[/Barrin]
            
            Call DibujarMenuMacros
            Call Inventory_Render
            
            'Barrin
            Call Send_Data(Loading_Finished, Byte_To_String(MensajesGlobales))
        
        Case Navigating_Toggle
            CurrentUser.Navegando = Not CurrentUser.Navegando
            Engine.Char_Current_OverWater_Set (CurrentUser.Navegando)
        
        Case Horse_Toggle
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
        
        'Case Fly_Toggle
        '    CurrentUser.Volando = Not CurrentUser.Volando

        Case End_Succ
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

        Case Comm_End_Succ
            Unload frmComerciar
            CurrentUser.Comerciando = False
        
        Case Bank_End_Succ
            Unload frmBancoObj
            CurrentUser.Comerciando = False

        Case Comm_Init
            i = 1
            Do While i <= UBound(UserInventory)
                If UserInventory(i).OBJIndex <> 0 Then
                    frmComerciar.List1(1).AddItem UserInventory(i).Name
                Else
                    frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            
            i = 1
            Do While i <= UBound(NPCInventory)
                If NPCInventory(i).OBJIndex <> 0 Then
                    frmComerciar.List1(0).AddItem NPCInventory(i).Name
                Else
                    frmComerciar.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            
            CurrentUser.Comerciando = True
            frmComerciar.Show vbModeless, frmMain
        
        Case Bank_Data
            i = 1
            Do While i <= UBound(UserInventory)
                If UserInventory(i).OBJIndex <> 0 Then
                    frmBancoObj.List1(1).AddItem UserInventory(i).Name
                Else
                    frmBancoObj.List1(1).AddItem "Nada"
                End If
                i = i + 1
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
        
        Case Herrero_Init
            frmHerrero.Show vbModeless, frmMain
        
        Case Carp_Init
            frmCarp.Show vbModeless, frmMain
        
        Case Druid_Init
            frmDruida.Show vbModeless, frmMain
        
        Case Tailor_Init
            frmSastre.Show vbModeless, frmMain
        
        Case Rol_Toggle
        
        Case Message_Revived
            Call AddtoRichTextBox(frmMain.RecTxt, "El cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar... ¡Has sido resucitado!", 0, 0, 0, 0, 0, 0, 4)
            Call Sound.Sound_Play(SND_RESUCITAR)
        
        Case Message_Cured
            Call AddtoRichTextBox(frmMain.RecTxt, "El sacerdote levanta sus manos, recita unas palabras, y comienzas a sentir un fuerte ardor. Luego ves como van cerrando tus heridas... ¡Has sido curado!", 0, 0, 0, 0, 0, 0, 4)
            Call Engine.Char_FX_Set(CurrentUser.CurrentChar, 9, 1)
            Call Sound.Sound_Play(SND_CURAR)
        
        Case Message_Piquete
            Call AddtoRichTextBox(frmMain.RecTxt, "Estás obstruyendo la vía pública, ¡muévete o serás encarcelado!", 0, 0, 0, 0, 0, 0, 4)
        
        Case Message_NPC_Failed
            Call AddtoRichTextBox(frmMain.RecTxt, "¡La criatura falló el golpe!", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_NPC_Killed
            Call AddtoRichTextBox(frmMain.RecTxt, "¡La criatura te ha matado!", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_Shield_Blocked
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has rechazado el ataque con el escudo!", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_Shield_Blocked_Our
            Call AddtoRichTextBox(frmMain.RecTxt, "¡El usuario rechazó el ataque con su escudo!", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_Less_Tired
            Call AddtoRichTextBox(frmMain.RecTxt, "Te sentis menos cansado.", 0, 0, 0, 0, 0, 0, 4)
        
        Case Message_Not_Hide
            Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado esconderte!", 0, 0, 0, 0, 0, 0, 4)
        
        Case Message_Offline_User
            Call AddtoRichTextBox(frmMain.RecTxt, "Usuario offline.", 0, 0, 0, 0, 0, 0, 8)
        
        Case Message_Nothing
            Call AddtoRichTextBox(frmMain.RecTxt, "No ves nada interesante.", 0, 0, 0, 0, 0, 0, 4)
        
        Case Message_Select_Char
            Call AddtoRichTextBox(frmMain.RecTxt, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", 0, 0, 0, 0, 0, 0, 4)
            
        Case Message_Not_Amount
            Call AddtoRichTextBox(frmMain.RecTxt, "No tenés esa cantidad.", 0, 0, 0, 0, 0, 0, 4)
            
        Case Message_NPC_Die
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has matado a la criatura!", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_Dead
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Estás muerto! Ve al sacerdote más cercano para que puedas ser revivido.", 0, 0, 0, 0, 0, 0, 4)
        
        Case Message_Not_Special_Move
            Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado realizar un movimiento especial!", 0, 0, 0, 0, 0, 0, 16)
        
        Case Message_Not_Stab
            Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado apuñalar a tu enemigo!", 0, 0, 0, 0, 0, 0, 16)
        
        Case Message_User_Not_Exist
            Call AddtoRichTextBox(frmMain.RecTxt, "El personaje no existe.", 0, 0, 0, 0, 0, 0, 8)
        
        Case Message_Reviving
            Call AddtoRichTextBox(frmMain.RecTxt, "Tu cuerpo comienza a tomar forma...", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_Corpse_Formed
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Has vuelto a la vida!", 0, 0, 0, 0, 0, 0, 2)
        
        Case Message_Invalid_Rules
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Tu consulta ha sido ignorada debido a que invalida el reglamento del juego! Te recomendamos leerlo en el sitio oficial: www.imperiumao.com.ar", 0, 0, 0, 0, 0, 0, 8)
            
        Case Message_Invalid_Manual
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Tu consulta ha sido ignorada debido a que está respondida en el manual básico o bien las preguntas frecuentes (FAQ)! Te recomendamos leer los textos en el sitio oficial: www.imperiumao.com.ar", 0, 0, 0, 0, 0, 0, 8)
        
        Case Message_GameMasters_Sent
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Gracias por tu mensaje! Será respondido a la brevedad.", 0, 0, 0, 0, 0, 0, 8)
        
        Case Message_GameMasters_Already_Sent
            Call AddtoRichTextBox(frmMain.RecTxt, "Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola.", 0, 0, 0, 0, 0, 0, 8)
        
        Case Message_Not_Gold
            Call AddtoRichTextBox(frmMain.RecTxt, "Oro insuficiente.", 0, 0, 0, 0, 0, 0, 4)
        
        Case Message_Failed
            Call Engine.Char_Dialog_Set(CurrentUser.CurrentChar, "*Fallas*", COLOR_ATAQUE, 5, 2)
            
        Case Message_Safe_Exit
            Call AddtoRichTextBox(frmMain.RecTxt, "Estás saliendo de una zona segura. Recuerda que aquí corres riesgo de ser atacado por otros.", 0, 0, 0, 0, 0, 0, 3)
        
        Case Ping_Received
            CurrentUser.Ping = GetTickCount - CurrentUser.Ping
            Call AddtoRichTextBox(frmMain.RecTxt, "Ping: " & CurrentUser.Ping & "ms", 0, 0, 0, 0, 0, 0, 4)
            Call Send_Data(Ping_Request)
        
        Case Map_Load_Request
            frmIniciando.Visible = True
            frmMain.Visible = False
            CurrentUser.CurrentChar = 0
            Call Map_Load(String_To_Integer(rData, 1), String_To_Byte(rData, 3), String_To_Long(rData, 4), String_To_Byte(rData, 4))
            Call Send_Data(Loading_Finished)

        Case Main_Map_Data
            t = Split(rData, "¬")
            
            For X = 0 To UBound(t)
                Call Handle_Data(t(X))
            Next X

        Case Closing
            Call AddtoRichTextBox(frmMain.RecTxt, "En " & String_To_Byte(rData, 1) & " segundos se cerrará el juego...", 65, 190, 156, False, True, False)
            CurrentUser.Saliendo = True
        
        Case Close_Cancel
            Call AddtoRichTextBox(frmMain.RecTxt, "¡El cierre del juego ha sido cancelado!", 0, 0, 0, False, False, False, 14)
            CurrentUser.Saliendo = False
        
        Case Exp_Won
            tempLng = String_To_Long(rData, 1)
            
            If GuardarEXP = 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt, "¡Has ganado " & tempLng & " puntos de experiencia!", 51, 183, 247, True, False, False)
            Else
                CurrentUser.ExpCount = CurrentUser.ExpCount + tempLng
            End If
            
        Case Group_Exp_Won
            Call AddtoRichTextBox(frmMain.RecTxt, "¡El grupo ha ganado " & String_To_Long(rData, 1) & " puntos de experiencia!", 51, 183, 247, True, False, False)

        Case Dialog_Hit_Spell
            tempLng = String_To_Long(rData, 1)
            
            If tempLng > 200 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & tempLng & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, CStr(tempLng), COLOR_ATAQUE, 5, 2
            End If
        
        Case Dialog_Hit_Weap
            tempLng = String_To_Long(rData, 1)
            
            If TempInt > 200 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & tempLng & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, CStr(tempLng), COLOR_ATAQUE, 5, 2
            End If
        
        Case Position_Current
            X = String_To_Integer(rData, 1)
            Y = String_To_Integer(rData, 3)
            Call Engine.Char_Current_Pos_Refresh(X, Y)
            Call Engine.Engine_View_Pos_Set(X, Y)
        
        Case Position_Other
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            X = String_To_Integer(rData, 3)
            Y = String_To_Integer(rData, 5)
            Call Engine.Char_Pos_Set(CharIndex, X, Y)
        
        Case Fly_Toggle
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Fly_Set(CharIndex, String_To_Byte(rData, 3))
        
        Case NPC_Hit
            tempByte = String_To_Byte(rData, 1)
            TempInt = String_To_Integer(rData, 2)
            
            Select Case tempByte
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado en la cabeza por " & TempInt, 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado el brazo izquierdo por " & TempInt, 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado el brazo derecho por " & TempInt, 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado la pierna izquierda por " & TempInt, 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado la pierna derecha por " & TempInt, 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "La criatura te ha pegado en el torso por " & TempInt, 255, 0, 0, True, False, False)
            End Select
        
        Case User_NPC_Hit
            tempLng = String_To_Long(rData, 1)
            
            If tempLng > 150 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & tempLng & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, CStr(tempLng), COLOR_ATAQUE, 5, 2
            End If
        
        Case User_Failed_User
            TempInt = String_To_Integer(rData, 1)
            tempStr = Engine.Char_Name_Get_No_Guild(TempInt)
            Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ataco y fallo!", 255, 0, 0, True, False, False)
        
        Case User_Hit_Current_User
            tempByte = String_To_Byte(rData, 1)
            tempLng = String_To_Long(rData, 2)
            TempInt = String_To_Integer(rData, 6)
            tempStr = Engine.Char_Name_Get_No_Guild(TempInt)
            
            Select Case tempByte
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ha pegado en la cabeza por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ha pegado el brazo izquierdo por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ha pegado el brazo derecho por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ha pegado la pierna izquierda por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ha pegado la pierna derecha por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡" & tempStr & " te ha pegado en el torso por " & tempLng & "!", 255, 0, 0, True, False, False)
            End Select
        
            If tempLng > 150 Then
                Engine.Char_Dialog_Set TempInt, "¡" & tempLng & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set TempInt, tempLng, COLOR_ATAQUE, 5, 2
            End If
        
        Case User_Hit_User
            tempByte = String_To_Byte(rData, 1)
            tempLng = String_To_Long(rData, 2)
            tempStr = Engine.Char_Name_Get_No_Guild(String_To_Integer(rData, 6))
            
            Select Case tempByte
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & tempStr & " en la cabeza por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & tempStr & " en el brazo izquierdo por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & tempStr & " en el brazo derecho por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & tempStr & " en la pierna izquierda por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & tempStr & " en la pierna derecha por " & tempLng & "!", 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, "¡Le has pegado a " & tempStr & " en el torso por " & tempLng & "!", 255, 0, 0, True, False, False)
            End Select
            
            If tempLng > 150 Then
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, "¡" & tempLng & "!", COLOR_ATAQUE, 5, 2
            Else
                Engine.Char_Dialog_Set CurrentUser.CurrentChar, tempLng, COLOR_ATAQUE, 5, 2
            End If

        Case Main_Dialog
            CharIndex = Val((General_Field_Read(3, rData, "°")))
            
            If CharIndex <> 0 Then
                CharIndex = Engine.Char_Find(Val(General_Field_Read(3, rData, "°")))
                tempStr = Engine.Char_Name_Get(CharIndex)
                
                If InStr(tempStr, "<") Then
                    cad$ = General_Field_Read(1, tempStr, "<")
                    cad$ = left$(cad$, Len(cad$) - 1)
                Else
                    cad$ = tempStr
                End If
                
                If Not NickIgnorado(cad$) Then
                    Engine.Char_Dialog_Set CharIndex, General_Field_Read(2, rData, "°"), Val(General_Field_Read(1, rData, "°")), 10
                    
                    If CopiarDialogos = 1 And tempStr <> "" Then
                        Call CopiarDialogoAConsola(tempStr, General_Field_Read(2, rData, "°"), Val(General_Field_Read(1, rData, "°")))
                    End If
                End If
                
            Else
                tempStr = General_Field_Read(1, rData, ">")
                
                If Not NickIgnorado(tempStr) Then
                    If InStr(rData, "~") = 0 Then
                        tempByte = String_To_Byte(rData, Len(rData))
                        AddtoRichTextBox frmMain.RecTxt, left$(rData, Len(rData) - 1), 0, 0, 0, 0, 0, False, tempByte
                    Else
                        AddtoRichTextBox frmMain.RecTxt, General_Field_Read(1, rData, "~"), Val(General_Field_Read(2, rData, "~")), Val(General_Field_Read(3, rData, "~")), Val(General_Field_Read(4, rData, "~")), Val(General_Field_Read(5, rData, "~")), Val(General_Field_Read(6, rData, "~"))
                    End If
                End If
                
            End If
        
        Case Magic_Words
            CharIndex = Val((General_Field_Read(3, rData, "°")))
            CharIndex = Engine.Char_Find(Val(General_Field_Read(3, rData, "°")))
            
            If CharIndex > 0 Then _
                Engine.Char_Dialog_Set CharIndex, General_Field_Read(2, rData, "°"), Val(General_Field_Read(1, rData, "°")), 5

        Case Global_Chat_Message
            If MensajesGlobales = 1 Then
                tempStr = General_Field_Read(1, rData, ">")
                If Not NickIgnorado(tempStr) Then _
                    Call AddtoRichTextBox(frmMain.RecTxt, rData, 0, 0, 0, 0, 0, 0, 23)
            End If
        
        Case Stats_HP
            Call ActualizarEst(, String_To_Integer(rData, 1))
        
        Case Stats_STA
            Call ActualizarEst(, , , , , String_To_Integer(rData, 1))
        
        Case Stats_MAN
            Call ActualizarEst(, , , String_To_Integer(rData, 1))
        
        Case Stats_GLD
            Call ActualizarEst(, , , , , , String_To_Long(rData, 1))
        
        Case Stats_EXP
            Call ActualizarEst(, , , , , , , String_To_Integer(rData, 1), String_To_Long(rData, 3), String_To_Long(rData, 7))
        
        Case Stats_ATT
            Call ActualizarEst(, , , , , , , , , , String_To_Integer(rData, 1), String_To_Integer(rData, 3))
        
        Case Server_Message
            frmMensaje.msg.Caption = rData
            frmMensaje.Show vbModeless, frmMain
            Call Sound.Sound_Play(118)

        Case Create_Char
            tempByte = String_To_Byte(rData, 21)
            tempStr = mid$(rData, 22 + (tempByte * 4))
            
            If tempByte > 0 Then
                ReDim part_type(1 To tempByte) As Integer
                ReDim part_life(1 To tempByte) As Long
            End If
            
            'Primer byte = 21
            'Último byte = 22 + (tempByte * 4)
            k = -2
            For i = 1 To tempByte
                k = k + 2
                part_type(i) = String_To_Integer(rData, 22 + k)
                part_life(i) = CLng(String_To_Integer(rData, 24 + k))
            Next i
        
            Call Engine.Char_Create(String_To_Integer(rData, 8), String_To_Integer(rData, 10), String_To_Byte(rData, 5), String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 16), String_To_Integer(rData, 12), String_To_Integer(rData, 14), tempStr, _
                String_To_Byte(rData, 18), 0, 0, String_To_Integer(rData, 6), _
                String_To_Byte(rData, 19), String_To_Byte(rData, 20), tempByte, part_life(), part_type())
                    
        Case Create_Char_Current
            tempByte = String_To_Byte(rData, 21)
            tempStr = mid$(rData, 22 + (tempByte * 4))
            
            If tempByte > 0 Then
                ReDim part_type(1 To tempByte) As Integer
                ReDim part_life(1 To tempByte) As Long
            End If
            
            'Primer byte = 21
            'Último byte = 22 + (tempByte * 4)
            k = -2
            For i = 1 To tempByte
                k = k + 2
                part_type(i) = String_To_Integer(rData, 22 + k)
                part_life(i) = CLng(String_To_Integer(rData, 24 + k))
            Next i
        
            CharIndex = Engine.Char_Create(String_To_Integer(rData, 8), String_To_Integer(rData, 10), String_To_Byte(rData, 5), String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 16), String_To_Integer(rData, 12), String_To_Integer(rData, 14), tempStr, _
                String_To_Byte(rData, 18), 0, 0, String_To_Integer(rData, 6), _
                String_To_Byte(rData, 19), String_To_Byte(rData, 20), tempByte, part_life(), part_type())
            
            CurrentUser.CurrentChar = CharIndex
            Call Engine.Char_Current_Set(CharIndex)
            Call Engine.Char_Pos_Get(CharIndex, X, Y)
            Call Engine.Engine_View_Pos_Set(X, Y)
            
            If Not EngineRun Then
                Call SetConnected
            Else
                frmMain.Visible = True
                frmIniciando.Visible = False
            End If
        
        Case Create_Char_NPC
            Call Engine.Char_Create(String_To_Integer(rData, 8), String_To_Integer(rData, 10), String_To_Byte(rData, 5), String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 16), String_To_Integer(rData, 12), String_To_Integer(rData, 14), "", _
                0, 0, 0, String_To_Integer(rData, 6), _
                0, 0, 0, part_life(), part_type())
        
        Case Change_Heading_North
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Heading_Set(CharIndex, NORTH)
        
        Case Change_Heading_South
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Heading_Set(CharIndex, SOUTH)
        
        Case Change_Heading_East
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Heading_Set(CharIndex, EAST)
        
        Case Change_Heading_West
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Heading_Set(CharIndex, WEST)
        
        Case Change_Type
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Type_Set(CharIndex, String_To_Byte(rData, 3))
        
        Case Change_Name
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            tempStr = mid$(rData, 3)
            Call Engine.Char_Name_Set(CharIndex, tempStr)
            
            If CharIndex = CurrentUser.CurrentChar Then _
                frmMensaje.mnuClan.Enabled = (InStr(tempStr, "<") > 0)
                        
        Case Change_Group
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Group_Set(CharIndex, String_To_Integer(rData, 3))
                
        Case Erase_Char
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Remove(CharIndex)
        
        Case Guild_Info
        
        Case Move_Char_North
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call Engine.Char_Move(CharIndex, NORTH)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Move_Char_East
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call Engine.Char_Move(CharIndex, EAST)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Move_Char_South
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call Engine.Char_Move(CharIndex, SOUTH)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Move_Char_West
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            If CharIndex <= 0 Then Exit Sub
            Call Engine.Char_Move(CharIndex, WEST)
            If Audio = 1 Then Call DoPasosFx(CharIndex)
        
        Case Change_Char
            Call Engine.Char_Change(String_To_Integer(rData, 1), String_To_Byte(rData, 7), _
            String_To_Integer(rData, 3), String_To_Integer(rData, 5), String_To_Integer(rData, 12), _
            String_To_Integer(rData, 8), String_To_Integer(rData, 10), 0, _
            0)
        
        Case Item_Create
            X = String_To_Integer(rData, 5)
            Y = String_To_Integer(rData, 7)
            Call Engine.Map_Item_Grh_Add(X, Y, String_To_Long(rData, 1))
        
        Case Item_Erase
            X = String_To_Integer(rData, 1)
            Y = String_To_Integer(rData, 3)
            Call Engine.Map_Item_Grh_Remove(X, Y)
        
        Case Pos_Bloq
            Call Engine.Map_Blocked_Set(String_To_Integer(rData, 1), String_To_Integer(rData, 3), String_To_Byte(rData, 5))

        Case Music_Change
            TempInt = String_To_Integer(rData, 1)
            If (Musica <> CONST_DESHABILITADA) And (DefMidi = 1) Then
                If TempInt <> 0 Then
                    Sound.NextMusic = TempInt
                    Sound.Fading = 200
                End If
            End If

        Case Ambient_Change
            X = String_To_Integer(rData, 1)
            If X <> 0 Then Sound.AmbienteActual = X
        
        Case Sound_Play
            If Audio = 1 Then
                TempInt = String_To_Integer(rData, 1)
                X = String_To_Integer(rData, 3)
                Y = String_To_Integer(rData, 5)
                
                If Engine.Map_In_Bounds(X, Y) = False Then
                    Call Sound.Sound_Play(TempInt)
                Else
                    Call Sound.Sound_Play(TempInt, , Sound.Calculate_Volume(X, Y), Sound.Calculate_Pan(X, Y))
                End If
            End If
        
        Case Guild_List
            Call frmGuildAdm.ParseGuildList(rData)
            
        Case Commerce_Data
            Call frmComercioSeguro.ParseData(rData)
        
        Case Main_Pause
            CurrentUser.Pausa = Not CurrentUser.Pausa
        
        Case Client_Validation
            bRK = String_To_Long(rData, 1)
            bK = String_To_Long(rData, 5)
            Call Login(ValidarLoginMSG(CInt(bRK)))
        
        Case Rain_Toggle
            If Meteo_Engine.SecondaryStatus = 2 Then
                Meteo_Engine.SecondaryStatus = 0
            Else
                Meteo_Engine.SecondaryStatus = 2
            End If
        
        Case Snow_Toggle
            If Meteo_Engine.SecondaryStatus = 3 Then
                Meteo_Engine.SecondaryStatus = 0
            Else
                Meteo_Engine.SecondaryStatus = 3
            End If
        
        Case Hour_Night_Toggle
            tempByte = String_To_Byte(rData, 1)
            X = String_To_Integer(rData, 2)
            Call Meteo_Engine.SetNuevoEstado(tempByte)
            If X <> 0 Then Call Sound.Sound_Play(X)

        Case Dialog_Erase
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Dialog_Remove(CharIndex)

        Case FX_Create
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            TempInt = String_To_Integer(rData, 7)
            Call Engine.Char_FX_Set(CharIndex, String_To_Integer(rData, 3), String_To_Integer(rData, 5))
            
            If TempInt <> 0 Then
                If Engine.Char_Pos_Get(CharIndex, X, Y) Then
                    Call Sound.Sound_Play(TempInt, , Sound.Calculate_Volume(X, Y), Sound.Calculate_Pan(X, Y))
                End If
            End If
        
        Case Map_FX_Create
            X = String_To_Integer(rData, 3)
            Y = String_To_Integer(rData, 5)
            TempInt = String_To_Integer(rData, 7)
            Call Engine.Map_Fx_Add(X, Y, String_To_Integer(rData, 1))
            If TempInt <> 0 Then Call Sound.Sound_Play(TempInt, , Sound.Calculate_Volume(X, Y), Sound.Calculate_Pan(X, Y))
        
        Case Particle_Char_Add
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call General_Char_Particle_Create(String_To_Integer(rData, 3), CharIndex, String_To_Integer(rData, 5))
        
        Case Particle_Map_Add
            X = String_To_Integer(rData, 5)
            Y = String_To_Integer(rData, 7)
            Call General_Particle_Create(String_To_Integer(rData, 1), X, Y, String_To_Integer(rData, 3))
        
        Case Particle_Char_Remove
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Particle_Group_Remove(CharIndex, String_To_Integer(rData, 3))

        Case Particle_All_Char_Remove
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Particle_Group_Remove_All(CharIndex)

        Case GM_Panel
            'Dim n As String, n2 As String
            'n = General_Field_Read(2, rData, "°")
            'n2 = General_Field_Read(1, rData, "°")
            frmPanelGm.Show vbModeless, frmMain

        Case User_Stats
            Call ActualizarEst(String_To_Integer(rData, 1), String_To_Integer(rData, 3), _
                String_To_Integer(rData, 5), String_To_Integer(rData, 7), _
                String_To_Integer(rData, 9), String_To_Integer(rData, 11), _
                String_To_Long(rData, 13), String_To_Integer(rData, 17), _
                String_To_Long(rData, 19), String_To_Long(rData, 23), _
                String_To_Integer(rData, 27), String_To_Integer(rData, 29), True)
        
        Case Quest_Info
            If rData = "" Then
                Call AddtoRichTextBox(frmMain.RecTxt, "¡Aún no has aceptado ninguna propuesta!", 204, 193, 115, 0, 1)
            Else
                frmQuestActual.Show vbModeless, frmMain
                frmQuestActual.ParseQuestInfo (rData)
            End If
        
        Case Working_Main
            CurrentUser.UsingSkill = String_To_Byte(rData, 1)
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
        
        Case Inventory_Update
            tempByte = String_To_Byte(rData, 1)

            If tempByte <> 111 Then
                tempStr = General_Field_Read(2, rData, ",")
                tempByte = String_To_Byte(tempStr, 1)
                UserInventory(tempByte).Name = General_Field_Read(1, rData, ",")
                UserInventory(tempByte).OBJIndex = String_To_Integer(tempStr, 2)
                UserInventory(tempByte).Amount = String_To_Integer(tempStr, 4)
                UserInventory(tempByte).Equipped = String_To_Byte(tempStr, 6)
                UserInventory(tempByte).GrhIndex = String_To_Long(tempStr, 7)
                UserInventory(tempByte).ObjType = String_To_Integer(tempStr, 11)
                UserInventory(tempByte).MaxHIT = String_To_Integer(tempStr, 13)
                UserInventory(tempByte).MinHIT = String_To_Integer(tempStr, 15)
                UserInventory(tempByte).Def = String_To_Integer(tempStr, 17)
                UserInventory(tempByte).Valor = String_To_Long(tempStr, 19)
            Else
                tempByte = String_To_Byte(rData, 2)
                UserInventory(tempByte).Name = "Nada"
                UserInventory(tempByte).OBJIndex = 0
                UserInventory(tempByte).Amount = 0
                UserInventory(tempByte).Equipped = 0
                UserInventory(tempByte).GrhIndex = 0
                UserInventory(tempByte).ObjType = 0
                UserInventory(tempByte).MaxHIT = 0
                UserInventory(tempByte).MinHIT = 0
                UserInventory(tempByte).Def = 0
                UserInventory(tempByte).Valor = 0
            End If
            
            If CurrentUser.Logged Then
                Engine.Engine_Inventory_Render_Set
                
                If frmComerciar.Visible Then
                    frmComerciar.List1(1).List(tempByte - 1) = UserInventory(tempByte).Name
                    Call frmComerciar.Refresh_User_Inventory
                ElseIf frmBancoObj.Visible Then
                    frmBancoObj.List1(1).List(tempByte - 1) = UserInventory(tempByte).Name
                    Call frmBancoObj.Refresh_User_Inventory
                End If
            End If
            
        Case Bank_Update
            tempByte = String_To_Byte(rData, 1)
            
            If tempByte <> 111 Then
                UserBancoInventory(tempByte).Name = mid$(rData, 20)
                UserBancoInventory(tempByte).OBJIndex = String_To_Integer(rData, 2)
                UserBancoInventory(tempByte).Amount = String_To_Long(rData, 4)
                UserBancoInventory(tempByte).GrhIndex = String_To_Long(rData, 8)
                UserBancoInventory(tempByte).ObjType = String_To_Integer(rData, 12)
                UserBancoInventory(tempByte).MaxHIT = String_To_Integer(rData, 14)
                UserBancoInventory(tempByte).MinHIT = String_To_Integer(rData, 16)
                UserBancoInventory(tempByte).Def = String_To_Integer(rData, 18)
            Else
                tempByte = String_To_Byte(tempStr, 2)
                UserBancoInventory(tempByte).Name = "Nada"
                UserBancoInventory(tempByte).OBJIndex = 0
                UserBancoInventory(tempByte).Amount = 0
                UserBancoInventory(tempByte).GrhIndex = 0
                UserBancoInventory(tempByte).ObjType = 0
                UserBancoInventory(tempByte).MaxHIT = 0
                UserBancoInventory(tempByte).MinHIT = 0
                UserBancoInventory(tempByte).Def = 0
            End If
        
            If frmBancoObj.Visible Then
                frmBancoObj.List1(0).List(tempByte - 1) = UserBancoInventory(tempByte).Name
                Call frmBancoObj.Refresh_Bank_Inventory
            End If
        
        Case Spell_Add
            tempByte = String_To_Byte(rData, 1)
            CurrentUser.UserHechizos(tempByte) = String_To_Integer(rData, 2)
            
            If tempByte > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem IIf(CurrentUser.UserHechizos(tempByte) = 0, "Nada", Right$(rData, Len(rData) - 3))
            Else
                frmMain.hlst.List(TempInt - 1) = IIf(CurrentUser.UserHechizos(tempByte) = 0, "Nada", Right$(rData, Len(rData) - 3))
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
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ArmasHerrero(1 To m) As Integer
            
            k = 0
            For i = 1 To m
                k = k + 1
                If i <> 1 Then k = k + 1
                cad$ = General_Field_Read(k, rData, ",")
                ArmasHerrero(i) = Val(General_Field_Read(k + 1, rData, ","))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
            Next i
        
        Case Armor_List
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ArmadurasHerrero(1 To m) As Integer
                        
            k = 0
            For i = 1 To m
                k = k + 1
                If i <> 1 Then k = k + 1
                cad$ = General_Field_Read(k, rData, ",")
                ArmadurasHerrero(i) = Val(General_Field_Read(k + 1, rData, ","))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
            Next i
        
        Case Potion_List
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ObjDruida(1 To m) As Integer
                        
            k = 0
            For i = 1 To m
                k = k + 1
                If i <> 1 Then k = k + 1
                cad$ = General_Field_Read(k, rData, ",")
                ObjDruida(i) = Val(General_Field_Read(k + 1, rData, ","))
                If cad$ <> "" Then frmDruida.lstPociones.AddItem cad$
            Next i
        
        Case Tailor_List
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ObjSastre(1 To m) As Integer
                 
            k = 0
            For i = 1 To m
                k = k + 1
                If i <> 1 Then k = k + 1
                cad$ = General_Field_Read(k, rData, ",")
                ObjSastre(i) = Val(General_Field_Read(k + 1, rData, ","))
                If cad$ <> "" Then frmSastre.lstRopas.AddItem cad$
            Next i
        
        Case Carpenter_List
            m = (General_Field_Count(rData, 44) / 2)
            ReDim ObjCarpintero(1 To m) As Integer
            
            k = 0
            For i = 1 To m
                k = k + 1
                If i <> 1 Then k = k + 1
                cad$ = General_Field_Read(k, rData, ",")
                ObjCarpintero(i) = Val(General_Field_Read(k + 1, rData, ","))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
            Next i
        
        Case Rest_Toggle
            CurrentUser.Descansando = Not CurrentUser.Descansando
        
        Case GM_Spawn_List
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmSpawnList.lstCriaturas.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmSpawnList.Show vbModeless, frmMain
        
        Case Error_Message
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
        
        Case Challenge_Data
            If rData = "JOYA" Then
                Unload frmTorneoCrear
                Exit Sub
            End If
            MensajeAdvertencia rData
        
        Case Stupid_True
            CurrentUser.Estupido = True
        
        Case Letter_Show
            Call Engine.Letter_Set(CInt(General_Field_Read(2, rData, "°")), General_Field_Read(1, rData, "°"))
        
        Case NPC_Inv_Receive
            tempByte = String_To_Byte(rData, 1)
            
            If tempByte <> 111 Then
                NPCInventory(tempByte).Amount = String_To_Integer(rData, 2)
                NPCInventory(tempByte).Valor = String_To_Long(rData, 4)
                NPCInventory(tempByte).GrhIndex = String_To_Long(rData, 8)
                NPCInventory(tempByte).OBJIndex = String_To_Integer(rData, 12)
                NPCInventory(tempByte).ObjType = String_To_Integer(rData, 14)
                NPCInventory(tempByte).MaxHIT = String_To_Integer(rData, 16)
                NPCInventory(tempByte).MinHIT = String_To_Integer(rData, 18)
                NPCInventory(tempByte).Def = String_To_Integer(rData, 20)
                NPCInventory(tempByte).Name = mid$(rData, 22)
            Else
                tempByte = String_To_Byte(rData, 2)
                NPCInventory(tempByte).Name = "Nada"
                NPCInventory(tempByte).Amount = 0
                NPCInventory(tempByte).Valor = 0
                NPCInventory(tempByte).GrhIndex = 0
                NPCInventory(tempByte).OBJIndex = 0
                NPCInventory(tempByte).ObjType = 0
                NPCInventory(tempByte).MaxHIT = 0
                NPCInventory(tempByte).MinHIT = 0
                NPCInventory(tempByte).Def = 0
            End If
        
            If frmComerciar.Visible Then
                frmComerciar.List1(0).List(tempByte - 1) = NPCInventory(tempByte).Name
                Call frmComerciar.Refresh_NPC_Inventory
            End If
        
        Case HYS_Update
            CurrentUser.UserMinAGU = String_To_Integer(rData, 1)
            CurrentUser.UserMinHAM = String_To_Integer(rData, 3)
            CurrentUser.UserMaxAGU = 100
            CurrentUser.UserMaxHAM = 100
            frmMain.AGUAsp.Width = (((CurrentUser.UserMinAGU / 100) / (CurrentUser.UserMaxAGU / 100)) * 91)
            frmMain.COMIDAsp.Width = (((CurrentUser.UserMinHAM / 100) / (CurrentUser.UserMaxHAM / 100)) * 91)
            frmMain.lblHAM.Caption = CurrentUser.UserMinHAM & "/" & CurrentUser.UserMaxHAM
            frmMain.lblSED.Caption = CurrentUser.UserMinAGU & "/" & CurrentUser.UserMaxAGU
        
        Case Fame_Update
            CurrentUser.UserReputacion.AsesinoRep = String_To_Long(rData, 1)
            CurrentUser.UserReputacion.BandidoRep = String_To_Long(rData, 5)
            CurrentUser.UserReputacion.BurguesRep = String_To_Long(rData, 9)
            CurrentUser.UserReputacion.LadronesRep = String_To_Long(rData, 13)
            CurrentUser.UserReputacion.NobleRep = String_To_Long(rData, 17)
            CurrentUser.UserReputacion.PlebeRep = String_To_Long(rData, 21)
            CurrentUser.UserReputacion.Promedio = String_To_Long(rData, 25)
            CurrentUser.UserReputacion.Culpabilidad = String_To_Long(rData, 29)
            LlegoFama = True
            Call MostrarEstadisticas

        Case Skill_Data
            CurrentUser.SkillPoints = String_To_Integer(rData, 1)
        
        Case Number_Enemys
            AddtoRichTextBox frmMain.RecTxt, "Hay " & String_To_Integer(rData, 1) & " NPCs.", 255, 255, 255, 0, 0
        
        Case GM_Sos_Message
            TempInt = InStr(1, rData, "µ")
            k = Val(General_Field_Read(4, rData, "µ"))
                        
            If TempInt > 0 Then
                If k = 1 Then 'Barrin: Está online!
                    tempStr = left(rData, TempInt - 1) & " (" & General_Field_Read(3, rData, "µ") & ")"
                    frmPanelGm.List1.AddItem tempStr
                    frmPanelGm.MensajePoner tempStr, General_Field_Read(2, rData, "µ")
                Else
                    tempStr = left(rData, TempInt - 1) & " (" & General_Field_Read(3, rData, "µ") & ")"
                    frmPanelGm.List2.AddItem tempStr
                    frmPanelGm.MensajePoner tempStr, General_Field_Read(2, rData, "µ")
                End If
            Else
                frmPanelGm.List1.AddItem rData
            End If
        
        Case GM_Sos_Show
            frmPanelGm.Show vbModeless, frmMain
        
        Case Forum_Main
            frmForo.List.AddItem General_Field_Read(1, rData, "°")
            frmForo.Text(frmForo.List.ListCount - 1).Text = General_Field_Read(2, rData, "°")
            Load frmForo.Text(frmForo.List.ListCount)
        
        Case Forum_Show
            If Not frmForo.Visible Then
                frmForo.Show vbModeless, frmMain
            End If
        
        Case GameMaster_Message_Show
            If Val(rData) = 1 Then
                frmGMAyuda.Label2.Caption = "Deje los siguientes datos: nombre del clan (sensitivo a mayúsculas y minúsculas), alineamiento y otros datos que considere de importancia."
                frmGMAyuda.optConsulta.Item(6).Value = True
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
            Call Engine.Char_Current_Blind_Set(True)
        
        Case Stats_Mini
            CurrentUser.UserStats.CiudasMatados = String_To_Long(rData, 1)
            CurrentUser.UserStats.CrimisMatados = String_To_Long(rData, 5)
            CurrentUser.UserStats.NPCsMatados = String_To_Long(rData, 9)
            CurrentUser.UserStats.Clase = String_To_Byte(rData, 13)
            CurrentUser.UserStats.TimesKilled = String_To_Long(rData, 14)
            CurrentUser.UserStats.Raza = String_To_Byte(rData, 18)
            CurrentUser.UserStats.Genero = String_To_Byte(rData, 19)
            LlegoEst = True
            Call MostrarEstadisticas

        Case Challenge_Can
            PuedeTorneo = String_To_Byte(rData, 1)
        
        Case Challenge_Info
            Dim TInscriptos As Integer
            Dim loopc As Integer
            If rData = "LI" Then
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
        
        Case Meditating_Toggle
            CurrentUser.Meditando = Not CurrentUser.Meditando
        
        Case char_invisible
            CharIndex = Engine.Char_Find(String_To_Integer(rData, 1))
            Call Engine.Char_Invisible_Set(CharIndex, String_To_Byte(rData, 3))
        
        Case GM_Motd_Info
            frmCambiaMotd.Show vbModeless, frmMain
            frmCambiaMotd.txtMotd.Text = rData
        
        Case Familiar_Adopt_Show
            frmSeleccionFamiliar.Show vbModeless, frmMain
        
        Case Quest_Data
            frmQuest.Show vbModeless, frmMain
            frmQuest.ParseQuestInfo (rData)
        
        Case Blind_False
            Call Engine.Char_Current_Blind_Set(False)
        
        Case Stupid_False
            CurrentUser.Estupido = False
        
        Case Skills_Rec
            k = -1
            
            For i = 1 To NUMSKILLS
                k = k + 2
                CurrentUser.UserSkills(i) = String_To_Integer(rData, k)
            Next i
            
            LlegaronSkills = True
            Call MostrarEstadisticas
        
        Case Train_NPC_List
            For i = 1 To Val(General_Field_Read(1, rData, ","))
                frmEntrenador.lstCriaturas.AddItem General_Field_Read(i + 1, rData, ",")
            Next i
            frmEntrenador.Show vbModeless, frmMain

        Case GM_User_List
            t = Split(rData, ",")
            frmPanelGm.cboListaUsus.Clear
            For i = LBound(t) To UBound(t)
                frmPanelGm.cboListaUsus.AddItem t(i)
            Next i
            If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
        
        Case Freezed_Toggle
            CurrentUser.Paralizado = Not CurrentUser.Paralizado
        
            If CurrentUser.Paralizado Then
                X = String_To_Integer(rData, 1)
                Y = String_To_Integer(rData, 3)
                Call Engine.Char_Current_Pos_Refresh(X, Y)
                Call Engine.Engine_View_Pos_Set(X, Y)
            End If
        
        Case Guild_News
            Call frmGuildNews.ParseGuildNews(rData)
        
        Case Guild_Petition 'Para dos
            Call frmUserRequest.recievePeticion(rData)
        
        Case Guild_Peace_Prop
            Call frmPeaceProp.ParsePeaceOffers(rData)
        
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
            frmGuildFoundation.txtClanName = rData
            Call frmGuildFoundation.Show(vbModeless, frmMain)
        
        Case Metam_Toggle
            CurrentUser.Transformado = Not CurrentUser.Transformado
            Call Engine.Char_Current_OnHorse_Set(CurrentUser.Transformado)
        
        'Quedaron fuera de uso con la remodelación del protocolo...
        Case Commerce_OK
        Case Bank_OK
        
    End Select

End Sub

Public Sub Send_Data(ByVal sndPaquet As Client_Paquet_ID, Optional ByVal sndData As String = "")

Dim PackSize As Integer

If frmMain.mainWinsock.State = sckConnected Then
    If bK <> 0 Then
        sndData = Encrypt_Data(Chr$(sndPaquet) & sndData)
    Else
        sndData = Encrypt_Data(Chr$(sndPaquet) & sndData)
    End If
    
    sndData = " " & sndData
    PackSize = Len(sndData)
    CopyMemory ByVal StrPtr(sndData), ByVal VarPtr(PackSize), 2

    Call frmMain.mainWinsock.SendData(sndData)
End If

End Sub

Public Sub Parse_Command_Str(ByVal strCmd As String)

On Error Resume Next

Dim strArg As String
Dim strCompa As String
Dim bFound As Boolean

bFound = True
strCompa = UCase$(strCmd)

Select Case strCompa
    Case "/PING"
        CurrentUser.Ping = GetTickCount
        Call Send_Data_Command(cmdPing)
    Case "/SALIR"
        Call Send_Data_Command(cmdSalir)
    Case "/MEDITAR"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdMeditar)
    Case "/CURAR"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdCurar)
    Case "/ONLINE"
        Call Send_Data_Command(cmdOnline)
    Case "/ONLINECLAN"
        Call Send_Data_Command(cmdOnlineClan)
    Case "/GM"
        Call Send_Data_Command(cmdGM)
    Case "/EST"
        Call Send_Data_Command(cmdEst)
    Case "/SUBASTA"
        Call Send_Data_Command(cmdSubasta)
    Case "/FAMI"
        Call Send_Data_Command(cmdFami)
    Case "/AYUDA"
        Call Send_Data_Command(cmdAyuda)
    Case "/BOVEDA"
        Call Send_Data_Command(cmdBoveda)
    Case "/COMERCIAR"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdComerciar)
    Case "/CUANTO"
        Call Send_Data_Command(cmdCuanto)
    Case "/RECHAZARMATRIMONIO"
        Call Send_Data_Command(cmdRechazarMatrimonio)
    Case "/ACOMPAÑAR"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdAcompañar)
    Case "/INFORMACION"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdInformacion)
    Case "/TORNEOS"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdTorneos)
    Case "/BALANCE"
        If DeadCheck Then Exit Sub
        Call Send_Data_Command(cmdBalance)
    Case "/FUNDARCLAN"
        Call Send_Data_Command(cmdFundarClan)
    Case "/SALIRCLAN"
        Call Send_Data_Command(cmdSalirClan)
    Case "/DIVORCIO"
        Call Send_Data_Command(cmdDivorcio)
    Case "/PENA"
        Call Send_Data_Command(cmdPena)
    Case Else
        bFound = False
End Select

If bFound Then Exit Sub

bFound = True
strCompa = UCase$(General_Field_Read(1, strCmd, " "))

Select Case strCompa
    Case "/PASSWD"
        If Len(strArg) < 6 Then
            Call AddtoRichTextBox(frmMain.RecTxt, "El password debe tener al menos 6 caracteres.", 0, 0, 0, 0, 0, 0, 1)
            Exit Sub
        End If
        
        strArg = MD5String(Right$(strCmd, Len(strCmd) - 8))
        Call Send_Data_Command(cmdPasswd, strArg)
    Case "/GMSG"
        CurrentUser.SendingType = 5
        Call Send_Data_Command_GM(cmdGMSG, Right$(strCmd, Len(strCmd) - 6))
    Case "/CMSG"
        If DeadCheck Then Exit Sub
        CurrentUser.SendingType = 4
        Call Send_Data_Command(cmdCMSG, Right$(strCmd, Len(strCmd) - 6))
    Case "/GRMG"
        If DeadCheck Then Exit Sub
        CurrentUser.SendingType = 6
        Call Send_Data_Command(cmdGRMG, Right$(strCmd, Len(strCmd) - 6))
    Case "/IGNORAR"
        strArg = Right$(strCmd, Len(strCmd) - 9)
        If strArg <> "" Then Call frmOpciones.AgregarIgnorado(strArg)
    Case "/RETIRAR"
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - 9)
        Call Send_Data_Command(cmdRetirar, strArg)
    Case "/OFERTAR"
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - 9)
        Call Send_Data_Command(cmdOfertar, strArg)
    Case "/OFERTAINICIAL"
        strArg = Right$(strCmd, Len(strCmd) - 15)
        Call Send_Data_Command(cmdOfertaInicial, strArg)
    Case "/DEPOSITAR"
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - 11)
        Call Send_Data_Command(cmdDepositar, strArg)
    Case "/APOSTAR"
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - 9)
        Call Send_Data_Command(cmdApostar, strArg)
    Case "/BUG"
        strArg = Right$(strCmd, Len(strCmd) - 5)
        Call Send_Data_Command(cmdBug, strArg)
    Case "/PROPONER"
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - 10)
        Call Send_Data_Command(cmdProponer, strArg)
    Case "/CENTINELA"
        strArg = Right$(strCmd, Len(strCmd) - 11)
        Call Send_Data_Command(cmdCentinela, strArg)
    Case "/DESC"
        If DeadCheck Then Exit Sub
        strArg = Right$(strCmd, Len(strCmd) - 6)
        Call Send_Data_Command(cmdDesc, strArg)
    Case "/ERCMD666"
        strArg = Right$(strCmd, Len(strCmd) - 10)
        Call Send_Data_Command(cmdErcmd666, strArg)
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
    Case "/CAOS"
        Call Send_Data_Command_GM(cmdCaos)
    Case "/REAL"
        Call Send_Data_Command_GM(cmdReal)
    Case "/ONLINEGM"
        Call Send_Data_Command_GM(cmdOnlineGM)
    Case "/CS"
        Call Send_Data_Command_GM(cmdCS)
    Case "/MOTDCAMBIA"
        Call Send_Data_Command_GM(cmdMotdCambia)
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
    Case "/CS"
        Call Send_Data_Command_GM(cmdCS, strArg)
    Case "/BANINFO"
        Call Send_Data_Command_GM(cmdBanInfo, strArg)
    Case "/REVIVIR"
        Call Send_Data_Command_GM(cmdRevivir, strArg)
    Case "/LASTIP"
        Call Send_Data_Command_GM(cmdLastIP, strArg)
    Case "/MOTDCAMBIA"
        Call Send_Data_Command_GM(cmdMotdCambia, strArg)
    Case "/INFO"
        Call Send_Data_Command_GM(cmdInfo, strArg)
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
    Case "/AUTORIZAR"
        Call Send_Data_Command_GM(cmdAutorizar, strArg)
    Case "/UNBAN"
        Call Send_Data_Command_GM(cmdUnBan, strArg)
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
        Call AddtoRichTextBox(frmMain.RecTxt, "¡Comando no reconocido!", 0, 0, 0, 0, 0, 0, 4)
End Select

End Sub

Public Sub Send_Data_Command(ByVal sndCmd As Client_User_Command_ID, Optional ByVal sndData As String = "")

Dim PackSize As Integer

If frmMain.mainWinsock.State = sckConnected Then
    sndData = Encrypt_Data(Chr$(Main_Command) & Chr$(sndCmd) & sndData)
    sndData = " " & sndData
    PackSize = Len(sndData)
    CopyMemory ByVal StrPtr(sndData), ByVal VarPtr(PackSize), 2
    Call frmMain.mainWinsock.SendData(sndData)
End If

End Sub

Public Sub Send_Data_Command_GM(ByVal sndCmd As Client_GM_Command_ID, Optional ByVal sndData As String = "")

Dim PackSize As Integer

If frmMain.mainWinsock.State = sckConnected Then
    sndData = Encrypt_Data(Chr$(Main_GM_Command) & Chr$(sndCmd) & sndData)
    sndData = " " & sndData
    PackSize = Len(sndData)
    CopyMemory ByVal StrPtr(sndData), ByVal VarPtr(PackSize), 2
    Call frmMain.mainWinsock.SendData(sndData)
End If

End Sub

Public Sub Login(ByVal valcode As Integer)

'Personaje grabado
If EstadoLogin = NORMAL Then
    Call Send_Data(Old_Login, CurrentUser.UserName & "," & CurrentUser.UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
'Crear personaje
ElseIf EstadoLogin = CrearNuevoPj Then
    If CurrentUser.UserClase = MAGO Or CurrentUser.UserClase = DRUIDA Or CurrentUser.UserClase = CAZADOR Then
        Call Send_Data(New_Login, CurrentUser.UserName & "," & CurrentUser.UserPassword _
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
        Call Send_Data(New_Login, CurrentUser.UserName & "," & CurrentUser.UserPassword _
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

Private Sub CopiarDialogoAConsola(ByVal NickName As String, ByVal Dialogo As String, ByVal color As Long)

If NickName = "" Then Exit Sub
If Right$(Dialogo, 1) = " " Then Exit Sub

If InStr(NickName, "<") Then
    NickName = left$(NickName, InStr(NickName, "<") - 2)
End If

Select Case color
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
    Case Else
        Call AddtoRichTextBox(frmMain.RecTxt, NickName & "> " & Dialogo, 255, 255, 255, False, True, False)
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

Public Function Decrypt_Data(ByVal strText As String, ByVal strPwd As String) As String
    
    Decrypt_Data = ""

End Function

Private Function Encrypt_Data(ByVal strText As String, ByVal strPwd As String) As String
    
    Encrypt_Data = ""

End Function

Private Function ActualizarEst(Optional ByVal MaxHP As Integer = -1, Optional ByVal MinHP As Integer = -1, Optional ByVal MaxMAN As Integer = -1, _
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

If CurrentUser.UserMinHP <= 0 Then
    CurrentUser.UserMinHP = 0
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

Public Function DeadCheck() As Boolean

If CurrentUser.Muerto Then
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Estás muerto! Ve al sacerdote más cercano para que puedas ser revivido.", 0, 0, 0, 0, 0, 0, 4)
    DeadCheck = True
End If

End Function

Public Function CombateCheck() As Boolean

If CurrentUser.Combate = False Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Debes estar en modo combate para realizar esa acción", 0, 0, 0, 0, 0, 0, 4)
    CombateCheck = True
End If

End Function

Public Function StealCheck() As Boolean

If CurrentUser.Seguro = True Then
    Call AddtoRichTextBox(frmMain.RecTxt, "No podes robar, para hacerlo debes desactivar el seguro.", 0, 0, 0, 0, 0, 0, 4)
    StealCheck = True
End If

End Function
