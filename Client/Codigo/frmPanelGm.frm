VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4380
      Width           =   975
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "Seleccionar personaje"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4380
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   630
      Width           =   4560
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Height          =   1035
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   110
      X2              =   4680
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Menu mnuUsuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuIra 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnuTraer 
         Caption         =   "Traer el usuario"
      End
      Begin VB.Menu mnuInvalida 
         Caption         =   "Inv�lida"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual/FAQ"
      End
   End
   Begin VB.Menu mnuChar 
      Caption         =   "Personaje"
      Begin VB.Menu cmdAccion 
         Caption         =   "Echar"
         Index           =   0
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Sumonear"
         Index           =   2
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir a"
         Index           =   3
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ubicaci�n"
         Index           =   6
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Desbanear"
         Index           =   12
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "IP del personaje"
         Index           =   13
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Revivir"
         Index           =   21
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Modo rol"
         Index           =   22
      End
      Begin VB.Menu cmdBanMenu 
         Caption         =   "Banear"
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje"
            Index           =   1
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje e IP"
            Index           =   19
         End
      End
      Begin VB.Menu mnuEncarcelar 
         Caption         =   "Encarcelar"
         Begin VB.Menu mnuCarcel 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Informaci�n"
         Begin VB.Menu mnuAccion 
            Caption         =   "General"
            Index           =   8
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Inventario"
            Index           =   9
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Skills"
            Index           =   10
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Atributos"
            Index           =   16
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "B�veda"
            Index           =   18
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Familiar o mascota"
            Index           =   20
         End
      End
      Begin VB.Menu mnuSilenciar 
         Caption         =   "Silenciar"
         Begin VB.Menu mnuSilencio 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
   End
   Begin VB.Menu cmdHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Insertar comentario"
         Index           =   4
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enviar hora"
         Index           =   5
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enemigos en mapa"
         Index           =   7
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Limpiar Mapa"
         Index           =   15
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios trabajando"
         Index           =   23
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios en grupo"
         Index           =   24
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Bloquear tile"
         Index           =   26
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios en el mapa"
         Index           =   30
      End
      Begin VB.Menu IP 
         Caption         =   "Direcci�nes de IP"
         Index           =   0
         Begin VB.Menu mnuIP 
            Caption         =   "Buscar IP's Coincidentes"
            Index           =   14
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Banear una IP"
            Index           =   17
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Lista de IPs baneadas"
            Index           =   25
         End
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Administraci�n"
      Index           =   0
      Begin VB.Menu mnuAdmin 
         Caption         =   "Apagar servidor"
         Index           =   27
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Grabar personajes"
         Index           =   28
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Iniciar WorldSave"
         Index           =   29
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Detener o reanudar el mundo"
         Index           =   33
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Limpiar el mundo"
         Index           =   34
      End
      Begin VB.Menu mnuRecargar 
         Caption         =   "Actualizar"
         Index           =   35
         Begin VB.Menu mnuReload 
            Caption         =   "Objetos"
            Index           =   1
         End
         Begin VB.Menu mnuReload 
            Caption         =   "General"
            Index           =   2
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Mapas"
            Index           =   3
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Hechizos"
            Index           =   4
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Motd"
            Index           =   5
         End
         Begin VB.Menu mnuReload 
            Caption         =   "NPCs"
            Index           =   6
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Sockets"
            Index           =   7
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Lista de clanes"
            Index           =   9
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Otros"
            Index           =   10
         End
      End
      Begin VB.Menu Ambiente 
         Caption         =   "Estado clim�tico"
         Index           =   0
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Iniciar o detener una lluvia"
            Index           =   31
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Iniciar o detener una nevada"
            Index           =   32
         End
      End
      Begin VB.Menu mnuCompressChars 
         Caption         =   "Comprimir personajes"
      End
      Begin VB.Menu mnuStartUp 
         Caption         =   "Iniciar aplicaci�n"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Matar proceso"
      End
      Begin VB.Menu mnuSQLQuery 
         Caption         =   "Correr consulta"
      End
   End
   Begin VB.Menu mnuOtros 
      Caption         =   "Otros"
      Begin VB.Menu mnuSpeed 
         Caption         =   "Velocidad"
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRapida 
            Caption         =   "R�pida"
         End
         Begin VB.Menu mnuMuy 
            Caption         =   "Muy r�pida"
         End
      End
      Begin VB.Menu mnuTransparencia 
         Caption         =   "Transparencia"
         Begin VB.Menu mnuTrans 
            Caption         =   "Ninguna"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "Baja"
            Index           =   2
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "Mediana"
            Index           =   3
         End
         Begin VB.Menu mnuTrans 
            Caption         =   "Alta"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmPanelGm - ImperiumAO - v1.4.5 R5
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
'Augusto Jos� Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Dim lista As New Collection
Dim Nick As String

Private Sub cmdAccion_Click(Index As Integer)

Dim tmp As String

Nick = cboListaUsus.Text

Select Case Index

Case 0 '/ECHAR nick
    Call ClientTCP.Send_Data_Command_GM(cmdEchar, Nick)
Case 1 '/ban motivo@nick
    tmp = InputBox("�Motivo?", "Ingrese el motivo")
    If MsgBox("�Est� seguro que desea banear al personaje " & cboListaUsus.Text & "?", vbYesNo + vbQuestion) = vbYes Then
        Call ClientTCP.Send_Data_Command_GM(cmdBan, tmp & "@" & Nick)
    End If
Case 2 '/sum nick
    Call ClientTCP.Send_Data_Command_GM(cmdSum, Nick)
Case 3 '/ira nick
    Call ClientTCP.Send_Data_Command_GM(cmdIra, Nick)
Case 4 '/rem
    tmp = InputBox("�Comentario?", "Ingrese comentario")
    Call ClientTCP.Send_Data_Command_GM(cmdRem, tmp)
Case 5 '/hora
    'Call ClientTCP.Send_Data_Command_GM(cmdHora)
Case 6 '/donde nick
    Call ClientTCP.Send_Data_Command_GM(cmdDonde, Nick)
Case 7 '/nene
    tmp = InputBox("�En qu� mapa?", vbNullString)
    Call ClientTCP.Send_Data_Command_GM(cmdNene, Trim(tmp))
Case 8 '/info nick
    Call ClientTCP.Send_Data_Command_GM(cmdInfo, Nick)
Case 9 '/inv nick
    Call ClientTCP.Send_Data_Command_GM(cmdInv, Nick)
Case 10 '/skills nick
    Call ClientTCP.Send_Data_Command_GM(cmdSkills, Nick)
Case 11 '/carcel minutos nick
    tmp = InputBox("�Minutos a encarcelar? (hasta 60)", vbNullString)
    If MsgBox("�Esta seguro que desea encarcelar al personaje vbNullString" & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
        Call ClientTCP.Send_Data_Command_GM(cmdCarcel, tmp & " " & Nick)
    End If
Case 12 '/unban nick
    If MsgBox("�Esta seguro que desea removerle el ban al personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
        Call ClientTCP.Send_Data_Command_GM(cmdUnBan, Nick)
    End If
Case 13 '/nick2ip nick
    Call ClientTCP.Send_Data_Command_GM(cmdNick2IP, Nick)
Case 14 '/sameip nick
    Call ClientTCP.Send_Data_Command_GM(cmdSameIP, Nick)
Case 15
    tmp = InputBox("�Mapa?", vbNullString)
    Call ClientTCP.Send_Data_Command_GM(cmdCleanMap, Trim(tmp))
Case 16 '/att nick
    Call ClientTCP.Send_Data_Command_GM(cmdAtt, Nick)
Case 17
    tmp = InputBox("Escriba la direcci�n IP a banear", vbNullString)
    If MsgBox("�Esta seguro que desea banear la IP " & tmp & "?", vbYesNo + vbQuestion) = vbYes Then
        Call ClientTCP.Send_Data_Command_GM(cmdBanIP, tmp)
    End If
Case 18 '/bov nick
    Call ClientTCP.Send_Data_Command_GM(cmdBov, Nick)
Case 19
    If MsgBox("�Esta seguro que desea banear la IP del personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
        Call ClientTCP.Send_Data_Command_GM(cmdBanIP, "BANIP " & Nick)
    End If
Case 20 '/infofami nick
    Call ClientTCP.Send_Data_Command_GM(cmdInfoFami, Nick)
Case 21 '/revivir nick
    Call ClientTCP.Send_Data_Command_GM(cmdRevivir, Nick)
Case 22
    Call ClientTCP.Send_Data_Command_GM(cmdHmr, Nick)
Case 23
    Call ClientTCP.Send_Data_Command_GM(cmdTrabajando)
Case 24
    Call ClientTCP.Send_Data_Command_GM(cmdEnGrupo)
Case 25
    Call ClientTCP.Send_Data_Command_GM(cmdBanIPList)
Case 26
    Call ClientTCP.Send_Data_Command_GM(cmdBloq)
Case 27
    Call ClientTCP.Send_Data_Command_GM(cmdApagar, "1")
Case 28
    Call ClientTCP.Send_Data_Command_GM(cmdGrabar)
Case 29
    Call ClientTCP.Send_Data_Command_GM(cmdDoBackUP)
Case 30
    Call ClientTCP.Send_Data_Command_GM(cmdOnlineMap)
Case 31
    Call ClientTCP.Send_Data_Command_GM(cmdLluvia)
Case 32
    Call ClientTCP.Send_Data_Command_GM(cmdNieve)
Case 34
    Call ClientTCP.Send_Data_Command_GM(cmdLimpiar)
Case 35 '/silencio minutos nick
    tmp = InputBox("�Minutos a silenciar? (hasta 60)", vbNullString)
    If MsgBox("�Esta seguro que desea silenciar al personaje " & Nick & "?", vbYesNo + vbQuestion) = vbYes Then
        Call ClientTCP.Send_Data_Command_GM(cmdSilencio, tmp & " " & Nick)
    End If
End Select

Nick = vbNullString

End Sub

Private Sub cmdActualiza_Click()
Call ClientTCP.Send_Data(GM_User_List_Cl)
End Sub

Private Sub cmdCerrar_Click()
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
End Sub

Private Sub cmdTarget_Click()

Call PrintToConsole(Locale_GUI_Frase(353), 100, 100, 120, 0, 0)

Call FormParser.Parse_Form(frmMain, E_WAIT)
CurrentUser.UsingSkill = GM_SELECT

End Sub

Private Sub Form_Load()

List1.Clear
txtMsg.Text = vbNullString

Select Case CurrentUser.CurrentSpeed
    Case VelNormal
        mnuNormal.Checked = True
        mnuRapida.Checked = False
        mnuMuy.Checked = False
    Case VelRapida
        mnuNormal.Checked = False
        mnuRapida.Checked = True
        mnuMuy.Checked = False
    Case VelUltra
        mnuNormal.Checked = False
        mnuRapida.Checked = False
        mnuMuy.Checked = True
End Select

Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call MensajeBorrarTodos
Me.Visible = False
List1.Clear
txtMsg.Text = vbNullString
End Sub

Private Sub mnuAccion_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuAdmin_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuAmbiente_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuBan_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuCarcel_Click(Index As Integer)

If Index = 60 Then
    Call cmdAccion_Click(11)
    Exit Sub
End If

Call ClientTCP.Send_Data_Command_GM(cmdCarcel, Index & " " & cboListaUsus.Text)

End Sub

Private Sub mnuSilencio_Click(Index As Integer)

If Index = 60 Then
    Call cmdAccion_Click(35)
    Exit Sub
End If

Call ClientTCP.Send_Data_Command_GM(cmdSilencio, Index & " " & cboListaUsus.Text)

End Sub

Private Sub mnuHerramientas_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Public Sub MensajePoner(ByVal Nick As String, ByVal Mensaje As String)
On Error Resume Next
lista.Add Mensaje, Nick
End Sub

Public Sub MensajeBorrarTodos()
Do While lista.Count > 0
    Call lista.Remove(lista.Count)
Loop
End Sub

Private Sub List1_Click()
On Error Resume Next
txtMsg.Text = lista.Item(List1.Text)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    PopupMenu mnuUsuario
End If

End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    PopupMenu mnuUsuario
End If

End Sub

Private Sub mnuBorrar_Click()

Call ReadNick

If List1.ListIndex < 0 Then Exit Sub
Call ClientTCP.Send_Data(GM_Sos_Erase, Long_To_String(List1.ItemData(List1.ListIndex)))
List1.RemoveItem List1.ListIndex
txtMsg.Text = vbNullString

End Sub

Private Sub mnuIP_Click(Index As Integer)
Call cmdAccion_Click(Index)
End Sub

Private Sub mnuIRa_Click()

Call ReadNick

If List1.Visible Then
    Call ClientTCP.Send_Data_Command_GM(cmdIra, Nick)
End If

End Sub

Private Sub mnuInvalida_Click()

Call ReadNick

If List1.ListIndex < 0 Then Exit Sub
Call ClientTCP.Send_Data(GM_Sos_Inv, Long_To_String(List1.ItemData(List1.ListIndex)))
List1.RemoveItem List1.ListIndex
txtMsg.Text = vbNullString

End Sub

Private Sub mnuManual_Click()

Call ReadNick

If List1.ListIndex < 0 Then Exit Sub
Call ClientTCP.Send_Data(GM_Sos_Manual, Long_To_String(List1.ItemData(List1.ListIndex)))
List1.RemoveItem List1.ListIndex
txtMsg.Text = vbNullString

End Sub

Private Sub mnuMuy_Click()
CurrentUser.CurrentSpeed = VelUltra
frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
mnuNormal.Checked = False
mnuMuy.Checked = True
mnuRapida.Checked = False
End Sub

Private Sub mnuNormal_Click()
CurrentUser.CurrentSpeed = VelNormal
frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
mnuNormal.Checked = True
mnuMuy.Checked = False
mnuRapida.Checked = False
End Sub

Private Sub mnuRapida_Click()
CurrentUser.CurrentSpeed = VelRapida
frmMain.Engine.Engine_Scroll_Pixels_Set (CurrentUser.CurrentSpeed)
mnuNormal.Checked = False
mnuMuy.Checked = False
mnuRapida.Checked = True
End Sub

Private Sub mnuReload_Click(Index As Integer)

Select Case Index
    Case 1 'Reload objetos
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "OBJ")
    Case 2 'Reload server.ini
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "SINI")
    Case 3 'Reload mapas
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "MAP")
    Case 4 'Reload hechizos
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "SPE")
    Case 5 'Reload motd
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "MOTD")
    Case 6 'Reload npcs
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "NPC")
    Case 7 'Reload sockets
        If MsgBox("Al realizar esta acci�n reiniciar� la API de Winsock. Se cerrar�n todas las conexi�nes.", vbYesNo, "Advertencia") = vbYes Then _
            Call ClientTCP.Send_Data_Command_GM(cmdReload, "SOCK")
    Case 9 'Reload Guilds
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "GUILDS")
    Case 10 'Reload otros
        Call ClientTCP.Send_Data_Command_GM(cmdReload, "OTROS")
End Select

End Sub

Private Sub mnuSQLQuery_Click()

Dim TempSQL As String
TempSQL = InputBox("Por favor ingrese el string SQL a ejecutar", vbNullString)
Call ClientTCP.Send_Data_Command_GM(cmdQuery, TempSQL)

End Sub

Private Sub mnuStartUp_Click()

Dim TempApp As String
TempApp = InputBox("Ingrese el nombre del ejecutable que desea iniciar en el servidor.", vbNullString)
Call ClientTCP.Send_Data_Command_GM(cmdIniciar, TempApp)

End Sub

Private Sub mnuKill_Click()

Dim TempApp As String
TempApp = InputBox("Ingrese el nombre del proceso que desea matar en el servidor.", vbNullString)
Call ClientTCP.Send_Data_Command_GM(cmdKilLApp, TempApp)

End Sub

Private Sub mnutraer_Click()

Call ReadNick

If List1.Visible Then
    Call ClientTCP.Send_Data_Command_GM(cmdSum, Nick)
Else
    Call ClientTCP.Send_Data_Command_GM(cmdSum, Nick)
End If

End Sub

Private Sub list1_dblClick()
On Error Resume Next

Call ReadNick
Call ClientTCP.Send_Data_Command_GM(cmdIra, Nick)
Call ClientTCP.Send_Data(GM_Sos_Erase, Long_To_String(List1.ItemData(List1.ListIndex)))
List1.Clear
Me.Visible = False
txtMsg.Text = vbNullString

End Sub

Private Sub ReadNick()

Nick = General_Field_Read(1, List1.List(List1.ListIndex), "(")
If Nick = vbNullString Then Exit Sub
Nick = left$(Nick, Len(Nick) - 1)

End Sub

Private Sub mnuTrans_Click(Index As Integer)

Dim i As Integer

Select Case Index
    Case 1
        Call Make_Transparent_Form(Me.hwnd, 255)
    Case 2
        Call Make_Transparent_Form(Me.hwnd, 210)
    Case 3
        Call Make_Transparent_Form(Me.hwnd, 200)
    Case 4
        Call Make_Transparent_Form(Me.hwnd, 150)
End Select

For i = 1 To mnuTrans.UBound
    mnuTrans(i).Checked = (i = Index)
Next i

End Sub
