Attribute VB_Name = "modClan"
'*****************************************************************
'modClan - ImperiumAO - v1.3.0
'
'Client side guild functions.
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
'   - First Relase
'*****************************************************************

Option Explicit

Public Const CV2_CHAR_SEP_PARAMETROS As String * 1 = "€"
Public Const CV2_CHAR_SEP_VALORES As String * 1 = "ø"

Public Const CV2_VENTANACLANES_FLAG_BTN_Votar As Long = &H1
Public Const CV2_VENTANACLANES_FLAG_BTN_Aceptar As Long = &H2
Public Const CV2_VENTANACLANES_FLAG_BTN_Informacion As Long = &H4
Public Const CV2_VENTANACLANES_FLAG_BTN_SolicitarIngreso As Long = &H8
Public Const CV2_VENTANACLANES_FLAG_BTN_Cancelar As Long = &H10
Public Const CV2_VENTANACLANES_FLAG_BTN_Politicas As Long = &H20
Public Const CV2_VENTANACLANES_FLAG_BTN_Administrar As Long = &H40
Public Const CV2_VENTANACLANES_FLAG_BTN_AdministrarGM As Long = &H80
Public Const CV2_VENTANACLANES_FLAG_BTN_SalirClan As Long = &H100
Public Const CV2_VENTANACLANES_FLAG_BTN_FundarClan As Long = &H200

Public VentanaClanesVisible As Boolean

Public Sub HandleClanData(sData As String)
Dim flags As Long
Dim parametros() As String
Dim valores() As String
Dim i As Long

If Len(sData) < 7 Then Exit Sub
'nada debería llegar aca con menos caracteres de los requeridos

Select Case UCase$(mid$(sData, 4, 3))
    Case "NLD"
        'nuevo lider :)
        ReDim parametros(1 To 2)
        parametros(2) = General_Field_Read(1, mid(sData, 7), Asc(CV2_CHAR_SEP_PARAMETROS))
        parametros(1) = General_Field_Read(2, mid(sData, 7), Asc(CV2_CHAR_SEP_PARAMETROS))
        AddtoRichTextBox frmMain.RecTxt, parametros(1) & " es el nuevo lider del Clan " & parametros(2), 2, 51, 223, 1, 1
    Case "LTP"
        frmClanes.HandleTratadosPendientes mid$(sData, 7)
    Case "TRG"
        frmClanes.ShowRelacionesGlobales mid$(sData, 7)
    Case "TRI"
        frmClanes.ShowRelacionesEntreClanes mid$(sData, 7)
    Case "NDL"
        parametros = Split(mid(sData, 7), CV2_CHAR_SEP_PARAMETROS)
        AddtoRichTextBox frmMain.RecTxt, "El Clan " & parametros(0) & " tiene un nuevo lider: " & parametros(1), 2, 51, 223, 1, 1
    Case "AID"
        If mid$(sData, 7) = "+" Then
            AddtoRichTextBox frmMain.RecTxt, "Mundo en pausa. Actualizando datos de clanes", 2, 51, 223, 1, 1
        Else
            AddtoRichTextBox frmMain.RecTxt, "Listo, todo hecho.", 2, 51, 223, 1, 1
        End If
    Case "TLC"
        frmClanes.ShowVotar mid$(sData, 7)
    Case "DRC"
        'datos de recursos del clan
        frmClanes.ShowRecursos mid$(sData, 7)
    Case "ADM"
        'lider administra miembros
        frmClanes.ParseADM mid$(sData, 7)
    Case "ADC"
        'lider administra clan
        frmClanes.ParseAdminClanInfo mid$(sData, 7)
    Case "IEC"
    'información de edición de clan vía GM
        frmClanes.ParseGmEdit mid$(sData, 7)
    Case "VAG"
        'admingm
    Case "ERR"
        'ante cualquier error mejor cerrar esto :P
        Unload frmClanes
        ShowErrMsg Right$(sData, 1)
    Case "MIF"
    'llego info de clan
        frmClanes.ParseInfoClan mid$(sData, 7)
        Exit Sub
    Case "AVE"
    'abrir ventana
        parametros = Split(mid$(sData, 7), CV2_CHAR_SEP_PARAMETROS)
        If (LBound(parametros) <> 0) Or (UBound(parametros) <> 1) Then
            'error llegaron mal datos
            Exit Sub
        End If
        flags = Val(parametros(1))
        frmClanes.Show
        frmClanes.HandleListaClanes parametros(0)
        frmClanes.ShowList False
        frmClanes.Show
        frmClanes.showBotones flags
End Select
    
End Sub

Private Sub ShowErrMsg(ErrCode As String)
Select Case ErrCode
    Case "0"
        AddtoRichTextBox frmMain.RecTxt, "No tiene suficientes skillpoints", 2, 51, 223, 1, 1
    Case "1"
        AddtoRichTextBox frmMain.RecTxt, "Ya fundo clan", 2, 51, 223, 1, 1
    Case "2"
        AddtoRichTextBox frmMain.RecTxt, "Nombre de clan repetido", 2, 51, 223, 1, 1
    Case "3"
        AddtoRichTextBox frmMain.RecTxt, "Formas parte de un clan, primero sali", 2, 51, 223, 1, 1
    Case "4"
        AddtoRichTextBox frmMain.RecTxt, "Ya mandaste solicitud anteriormente!!", 2, 51, 223, 1, 1
    Case "5"
        '-
    Case "6"
        '-
    Case "7"
        '-
    Case "8"
        AddtoRichTextBox frmMain.RecTxt, "No tenes oro", 2, 51, 223, 1, 1
    Case "9"
        AddtoRichTextBox frmMain.RecTxt, "Clan no habilidato por GMs", 2, 51, 223, 1, 1
    Case "A"
        AddtoRichTextBox frmMain.RecTxt, "Clan cerrado", 2, 51, 223, 1, 1
    Case "B"
        AddtoRichTextBox frmMain.RecTxt, "Clan con muchas solicitudes pendientes", 2, 51, 223, 1, 1
    Case "C"
        AddtoRichTextBox frmMain.RecTxt, "Ya fundaste clan anteriormente", 2, 51, 223, 1, 1
    Case "D"
        AddtoRichTextBox frmMain.RecTxt, "Te garchadon de dorapa", 2, 51, 223, 1, 1
    Case "E"
        AddtoRichTextBox frmMain.RecTxt, "No existe el clan", 2, 51, 223, 1, 1
    Case "F"
        AddtoRichTextBox frmMain.RecTxt, "Todo oka", 2, 51, 223, 1, 1
    Case "G"
        AddtoRichTextBox frmMain.RecTxt, "No se encontro al miembro a rajar... cosa rara :(", 2, 51, 223, 1, 1
    Case "H"
        AddtoRichTextBox frmMain.RecTxt, "No podes rajarte mientras seas lider :S", 2, 51, 223, 1, 1
    Case "I"
        AddtoRichTextBox frmMain.RecTxt, "No tenes suficientes recursos", 2, 51, 223, 1, 1
    Case "J"
        AddtoRichTextBox frmMain.RecTxt, "Hoy no es día de votación!", 2, 51, 223, 1, 1
    Case "K"
        AddtoRichTextBox frmMain.RecTxt, "El candidato que votaste no es miembro del clan :S", 2, 51, 223, 1, 1
    Case "L"
        AddtoRichTextBox frmMain.RecTxt, "El pedido del tratado que no fue aceptado ni rechazado fue reemplazado por el actual", 2, 51, 223, 1, 1
    Case "M"
        AddtoRichTextBox frmMain.RecTxt, "No podes hacer tratados con tu propio clan", 2, 51, 223, 1, 1
    Case "N"
        AddtoRichTextBox frmMain.RecTxt, "Ya Votaste. No se computa el voto nuevo", 2, 51, 223, 1, 1
    Case "Y"
        AddtoRichTextBox frmMain.RecTxt, "No podés salir del clan si sos el líder.", 2, 51, 223, 1, 1
    Case "Z"
        AddtoRichTextBox frmMain.RecTxt, "No podés salir de ningún clan ya que no perteneces a ningúno.", 2, 51, 223, 1, 1
End Select
End Sub
