Attribute VB_Name = "modBindKeys"
'*****************************************************************
'modBindKeys - ImperiumAO - v1.3.0
'
'User input functions.
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

Type tBoton
    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer
End Type

Type tBindedKey
    KeyCode As Integer
    Name As String
End Type

Public NUMBOTONES As Integer
Public NUMBINDS As Integer

Public MacroKeys() As tBoton
Public BindKeys() As tBindedKey
Public BotonElegido As Integer

Public Function Accionar(ByVal KeyCode As Integer) As Boolean

    If KeyCode = vbKeyMultiply Then
        Engine.Engine_Stats_Show_Toggle
    
    ElseIf KeyCode = BindKeys(1).KeyCode Then
        If (IntervaloPermiteAtacar) And _
           (Not CurrentUser.Descansando) And _
           (Not CurrentUser.Meditando) Then
                SendData "AT9"
        End If
    
    ElseIf KeyCode = BindKeys(2).KeyCode Then
        If Not CurrentUser.Comerciando Then
            Call AgarrarItem
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, "No podes agarrar objetos mientras comercias", 255, 0, 32, False, False, False)
        End If
    
    ElseIf KeyCode = BindKeys(3).KeyCode Then
        If Not CurrentUser.Comerciando Then
            Call TirarItem
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, "No podes tirar objetos mientras comercias", 255, 0, 32, False, False, False)
        End If
    
    ElseIf KeyCode = BindKeys(6).KeyCode Then
        Call SendData("SEG")
        CurrentUser.Seguro = Not CurrentUser.Seguro
        frmMain.modoseguro.Visible = Not frmMain.modoseguro.Visible
        frmMain.nomodoseguro.Visible = Not frmMain.nomodoseguro.Visible
    
    ElseIf KeyCode = BindKeys(12).KeyCode Then
        Call SendData("TAB")
        CurrentUser.Combate = Not CurrentUser.Combate
        frmMain.modocombate.Visible = Not frmMain.modocombate.Visible
        frmMain.nomodocombate.Visible = Not frmMain.nomodocombate.Visible

    ElseIf KeyCode = BindKeys(7).KeyCode Then
        Engine.Engine_Label_Render_Set
    
    ElseIf KeyCode = BindKeys(8).KeyCode Then
        Call SendData("UK" & Domar)
    
    ElseIf KeyCode = BindKeys(9).KeyCode Then
        Call SendData("UK" & Robar)
    
    ElseIf KeyCode = BindKeys(5).KeyCode Then
        Call EquiparItem
    
    ElseIf KeyCode = BindKeys(4).KeyCode Then
        If IntervaloPermiteUsar Then Call UsarItem
    
    ElseIf KeyCode = BindKeys(10).KeyCode Then
        If CurrentUser.UserPuedeRefrescar Then
            Call SendData("RPU")
            CurrentUser.UserPuedeRefrescar = False
            Beep
        End If
    
    ElseIf KeyCode = BindKeys(11).KeyCode Then
        Call SendData("UK" & Ocultarse)
        
    ElseIf KeyCode = BindKeys(13).KeyCode Then
        Call SendData("RPG")
        CurrentUser.Rol = Not CurrentUser.Rol
        frmMain.modorol.Visible = Not frmMain.modorol.Visible
        frmMain.nomodorol.Visible = Not frmMain.nomodorol.Visible
        
    Else
        Accionar = False
        Exit Function
    End If

    Accionar = True

End Function

Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido <= MAX_INVENTORY_SLOTS) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show vbModeless, frmMain
           End If
        End If
    End If
End Sub

Sub AgarrarItem()
    SendData "AG"
End Sub

Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then SendData "USA" & ItemElegido
End Sub

Sub EquiparItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then _
        SendData "EQUI" & ItemElegido
End Sub

Sub LoadDefaultBinds()

Dim Arch As String, lc As Integer
Arch = App.Path & "\init\" & "ImpAoInit.bnd"

NUMBINDS = Val(General_Var_Get(Arch, "INIT", "NumBinds"))
ReDim Preserve BindKeys(1 To NUMBINDS) As tBindedKey

For lc = 1 To NUMBINDS
    BindKeys(lc).KeyCode = Val(General_Field_Read(1, General_Var_Get(Arch, "DEFAULTS", Str(lc)), ","))
    BindKeys(lc).Name = General_Field_Read(2, General_Var_Get(Arch, "DEFAULTS", Str(lc)), ",")
Next lc

End Sub

Public Sub MouseLeftClick(ByVal tX As Integer, ByVal tY As Integer)

If Not CBool((GetKeyState(vbKeyShift) Or 1) Mod -127) Then
    Call SendData("/TP" & tX & ";" & tY)
    Exit Sub
End If

If CurrentUser.UsingSkill = 0 Then
    SendData "LC" & tX & "," & tY
Else
    Select Case CurrentUser.UsingSkill
        Case Magia
            'If Not IntervaloPermiteLanzarSpell Then Exit Sub
        Case Proyectiles, Arrojadizas
            If Not IntervaloPermiteAtacar Then Exit Sub
        Case Else
            If Not IntervaloPermiteTrabajar Then Exit Sub
    End Select
    
    frmMain.MousePointer = vbDefault
    SendData "WLC" & tX & "," & tY & "," & CurrentUser.UsingSkill
    CurrentUser.UsingSkill = 0

End If

End Sub

Public Sub MouseLeftDoubleClick(ByVal tX As Integer, ByVal tY As Integer)

If Not frmForo.Visible Then
    SendData "RC" & tX & "," & tY
End If

End Sub
