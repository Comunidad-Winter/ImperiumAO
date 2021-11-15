Attribute VB_Name = "modInventoryRender"
'*****************************************************************
'modInventoryRender - ImperiumAO - v1.4.5 R5
'
'User inventory rendering logic.
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

Private Const XCantItems As Integer = 5
Public ItemElegido As Integer

Public Sub ItemClick(ByVal x As Single, ByVal y As Single)

If ItemElegido = (x \ 32) + 1 + (y \ 32) * XCantItems Then
    Exit Sub
Else
    ItemElegido = (x \ 32) + 1 + (y \ 32) * XCantItems
    If ItemElegido > MAX_INVENTORY_SLOTS Or _
        ItemElegido < 1 Then
        ItemElegido = 0
    Else
        If UserInventory(ItemElegido).GrhIndex > 0 Then _
            Inventory_Render
    End If
End If

End Sub

Sub Inventory_Render()

frmMain.Engine.Inventory_Render_Start
Inventory_Render_All
frmMain.Engine.Inventory_Render_End frmMain.picInv.hwnd

End Sub

Private Function RenderInvItem(ByVal InvIndex As Integer)

Dim ItemX As Single
Dim ItemY As Single
Dim temp_array(3) As Long

ItemX = ((InvIndex - 1) Mod XCantItems)
ItemY = ((InvIndex - 1) \ XCantItems)

If UserInventory(InvIndex).PuedeUsar Then
    temp_array(0) = &HFFFFFFFF
    temp_array(1) = &HFFFFFFFF
    temp_array(2) = &HFFFFFFFF
    temp_array(3) = &HFFFFFFFF
Else
    temp_array(0) = -1763311616
    temp_array(1) = -1763311616
    temp_array(2) = -1763311616
    temp_array(3) = -1763311616
End If

frmMain.Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array


If UserInventory(InvIndex).Equipped Then
    frmMain.Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array
    
    temp_array(0) = &HFF0000
    temp_array(1) = &HFF0000
    temp_array(2) = &HFF0000
    temp_array(3) = &HFF0000
    
    frmMain.Engine.Engine_Text_Render "+", ItemX * 32 + 22, ItemY * 32 - 2, temp_array
    
    temp_array(0) = &HFFFFFF
    temp_array(1) = &HFFFFFF
    temp_array(2) = &HFFFFFF
    temp_array(3) = &HFFFFFF
    
Else
    frmMain.Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array
End If

End Function

Public Function Inventory_Render_All()

Dim i As Integer
Dim x As Single
Dim y As Single
Dim tmp As String
Dim tempito(3) As Long

tempito(0) = &HFFFFFFFF
tempito(1) = &HFFFFFFFF
tempito(2) = &HFFFFFFFF
tempito(3) = &HFFFFFFFF

For i = 1 To MAX_INVENTORY_SLOTS
    If UserInventory(i).Amount > 0 Then
        RenderInvItem i
        x = ((i - 1) Mod XCantItems) * 32
        y = (((i - 1) \ XCantItems) + 0.75) * 32 - 3
        
        tmp = CStr(UserInventory(i).Amount)

        frmMain.Engine.Engine_Text_Render tmp, Int(x - 1), Int(y - 2), tempito
        
    End If
Next i

If ItemElegido > 0 Then _
    frmMain.Engine.Grh_Inventory_Render 2, ((ItemElegido - 1) Mod XCantItems) * 32, ((ItemElegido - 1) \ XCantItems) * 32, tempito

End Function

Public Sub DibujarMenuMacros(Optional ActualizarCual As Integer = 0, Optional AlphaEffect As Byte = 0)

Dim i As Integer

If ActualizarCual <= 0 Then

    For i = 1 To NUMBOTONES
        Select Case MacroKeys(i).TipoAccion
            Case 1 'Envia comando
                Call frmMain.Engine.Grh_Render_To_Hdc(17506, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = Locale_GUI_Frase(8) & ": " & MacroKeys(i).SendString
            Case 2 'Lanza hechizo
                Call frmMain.Engine.Grh_Render_To_Hdc(609, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = Locale_GUI_Frase(9) & ": " & frmMain.hlst.List(MacroKeys(i).hlist - 1)
            Case 3 'Equipa
                Call frmMain.Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(i).invslot).GrhIndex, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = Locale_GUI_Frase(401) & ": " & UserInventory(MacroKeys(i).invslot).name
            Case 4 'Usa
                Call frmMain.Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(i).invslot).GrhIndex, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = Locale_GUI_Frase(400) & ": " & UserInventory(MacroKeys(i).invslot).name
            End Select
    Next i

Else

    Select Case MacroKeys(ActualizarCual).TipoAccion
        Case 1 'Envia comando
            Call frmMain.Engine.Grh_Render_To_Hdc(17506, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = Locale_GUI_Frase(8) & ": " & MacroKeys(ActualizarCual).SendString
        Case 2 'Lanza hechizo
            Call frmMain.Engine.Grh_Render_To_Hdc(609, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = Locale_GUI_Frase(9) & ": " & frmMain.hlst.List(MacroKeys(ActualizarCual).hlist - 1)
        Case 3 'Equipa
            Call frmMain.Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(ActualizarCual).invslot).GrhIndex, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = Locale_GUI_Frase(401) & ": " & UserInventory(MacroKeys(ActualizarCual).invslot).name
        Case 4 'Usa
            Call frmMain.Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(ActualizarCual).invslot).GrhIndex, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = Locale_GUI_Frase(400) & ": " & UserInventory(MacroKeys(ActualizarCual).invslot).name
    End Select

    frmMain.picMacro(ActualizarCual - 1).Refresh

End If

End Sub
