Attribute VB_Name = "modInventoryRender"
'*****************************************************************
'modInventoryRender - ImperiumAO - v1.3.0
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
Engine.Inventory_Render_Start
Inventory_Render_All
Engine.Inventory_Render_End frmMain.picInv.hwnd
End Sub

Private Function RenderInvItem(InvIndex As Integer)

Dim ItemX As Single
Dim ItemY As Single

ItemX = ((InvIndex - 1) Mod XCantItems)
ItemY = ((InvIndex - 1) \ XCantItems)

Dim temp_array(3) As Long
temp_array(0) = &HFFFFFF
temp_array(1) = &HFFFFFF
temp_array(2) = &HFFFFFF
temp_array(3) = &HFFFFFF

Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array

If ItemElegido = InvIndex Then
    If UserInventory(InvIndex).Equipped Then
        
        Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array
        
        temp_array(0) = &HFF0000
        temp_array(1) = &HFF0000
        temp_array(2) = &HFF0000
        temp_array(3) = &HFF0000
        
        Engine.Engine_Text_Render "+", ItemX * 32 + 22, ItemY * 32 - 2, temp_array
        
        temp_array(0) = &HFFFFFF
        temp_array(1) = &HFFFFFF
        temp_array(2) = &HFFFFFF
        temp_array(3) = &HFFFFFF
    Else
        Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array
    End If
    
    Engine.Grh_Inventory_Render 2, ItemX * 32, ItemY * 32, temp_array

Else
    If UserInventory(InvIndex).Equipped Then
        Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array
        
        temp_array(0) = &HFF0000
        temp_array(1) = &HFF0000
        temp_array(2) = &HFF0000
        temp_array(3) = &HFF0000
        
        Engine.Engine_Text_Render "+", ItemX * 32 + 22, ItemY * 32 - 2, temp_array
        
        temp_array(0) = &HFFFFFF
        temp_array(1) = &HFFFFFF
        temp_array(2) = &HFFFFFF
        temp_array(3) = &HFFFFFF
    Else
        Engine.Grh_Inventory_Render UserInventory(InvIndex).GrhIndex, ItemX * 32, ItemY * 32, temp_array
    End If
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
        
        If UserInventory(i).Amount = 10000 Then
            tmp = "10000"
        Else
            tmp = Str(UserInventory(i).Amount)
        End If

        Engine.Engine_Text_Render tmp, Int(x - 3), Int(y - 2), tempito
        
    End If
Next i

End Function

Public Sub DibujarMenuMacros(Optional ActualizarCual As Integer = 0, Optional AlphaEffect As Byte = 0)

Dim i As Integer

If ActualizarCual <= 0 Then

    For i = 1 To NUMBOTONES
        Select Case MacroKeys(i).TipoAccion
            Case 1 'Envia comando
                Call Engine.Grh_Render_To_Hdc(17506, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Enviar comando: " & MacroKeys(i).SendString
            Case 2 'Lanza hechizo
                Call Engine.Grh_Render_To_Hdc(609, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Lanzar hechizo: " & frmMain.hlst.List(MacroKeys(i).hlist - 1)
            Case 3 'Trabaja
                Call Engine.Grh_Render_To_Hdc(505, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Trabajar"
            Case 4 'Equipa
                Call Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(i).invslot).GrhIndex, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Equipar objeto: " & UserInventory(MacroKeys(i).invslot).Name
            Case 5 'Usa
                Call Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(i).invslot).GrhIndex, frmMain.picMacro(i - 1).hDC, 0, 0)
                frmMain.picMacro(i - 1).ToolTipText = "Usar objeto: " & UserInventory(MacroKeys(i).invslot).Name
            End Select
    Next i

Else

    Select Case MacroKeys(ActualizarCual).TipoAccion
        Case 1 'Envia comando
            Call Engine.Grh_Render_To_Hdc(17506, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Enviar comando: " & MacroKeys(ActualizarCual).SendString
        Case 2 'Lanza hechizo
            Call Engine.Grh_Render_To_Hdc(609, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Lanzar hechizo: " & frmMain.hlst.List(MacroKeys(ActualizarCual).hlist - 1)
        Case 3 'Trabaja
            Call Engine.Grh_Render_To_Hdc(505, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Trabajar"
        Case 4 'Equipa
            Call Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(ActualizarCual).invslot).GrhIndex, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Equipar objeto: " & UserInventory(MacroKeys(ActualizarCual).invslot).Name
        Case 5 'Usa
            Call Engine.Grh_Render_To_Hdc(UserInventory(MacroKeys(ActualizarCual).invslot).GrhIndex, frmMain.picMacro(ActualizarCual - 1).hDC, 0, 0)
            frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Usar objeto: " & UserInventory(MacroKeys(ActualizarCual).invslot).Name
    End Select

    frmMain.picMacro(ActualizarCual - 1).Refresh

End If

End Sub
