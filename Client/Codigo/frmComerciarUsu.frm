VERSION 5.00
Begin VB.Form frmComercioSeguro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9330
   ControlBox      =   0   'False
   Icon            =   "frmComerciarUsu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "$2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7920
      TabIndex        =   11
      Top             =   150
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "$16"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5955
      Left            =   3450
      TabIndex        =   14
      Top             =   750
      Width           =   5745
      Begin VB.CommandButton cmdAgregarOro 
         Caption         =   "$19"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         TabIndex        =   9
         Top             =   4920
         Width           =   2505
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2700
         TabIndex        =   7
         Top             =   2880
         Width           =   315
      End
      Begin VB.CommandButton cmdOfrecer 
         Caption         =   "$20"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         TabIndex        =   10
         Top             =   5400
         Width           =   2505
      End
      Begin VB.ListBox lstComercioSeguro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4155
         Index           =   2
         Left            =   3060
         TabIndex        =   8
         Top             =   660
         Width           =   2505
      End
      Begin VB.TextBox txtCant 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "1"
         Top             =   5460
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2700
         TabIndex        =   6
         Top             =   2580
         Width           =   315
      End
      Begin VB.ListBox lstComercioSeguro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   660
         Width           =   2490
      End
      Begin VB.Label Label3 
         Caption         =   "$18"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3060
         TabIndex        =   18
         Top             =   390
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "$17"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   390
         Width           =   2505
      End
      Begin VB.Label lblCant 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$26"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   5500
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5985
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   3315
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "$21"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   5460
         Width           =   1440
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "$1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   5460
         Width           =   1410
      End
      Begin VB.ListBox lstComercioSeguro 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   2970
      End
      Begin VB.Label lblOro 
         Caption         =   "Oro: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   2985
      End
   End
   Begin VB.PictureBox picInv 
      BackColor       =   &H00000000&
      Height          =   540
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lblCantidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   20
      Top             =   360
      Width           =   60
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   19
      Top             =   90
      Width           =   60
   End
   Begin VB.Label lblEstadoResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$270"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   6750
      Width           =   9150
   End
End
Attribute VB_Name = "frmComercioSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmComercioSeguro - ImperiumAO - v1.4.5 R5
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

Private Const LISTA_OTRO As Integer = 0
Private Const LISTA_INVENTARIO As Integer = 1
Private Const LISTA_OFERTA As Integer = 2

Private Sub cmdAceptar_Click()

Call ClientTCP.Send_Data(User_Commerce, "2")
cmdAceptar.Enabled = False
cmdRechazar.Enabled = False
lblEstadoResp.Caption = Locale_GUI_Frase(209)

End Sub

Private Sub cmdAgregar_Click()

Dim Cantidad As Long, Slot As Integer, CantidadActual As Long, IndiceListaActual As Integer, i As Integer

IndiceListaActual = -1
Cantidad = Val(Trim(txtCant.Text))

If lstComercioSeguro(LISTA_INVENTARIO).ListIndex < 0 Then Exit Sub
If lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).ListIndex) <= 0 Then Exit Sub
    
CantidadActual = lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).ListIndex)
Slot = DarItemSlot(lstComercioSeguro(LISTA_INVENTARIO).List(lstComercioSeguro(LISTA_INVENTARIO).ListIndex))
If Slot = -1 Then Exit Sub

If Cantidad <= CantidadActual Then
    
    For i = 0 To lstComercioSeguro(LISTA_OFERTA).ListCount
        If lstComercioSeguro(LISTA_OFERTA).List(i) = lstComercioSeguro(LISTA_INVENTARIO).List(lstComercioSeguro(LISTA_INVENTARIO).ListIndex) Then
            IndiceListaActual = i
            Exit For
        End If
    Next i
    
    If IndiceListaActual = -1 Then
        Call ClientTCP.Send_Data(User_Commerce, "3" & Slot & "�" & Cantidad)
        lstComercioSeguro(LISTA_OFERTA).AddItem lstComercioSeguro(LISTA_INVENTARIO).List(lstComercioSeguro(LISTA_INVENTARIO).ListIndex)
        lstComercioSeguro(LISTA_OFERTA).ItemData(lstComercioSeguro(LISTA_OFERTA).NewIndex) = Cantidad
    Else
        Call ClientTCP.Send_Data(User_Commerce, "4" & UserInventory(Slot).ObjIndex)
        Call ClientTCP.Send_Data(User_Commerce, "3" & Slot & "�" & (lstComercioSeguro(LISTA_OFERTA).ItemData(IndiceListaActual) + Cantidad))
        lstComercioSeguro(LISTA_OFERTA).ItemData(IndiceListaActual) = lstComercioSeguro(LISTA_OFERTA).ItemData(IndiceListaActual) + Cantidad
    End If
    
    If Cantidad = CantidadActual Then
        Call lstComercioSeguro(LISTA_INVENTARIO).RemoveItem(lstComercioSeguro(LISTA_INVENTARIO).ListIndex)
    Else
        lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).ListIndex) = CantidadActual - Cantidad
    End If
    
Else
    txtCant.ForeColor = vbRed
End If

End Sub

Private Sub cmdAgregarOro_Click()

Dim Cantidad As Long, Slot As Integer
Cantidad = Val(Trim(txtCant.Text))

If Cantidad <= CurrentUser.UserGLD Then
    lstComercioSeguro(LISTA_OFERTA).AddItem "Oro (Billetera)"
    lstComercioSeguro(LISTA_OFERTA).ItemData(lstComercioSeguro(LISTA_OFERTA).NewIndex) = Cantidad
    Call ClientTCP.Send_Data(User_Commerce, "7" & Cantidad)
    cmdAgregarOro.Enabled = False
Else
    txtCant.ForeColor = vbRed
End If

End Sub

Private Sub cmdQuitar_Click()

Dim nombre As String
Dim Slot As Integer
Dim IndiceListaActual As Integer
Dim i As Integer

If lstComercioSeguro(LISTA_OFERTA).ListIndex < 0 Then Exit Sub
If lstComercioSeguro(LISTA_OFERTA).ItemData(lstComercioSeguro(LISTA_OFERTA).ListIndex) <= 0 Then Exit Sub

nombre = lstComercioSeguro(LISTA_OFERTA).List(lstComercioSeguro(LISTA_OFERTA).ListIndex)
IndiceListaActual = -1

If nombre = "Oro (Billetera)" Then
    Call ClientTCP.Send_Data(User_Commerce, "8")
    cmdAgregarOro.Enabled = True
Else

    Slot = DarItemSlot(lstComercioSeguro(LISTA_OFERTA).List(lstComercioSeguro(LISTA_OFERTA).ListIndex))
    If Slot = -1 Then Exit Sub

    For i = 0 To lstComercioSeguro(LISTA_INVENTARIO).ListCount
        If lstComercioSeguro(LISTA_INVENTARIO).List(i) = lstComercioSeguro(LISTA_OFERTA).List(lstComercioSeguro(LISTA_OFERTA).ListIndex) Then
            IndiceListaActual = i
            Exit For
        End If
    Next i

    If IndiceListaActual = -1 Then
        If UserInventory(Slot).ObjIndex <> 0 Then
            lstComercioSeguro(LISTA_INVENTARIO).AddItem UserInventory(Slot).name
            lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).NewIndex) = UserInventory(Slot).Amount
            Call ClientTCP.Send_Data(User_Commerce, "4" & UserInventory(Slot).ObjIndex)
        Else
            MensajeAdvertencia "Critical error. Report to Barrin. Errorcode: 9"
            lstComercioSeguro(LISTA_INVENTARIO).AddItem Locale_GUI_Frase(269)
            lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).NewIndex) = 0
        End If
    Else
        lstComercioSeguro(LISTA_INVENTARIO).ItemData(IndiceListaActual) = UserInventory(Slot).Amount
        Call ClientTCP.Send_Data(User_Commerce, "4" & UserInventory(Slot).ObjIndex)
    End If

End If

Call lstComercioSeguro(LISTA_OFERTA).RemoveItem(lstComercioSeguro(LISTA_OFERTA).ListIndex)

End Sub

Private Sub cmdOfrecer_Click()

Call ClientTCP.Send_Data(User_Commerce, "1")
cmdAgregar.Enabled = False
cmdQuitar.Enabled = False
cmdOfrecer.Enabled = False
cmdAgregarOro.Enabled = False
txtCant.Enabled = False
lstComercioSeguro(LISTA_INVENTARIO).Enabled = False
lstComercioSeguro(LISTA_OFERTA).Enabled = False
lblEstadoResp.Caption = Locale_GUI_Frase(210)

End Sub

Private Sub cmdRechazar_Click()
Call ClientTCP.Send_Data(User_Commerce, "5")
cmdAceptar.Enabled = False
cmdRechazar.Enabled = False
lblEstadoResp.Caption = Locale_GUI_Frase(211)
End Sub

Private Sub Command2_Click()
Call ClientTCP.Send_Data(User_Commerce, "6")
cmdAceptar.Enabled = False
cmdRechazar.Enabled = False
cmdAgregar.Enabled = False
cmdQuitar.Enabled = False
txtCant.Enabled = False
lstComercioSeguro(LISTA_INVENTARIO).Enabled = False
lstComercioSeguro(LISTA_OTRO).Enabled = False
lstComercioSeguro(LISTA_OFERTA).Enabled = False
lblEstadoResp.Caption = Locale_GUI_Frase(211)
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me, E_NORMAL)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
Call FormParser.Parse_Form(Me)
End Sub

Private Sub DibujaGrh(ByVal grh As Long)
Call frmMain.Engine.Grh_Render_To_Hdc(grh, picInv.hDC, 0, 0)
End Sub

Private Sub lstComercioSeguro_Click(Index As Integer)

Dim Slot As Integer

Select Case Index
    Case LISTA_INVENTARIO
        Slot = DarItemSlot(lstComercioSeguro(LISTA_INVENTARIO).List(lstComercioSeguro(LISTA_INVENTARIO).ListIndex))
        If Slot <> -1 Then
            Call DibujaGrh(UserInventory(Slot).GrhIndex)
            lblName.Caption = UserInventory(Slot).name
            lblCantidad.Caption = Locale_GUI_Frase(206) & ": " & lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).ListIndex)
        End If
    Case LISTA_OFERTA
        If lstComercioSeguro(LISTA_OFERTA).List(lstComercioSeguro(LISTA_OFERTA).ListIndex) <> "Oro (Billetera)" Then
            Slot = DarItemSlot(lstComercioSeguro(LISTA_OFERTA).List(lstComercioSeguro(LISTA_OFERTA).ListIndex))
            If Slot <> -1 Then
                Call DibujaGrh(UserInventory(Slot).GrhIndex)
                lblName.Caption = UserInventory(Slot).name
                lblCantidad.Caption = Locale_GUI_Frase(206) & ": " & lstComercioSeguro(LISTA_OFERTA).ItemData(lstComercioSeguro(LISTA_OFERTA).ListIndex)
            End If
        Else
            Call DibujaGrh(GRH_ORO)
            lblName.Caption = Locale_GUI_Frase(212)
            lblCantidad.Caption = Locale_GUI_Frase(206) & ": " & lstComercioSeguro(LISTA_OFERTA).ItemData(lstComercioSeguro(LISTA_OFERTA).ListIndex)
        End If
                
    Case LISTA_OTRO
        Call DibujaGrh(OtroInventario(lstComercioSeguro(LISTA_OTRO).ListIndex + 1).GrhIndex)
        lblName.Caption = OtroInventario(lstComercioSeguro(LISTA_OTRO).ListIndex + 1).name
        lblCantidad.Caption = Locale_GUI_Frase(206) & ": " & OtroInventario(lstComercioSeguro(LISTA_OTRO).ListIndex + 1).Amount
End Select

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)

txtCant.ForeColor = vbBlack

If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    KeyAscii = 0
End If

End Sub

Public Sub ParseData(ByVal Argumentos As String)

Dim Grafico As Integer, nombre As String, Cantidad As Long, i As Integer, Slot As Integer, Identificador As Byte

Identificador = String_To_Byte(Argumentos, 1)
If Identificador <= 0 Then Exit Sub

Select Case Identificador
    Case 1 'Abrir ventana
        cmdRechazar.Enabled = False
        cmdAceptar.Enabled = False
        
        For i = 1 To MAX_INVENTORY_SLOTS
            If UserInventory(i).ObjIndex <> 0 Then
                lstComercioSeguro(LISTA_INVENTARIO).AddItem UserInventory(i).name
                lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).NewIndex) = UserInventory(i).Amount
            Else
                lstComercioSeguro(LISTA_INVENTARIO).AddItem Locale_GUI_Frase(269)
                lstComercioSeguro(LISTA_INVENTARIO).ItemData(lstComercioSeguro(LISTA_INVENTARIO).NewIndex) = 0
            End If
        Next i
        
        CurrentUser.Comerciando = True
        Me.Show vbModeless, frmMain
        
    Case 2 'Cerrar ventana
        CurrentUser.Comerciando = False
        Unload Me
    Case 3 'Ofertas confirmadas
        cmdRechazar.Enabled = True
        cmdAceptar.Enabled = True
        lblEstadoResp.Caption = Locale_GUI_Frase(213)
    Case 4 'Quitar item (lista del otro)
        Slot = String_To_Integer(Argumentos, 2)
        If Slot <= 0 Then Exit Sub
        
        For i = 0 To lstComercioSeguro(LISTA_OTRO).ListCount
            If lstComercioSeguro(LISTA_OTRO).List(i) = OtroInventario(Slot).name Then
                lstComercioSeguro(LISTA_OTRO).RemoveItem (i)
                Exit For
            End If
        Next i
        
        OtroInventario(Slot).GrhIndex = 0
        OtroInventario(Slot).Amount = 0
        OtroInventario(Slot).name = vbNullString
        
    Case 5 'Agregar item (lista del otro)
        Grafico = String_To_Long(Argumentos, 2)
        Cantidad = String_To_Long(Argumentos, 6)
        Slot = String_To_Integer(Argumentos, 10)
        nombre = mid$(Argumentos, 12)
        
        lstComercioSeguro(LISTA_OTRO).AddItem nombre, (Slot - 1)
        lstComercioSeguro(LISTA_OTRO).ItemData(lstComercioSeguro(LISTA_OTRO).NewIndex) = Cantidad
        
        OtroInventario(Slot).GrhIndex = Grafico
        OtroInventario(Slot).Amount = Cantidad
        OtroInventario(Slot).name = nombre
    Case 6 'Cambio en el oro
        Cantidad = String_To_Long(Argumentos, 2)
        If Cantidad < 0 Then Exit Sub
        lblOro.Caption = Locale_GUI_Frase(159) & ": " & Cantidad
        lblOro.ForeColor = vbRed
End Select

End Sub

Private Function DarItemSlot(ByVal ItemName As String) As Integer

Dim i As Long

For i = 1 To MAX_INVENTORY_SLOTS
    If UserInventory(i).name = ItemName Then
        DarItemSlot = i
        Exit Function
    End If
Next i

DarItemSlot = -1

End Function

