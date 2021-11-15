VERSION 5.00
Begin VB.Form frmSSCom 
   BorderStyle     =   0  'None
   Caption         =   "Comercio seguro"
   ClientHeight    =   7290
   ClientLeft      =   195
   ClientTop       =   3120
   ClientWidth     =   6960
   ControlBox      =   0   'False
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
   ScaleHeight     =   7290
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdListo 
      Caption         =   "Listo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3780
      TabIndex        =   12
      Top             =   6240
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   490
      Left            =   480
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   780
      Width           =   490
   End
   Begin VB.ListBox lstMeOfrecen 
      Height          =   1815
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      Left            =   3780
      TabIndex        =   7
      Top             =   6780
      Width           =   1575
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ListBox lstMiOferta 
      Height          =   1815
      Left            =   3840
      TabIndex        =   4
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
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
      Left            =   840
      TabIndex        =   2
      Tag             =   "0"
      Text            =   "0"
      Top             =   6200
      Width           =   2055
   End
   Begin VB.ListBox lstInventario 
      Height          =   3960
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblEstado 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando ofertas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4845
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Me ofrecen"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Mi oferta"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Inventario"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmSSCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAceptar_Click()
Call SendData("SSCOMOK")
End Sub

Private Sub CmdAgregar_Click()
If lstInventario.ListIndex = -1 Then Exit Sub
If Val(txtCantidad) < 1 Then Exit Sub
If lstInventario.List(lstInventario.ListIndex) = "nada" Then Exit Sub
Call SendData("SSAGREG" & lstInventario.ListIndex & "," & txtCantidad)
Dim i As Integer
i = lstInventario.ListIndex
lstInventario.RemoveItem i
lstInventario.AddItem "Nada", i
End Sub

Private Sub CmdCancelar_Click()
Call SendData("SSCANCELA")
End Sub

Private Sub CmdListo_Click()
Me.CmdListo.Enabled = False
Me.CmdAgregar.Enabled = False
If lstMeOfrecen.ListCount > 0 Then
    Me.lblEstado.Caption = "Esperando confirmación"
Else
    Me.lblEstado.Caption = "Esperando ofertas"
End If
Call SendData("SSCAMBIOT")
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
Picture1.SetFocus
End Sub

Private Sub Form_Load()
'El inventario debería llenarse con los datos del cliente

Me.Picture = LoadPicture(App.Path & "\Graficos\ComerSeguro.jpg")

Dim i As Integer
Me.lstInventario.AddItem "Oro: " & UserGLD
lstInventario.ItemData(lstInventario.ListCount - 1) = UserGLD
For i = 1 To MAX_INVENTORY_SLOTS
If UserInventory(i).Amount > 0 Then
    Me.lstInventario.AddItem UserInventory(i).Name & " " & UserInventory(i).Amount
    lstInventario.ItemData(lstInventario.ListCount - 1) = UserInventory(i).Amount
Else
    Me.lstInventario.AddItem "nada"
    lstInventario.ItemData(lstInventario.ListCount - 1) = 0
End If
Next i

SSComerciando = True
End Sub

Private Sub Form_LostFocus()
Me.SetFocus
Picture1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
SSComerciando = False
End Sub

Private Sub lstInventario_Click()

txtCantidad.Text = lstInventario.ItemData(lstInventario.ListIndex)
txtCantidad.Tag = lstInventario.ItemData(lstInventario.ListIndex)

If lstInventario.ListIndex > 0 And Val(txtCantidad.Tag) > 0 Then
    DibujaGrh UserInventory(lstInventario.ListIndex).GrhIndex
Else
    If lstInventario.ListIndex = 0 Then
        DibujaGrh 511
    Else
        Picture1.Cls
    End If
End If

End Sub



Private Sub txtCantidad_Validate(Cancel As Boolean)
If Val(txtCantidad.Text) > Val(txtCantidad.Tag) Then
    
    Call AddtoRichTextBox(frmMain.RecTxt, "No tienes esa cantidad", 32, 51, 233, 1, 1)
    txtCantidad.Text = txtCantidad.Tag
End If
If Val(txtCantidad.Text) < 0 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "No puedes comerciar por cantidades negativas", 32, 51, 233, 1, 1)
    txtCantidad.Text = 0
End If
End Sub

Public Sub DibujaGrh(Grh As Integer)
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32
Call DrawGrhtoHdc(Picture1.hWnd, Picture1.Hdc, Grh, SR, DR)
End Sub

