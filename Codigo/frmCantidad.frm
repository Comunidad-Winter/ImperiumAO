VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   2220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCant 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   300
      TabIndex        =   0
      Top             =   540
      Width           =   1470
   End
   Begin VB.Image imgMenos 
      Height          =   135
      Left            =   1800
      Top             =   630
      Width           =   195
   End
   Begin VB.Image imgMas 
      Height          =   135
      Left            =   1800
      Top             =   510
      Width           =   195
   End
   Begin VB.Image imgCerrar 
      Height          =   330
      Left            =   1890
      Tag             =   "0"
      Top             =   0
      Width           =   315
   End
   Begin VB.Image imgTodo 
      Height          =   405
      Left            =   1125
      Tag             =   "0"
      Top             =   840
      Width           =   945
   End
   Begin VB.Image imgAceptar 
      Height          =   405
      Left            =   150
      Tag             =   "0"
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmCantidad - ImperiumAO - v1.3.0
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
'*****************************************************************

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("cantidad.bmp")
Call Make_Transparent_Form(Me.hwnd, 210)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If imgAceptar.Tag = "1" Then
    imgAceptar.Picture = Nothing
    imgAceptar.Tag = "0"
End If

If imgTodo.Tag = "1" Then
    imgTodo.Picture = Nothing
    imgTodo.Tag = "0"
End If

If imgCerrar.Tag = "1" Then
    imgCerrar.Picture = Nothing
    imgCerrar.Tag = "0"
End If

End Sub

Private Sub imgCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgCerrar.Picture = General_Load_Picture_From_Resource("cerrarcantdown.bmp")
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If imgAceptar.Tag = "1" Then
    imgAceptar.Picture = Nothing
    imgAceptar.Tag = "0"
End If

If imgTodo.Tag = "1" Then
    imgTodo.Picture = Nothing
    imgTodo.Tag = "0"
End If

If imgCerrar.Tag = "0" Then
    imgCerrar.Picture = General_Load_Picture_From_Resource("cerrarcantover.bmp")
    imgCerrar.Tag = "1"
End If

End Sub

Private Sub imgCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Unload Me
End Sub

Private Sub imgAceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgAceptar.Picture = General_Load_Picture_From_Resource("dejardown.bmp")
End Sub

Private Sub imgAceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If imgTodo.Tag = "1" Then
    imgTodo.Picture = Nothing
    imgTodo.Tag = "0"
End If

If imgAceptar.Tag = "0" Then
    imgAceptar.Picture = General_Load_Picture_From_Resource("dejarover.bmp")
    imgAceptar.Tag = "1"
End If

End Sub

Private Sub imgMas_Click()
txtCant.Text = Val(txtCant.Text) + 1
End Sub

Private Sub imgMas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub imgMenos_Click()

If Val(txtCant.Text) > 0 Then _
    txtCant.Text = Val(txtCant.Text) - 1

End Sub

Private Sub imgMenos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub imgTodo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
imgTodo.Picture = General_Load_Picture_From_Resource("dejartododown.bmp")
End Sub

Private Sub imgTodo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If imgTodo.Tag = "0" Then
    imgTodo.Picture = General_Load_Picture_From_Resource("dejartodoover.bmp")
    imgTodo.Tag = "1"
End If

If imgAceptar.Tag = "1" Then
    imgAceptar.Picture = Nothing
    imgAceptar.Tag = "0"
End If

End Sub

Private Sub imgTodo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Me.Visible = False

If ItemElegido <> FLAGORO Then
    SendData "TI" & ItemElegido & "," & UserInventory(ItemElegido).Amount
Else
    SendData "TI" & ItemElegido & "," & IIf(CurrentUser.UserGLD <= 100000, CurrentUser.UserGLD, 100000)
End If

txtCant.Text = "0"

End Sub

Private Sub imgAceptar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Me.Visible = False
SendData "TI" & ItemElegido & "," & txtCant.Text
txtCant.Text = "0"

End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)

If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Or (KeyAscii >= 37 And KeyAscii <= 40)) Then
    KeyAscii = 0
End If

End Sub

Private Sub txtCant_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub
