VERSION 5.00
Begin VB.Form frmPregunta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   15
   ClientTop       =   45
   ClientWidth     =   3915
   ClipControls    =   0   'False
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
   Icon            =   "frmPregunta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdCancelar 
      Height          =   465
      Left            =   2010
      Tag             =   "1"
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Image cmdAceptar 
      Height          =   465
      Left            =   720
      Tag             =   "1"
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3675
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMensaje 
      Caption         =   "Mensaje"
      Visible         =   0   'False
      Begin VB.Menu mnuNormal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mnuPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu mnuGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mnuGrupo 
         Caption         =   "Grupo"
      End
      Begin VB.Menu mnuGMs 
         Caption         =   "GMs"
      End
      Begin VB.Menu mnuClan 
         Caption         =   "Clan"
      End
      Begin VB.Menu mnuFaccion 
         Caption         =   "Faccion"
      End
   End
End
Attribute VB_Name = "frmPregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmMensaje - ImperiumAO - v1.4.5 R5
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

Public Unload_And_Update As Boolean

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)
cmdAceptar.Picture = General_Load_Picture_From_Resource_Ex("pregaceptardown")
cmdAceptar.Tag = "1"

End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdAceptar.Tag = "0" Then
    cmdAceptar.Picture = General_Load_Picture_From_Resource_Ex("pregaceptarover")
    cmdAceptar.Tag = "1"
End If

End Sub

Private Sub cmdAceptar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Form_MouseMove(Button, Shift, x, y)

If Unload_And_Update Then
    Call EndGame(True, True)
Else
    Me.msg.Caption = vbNullString
    Me.Visible = False
End If

End Sub

Private Sub cmdCancelar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)
cmdCancelar.Picture = General_Load_Picture_From_Resource_Ex("pregcancelardown")
cmdCancelar.Tag = "1"

End Sub

Private Sub cmdCancelar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdCancelar.Tag = "0" Then
    cmdCancelar.Picture = General_Load_Picture_From_Resource_Ex("pregcancelarover")
    cmdCancelar.Tag = "1"
End If

End Sub

Private Sub cmdCancelar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Form_MouseMove(Button, Shift, x, y)

Me.msg.Caption = vbNullString
Me.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    If Unload_And_Update Then
        Call EndGame(True, True)
    Else
        Me.msg.Caption = vbNullString
        Me.Visible = False
    End If
End If

End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource_Ex("preg")
Call Make_Transparent_Form(Me.hwnd, 210)
Call FormParser.Parse_Form(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdAceptar.Tag = "1" Then
    cmdAceptar.Picture = Nothing
    cmdAceptar.Tag = "0"
End If

If cmdCancelar.Tag = "1" Then
    cmdCancelar.Picture = Nothing
    cmdCancelar.Tag = "0"
End If

End Sub

'[Barrin]

Private Sub msg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Me.Visible = False
End If

End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

'[/Barrin]
