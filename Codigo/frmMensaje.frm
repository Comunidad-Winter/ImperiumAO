VERSION 5.00
Begin VB.Form frmMensaje 
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdAceptar 
      Height          =   465
      Left            =   1365
      Tag             =   "1"
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   3315
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
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmMensaje - ImperiumAO - v1.3.0
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

Private Sub cmdAceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)
cmdAceptar.Picture = General_Load_Picture_From_Resource("infoaceptardown.bmp")
cmdAceptar.Tag = "1"

End Sub

Private Sub cmdAceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If cmdAceptar.Tag = "0" Then
    cmdAceptar.Picture = General_Load_Picture_From_Resource("infoaceptarover.bmp")
    cmdAceptar.Tag = "1"
End If

End Sub

Private Sub cmdAceptar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = General_Load_Picture_From_Resource("info.bmp")
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

If cmdAceptar.Tag = "1" Then
    cmdAceptar.Picture = Nothing
    cmdAceptar.Tag = "0"
End If

End Sub

Public Sub PopupMenuMensaje()

Select Case CurrentUser.SendingType
    Case 1
        mnuNormal.Checked = True
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGMs.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
    Case 2
        mnuNormal.Checked = False
        mnuGritar.Checked = True
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGMs.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
    Case 3
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = True
        mnuClan.Checked = False
        mnuGMs.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
    Case 4
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = True
        mnuGMs.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
    Case 5
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGMs.Checked = True
        mnuGrupo.Checked = False
        mnuGlobal.Checked = False
    Case 6
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGMs.Checked = False
        mnuGrupo.Checked = True
        mnuGlobal.Checked = False
    Case 7
        mnuNormal.Checked = False
        mnuGritar.Checked = False
        mnuPrivado.Checked = False
        mnuClan.Checked = False
        mnuGMs.Checked = False
        mnuGrupo.Checked = False
        mnuGlobal.Checked = True
End Select

PopupMenu mnuMensaje

End Sub

'[Barrin]
'Moví este menú acá para que se pueda ver el caption del
'frmMain sin que se tenga que ver el ControlBox

Private Sub mnuNormal_Click()

CurrentUser.SendingType = 1
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGritar_click()

CurrentUser.SendingType = 2
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuPrivado_click()

CurrentUser.sndPrivateTo = InputBox("Escriba el usuario con el que desea iniciar una conversación privada", "")

If CurrentUser.sndPrivateTo <> "" Then
    CurrentUser.SendingType = 3
    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
Else
    Call MensajeAdvertencia("Debes escribir un usuario válido")
End If

End Sub

Private Sub mnuClan_click()

CurrentUser.SendingType = 4
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGMs_click()

CurrentUser.SendingType = 5
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGrupo_click()

CurrentUser.SendingType = 6
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub mnuGlobal_Click()

CurrentUser.SendingType = 7
If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

End Sub

Private Sub msg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If (Button = vbLeftButton) Then
    Call Auto_Drag(Me.hwnd)
Else
    Unload Me
End If

End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call Form_MouseMove(Button, Shift, x, y)
End Sub

'[/Barrin]
