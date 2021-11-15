VERSION 5.00
Begin VB.Form frmChangeBind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar una configuración de controles"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtChange 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   1
      TabIndex        =   0
      Top             =   720
      Width           =   4515
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Presione en la caja de texto la nueva tecla por favor..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmChangeBind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmChangeBind - ImperiumAO - v1.3.0
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
'Augusto José Rando (augustorando@gmail.com)
'   - First Relase
'*****************************************************************

Option Explicit

Public CurrentIndex As Integer

Private Sub txtChange_Click()
txtChange.Text = ""
End Sub

Private Sub txtChange_KeyUp(KeyCode As Integer, Shift As Integer)

Dim Name As String
Name = txtChange.Text

If KeyCode > 0 Then
    
    If frmReBind.AlreadyBinded(KeyCode) Then
        lblInfo.Caption = "¡La tecla ya se encuentra asignada! Presione una diferente."
        Beep
        Exit Sub
    End If
    
    If Name = "" Then
        If KeyCode = vbKeyShift Then
            Name = "Shift"
        ElseIf KeyCode = vbKeyLeft Then
            Name = "Flecha Izquierda"
        ElseIf KeyCode = vbKeyRight Then
            Name = "Flecha Derecha"
        ElseIf KeyCode = vbKeyDown Then
            Name = "Flecha Abajo"
        ElseIf KeyCode = vbKeyUp Then
            Name = "Flecha Arriba"
        ElseIf KeyCode = vbKeyControl Then
            Name = "Control"
        ElseIf KeyCode = vbKeyPageDown Then
            Name = "Page Down"
        ElseIf KeyCode = vbKeyPageUp Then
            Name = "Page Up"
        ElseIf KeyCode = vbKeySeparator Then 'Enter teclado numerico
            Name = "Intro"
        ElseIf KeyCode = vbKeySpace Then
            Name = "Barra Espaciadora"
        ElseIf KeyCode = vbKeyDelete Then
            Name = "Delete"
        ElseIf KeyCode = vbKeyEnd Then
            Name = "Fin"
        ElseIf KeyCode = vbKeyHome Then
            Name = "Inicio"
        ElseIf KeyCode = vbKeyInsert Then
            Name = "Insert"
        Else
            Name = "Desconocido (Código " & KeyCode & ")"
        End If
    End If
    
    Call frmReBind.Change_TempKey(CurrentIndex, KeyCode, Name)
    Unload Me
    lblInfo.Caption = "Presione en la caja de texto la nueva tecla por favor..."
End If

End Sub
