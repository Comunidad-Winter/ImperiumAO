VERSION 5.00
Begin VB.Form frmSastre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$422"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4935
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
   Icon            =   "frmSastre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "1"
      Top             =   2640
      Width           =   4665
   End
   Begin VB.ListBox lstRopas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4665
   End
   Begin VB.CommandButton Command3 
      Caption         =   "$1"
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
      Left            =   3090
      MouseIcon       =   "frmSastre.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3030
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
      Caption         =   "$25"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmSastre.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3030
      Width           =   1710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4665
   End
End
Attribute VB_Name = "frmSastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmSastre - ImperiumAO - v1.4.5 R5
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
'Kevin Neb (kbneb@hotmail.com)
'   - First Relase
'*****************************************************************

Private Sub Command3_Click()

If lstRopas.ListIndex = -1 Then Exit Sub
Call ClientTCP.Send_Data(Tailor_Build_Item, Integer_To_String(ObjSastre(lstRopas.ListIndex + 1)) & Long_To_String(CLng(txtCantidad.Text)))
Unload Me

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub txtCantidad_Change()

If Val(txtCantidad.Text) < 0 Then
    txtCantidad.Text = 1
End If

End Sub
