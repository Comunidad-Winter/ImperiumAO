VERSION 5.00
Begin VB.Form frmGrupo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$437"
   ClientHeight    =   3675
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
   Icon            =   "frmGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbandonar 
      Caption         =   "$32"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3180
      Width           =   3375
   End
   Begin VB.CommandButton cmdExpulsar 
      Caption         =   "$31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1860
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2730
      Width           =   1635
   End
   Begin VB.CommandButton cmdInvitar 
      Caption         =   "$30"
      Height          =   390
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2730
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "$25"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3570
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2730
      Width           =   1230
   End
   Begin VB.ListBox lstGrupo 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "$33"
      Height          =   1050
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   4635
   End
End
Attribute VB_Name = "frmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGrupo - ImperiumAO - v1.4.5 R5
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

Private Sub cmdAbandonar_Click()
Call ClientTCP.Send_Data(Group_Remove_Self)
End Sub

Private Sub cmdExpulsar_Click()
Call ClientTCP.Send_Data(Group_Remove_Member, lstGrupo.Text)
Me.Visible = False
End Sub

Private Sub cmdInvitar_Click()
Call ClientTCP.Send_Data(Group_Toggle)
Me.Visible = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
