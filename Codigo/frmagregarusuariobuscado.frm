VERSION 5.00
Begin VB.Form frmAgregarUsuarioBuscado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ofrecer una recompensa"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recompensa ofrecida"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   $"frmagregarusuariobuscado.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nombre"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtBuscado 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   $"frmagregarusuariobuscado.frx":00AD
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmAgregarUsuarioBuscado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmAgregarUsuarioBuscado - ImperiumAO - v1.3.0
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

Private Sub Command1_Click()
SendData ("ADDB" & Text2.Text & "," & txtBuscado.Text)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
