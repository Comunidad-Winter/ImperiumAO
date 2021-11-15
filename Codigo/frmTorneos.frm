VERSION 5.00
Begin VB.Form frmTorneosLider 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de torneos"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6135
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
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Actualizar &lista de participantes"
      Height          =   315
      Left            =   3120
      MouseIcon       =   "frmTorneos.frx":0000
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Comenzar Torneo"
      Height          =   315
      Left            =   3120
      MouseIcon       =   "frmTorneos.frx":0152
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   3120
      MouseIcon       =   "frmTorneos.frx":02A4
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Frame txtnews 
      Caption         =   "Descripcion del Torneo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmTorneos.frx":03F6
         TabIndex        =   4
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Concursantes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox members 
         Height          =   1815
         ItemData        =   "frmTorneos.frx":0548
         Left            =   120
         List            =   "frmTorneos.frx":054A
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Recuerda que al ingresar ""Comenzar Torneo"" se cerrará la inscripción y el torneo comenzará en los próximos diez minutos"
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
      Left            =   3180
      TabIndex        =   7
      Top             =   300
      Width           =   2835
   End
End
Attribute VB_Name = "frmTorneosLider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmTorneosLider - ImperiumAO - v1.3.0
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

Private Sub Command1_Click()
Call SendData("/TORNEOS")
End Sub

Private Sub Command2_Click()
On Error Resume Next
Call SendData("TRUN")
Unload Me
frmMain.SetFocus
End Sub

Private Sub Command3_Click()
SendData "TACT" & txtguildnews
End Sub

Private Sub Command8_Click()
On Error Resume Next
Unload Me
frmMain.SetFocus
End Sub
