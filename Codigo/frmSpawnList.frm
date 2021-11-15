VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sumonear"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSpawn 
      Caption         =   "¿&Respawn automático?"
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
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Spawn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   345
      TabIndex        =   2
      Top             =   3360
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sa&lir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2040
      TabIndex        =   1
      Top             =   3360
      Width           =   810
   End
   Begin VB.ListBox lstCriaturas 
      Height          =   2400
      Left            =   345
      TabIndex        =   0
      Top             =   480
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selecciona la criatura:"
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
      Left            =   645
      TabIndex        =   3
      Top             =   75
      Width           =   1860
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmSpawnList - ImperiumAO - v1.3.0
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

If chkSpawn.value Then _
    If MsgBox("Recuerde que el utilizar re-spawn automático puede traer graves consecuencias y no es recomendado utilizarlo ¿Está seguro?", vbYesNo) = vbNo Then Exit Sub

Call SendData("SPA" & lstCriaturas.ListIndex + 1 & "," & chkSpawn.value)

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
