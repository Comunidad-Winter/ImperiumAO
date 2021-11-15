VERSION 5.00
Begin VB.Form frmQuest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de la quest"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuestAceptar 
      Caption         =   "Aceptar Propuesta"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lbQuestInfo 
      AutoSize        =   -1  'True
      Caption         =   "QuestInfo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label lbQuestInfo 
      AutoSize        =   -1  'True
      Caption         =   "QuestInfo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label lbDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la quest"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmQuest - ImperiumAO - v1.3.0
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

Option Explicit

Private Sub cmdQuestAceptar_Click()
Call SendData("ACQU")
Unload Me
End Sub

Private Sub cmdRechazar_Click()
Unload Me
End Sub

Sub ParseQuestInfo(ByVal Datos As String)

lbDesc.Caption = General_Field_Read(1, Datos, "¬")
lbQuestInfo(0).Caption = "Premio: " & General_Field_Read(2, Datos, "¬")
lbQuestInfo(1).Caption = "Cantidad: " & General_Field_Read(3, Datos, "¬")

End Sub
