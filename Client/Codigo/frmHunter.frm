VERSION 5.00
Begin VB.Form frmHunter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$428"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmHunter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "$64"
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
      Left            =   1920
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$63"
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
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
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
      Height          =   390
      Left            =   3600
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Width           =   810
   End
   Begin VB.ListBox lstBuscados 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$62"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   90
      Width           =   450
   End
End
Attribute VB_Name = "frmHunter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmHunter - ImperiumAO - v1.4.5 R5
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
'Augusto Jos� Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Private Sub Command1_Click()
Call ClientTCP.Send_Data(Bounty_Info_Cl, lstBuscados.ListIndex + 1)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
frmAgregarUsuarioBuscado.Show vbModeless, frmHunter
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub lstBuscados_Click()
Call ClientTCP.Send_Data(Bounty_Info_Cl, lstBuscados.ListIndex + 1)
End Sub
