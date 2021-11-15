VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración del clan"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6090
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
   ScaleHeight     =   6675
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3150
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Propuestas de paz"
      Height          =   495
      Left            =   3150
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar URL de la web del clan"
      Height          =   495
      Left            =   3150
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   495
      Left            =   3150
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0000
         Left            =   120
         List            =   "frmGuildLeader.frx":0002
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
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
      TabIndex        =   6
      Top             =   2460
      Width           =   5805
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2970
      TabIndex        =   3
      Top             =   90
      Width           =   2985
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0004
         Left            =   120
         List            =   "frmGuildLeader.frx":0006
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   150
      TabIndex        =   0
      Top             =   4110
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   1230
         ItemData        =   "frmGuildLeader.frx":0008
         Left            =   120
         List            =   "frmGuildLeader.frx":000A
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Alignment       =   2  'Center
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildLeader - ImperiumAO - v1.3.0
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

frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.ListIndex))

End Sub

Private Sub Command2_Click()

frmCharInfo.frmmiembros = True
Call SendData("1HRINFO<" & members.List(members.ListIndex))

End Sub

Private Sub Command3_Click()

Dim k$
k$ = Replace(txtguildnews, vbCrLf, "º")
Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command4_Click()

frmGuildBrief.EsLeader = True
Call SendData("CLANDETAI" & guildslist.List(guildslist.ListIndex))

End Sub

Private Sub Command5_Click()

frmGuildDetails.framAlign.Visible = False
Call frmGuildDetails.Show(vbModeless, frmGuildLeader)

End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
End Sub

Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub

Private Sub Command8_Click()
Unload Me
End Sub


Public Sub ParseLeaderInfo(ByVal data As String)

If Me.Visible Then Exit Sub

Dim r%, t%

r% = Val(General_Field_Read(1, data, "¬"))

For t% = 1 To r%
    guildslist.AddItem General_Field_Read(1 + t%, data, "¬")
Next t%

r% = Val(General_Field_Read(t% + 1, data, "¬"))
Miembros.Caption = IIf(r% > 1, "El clan cuenta con " & r% & " miembros.", "El clan cuenta con un miembro.")

Dim k%

For k% = 1 To r%
    members.AddItem General_Field_Read(t% + 1 + k%, data, "¬")
Next k%

txtguildnews = Replace(General_Field_Read(t% + k% + 1, data, "¬"), "º", vbCrLf)

t% = t% + k% + 2

r% = Val(General_Field_Read(t%, data, "¬"))

For k% = 1 To r%
    solicitudes.AddItem General_Field_Read(t% + k%, data, "¬")
Next k%

Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Deactivate()
On Error Resume Next
If Me.Visible And Not frmGuildURL.Visible _
And Not frmGuildBrief.Visible _
And Not frmCharInfo.Visible _
Then Me.SetFocus
End Sub

