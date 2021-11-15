VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7530
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
   ScaleHeight     =   7680
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Ofrecer Paz"
      Height          =   375
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton aliado 
      Caption         =   "Declarar Aliado"
      Height          =   375
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Guerra 
      Caption         =   "Declarar Guerra"
      Height          =   375
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   6000
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   7215
      Begin VB.TextBox Desc 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Codex"
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
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   7215
      Begin VB.Label Codex 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Label Alineamiento 
         Caption         =   "Alineamiento:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   6975
      End
      Begin VB.Label Aliados 
         Caption         =   "Clanes Aliados:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   6975
      End
      Begin VB.Label Enemigos 
         Caption         =   "Clanes Enemigos:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   6975
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label eleccion 
         Caption         =   "Dias para proxima eleccion de lider:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   6975
      End
      Begin VB.Label Miembros 
         Caption         =   "Miembros:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label web 
         Caption         =   "Web site:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Label lider 
         Caption         =   "Lider:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label creacion 
         Caption         =   "Fecha de creacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   6975
      End
      Begin VB.Label fundador 
         Caption         =   "Fundador:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildBrief - ImperiumAO - v1.3.0
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

Public EsLeader As Boolean


Public Sub ParseGuildInfo(ByVal buffer As String)

If Not EsLeader Then
    guerra.Visible = False
    aliado.Visible = False
    Command3.Visible = False
Else
    guerra.Visible = True
    aliado.Visible = True
    Command3.Visible = True
End If

nombre.Caption = "Nombre:" & General_Field_Read(1, buffer, "¬")
fundador.Caption = "Fundador:" & General_Field_Read(2, buffer, "¬")
creacion.Caption = "Fecha de creacion:" & General_Field_Read(3, buffer, "¬")
lider.Caption = "Lider:" & General_Field_Read(4, buffer, "¬")
web.Caption = "Web site:" & General_Field_Read(5, buffer, "¬")
Miembros.Caption = "Miembros:" & General_Field_Read(6, buffer, "¬")
eleccion.Caption = "Dias para proxima eleccion de lider:" & General_Field_Read(7, buffer, "¬")
Oro.Caption = "Oro:" & General_Field_Read(8, buffer, "¬")
Enemigos.Caption = "Clanes enemigos:" & General_Field_Read(9, buffer, "¬")
aliados.Caption = "Clanes aliados:" & General_Field_Read(10, buffer, "¬")

Select Case Val(General_Field_Read(11, buffer, "¬"))

Case Neutral
    Alineamiento.Caption = "Alineamiento: Neutro"
    Alineamiento.ForeColor = &H808080
Case Legal
    Alineamiento.Caption = "Alineamiento: Legal"
    Alineamiento.ForeColor = &HC00000
Case Caotico
    Alineamiento.Caption = "Alineamiento: Caótico"
    Alineamiento.ForeColor = &HC0&

End Select

Dim t%, k%
k% = Val(General_Field_Read(12, buffer, "¬"))

For t% = 1 To k%
    Codex(t% - 1).Caption = General_Field_Read(12 + t%, buffer, "¬")
Next t%


Dim des$

des$ = General_Field_Read(12 + t%, buffer, "¬")

Desc = Replace(des$, "º", vbCrLf)

Me.Show vbModeless, frmMain

End Sub

Private Sub aliado_Click()
Call SendData("DECALIAD" & Right(nombre, Len(nombre) - 7))
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Call frmGuildSol.RecieveSolicitud(Right$(nombre, Len(nombre) - 7))
Call frmGuildSol.Show(vbModeless, frmGuildBrief)

End Sub

Private Sub Command3_Click()
frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
Call frmCommet.Show(vbModeless, frmGuildBrief)
End Sub

Private Sub Form_Deactivate()
If Not frmCommet.Visible And _
Not frmGuildSol.Visible _
And Me.Visible Then Me.SetFocus
End Sub

Private Sub Guerra_Click()
Call SendData("DECGUERR" & Right(nombre.Caption, Len(nombre.Caption) - 7))
End Sub
