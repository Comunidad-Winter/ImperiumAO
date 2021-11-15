VERSION 5.00
Begin VB.Form frmBuscadoInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuario buscado"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5385
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
   ScaleHeight     =   4890
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameRecompensa 
      Caption         =   "Recompensa"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox txtCantidad 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Text            =   "1"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Recompensado 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Se han ofrecido X monedas de oro por su cabeza..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "¿Deseas ofrecer más? Tu oferta será sumada a la ya hecha."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   4575
      End
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Ofrecer"
      Height          =   495
      Left            =   4200
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame charinfo 
      Caption         =   "Información"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Label Faccion 
         BackStyle       =   0  'Transparent
         Caption         =   "Facción:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   4815
      End
      Begin VB.Label Status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label LastPlace 
         Caption         =   "Último lugar en el que fue visto:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Genero 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Raza 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Clase 
         Caption         =   "Clase:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Peligrosidad 
         Caption         =   "Peligrosidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1680
      Width           =   4815
   End
End
Attribute VB_Name = "frmBuscadoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmBuscadoInfo - ImperiumAO - v1.3.0
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

Private Sub Aceptar_Click()
SendData ("OFBB" & "," & frmHunter.lstBuscados.ListIndex + 1 & "," & txtCantidad.Text)
Unload Me
frmHunter.Visible = True
End Sub

Private Sub Command1_Click()
Unload Me
frmHunter.Visible = True
End Sub

Public Sub parseBuscadoInfo(ByVal rdata As String)

nombre.Caption = "Nombre: " & General_Field_Read(1, rdata, ",")
Raza.Caption = "Raza: " & General_Field_Read(2, rdata, ",")
Clase.Caption = "Clase: " & General_Field_Read(3, rdata, ",")
Genero.Caption = "Genero: " & General_Field_Read(4, rdata, ",")
Peligrosidad.Caption = "Peligrosidad: " & CuanPeligrosoEs(Val(General_Field_Read(5, rdata, ",")))
LastPlace.Caption = "Último lugar en el que fue visto: " & General_Field_Read(6, rdata, ",")
Recompensado.Caption = "Se han ofrecido " & Val(General_Field_Read(7, rdata, ",")) & " monedas de oro por su cabeza..."

Dim y As Long, k As Long

y = Val(General_Field_Read(8, rdata, ","))

If y > 0 Then
    status.Caption = "Status: Ciudadano"
ElseIf y < 0 Then
    status.Caption = "Status: Criminal"
Else
    status.Caption = "Status: Neutro"
End If

y = Val(General_Field_Read(9, rdata, ","))

If y = 1 Then
    faccion.Caption = "Facción: Ejercito Real"
Else
    k = Val(General_Field_Read(10, rdata, ","))
    If k = 1 Then
        faccion.Caption = "Facción: Fuerzas del caos"
    Else
        faccion.Caption = "Facción: Ninguna"
    End If
End If

Me.Show vbModeless, frmHunter

End Sub

Public Function CuanPeligrosoEs(NivelPeligroso As Integer) As String

Select Case NivelPeligroso

Case 1
    CuanPeligrosoEs = "Facilísimo"
Case 2
    CuanPeligrosoEs = "Fácil"
Case 3
    CuanPeligrosoEs = "Normal"
Case 4
    CuanPeligrosoEs = "Dará batalla"
Case 5
    CuanPeligrosoEs = "Difícil"
Case 6
    CuanPeligrosoEs = "Muy Complicado"
Case Else
    CuanPeligrosoEs = "Imposible"

End Select

End Function

Private Sub txtCantidad_Change()
Aceptar.Enabled = (Val(txtCantidad.Text) > 0 And Val(txtCantidad.Text) < CurrentUser.UserGLD)
End Sub

