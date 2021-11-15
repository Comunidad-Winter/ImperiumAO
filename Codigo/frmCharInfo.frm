VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del usuario"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5325
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
   ScaleHeight     =   6195
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton desc 
      Caption         =   "Peticion"
      Height          =   495
      Left            =   2100
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Echar 
      Caption         =   "Echar"
      Height          =   495
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4200
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Rechazar 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   5640
      Width           =   855
   End
   Begin VB.Frame rep 
      Caption         =   "Reputacion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   4470
      Width           =   5055
      Begin VB.Label reputacion 
         Caption         =   "Reputacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label criminales 
         Caption         =   "Criminales asesinados:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Ciudadanos 
         Caption         =   "Ciudadanos asesinados:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2610
      Width           =   5055
      Begin VB.Label faccion 
         Caption         =   "Faccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label integro 
         Caption         =   "Clanes que integro:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lider 
         Caption         =   "Veces fue lider de clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label fundo 
         Caption         =   "Fundo el clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label solicitudesRechazadas 
         Caption         =   "Solicitudes rechazadas:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Solicitudes 
         Caption         =   "Solicitudes para ingresar a clanes:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame charinfo 
      Caption         =   "General"
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
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.Label status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Banco 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
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
      Begin VB.Label Nivel 
         Caption         =   "Nivel:"
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
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmCharInfo - ImperiumAO - v1.3.0
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

Public frmmiembros As Boolean
Public frmsolicitudes As Boolean

Private Sub Aceptar_Click()
frmmiembros = False
frmsolicitudes = False
Call SendData("ACEPTARI" & Right(nombre.Caption, Len(nombre.Caption) - 7))
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


Public Sub parseCharInfo(ByVal rdata As String)

If frmmiembros Then
    Rechazar.Visible = False
    Aceptar.Visible = False
    Echar.Visible = True
    Desc.Visible = False
Else
    Rechazar.Visible = True
    Aceptar.Visible = True
    Echar.Visible = False
    Desc.Visible = True
End If

nombre.Caption = "Nombre:" & General_Field_Read(1, rdata, ",")
Raza.Caption = "Raza:" & General_Field_Read(2, rdata, ",")
Clase.Caption = "Clase:" & General_Field_Read(3, rdata, ",")
Genero.Caption = "Genero:" & General_Field_Read(4, rdata, ",")
Nivel.Caption = "Nivel:" & General_Field_Read(5, rdata, ",")
Oro.Caption = "Oro:" & General_Field_Read(6, rdata, ",")
Banco.Caption = "Banco:" & General_Field_Read(7, rdata, ",")

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

solicitudes.Caption = "Solicitudes para ingresar a clanes:" & General_Field_Read(12, rdata, ",")
solicitudesRechazadas.Caption = "Solicitudes rechazadas:" & General_Field_Read(13, rdata, ",")

If y = 1 Then
    fundo.Caption = "Fundo el clan: " & General_Field_Read(16, rdata, ",")
Else
    fundo.Caption = "Fundo el clan: Ninguno"
End If

lider.Caption = "Veces fue lider de clan:" & General_Field_Read(14, rdata, ",")
integro.Caption = "Clanes que integro:" & General_Field_Read(15, rdata, ",")

y = Val(General_Field_Read(18, rdata, ","))

If y = 1 Then
    faccion.Caption = "Faccion: Ejercito Real"
Else
    k = Val(General_Field_Read(19, rdata, ","))
    If k = 1 Then
        faccion.Caption = "Faccion: Fuerzas del caos"
    Else
        faccion.Caption = "Faccion: Ninguna"
    End If
End If

Ciudadanos.Caption = "Ciudadanos asesinados:" & General_Field_Read(20, rdata, ",")
criminales.Caption = "Criminales asesinados:" & General_Field_Read(21, rdata, ",")
reputacion.Caption = "Reputacion:" & Val(General_Field_Read(8, rdata, ","))
Me.Show vbModeless, frmMain

End Sub

Private Sub desc_Click()
Call SendData("ENVCOMEN" & Right(nombre, Len(nombre) - 7))
End Sub

Private Sub Echar_Click()
Call SendData("ECHARCLA" & Right(nombre, Len(nombre) - 7))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub

Private Sub Rechazar_Click()
Call SendData("RECHAZAR" & Right(nombre, Len(nombre) - 7))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call SendData("GLINFO")
Unload Me
End Sub
