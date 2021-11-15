VERSION 5.00
Begin VB.Form frmTorneoCrear 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Iniciar un nuevo torneo"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6120
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TXTNombreTorneo 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   30
      TabIndex        =   19
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos principales"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.CheckBox CheckIT 
         Caption         =   "Inscribirse en el Torneo"
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   2040
         Width           =   2865
      End
      Begin VB.TextBox TXTPrecio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TXTPjs 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Modo"
         Height          =   855
         Left            =   3120
         TabIndex        =   8
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton Modo1 
            Caption         =   "Uno contra uno"
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
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Modo2 
            Caption         =   """Free for all"""
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
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Opciónes"
         Height          =   1335
         Left            =   3120
         TabIndex        =   5
         Top             =   960
         Width           =   2655
         Begin VB.OptionButton Val1 
            Caption         =   "Sin modificar jugabilidad (Invisibilidad, Parálisis)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Val2 
            Caption         =   "Vale Todo"
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
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   2415
         End
      End
      Begin VB.TextBox TXTPR 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TXTGR 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Precio de Inscripcion"
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
         Left            =   100
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Máximo de participantes"
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
         Left            =   100
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Premio del ganador"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "% / recaudado"
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
         Left            =   1920
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Ganancias"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "% / recaudado"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      MouseIcon       =   "frmTorneoCrear.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Siguiente"
      Height          =   375
      Left            =   5040
      MouseIcon       =   "frmTorneoCrear.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del Torneo:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   2415
   End
End
Attribute VB_Name = "frmTorneoCrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmTorneoCrear - ImperiumAO - v1.3.0
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

Dim PIN As Byte

Private Sub CheckIT_Click()
If PIN = 1 Then
PIN = 0
Else
PIN = 1
End If
End Sub

Private Sub TXTPjs_Change()
TXTPjs = Val(TXTPjs)
If TXTPjs < 2 Then TXTPjs = 2
If TXTPjs > 26 Then TXTPjs = 26
End Sub
Private Sub TXTPR_Change()
TXTPR = Val(TXTPR)
If Val(TXTPR) > 100 Then
    TXTPR = 100
End If
End Sub
Private Sub TXTGR_Change()
TXTGR = Val(TXTGR)
If Val(TXTGR) > 100 Then
    TXTGR = 100
End If
End Sub
Private Sub TXTPrecio_Change()
On Error Resume Next
If TXTPrecio = "Gratuito" Then
    Exit Sub
End If
If TXTPrecio <> Val(TXTPrecio) Then TXTPrecio = 1
TXTPrecio = Val(TXTPrecio)
If TXTPrecio = "" Then TXTPrecio = 1
If TXTPrecio > 5000 Then TXTPrecio = 5000
If TXTPrecio <= 0 Then TXTPrecio = "Gratuito"
End Sub



Private Sub Command1_Click()

If Len(TXTNombreTorneo.Text) <= 30 Then
    If Not AsciiValidos(TXTNombreTorneo) Then
        Call MensajeAdvertencia("¡El nombre del torneo es inválido!")
        Exit Sub
    End If
Else
    MensajeAdvertencia "El nombre del torneo es demasiado extenso."
    Exit Sub
End If

If Val(TXTPR) + Val(TXTGR) <> 100 Then
    MensajeAdvertencia "La division de ganancias es inválida."
    Exit Sub
End If

Call SendData("CTOR" & TXTNombreTorneo & "," & Val(TXTPrecio) & "," & TXTPjs & "," & TXTPR & "," & TXTGR & "," & IIf(Val1.Value = True, "0", "1") & "," & IIf(Modo1.Value = True, "0", "1") & "," & PIN)

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

