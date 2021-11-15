VERSION 5.00
Begin VB.Form frmGMAyuda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optConsulta 
      Caption         =   "Pedido de fundación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   1980
      TabIndex        =   10
      Top             =   1890
      Width           =   1755
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Otro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   9
      Top             =   1890
      Width           =   855
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Reporte de bug"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   1650
      Width           =   1455
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Sugerencia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1980
      TabIndex        =   7
      Top             =   1650
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Acusación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3180
      TabIndex        =   6
      Top             =   1410
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Descargo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   5
      Top             =   1410
      Width           =   975
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Consulta regular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   1410
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   4215
   End
   Begin VB.TextBox txtMotivo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   233
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2220
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGMAyuda.frx":0000
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
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGMAyuda.frx":009D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmGMAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGMAyuda - ImperiumAO - v1.3.0
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

If txtMotivo.Text = "" Then
    Call MensajeAdvertencia("Debes escribir tu mensaje")
    Exit Sub
ElseIf DarIndiceElegido = -1 Then
    Call MensajeAdvertencia("Debes elegir el motivo de tu consulta")
    Exit Sub
Else
    SendData ("GMS" & txtMotivo.Text & "µ" & DarIndiceElegido)
    Unload Me
End If

End Sub

Private Sub Label1_Click()
frmHlp.Show vbModeless, frmGMAyuda
End Sub

Private Sub optConsulta_Click(Index As Integer)

Dim i As Integer

For i = 0 To 6
    If i <> Index Then
        optConsulta(i).Value = False
    Else
        optConsulta(i).Value = True
    End If
Next i

Select Case Index
    Case 0
        Label2.Caption = "¡Por favor explique correctamente el motivo de su consulta!"
    Case 1
        Label2.Caption = "Deje el nombre del personaje del que está pidiendo descargo por una medida, conjunto con el administrador que está relacionado con ella."
    Case 2
        Label2.Caption = "Se dará prioridad a su consulta enviando un mensaje a los administradores conectados, por favor utilize ésta opción responsablemente."
    Case 3
        Label2.Caption = "Su sugerencia SERÁ leída por un miembro del staff, y será tomada en cuenta para futuros cambios."
    Case 4
        Label2.Caption = "Explique de la forma más detallada la forma de repetir el error. El staff de programación lo resolverá lo antes posible."
    Case 5
        Label2.Caption = "Deje la mayor cantidad de datos posibles, esta opción es para consultas que no entran en otras secciónes."
    Case 6
        Label2.Caption = "Deje los siguientes datos: nombre del clan (sensitivo a mayúsculas y minúsculas), alineamiento, historia y otros datos que considere de importancia."
End Select

End Sub

Private Function DarIndiceElegido() As Integer

Dim i As Integer

For i = 0 To 6
    If optConsulta(i).Value = True Then
        DarIndiceElegido = i
        Exit Function
    End If
Next i

DarIndiceElegido = -1

End Function
