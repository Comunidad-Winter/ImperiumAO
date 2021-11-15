VERSION 5.00
Begin VB.Form frmGMAyuda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$296"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmGMAyuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optConsulta 
      Caption         =   "$295"
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
      Caption         =   "$294"
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
      Caption         =   "$292"
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
      Caption         =   "$293"
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
      Caption         =   "$291"
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
      Caption         =   "$290"
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
      Caption         =   "$289"
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
      Caption         =   "$28"
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
      Top             =   5220
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
      Caption         =   "$27"
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
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$26"
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
'frmGMAyuda - ImperiumAO - v1.4.5 R5
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

Dim selIndex As Integer

selIndex = DarIndiceElegido

If txtMotivo.Text = vbNullString Then
    Call MensajeAdvertencia(Locale_GUI_Frase(264))
    Exit Sub
ElseIf selIndex = -1 Then
    Call MensajeAdvertencia(Locale_GUI_Frase(265))
    Exit Sub
Else
    Call ClientTCP.Send_Data(Game_Master_Support, txtMotivo.Text & "µ" & DarIndiceElegido)
    Unload Me
End If

End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub Label1_Click()
frmHlp.Show vbModeless, frmGMAyuda
End Sub

Private Sub Label2_Click()
If optConsulta(1).value Then ShellExecute Me.hwnd, "open", Chr$(34) & "http://soporte.imperiumao.com.ar" & Chr$(34), vbNullString, vbNullString, 1
End Sub

Private Sub optConsulta_Click(Index As Integer)

Dim i As Integer

For i = 0 To 6
    If i <> Index Then
        optConsulta(i).value = False
    Else
        optConsulta(i).value = True
    End If
Next i

Select Case Index
    Case 0
        Label2.Caption = Locale_GUI_Frase(199)
    Case 1
        Label2.Caption = Locale_GUI_Frase(200)
    Case 2
        Label2.Caption = Locale_GUI_Frase(201)
    Case 3
        Label2.Caption = Locale_GUI_Frase(202)
    Case 4
        Label2.Caption = Locale_GUI_Frase(203)
    Case 5
        Label2.Caption = Locale_GUI_Frase(204)
    Case 6
        Label2.Caption = Locale_GUI_Frase(172)
End Select

Command1.Enabled = (Index <> 1)
txtMotivo.Enabled = (Index <> 1)

End Sub

Private Function DarIndiceElegido() As Integer

Dim i As Integer

For i = 0 To 6
    If optConsulta(i).value = True Then
        DarIndiceElegido = i
        Exit Function
    End If
Next i

DarIndiceElegido = -1

End Function
