VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$435"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   7470
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
   Icon            =   "frmGuildBrief.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "$36"
      Height          =   375
      Left            =   6000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5970
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "$25"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   5970
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "$39"
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
      TabIndex        =   15
      Top             =   4560
      Width           =   7215
      Begin VB.TextBox Desc 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "$38"
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
      TabIndex        =   6
      Top             =   2010
      Width           =   7215
      Begin VB.Label Codex 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   6735
      End
      Begin VB.Label Codex 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$37"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Label Alineamiento 
         Caption         =   "Alineamiento:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label Miembros 
         Caption         =   "Miembros:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6975
      End
      Begin VB.Label web 
         Caption         =   "Web site:"
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
'frmGuildBrief - ImperiumAO - v1.4.5 R5
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

Option Explicit

Public EsLeader As Boolean

Public Sub ParseGuildInfo(ByVal buffer As String)

nombre.Caption = Locale_GUI_Frase(3) & ": " & General_Field_Read(1, buffer, "¬")
fundador.Caption = Locale_GUI_Frase(164) & ": " & General_Field_Read(2, buffer, "¬")
creacion.Caption = Locale_GUI_Frase(163) & ": " & General_Field_Read(3, buffer, "¬")
web.Caption = Locale_GUI_Frase(46) & ": " & General_Field_Read(4, buffer, "¬")
Miembros.Caption = Locale_GUI_Frase(53) & ": " & General_Field_Read(5, buffer, "¬")

Select Case Val(General_Field_Read(6, buffer, "¬"))
    Case Republicano
        Alineamiento.Caption = Locale_GUI_Frase(41) & ": " & Locale_GUI_Frase(153)
        Alineamiento.ForeColor = &H808080
    Case Legal
        Alineamiento.Caption = Locale_GUI_Frase(41) & ": " & Locale_GUI_Frase(152)
        Alineamiento.ForeColor = &HC00000
    Case Caotico
        Alineamiento.Caption = Locale_GUI_Frase(41) & ": " & Locale_GUI_Frase(165)
        Alineamiento.ForeColor = &HC0&
End Select

Dim t%, k%
k% = Val(General_Field_Read(7, buffer, "¬")) + 1

For t% = 1 To k%
    Codex(t% - 1).Caption = General_Field_Read(7 + t%, buffer, "¬")
Next t%


Dim des$

des$ = General_Field_Read(7 + t%, buffer, "¬")

Desc = Replace$(des$, "º", vbCrLf)

Me.Show vbModeless, frmMain

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Call frmGuildSol.RecieveSolicitud(Right$(nombre, Len(nombre) - 7))
Call frmGuildSol.Show(vbModeless, frmGuildBrief)
End Sub

Private Sub Form_Deactivate()
If Not frmGuildSol.Visible _
And Me.Visible Then Me.SetFocus
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
