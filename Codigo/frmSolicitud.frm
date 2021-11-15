VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso al clan"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4530
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
   ScaleHeight     =   3495
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   3360
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1500
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmSolicitud.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildSol - ImperiumAO - v1.3.0
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

Dim CName As String
Dim firstprint As Boolean

Private Sub Command1_Click()
Dim f$

f$ = "SOLICITUD" & CName
f$ = f$ & "," & Replace(Text1, vbCrLf, "º")

Call SendData(f$)

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)

CName = GuildName

End Sub
