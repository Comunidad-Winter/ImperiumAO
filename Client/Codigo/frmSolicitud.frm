VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$431"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4440
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
   Icon            =   "frmSolicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "$2"
      Height          =   495
      Left            =   180
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2820
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "$1"
      Height          =   495
      Left            =   3300
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2850
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "$54"
      Height          =   1215
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4095
   End
End
Attribute VB_Name = "frmGuildSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildSol - ImperiumAO - v1.4.5 R5
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
'Pablo Ignacio M�rquez (morgolock@speedy.com.ar)
'   - First Relase
'*****************************************************************

Dim CName As String

Private Sub Command1_Click()

Dim f$

f$ = Trim$(CName)
f$ = f$ & "," & Replace(Text1, vbCrLf, "�")
Call ClientTCP.Send_Data(Guild_Sol_Get, f$)
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)
CName = GuildName
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
