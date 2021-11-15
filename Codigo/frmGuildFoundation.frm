VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creación de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4050
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   3000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Web site del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información basica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtClanName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   $"frmGuildFoundation.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   210
         TabIndex        =   3
         Top             =   600
         Width           =   3525
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del clan:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmGuildFoundation - ImperiumAO - v1.3.0
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
CurrentUser.ClanName = txtClanName
CurrentUser.Site = Text2
Unload Me
frmGuildDetails.framAlign.Visible = True
frmGuildDetails.Show vbModeless, frmMain
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

If Len(txtClanName.Text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        MensajeAdvertencia "Nombre invalido."
        Exit Sub
    End If
Else
    MensajeAdvertencia "Nombre demasiado extenso."
    Exit Sub
End If

End Sub
