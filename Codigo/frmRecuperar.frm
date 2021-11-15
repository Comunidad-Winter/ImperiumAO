VERSION 5.00
Begin VB.Form frmRecuperar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txtcorreo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton cmdRecuperar 
      Caption         =   "&Recuperar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3450
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   1170
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección de correo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   6
      Top             =   1650
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre del personaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   5
      Top             =   900
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmRecuperar.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   4500
   End
End
Attribute VB_Name = "frmRecuperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmRecuperar - ImperiumAO - v1.3.0
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

Private Sub cmdRecuperar_Click()

If frmMain.mainWinsock.State = sckConnected Or frmMain.mainWinsock.State = sckClosing Then
    Exit Sub
Else
    Me.MousePointer = 11
    EstadoLogin = RecuperarPass
    Call frmMain.mainWinsock.Connect(CurServerIp, CurServerPort)
End If

End Sub

Private Sub Command1_Click()
If Me.MousePointer <> 11 Then Me.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub
