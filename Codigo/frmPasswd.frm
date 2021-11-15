VERSION 5.00
Begin VB.Form frmPasswd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5025
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
   Moveable        =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCorreoCheck 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2250
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Volver"
      Height          =   495
      Left            =   90
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3825
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   3810
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3825
      Width           =   1095
   End
   Begin VB.TextBox txtPasswdCheck 
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3465
      Width           =   3510
   End
   Begin VB.TextBox txtPasswd 
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2835
      Width           =   3510
   End
   Begin VB.TextBox txtCorreo 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1620
      Width           =   3510
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Verificación del correo electronico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   12
      Top             =   2000
      Width           =   3510
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   45
      TabIndex        =   11
      Top             =   4500
      Width           =   4935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Verifiación del password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   705
      TabIndex        =   10
      Top             =   3255
      Width           =   3510
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   705
      TabIndex        =   9
      Top             =   2625
      Width           =   3525
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Dirección de correo electronico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   8
      Top             =   1340
      Width           =   3510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmPasswd.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   60
      TabIndex        =   7
      Top             =   405
      Width           =   4890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "¡¡¡¡CUIDADO!!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   4905
   End
End
Attribute VB_Name = "frmPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmPasswd - ImperiumAO - v1.3.0
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

Private Sub Command1_Click()

If Not General_Is_Valid_Email(txtCorreo.Text) Then
    txtCorreo.Text = ""
    txtCorreoCheck.Text = ""
    lblStatus.Caption = "La dirección de correo no se puede reconocer como válida. Por favor, complete el formulario con una dirección de correo real."
    Command1.Enabled = False
    Exit Sub
End If

If frmMain.mainWinsock.State Then
    lblStatus.Caption = "Advertencia: por favor espere, se está realizando la conexión con el servidor."
    Exit Sub
End If
    
CurrentUser.UserPassword = MD5String(txtPasswd.Text)
CurrentUser.UserEmail = txtCorreo.Text
EstadoLogin = CrearNuevoPj
Me.MousePointer = 11

Call frmMain.mainWinsock.Connect(CurServerIp, CurServerPort)
Call Login(ValidarLoginMSG(CInt(bRK)))
    
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub txtCorreo_Change()
Call VerificarDatos
End Sub

Private Sub txtCorreoCheck_Change()
Call VerificarDatos
End Sub

Private Sub txtPasswd_Change()
Call VerificarDatos
End Sub

Private Sub txtPasswdCheck_Change()
Call VerificarDatos
End Sub

Private Sub VerificarDatos()
Command1.Enabled = ((txtPasswd.Text <> "" And txtCorreo.Text <> "") And (txtPasswd.Text = txtPasswdCheck.Text) And (txtCorreo.Text = txtCorreoCheck.Text))
End Sub
