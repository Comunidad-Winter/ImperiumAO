VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar MOTD"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   4680
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
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   660
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtMotd 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   660
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡No te olvides de poner los codigos de colores al final de cada linea!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmCambiaMotd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmCambiaMotd - ImperiumAO - v1.3.0
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
'Alejandro Santos (alejandrosantos@fibertel.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Sub cmdOk_Click()

Dim t() As String
Dim i As Long, n As Long, Pos As Long

If Len(txtMotd.Text) >= 2 Then
    If Right(txtMotd.Text, 2) = vbCrLf Then txtMotd.Text = left(txtMotd.Text, Len(txtMotd.Text) - 2)
End If

t = Split(txtMotd.Text, vbCrLf)

For i = LBound(t) To UBound(t)
    n = 0
    Pos = InStr(1, t(i), "~")
    Do While Pos > 0 And Pos < Len(t(i))
        n = n + 1
        Pos = InStr(Pos + 1, t(i), "~")
    Loop
    If n <> 5 Then
        MensajeAdvertencia "Error en el formato de la linea " & i + 1 & "."
        Exit Sub
    End If
Next i

Call SendData("ZMOT" & txtMotd.Text)
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
