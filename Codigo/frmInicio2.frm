VERSION 5.00
Begin VB.Form frmInicio2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4920
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
   Icon            =   "frmInicio2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "v1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   3
      Left            =   1080
      Top             =   3360
      Width           =   2760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Manual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   2775
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Notas de lanzamiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   2
      Left            =   1080
      Top             =   2520
      Width           =   2760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Ejecutar el Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   1695
      TabIndex        =   1
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      Top             =   1680
      Width           =   2760
   End
   Begin VB.Image Image2 
      Height          =   810
      Index           =   0
      Left            =   1080
      Top             =   840
      Width           =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   730
      TabIndex        =   0
      Top             =   180
      Width           =   3375
   End
End
Attribute VB_Name = "frmInicio2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Image2_Click(Index As Integer)

Select Case Index
    Case 0
        ShellExecute Me.hWnd, "open", App.Path & "\AoSetup.exe", "", "", 1
    Case 1
        ShellExecute Me.hWnd, "open", App.Path & "\AoYa.txt", "", "", 1
    Case 2
        ShellExecute Me.hWnd, "open", "http://www.aoya.com.ar", "", "", 1
    Case 3
        Me.Visible = False
        frmInicio.Show
End Select
End Sub

Private Sub Form_Deactivate()
On Error Resume Next
Me.SetFocus
'Me.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2(0).Tag = "1" Then
            Image2(0).Tag = "0"
            Image2(0).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
End If
If Image2(1).Tag = "1" Then
            Image2(1).Tag = "0"
            Image2(1).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
End If
If Image2(2).Tag = "1" Then
            Image2(2).Tag = "0"
            Image2(2).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
End If
If Image2(3).Tag = "1" Then
            Image2(3).Tag = "0"
            Image2(3).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
End If
End Sub

Private Sub Form_Load()
'AlwaysOnTop Me.hWnd
Dim j
For Each j In Image2()
    j.Tag = "0"
Next
Image2(0).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
Image2(1).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
Image2(2).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
Image2(3).Picture = LoadPicture(App.Path & "\Graficos\BotonInicio.jpg")
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If Image2(0).Tag = "0" Then
            Image2(0).Tag = "1"
            Image2(0).Picture = LoadPicture(App.Path & "\Graficos\BotonInicioApretado.jpg")
        End If
    Case 1
        If Image2(1).Tag = "0" Then
            Image2(1).Tag = "1"
            Image2(1).Picture = LoadPicture(App.Path & "\Graficos\BotonInicioApretado.jpg")
        End If
    Case 2
        If Image2(2).Tag = "0" Then
            Image2(2).Tag = "1"
            Image2(2).Picture = LoadPicture(App.Path & "\Graficos\BotonInicioApretado.jpg")
        End If
    Case 3
        If Image2(3).Tag = "0" Then
            Image2(3).Tag = "1"
            Image2(3).Picture = LoadPicture(App.Path & "\Graficos\BotonInicioApretado.jpg")
        End If
End Select
End Sub

Private Sub Label2_Click(Index As Integer)

Select Case Index
    Case 0
        ShellExecute Me.hWnd, "open", App.Path & "\AoSetup.exe", "", "", 1
    Case 1
        ShellExecute Me.hWnd, "open", App.Path & "\AoYa.txt", "", "", 1
    Case 2
        ShellExecute Me.hWnd, "open", "http://www.aoya.com.ar", "", "", 1
    Case 3
        Me.Visible = False
        frmInicio.Show
End Select

End Sub

