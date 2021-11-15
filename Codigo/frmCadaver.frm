VERSION 5.00
Begin VB.Form frmCadaver 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6015
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
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   3960
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   2490
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2490
   End
   Begin VB.CommandButton cmdOro 
      Caption         =   "Tomar oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   240
      ScaleHeight     =   540
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   120
      Width           =   555
   End
   Begin VB.CommandButton cmdTomar 
      Caption         =   "Tomar Objeto"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   2505
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   3240
      MouseIcon       =   "frmCadaver.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5520
      Width           =   2505
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   392
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   3465
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   420
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   180
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cadáver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2265
      TabIndex        =   7
      Top             =   6750
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   6
      Top             =   975
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   3990
      TabIndex        =   5
      Top             =   630
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   2730
      TabIndex        =   4
      Top             =   1170
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   1125
      TabIndex        =   3
      Top             =   450
      Width           =   45
   End
End
Attribute VB_Name = "frmCadaver"
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

Option Explicit

Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub cmdOro_Click()
    Call SendData("GLT")
End Sub

Private Sub cmdTomar_Click()

Call IAO_SE.PlaySound(SND_CLICK)

If List1(0).List(List1(0).ListIndex) = "Nada" Or _
   List1(0).ListIndex < 0 Then Exit Sub

frmCadaver.List1(0).SetFocus
LastIndex1 = List1(0).ListIndex
        
SendData ("ARET" & List1(0).ListIndex + 1)
                                
List1(0).Clear

List1(1).Clear

NPCInvDim = 0

End Sub

Private Sub Command2_Click()
SendData ("FINCAD")
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub list1_Click(Index As Integer)

Select Case Index
    Case 0
        Label4(1).Caption = CadaverInventory(List1(0).ListIndex + 1).Name
        Label4(0).Caption = "Cantidad: " & CadaverInventory(List1(0).ListIndex + 1).Amount
        'Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hDC, CadaverInventory(List1(0).ListIndex + 1).GrhIndex, SR, DR)
    Case 1
        Label4(1).Caption = UserInventory(List1(1).ListIndex + 1).Name
        Label4(0).Caption = "Cantidad: " & UserInventory(List1(1).ListIndex + 1).Amount
        'Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hDC, UserInventory(List1(1).ListIndex + 1).GrhIndex, SR, DR)
End Select
Picture1.Refresh

End Sub

Private Sub Form_Load()

If CadaverOro > 0 Then
    Label5.Caption = "Oro: " & CadaverOro
    Label5.Visible = True
    cmdOro.Visible = True
End If
    
End Sub
