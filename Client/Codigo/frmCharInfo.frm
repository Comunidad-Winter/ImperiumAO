VERSION 5.00
Begin VB.Form frmCharInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$440"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5325
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
   Icon            =   "frmCharInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton desc 
      Caption         =   "$441"
      Height          =   495
      Left            =   2100
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6330
      Width           =   855
   End
   Begin VB.CommandButton Echar 
      Caption         =   "$442"
      Height          =   495
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   6330
      Width           =   855
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "$1"
      Height          =   495
      Left            =   4200
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   6330
      Width           =   975
   End
   Begin VB.CommandButton Rechazar 
      Caption         =   "$21"
      Height          =   495
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   6330
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "$2"
      Height          =   495
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   6330
      Width           =   855
   End
   Begin VB.Frame rep 
      Caption         =   "$443"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   120
      TabIndex        =   16
      Top             =   4470
      Width           =   5055
      Begin VB.Label Caoticos 
         Caption         =   "Caoticos matados:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Renegados 
         Caption         =   "Renegados matados:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label Armadas 
         Caption         =   "Armadas matados:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Milicianos 
         Caption         =   "Milicianos matados:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Republicanos 
         Caption         =   "Republicanos matados:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Imperiales 
         Caption         =   "Imperiales matados:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "$34"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2610
      Width           =   5055
      Begin VB.Label faccion 
         Caption         =   "Faccion:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label integro 
         Caption         =   "Clanes que integro:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lider 
         Caption         =   "Veces fue lider de clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label fundo 
         Caption         =   "Fundo el clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label solicitudesRechazadas 
         Caption         =   "Solicitudes rechazadas:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Solicitudes 
         Caption         =   "Solicitudes para ingresar a clanes:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame charinfo 
      Caption         =   "$68"
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
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.Label status 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Banco 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label Oro 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Genero 
         Caption         =   "Genero:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Raza 
         Caption         =   "Raza:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Clase 
         Caption         =   "Clase:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Nivel 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmCharInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmCharInfo - ImperiumAO - v1.4.5 R5
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

Public frmmiembros As Boolean
Public frmsolicitudes As Boolean

Private Sub Aceptar_Click()
frmmiembros = False
frmsolicitudes = False
Call ClientTCP.Send_Data(Guild_Accept_Member, Right$(nombre.Caption, Len(nombre.Caption) - 8))
Unload frmGuildLeader
Call ClientTCP.Send_Data(Guild_Info_Request)
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub parseCharInfo(ByVal Argumentos As String)

Dim tempLng As Long, tempByte As Byte, tempStr As String
tempStr = mid$(Argumentos, 60)

If frmmiembros Then
    Rechazar.Visible = False
    Aceptar.Visible = False
    Echar.Visible = True
    Desc.Visible = False
Else
    Rechazar.Visible = True
    Aceptar.Visible = True
    Echar.Visible = False
    Desc.Visible = True
End If

Raza.Caption = Locale_GUI_Frase(155) & ": " & RazaToString(String_To_Byte(Argumentos, 1))
Clase.Caption = Locale_GUI_Frase(156) & ": " & CharClaseValueToString(String_To_Byte(Argumentos, 2))
Genero.Caption = Locale_GUI_Frase(157) & ": " & IIf(String_To_Byte(Argumentos, 3) = Masculino, "Masculino", "Femenino")
Nivel.Caption = Locale_GUI_Frase(158) & ": " & String_To_Integer(Argumentos, 4)
Oro.Caption = Locale_GUI_Frase(159) & ": " & String_To_Long(Argumentos, 6)
Banco.Caption = Locale_GUI_Frase(160) & ": " & String_To_Long(Argumentos, 10)

tempLng = String_To_Long(Argumentos, 14)

If tempLng = eImperial Then
    status.Caption = "Status: " & Locale_GUI_Frase(152)
ElseIf tempLng = eRepublicano Then
    status.Caption = "Status: " & Locale_GUI_Frase(153)
Else
    status.Caption = "Status: " & Locale_GUI_Frase(154)
End If

tempByte = String_To_Byte(Argumentos, 18)

solicitudes.Caption = Locale_GUI_Frase(161) & ": " & String_To_Long(Argumentos, 19)
solicitudesRechazadas.Caption = Locale_GUI_Frase(162) & ": " & String_To_Long(Argumentos, 23)

If tempByte = 1 Then
    fundo.Caption = Locale_GUI_Frase(139) & ": " & General_Field_Read(2, tempStr, ",")
Else
    fundo.Caption = Locale_GUI_Frase(140)
End If

lider.Caption = Locale_GUI_Frase(137) & ": " & String_To_Long(Argumentos, 27)
integro.Caption = Locale_GUI_Frase(138) & ": " & String_To_Long(Argumentos, 31)

tempByte = String_To_Byte(Argumentos, 35)

If tempByte = eArImperial Then
    faccion.Caption = Locale_GUI_Frase(147) & ": " & Locale_GUI_Frase(148)
ElseIf tempByte = eArRepublicamo Then
    faccion.Caption = Locale_GUI_Frase(147) & ": " & Locale_GUI_Frase(149)
ElseIf tempByte = eArCaos Then
    faccion.Caption = Locale_GUI_Frase(147) & ": " & Locale_GUI_Frase(150)
Else
    faccion.Caption = Locale_GUI_Frase(147) & ": " & Locale_GUI_Frase(151)
End If

Imperiales.Caption = Locale_GUI_Frase(141) & ": " & String_To_Long(Argumentos, 36)
Republicanos.Caption = Locale_GUI_Frase(142) & ": " & String_To_Long(Argumentos, 40)
Armadas.Caption = Locale_GUI_Frase(143) & ": " & String_To_Long(Argumentos, 44)
Milicianos.Caption = Locale_GUI_Frase(144) & ": " & String_To_Long(Argumentos, 48)
Caoticos.Caption = Locale_GUI_Frase(145) & ": " & String_To_Long(Argumentos, 52)
Renegados.Caption = Locale_GUI_Frase(146) & ": " & String_To_Long(Argumentos, 56)

nombre.Caption = Locale_GUI_Frase(3) & ": " & General_Field_Read(1, tempStr, ",")

Me.Show vbModeless, frmMain

End Sub

Private Sub desc_Click()
Call ClientTCP.Send_Data(Guild_Send_Petition, Right$(nombre, Len(nombre) - 8))
End Sub

Private Sub Echar_Click()
Call ClientTCP.Send_Data(Guild_Remove_Member, Right$(nombre, Len(nombre) - 8))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call ClientTCP.Send_Data(Guild_Info_Request)
Unload Me
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub Rechazar_Click()
Call ClientTCP.Send_Data(Guild_Deny_Member, Right$(nombre, Len(nombre) - 8))
frmmiembros = False
frmsolicitudes = False
Unload frmGuildLeader
Call ClientTCP.Send_Data(Guild_Info_Request)
Unload Me
End Sub
