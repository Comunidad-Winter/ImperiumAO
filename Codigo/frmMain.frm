VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "ImperiumAO 1.3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   315
   ClientWidth     =   12000
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
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MiniMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   10200
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   29
      Top             =   7380
      Width           =   1455
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   10
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   9
      Left            =   5505
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   8
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   7
      Left            =   4335
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   6
      Left            =   3750
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6225
      Left            =   210
      ScaleHeight     =   6225
      ScaleWidth      =   8175
      TabIndex        =   20
      Top             =   2070
      Width           =   8175
   End
   Begin VB.Timer tmrExp 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   6270
      Top             =   210
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   5
      Left            =   3165
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   2580
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   9015
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   10
      Top             =   2220
      Width           =   2415
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   8865
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2085
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   1410
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   1995
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   825
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   8430
      Width           =   480
   End
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   8430
      Width           =   480
   End
   Begin MSWinsockLib.Winsock mainWinsock 
      Left            =   5790
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrMacro 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   6750
      Top             =   210
   End
   Begin VB.Timer sldTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   16200
      Top             =   16200
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7230
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   210
      MaxLength       =   500
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1755
      Visible         =   0   'False
      Width           =   7470
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   180
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":0ECA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgHora 
      Height          =   480
      Left            =   6675
      Top             =   8430
      Width           =   1695
   End
   Begin VB.Image imgMiniCerra 
      Enabled         =   0   'False
      Height          =   315
      Left            =   11325
      Top             =   150
      Width           =   510
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   4
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4350
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   0
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2010
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   1
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   2595
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   2
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3180
      Width           =   1890
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   3
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   3765
      Width           =   1890
   End
   Begin VB.Image cmdHechizos 
      Height          =   390
      Index           =   0
      Left            =   8775
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   1845
   End
   Begin VB.Image cmdMenu 
      Height          =   450
      Index           =   5
      Left            =   9270
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   1890
   End
   Begin VB.Image nomodorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":0F47
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Label lblAG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   28
      Top             =   8550
      Width           =   345
   End
   Begin VB.Label lblFU 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9330
      TabIndex        =   27
      Top             =   8340
      Width           =   345
   End
   Begin VB.Image cmdDropGold 
      Height          =   300
      Left            =   10260
      MousePointer    =   99  'Custom
      Top             =   5670
      Width           =   300
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   2
      Left            =   10740
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Label lblInvInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   9000
      TabIndex        =   21
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Index           =   1
      Left            =   8820
      TabIndex        =   19
      Top             =   870
      Width           =   1815
   End
   Begin VB.Shape ExpShp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8820
      Top             =   900
      Width           =   1815
   End
   Begin VB.Image cmdMinimizar 
      Height          =   225
      Left            =   11340
      Top             =   180
      Width           =   225
   End
   Begin VB.Image cmdCerrar 
      Height          =   225
      Left            =   11580
      Top             =   180
      Width           =   255
   End
   Begin VB.Image cmdMensaje 
      Height          =   255
      Left            =   7815
      Top             =   1740
      Width           =   555
   End
   Begin VB.Label lblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NickDelPersonaje"
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
      Left            =   8610
      TabIndex        =   18
      Top             =   180
      Width           =   2625
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   1
      Left            =   9660
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image imgCentros 
      Height          =   510
      Index           =   0
      Left            =   8580
      Top             =   1230
      Width           =   1065
   End
   Begin VB.Image cmdHechizos 
      Height          =   420
      Index           =   3
      Left            =   11460
      Top             =   3405
      Width           =   300
   End
   Begin VB.Image cmdHechizos 
      Height          =   420
      Index           =   2
      Left            =   11475
      Top             =   2910
      Width           =   300
   End
   Begin VB.Image cmdHechizos 
      Height          =   390
      Index           =   1
      Left            =   10650
      MousePointer    =   99  'Custom
      Top             =   4935
      Width           =   945
   End
   Begin VB.Label lblSED 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   17
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblHAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   10320
      TabIndex        =   16
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblST 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   15
      Top             =   6660
      Width           =   1365
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   14
      Top             =   6255
      Width           =   1365
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8745
      TabIndex        =   13
      Top             =   5850
      Width           =   1350
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10950
      TabIndex        =   9
      Top             =   870
      Width           =   435
   End
   Begin VB.Image InvEqu 
      Height          =   4275
      Left            =   8580
      Top             =   1230
      Width           =   3240
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H00000080&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8745
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8745
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8745
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      Height          =   135
      Left            =   10320
      Top             =   6285
      Width           =   1365
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      Height          =   135
      Left            =   10320
      Top             =   6690
      Width           =   1365
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   8
      Top             =   5745
      Width           =   1110
   End
   Begin VB.Image nomodoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":1385
      ToolTipText     =   "Seguro"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image nomodocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":17C3
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modorol 
      Height          =   255
      Left            =   9645
      Picture         =   "frmMain.frx":1C01
      ToolTipText     =   "Modo Rol"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa desconocido"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   8640
      TabIndex        =   2
      Top             =   7020
      Width           =   3105
   End
   Begin VB.Image modoseguro 
      Height          =   255
      Left            =   9240
      Picture         =   "frmMain.frx":2197
      ToolTipText     =   "Seguro"
      Top             =   7620
      Width           =   300
   End
   Begin VB.Image modocombate 
      Height          =   255
      Left            =   8820
      Picture         =   "frmMain.frx":25D5
      ToolTipText     =   "Modo Combate"
      Top             =   7620
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
'frmMain - ImperiumAO - v1.3.0
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
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - Complete recoding
'*****************************************************************

Option Explicit

Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

Public PedimosEst As Boolean

Dim endEvent As Long

'Barrin
Dim UltimoIndex As Integer
Public UltPos As Integer
Public UltPosInterface As Integer
Public UltPosSolapas As Integer

Private CentroActual As Byte

Private m_Jpeg As cJpeg
Private m_FileName As String

Private Sub cmdMensaje_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMensaje.Picture = General_Load_Picture_From_Resource("modotextodown")
End Sub

Private Sub cmdMensaje_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdMensaje.Picture = General_Load_Picture_From_Resource("modotextoover")
End Sub

Private Sub cmdMensaje_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call cmdMensaje_MouseMove(Button, Shift, X, Y)
frmMensaje.PopupMenuMensaje
cmdMensaje.Picture = General_Load_Picture_From_Resource("modotextoover")
End Sub

Private Sub cmdMinimizar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = General_Load_Picture_From_Resource("minimizardown")
End Sub

Private Sub cmdCerrar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call EndGame(True)
End Sub

Private Sub cmdMinimizar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = General_Load_Picture_From_Resource("minimizarover")
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = General_Load_Picture_From_Resource("cerrarover")
End Sub

Private Sub cmdMinimizar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = vbMinimized
imgMiniCerra.Picture = Nothing
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMiniCerra.Picture = General_Load_Picture_From_Resource("cerrardown")
End Sub

Private Sub Form_Activate()
    If SendTxt.Visible Then SendTxt.SetFocus
End Sub

Private Sub cmdDropGold_Click()

ItemElegido = FLAGORO

If Not CurrentUser.Comerciando Then
    If CurrentUser.UserGLD > 0 Then
        frmCantidad.Show vbModeless, frmMain
    End If
Else
    Call AddtoRichTextBox(frmMain.RecTxt, "¡No podes modificar tu inventario mientras comercias!", 255, 0, 32, False, False, False)
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If Not SendTxt.Visible Then
    If Not CurrentUser.Pausa And frmMain.Visible And Not frmForo.Visible And _
        Not frmComerciar.Visible And Not frmComercioSeguro.Visible And CurrentUser.Logged Then
    
        If Accionar(KeyCode) Then
            Exit Sub
        ElseIf KeyCode = vbKeyReturn Then
            If Not frmCantidad.Visible Then
                Call CompletarEnvioMensajes
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If
        ElseIf KeyCode = vbKeyF1 Then
            BotonElegido = 1
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF2 Then
            BotonElegido = 2
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF3 Then
            BotonElegido = 3
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF4 Then
            BotonElegido = 4
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF5 Then
            BotonElegido = 5
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF6 Then
            BotonElegido = 6
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF7 Then
            BotonElegido = 7
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF8 Then
            BotonElegido = 8
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF9 Then
            BotonElegido = 9
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF10 Then
            BotonElegido = 10
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF11 Then
            BotonElegido = 11
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = vbKeyF12 Then
            BotonElegido = 12
            If MacroKeys(BotonElegido).TipoAccion = 0 Then
                frmBindKey.Show vbModeless, frmMain
            Else
                Call Bind_Accion(BotonElegido)
            End If
        ElseIf KeyCode = 27 And CurrentUser.Saliendo Then
            Call ClientTCP.Send_Data(Cancel_Exit)
        End If
    End If
Else
    SendTxt.SetFocus
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) And (RunWindowed = 1) Then Call Auto_Drag(Me.hwnd)
End Sub

Private Sub hlst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgHora_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgHora.ToolTipText = "La hora en el mundo es: " & Meteo_Engine.Get_Time_String
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim map_x As Integer
Dim map_y As Integer

Call Engine.Char_Pos_Get(CurrentUser.CurrentChar, map_x, map_y)

If UltPos <> Index Then
    
    If UltPos >= 0 Then
        If Index = 1 Then
            Label2(Index).Caption = CurrentUser.UserPercExp & "%"
        Else
            Label2(Index).Caption = Engine.Map_Name_Get
        End If
    End If
    
    If Index = 1 Then
        Label2(Index).Caption = CurrentUser.UserExp & "/" & CurrentUser.UserPasarNivel
    Else
        Label2(Index).Caption = "Posición: " & CurrentUser.MapNum & ", " & map_x & ", " & map_y
    End If
    
    If CurrentUser.UserPasarNivel = 0 Then
        frmMain.Label2(1).Caption = "¡Nivel máximo!"
    End If
    
    UltPos = Index
End If

End Sub

Private Sub lbMensaje_Click()
PopupMenu frmMensaje.mnuMensaje
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim map_x As Integer
Dim map_y As Integer

Engine.Input_Mouse_Map_Get map_x, map_y

If Button = vbLeftButton Then
    If Engine.Map_In_Bounds(map_x, map_y) Then
        Call MouseLeftClick(map_x, map_y)
        Exit Sub
    End If
ElseIf Button = vbRightButton Then
    If Engine.Map_In_Bounds(map_x, map_y) Then
        Call MouseLeftDoubleClick(map_x, map_y)
        Exit Sub
    End If
End If

End Sub

Private Sub modocombate_Click()
    Call ClientTCP.Send_Data(Combat_Mode)
    CurrentUser.Combate = Not CurrentUser.Combate
    frmMain.modocombate.Visible = Not frmMain.modocombate.Visible
    frmMain.nomodocombate.Visible = Not frmMain.nomodocombate.Visible
End Sub

Private Sub modoseguro_Click()
    Call ClientTCP.Send_Data(Safe_Mode)
    CurrentUser.Seguro = Not CurrentUser.Seguro
    frmMain.modoseguro.Visible = Not frmMain.modoseguro.Visible
    frmMain.nomodoseguro.Visible = Not frmMain.nomodoseguro.Visible
End Sub

Private Sub modorol_Click()
    Call ClientTCP.Send_Data(Role_Mode)
    CurrentUser.Rol = Not CurrentUser.Rol
    frmMain.modorol.Visible = Not frmMain.modorol.Visible
    frmMain.nomodorol.Visible = Not frmMain.nomodorol.Visible
End Sub

Private Sub nomodocombate_Click()
    Call ClientTCP.Send_Data(Combat_Mode)
    CurrentUser.Combate = Not CurrentUser.Combate
    frmMain.modocombate.Visible = Not frmMain.modocombate.Visible
    frmMain.nomodocombate.Visible = Not frmMain.nomodocombate.Visible
End Sub

Private Sub nomodorol_Click()
    Call ClientTCP.Send_Data(Role_Mode)
    CurrentUser.Rol = Not CurrentUser.Rol
    frmMain.modorol.Visible = Not frmMain.modorol.Visible
    frmMain.nomodorol.Visible = Not frmMain.nomodorol.Visible
End Sub

Private Sub nomodoseguro_Click()
    Call ClientTCP.Send_Data(Safe_Mode)
    CurrentUser.Seguro = Not CurrentUser.Seguro
    frmMain.modoseguro.Visible = Not frmMain.modoseguro.Visible
    frmMain.nomodoseguro.Visible = Not frmMain.nomodoseguro.Visible
End Sub

Private Sub picInv_Paint()
Engine.Engine_Inventory_Render_Set
End Sub

Private Sub picMacro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

BotonElegido = Index + 1

If MacroKeys(BotonElegido).TipoAccion = 0 Or Button = vbRightButton Then
    frmBindKey.Show vbModeless, frmMain
Else
    Call Bind_Accion(Index + 1)
End If

End Sub

Private Sub picMacro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If UltimoIndex <> Index Then
    'If UltimoIndex >= 0 Then DibujarMenuMacros UltimoIndex + 1
    'DibujarMenuMacros Index + 1, 1
    UltimoIndex = Index
End If

End Sub

Private Sub tmrExp_Timer()
    
    If CurrentUser.ExpCount > 0 Then
        Call AddtoRichTextBox(frmMain.RecTxt, "¡Has ganado " & CurrentUser.ExpCount & " puntos de experiencia!", 51, 183, 247, True, False, False)
        CurrentUser.ExpCount = 0
    End If

End Sub

Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido <= MAX_INVENTORY_SLOTS) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            Call ClientTCP.Send_Data(Drop_Item, ItemElegido & "," & 1)
        Else
           If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show vbModeless, frmMain
           End If
        End If
    End If

End Sub

Private Sub AgarrarItem()
Call ClientTCP.Send_Data(Get_Item)
End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then
        Call ClientTCP.Send_Data(Use_Item, ItemElegido)
    End If
End Sub

Private Sub EquiparItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then _
        Call ClientTCP.Send_Data(Equip_Item, ItemElegido)
End Sub

Private Sub Form_Load()

Me.Picture = General_Load_Picture_From_Resource("todo")
Me.Caption = Form_Caption
Call Make_Transparent_Richtext(RecTxt.hwnd)
Call CambiaCentro(CentroInventario)

UltPos = -1
UltimoIndex = -1
UltPosInterface = -1
UltPosSolapas = -1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseX = X
    MouseY = Y
    
    If UltimoIndex >= 0 Then
        'DibujarMenuMacros UltimoIndex + 1
        UltimoIndex = -1
    End If
    
    If UltPos >= 0 Then
        If UltPos = 1 Then
            Label2(UltPos).Caption = CurrentUser.UserPercExp & "%"
        Else
            Label2(UltPos).Caption = Engine.Map_Name_Get
        End If
        
        If CurrentUser.UserPasarNivel = 0 Then
            frmMain.Label2(1).Caption = "¡Nivel máximo!"
        End If
        
        UltPos = -1
        
    End If
    
    Call RestaurarCentroActual
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub MostrarCentroInventario()
    InvEqu.Picture = General_Load_Picture_From_Resource("centroinventario")
    picInv.Visible = True
    lblInvInfo.Visible = True
    lblInvInfo = ""
End Sub

Private Sub OcultarCentroInventario()
    picInv.Visible = False
    lblInvInfo.Visible = False
End Sub

Private Sub MostrarCentroHechizos()
    InvEqu.Picture = General_Load_Picture_From_Resource("centrohechizos")
    cmdHechizos(0).Visible = True
    cmdHechizos(1).Visible = True
    cmdHechizos(2).Visible = True
    cmdHechizos(3).Visible = True
    hlst.Visible = True
End Sub

Private Sub OcultarCentroHechizos()
    hlst.Visible = False
    cmdHechizos(0).Visible = False
    cmdHechizos(1).Visible = False
    cmdHechizos(2).Visible = False
    cmdHechizos(3).Visible = False
End Sub

Private Sub MostrarCentroMenu()
    cmdMenu(0).Visible = True
    cmdMenu(1).Visible = True
    cmdMenu(2).Visible = True
    cmdMenu(3).Visible = True
    cmdMenu(4).Visible = True
    cmdMenu(5).Visible = True
    InvEqu.Picture = General_Load_Picture_From_Resource("centromenu")
End Sub

Private Sub OcultarCentroMenu()
    cmdMenu(0).Visible = False
    cmdMenu(1).Visible = False
    cmdMenu(2).Visible = False
    cmdMenu(3).Visible = False
    cmdMenu(4).Visible = False
    cmdMenu(5).Visible = False
End Sub

Private Sub CambiaCentro(NuevoCentro As Byte)

CentroActual = NuevoCentro

If NuevoCentro = CentroMenu Then
    Call MostrarCentroMenu
    Call OcultarCentroHechizos
    Call OcultarCentroInventario
ElseIf NuevoCentro = CentroHechizos Then
    Call MostrarCentroHechizos
    Call OcultarCentroMenu
    Call OcultarCentroInventario
Else
    Call MostrarCentroInventario
    Call OcultarCentroHechizos
    Call OcultarCentroMenu
End If

End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    If IntervaloPermiteUsar Then UsarItem
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
    Dim Mx As Integer
    Dim My As Integer
    Dim aux As Integer
    Mx = X \ 32 + 1
    My = Y \ 32
    aux = (Mx + My * 5)
    If aux > 0 And aux <= MAX_INVENTORY_SLOTS Then
        lblInvInfo = IIf(UserInventory(aux).Amount > 0, UserInventory(aux).Name & " - " & UserInventory(aux).Amount, "")
    End If
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ItemClick(CInt(X), CInt(Y))
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    ElseIf hlst.Visible Then
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
stxtBuffer = SendTxt.Text
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub CompletarEnvioMensajes()

Select Case CurrentUser.SendingType
    Case 1
        SendTxt.Text = ""
    Case 2
        SendTxt.Text = "-"
    Case 3
        SendTxt.Text = ("\" & CurrentUser.sndPrivateTo & " ")
    Case 4
        SendTxt.Text = "/CMSG "
    Case 5
        SendTxt.Text = "/GMSG "
    Case 6
        SendTxt.Text = "/GRMG "
    Case 7
        SendTxt.Text = ";"
End Select

stxtBuffer = SendTxt.Text
SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim str1 As String
    Dim str2 As String
    
    'Send text
    If KeyCode = vbKeyReturn Then
        If left$(stxtBuffer, 1) = "/" Then
            Call ClientTCP.Parse_Command_Str(stxtBuffer)
    
        'Shout
        ElseIf left$(stxtBuffer, 1) = "-" Then
            If Right$(stxtBuffer, Len(stxtBuffer) - 1) <> "" Then Call ClientTCP.Send_Data(Shout_Chat, Right$(stxtBuffer, Len(stxtBuffer) - 1))
            CurrentUser.SendingType = 2
            
        'Global
        ElseIf left$(stxtBuffer, 1) = ";" Then
            If Right$(stxtBuffer, Len(stxtBuffer) - 1) <> "" Then Call ClientTCP.Send_Data(Global_Chat, Right$(stxtBuffer, Len(stxtBuffer) - 1))
            CurrentUser.SendingType = 7

        'Privado
        ElseIf left$(stxtBuffer, 1) = "\" Then
            str1 = Right$(stxtBuffer, Len(stxtBuffer) - 1)
            str2 = General_Field_Read(1, str1, " ")
            If str1 <> "" Then Call ClientTCP.Send_Data(Private_Chat, str1)
            CurrentUser.sndPrivateTo = str2
            CurrentUser.SendingType = 3
                    
        'Say
        Else
            If stxtBuffer <> "" Then Call ClientTCP.Send_Data(Normal_Chat, stxtBuffer)
            CurrentUser.SendingType = 1
        End If

        stxtBuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub

'[Barrin]
Private Sub Bind_Accion(ByVal FNUM As Integer)

If MacroKeys(FNUM).TipoAccion = 0 Then Exit Sub

Select Case MacroKeys(FNUM).TipoAccion

Case 1 'Envia comando
    Call ClientTCP.Parse_Command_Str("/" & MacroKeys(FNUM).SendString)
Case 2 'Lanza hechizo
    If hlst.List(MacroKeys(FNUM).hlist - 1) <> "Nada" And CurrentUser.Descansando = False Then
        If Not HechizoInvalido(hlst.List(MacroKeys(FNUM).hlist - 1)) Then
            If ClientTCP.DeadCheck Then Exit Sub
            If IntervaloPermiteLanzarSpell Then Call ClientTCP.Send_Data(Cast_Spell, Byte_To_String(MacroKeys(FNUM).hlist))
        Else
            Call AddtoRichTextBox(frmMain.RecTxt, "Esa acción no está permitida con el hechizo que estás intentando lanzar.", 61, 142, 36, True, True, False)
        End If
    End If
Case 3 'Trabaja
    tmrMacro.Enabled = Not tmrMacro.Enabled
Case 4 'Equipa
    If ClientTCP.DeadCheck Then Exit Sub
    Call EquiparItemMacro(MacroKeys(FNUM).invslot)
Case 5 'Usa
    If ClientTCP.DeadCheck Then Exit Sub
    If IntervaloPermiteUsar Then Call UsarItemMacro(MacroKeys(FNUM).invslot)
End Select

End Sub

Private Sub tmrMacro_Timer()

If IntervaloPermiteUsar Then
    Call UsarItem
    Call MainViewPic_MouseUp(vbLeftButton, 0, 0, 0)
End If

End Sub

Private Sub EquiparItemMacro(SelectedItemSlot As Integer)
    If (SelectedItemSlot > 0) And (SelectedItemSlot <= MAX_INVENTORY_SLOTS) Then _
        Call ClientTCP.Send_Data(Equip_Item, SelectedItemSlot)
End Sub

Private Sub UsarItemMacro(SelectedItemSlot As Integer)
    If (SelectedItemSlot > 0) And (SelectedItemSlot <= MAX_INVENTORY_SLOTS) Then
        Call ClientTCP.Send_Data(Use_Item, SelectedItemSlot)
    End If
End Sub
'[/Barrin]

Private Sub mainWinsock_Connect()

Debug.Print "*** Conectado"

If EstadoLogin = CrearNuevoPj Then
    If frmPasswd.Visible Then frmPasswd.lblStatus.Caption = "Conectado. Enviando datos..."
End If

Select Case EstadoLogin
    Case NORMAL, CrearNuevoPj
        Call ClientTCP.Send_Data(Auth_Start)
End Select

End Sub

Private Sub mainWinsock_Close()

Debug.Print "*** Cerrado"

frmConnect.Visible = True
frmMain.Visible = False
frmIniciando.Visible = False
frmMensaje.Visible = False
frmConnect.MousePointer = 1

CurrentUser.Pausa = False
Call ResetCurrentUser

End Sub

Private Sub mainWinsock_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

frmConnect.MousePointer = 1
frmMain.mainWinsock.Close
frmMain.Visible = False

If CurrentUser.Logged Then
    
    Call ResetCurrentUser
    
    If Musica <> CONST_DESHABILITADA Then
        Sound.NextMusic = MUS_VolverInicio
        Sound.Fading = 200
    End If
    
End If

If Not frmCrearPersonaje.Visible Then
    If Not frmIniciando.Visible Then
        frmConnect.Show
        Call MsgBox("Ha ocurrido un error al conectar con el servidor solicitado. Le recomendamos verificar el estado de los servidores en www.imperiumao.com.ar, y asegurarse de estar conectado directamente a internet (" & Description & " - " & number & ")", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
    Else
        frmConnect.Show
        Unload frmIniciando
        Call MsgBox("Ha ocurrido un error al conectar con el servidor solicitado. Le recomendamos verificar el estado de los servidores en www.imperiumao.com.ar, y asegurarse de estar conectado directamente a internet (" & Description & " - " & number & ")", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error al conectar")
    End If
Else
    frmPasswd.lblStatus.Caption = "Error: " & Description
    frmPasswd.MousePointer = 0
End If

End Sub

Private Sub mainWinsock_DataArrival(ByVal BytesTotal As Long)

On Error Resume Next

Dim HayComandoCompleto As Boolean
Dim RD As String, SzComando As Integer

mainWinsock.GetData RD

If Len(CurrentUser.RDBuffer) > 0 Then
    CurrentUser.RDBuffer = CurrentUser.RDBuffer & RD
Else
    CurrentUser.RDBuffer = RD
End If

CopyMemory SzComando, ByVal StrPtr(CurrentUser.RDBuffer), 2
HayComandoCompleto = SzComando <= Len(CurrentUser.RDBuffer)

Do While HayComandoCompleto
    Call ClientTCP.Handle_Data(ClientTCP.Decrypt_Data(mid$(CurrentUser.RDBuffer, 2, SzComando - 1), ""))
    CurrentUser.RDBuffer = mid$(CurrentUser.RDBuffer, SzComando + 1)
    If Len(CurrentUser.RDBuffer) > 1 Then
        CopyMemory SzComando, ByVal StrPtr(CurrentUser.RDBuffer), 2
        HayComandoCompleto = SzComando <= Len(CurrentUser.RDBuffer)
    Else
        HayComandoCompleto = False
    End If
Loop

'On Error Resume Next
'
'Dim AuxPos As Long
'Dim LastPos As Long
'Dim RD As String
'
'mainWinsock.GetData RD
'
'RD = CurrentUser.RDBuffer & RD 'StrConv(RD, vbUnicode)
'LastPos = 1
'
'Do
'    AuxPos = InStr(LastPos, RD, ENDC)
'    If AuxPos = 0 Then Exit Do
'    Call ClientTCP.Handle_Data(ClientTCP.Decrypt_Data(mid$(RD, LastPos, AuxPos - LastPos), "WhaTis33thiS"))
'    LastPos = AuxPos + 1
'Loop
'
'LastPos = LastPos - 1
'If Len(RD) <> LastPos Then CurrentUser.RDBuffer = Right$(RD, Len(RD) - LastPos)
'

End Sub

Private Function HechizoInvalido(ByVal HechizoName As String) As Boolean

HechizoName = UCase$(HechizoName)

If HechizoName = "REMOVER PARALISIS" Or HechizoName = "DESENCANTAR" Or HechizoName = "SANAR" Then
    HechizoInvalido = True
    Exit Function
End If

End Function

'###########################################################
'                        GUI
'###########################################################

Private Sub cmdHechizos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If CentroActual <> CentroHechizos Then Exit Sub

Select Case Index
    Case 0 'Lanzar
        cmdHechizos(0).Picture = General_Load_Picture_From_Resource("[hechizos]lanzar-down")
    Case 1 'Info
        cmdHechizos(1).Picture = General_Load_Picture_From_Resource("[hechizos]info-down")
    Case 2 'Subir
        cmdHechizos(2).Picture = General_Load_Picture_From_Resource("[hechizos]flechaarriba-down")
    Case 3 'Bajar
        cmdHechizos(3).Picture = General_Load_Picture_From_Resource("[hechizos]flechaabajo-down")
End Select

End Sub

Private Sub cmdHechizos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If hlst.Visible = False Then Exit Sub
If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index
    Case 0 'lanzar
        cmdHechizos(0).Picture = General_Load_Picture_From_Resource("[hechizos]lanzar-over")
    Case 1 'info
        cmdHechizos(1).Picture = General_Load_Picture_From_Resource("[hechizos]info-over")
    Case 2 'Subir
        cmdHechizos(2).Picture = General_Load_Picture_From_Resource("[hechizos]flechaarriba-over")
    Case 3 'Bajar
        cmdHechizos(3).Picture = General_Load_Picture_From_Resource("[hechizos]flechaabajo-over")
End Select

End Sub

Private Sub cmdHechizos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If CentroActual <> CentroHechizos Then Exit Sub
Call Form_MouseMove(Button, Shift, X, Y)

If hlst.ListIndex = -1 Then Exit Sub

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0 'lanzar
        If hlst.List(hlst.ListIndex) <> "Nada" And CurrentUser.Descansando = False Then
            If ClientTCP.DeadCheck Then Exit Sub
            If IntervaloPermiteLanzarSpell Then Call ClientTCP.Send_Data(Cast_Spell, Byte_To_String(hlst.ListIndex + 1))
        End If
    Case 1 'info
        Call ClientTCP.Send_Data(Spell_Info_Request, hlst.ListIndex + 1)
    Case 2 'subir
        If hlst.ListIndex = 0 Then Exit Sub
        Call ClientTCP.Send_Data(Spell_Move, 1 & "," & hlst.ListIndex + 1)
        hlst.ListIndex = hlst.ListIndex - 1
    Case 3 'bajar
        If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        Call ClientTCP.Send_Data(Spell_Move, 2 & "," & hlst.ListIndex + 1)
        hlst.ListIndex = hlst.ListIndex + 1
End Select

End Sub

Private Sub CentroHechizosRestaurar(Index As Integer)

cmdHechizos(Index).Picture = Nothing

End Sub

Private Sub SolapasRestaurar(Index As Integer)

imgCentros(Index).Picture = Nothing
imgMiniCerra.Picture = Nothing
cmdMensaje.Picture = Nothing

End Sub

Private Sub imgCentros_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0 'Inv
        'imgCentros(0).Picture = General_Load_Picture_From_Resource("[Solapas]Inventario-Down")
    Case 1 'Hechizos
        'imgCentros(1).Picture = General_Load_Picture_From_Resource("[Solapas]Hechizos-Down")
    Case 2 'Menu
        'imgCentros(2).Picture = General_Load_Picture_From_Resource("[Solapas]Menu-Down")
End Select

End Sub

Private Sub imgCentros_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If UltPosSolapas = Index Then Exit Sub

If UltPosSolapas <> -1 Then Call RestaurarCentroActual
UltPosSolapas = Index

Select Case Index
    Case 0 'Inv
        imgCentros(0).Picture = General_Load_Picture_From_Resource("[solapas]inventario-over")
    Case 1 'Hechizos
        imgCentros(1).Picture = General_Load_Picture_From_Resource("[solapas]hechizos-over")
    Case 2 'Menu
        imgCentros(2).Picture = General_Load_Picture_From_Resource("[solapas]menu-over")
End Select

End Sub

Private Sub cmdMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0 'Grupo
        cmdMenu(0).Picture = General_Load_Picture_From_Resource("[menu]grupo-down")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = General_Load_Picture_From_Resource("[menu]estadisticas-down")
    Case 2 'Guild
        cmdMenu(2).Picture = General_Load_Picture_From_Resource("[menu]clanes-down")
    Case 3 'Quest
        cmdMenu(3).Picture = General_Load_Picture_From_Resource("[menu]quests-down")
    Case 4 'Torneos
        cmdMenu(4).Picture = General_Load_Picture_From_Resource("[menu]torneos-down")
    Case 5 'Opciones
        cmdMenu(5).Picture = General_Load_Picture_From_Resource("[menu]opciones-down")
End Select

End Sub

Private Sub cmdMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If UltPosInterface = Index Then Exit Sub

If UltPosInterface <> -1 Then Call RestaurarCentroActual
UltPosInterface = Index

Select Case Index

    Case 0 'Grupo
        cmdMenu(0).Picture = General_Load_Picture_From_Resource("[menu]grupo-over")
    Case 1 'Estadisticas
        cmdMenu(1).Picture = General_Load_Picture_From_Resource("[menu]estadisticas-over")
    Case 2 'Guild
        cmdMenu(2).Picture = General_Load_Picture_From_Resource("[menu]clanes-over")
    Case 3 'Quest
        cmdMenu(3).Picture = General_Load_Picture_From_Resource("[menu]quests-over")
    Case 4 'Torneos
        cmdMenu(4).Picture = General_Load_Picture_From_Resource("[menu]torneos-over")
    Case 5 'Opciones
        cmdMenu(5).Picture = General_Load_Picture_From_Resource("[menu]opciones-over")
End Select

End Sub

Private Sub cmdMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If CentroActual <> CentroMenu Then Exit Sub
Call Form_MouseMove(Button, Shift, X, Y)

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0 'Grupo
        Call ClientTCP.Send_Data(Group_Member_List)
    Case 1 'Estadisticas
        Call ClientTCP.Reset_Skill_Data
        Call ClientTCP.Send_Data(Stats_Att_Request)
        Call ClientTCP.Send_Data(Stats_Skills_Request)
        Call ClientTCP.Send_Data(Stats_Fame_Request)
        Call ClientTCP.Send_Data(Stats_Familiar_Request)
        Call ClientTCP.Send_Data(Stats_General_Request)
        PedimosEst = True
    Case 2 'Guild
        If Not (frmGuildLeader.Visible Or frmGuildAdm.Visible) Then _
            Call ClientTCP.Send_Data(Guild_Info_Request)
    Case 3 'Quest
        Call ClientTCP.Send_Data(Quest_Data_Cl)
    Case 4 'Torneos
        Call ClientTCP.Send_Data_Command(cmdTorneos)
    Case 5 'Opciones
        Call frmOpciones.Init
End Select

End Sub

Private Sub imgCentros_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Form_MouseMove(Button, Shift, X, Y)
Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0
        Call CambiaCentro(CentroInventario)
    Case 1
        Call CambiaCentro(CentroHechizos)
    Case 2
        Call CambiaCentro(CentroMenu)
End Select

End Sub

Private Sub CentroMenuRestaurar(Index As Integer)

cmdMenu(Index).Picture = Nothing

End Sub

Private Sub RestaurarCentroActual()

Select Case CentroActual
    Case CentroHechizos
        If UltPosInterface <> -1 Then Call CentroHechizosRestaurar(UltPosInterface)
    Case CentroInventario
    Case CentroMenu
        If UltPosInterface <> -1 Then Call CentroMenuRestaurar(UltPosInterface)
End Select

If UltPosSolapas <> -1 Then Call SolapasRestaurar(UltPosSolapas)

UltPosInterface = -1
UltPosSolapas = -1

imgMiniCerra.Picture = Nothing
cmdMensaje.Picture = Nothing
lblInvInfo.Caption = ""

End Sub

Public Sub Client_Screenshot()

On Error GoTo ErrorHandler

Dim i As Long
Dim Index As Long
i = 1

Set m_Jpeg = New cJpeg

'Sample the cImage by hDC
m_Jpeg.SampleHDC frmMain.hDC, 800, 600

m_FileName = App.Path & "\Fotos\ImpAO_Foto"

If Dir(App.Path & "\Fotos", vbDirectory) = "" Then
    MkDir (App.Path & "\Fotos")
End If

Do While Dir(m_FileName & Trim(str(i)) & ".jpg") <> ""
    i = i + 1
    DoEvents
Loop

Index = i

m_Jpeg.Comment = "Character: " & CurrentUser.UserName & " - " & format(Date, "dd/mm/yyyy") & " - " & format(Time, "hh:mm AM/PM")

'Save the JPG file
m_Jpeg.SaveFile m_FileName & Trim(str(Index)) & ".jpg"

Call AddtoRichTextBox(frmMain.RecTxt, "Screenshot grabada correctamente como " & m_FileName & Trim(str(Index)) & ".jpg", 65, 190, 156, False, True, False)

Set m_Jpeg = Nothing

Exit Sub

ErrorHandler:
    Call AddtoRichTextBox(frmMain.RecTxt, "Error al grabar el screenshot. Por favor intente nuevamente.", 65, 190, 156, False, True, False)

End Sub
